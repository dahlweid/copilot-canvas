// Smoke test for the per-instance viewer server: HTTP API, PDF serving with
// byte ranges, SSE, and live reload when the document is rewritten on disk.
//
// Run: node .github/extensions/office-canvas/test/integration/server-smoke.mjs

import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { promisify } from "node:util";
import { fileURLToPath } from "node:url";
import { RenderCache, normalizeDocPath } from "../../src/render-cache.mjs";
import { ViewerInstance } from "../../src/server.mjs";
import { codepoints } from "./docx-zip.mjs";
import { assertNoLeakedWord, wordPids } from "./word-pids.mjs";

const execFileAsync = promisify(execFile);
const HERE = path.dirname(fileURLToPath(import.meta.url));
const results = [];

async function check(name, fn) {
    try {
        await fn();
        results.push({ name, ok: true });
        process.stderr.write(`  ok   ${name}\n`);
    } catch (err) {
        results.push({ name, ok: false, err });
        process.stderr.write(`  FAIL ${name}\n       ${err.message}\n`);
    }
}

const makeFixture = (out, chapters) =>
    execFileAsync(
        "powershell.exe",
        [
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            path.join(HERE, "make-fixture.ps1"),
            "-Out",
            out,
            "-Chapters",
            String(chapters),
        ],
        { maxBuffer: 8 * 1024 * 1024 },
    );

const workRoot = await mkdtemp(path.join(tmpdir(), "word-server-test-"));
const fixture = path.join(workRoot, "docs", "report.docx");
const pidsBefore = await wordPids();

process.stderr.write("generating fixture...\n");
await makeFixture(fixture, 3);

const cache = new RenderCache({
    cacheRoot: path.join(workRoot, "artifacts"),
    log: (m) => process.stderr.write(`[cache] ${m}\n`),
});

const viewer = new ViewerInstance({
    cache,
    instanceId: "server-smoke-1",
    workspacePath: workRoot,
    log: (m) => process.stderr.write(`[viewer] ${m}\n`),
});

/** Fetch against the instance, with the loopback origin filled in. */
const get = (route, init) => fetch(new URL(route, viewer.url), init);

/** Collects named SSE events off /events until a predicate is satisfied. */
function listen(signal) {
    const seen = [];
    /** @type {{name: string, resolve: (e: any) => void}[]} */
    const waiters = [];
    const promise = (async () => {
        const res = await fetch(new URL("/events", viewer.url), { signal });
        const reader = res.body.getReader();
        const decoder = new TextDecoder();
        let buffer = "";
        while (true) {
            const { done, value } = await reader.read();
            if (done) break;
            buffer += decoder.decode(value, { stream: true });
            let split;
            while ((split = buffer.indexOf("\n\n")) !== -1) {
                const frame = buffer.slice(0, split);
                buffer = buffer.slice(split + 2);
                const event = /^event: (.+)$/m.exec(frame)?.[1];
                const data = /^data: (.+)$/m.exec(frame)?.[1];
                if (!event) continue;
                const entry = { event, data: data ? JSON.parse(data) : null };
                seen.push(entry);
                for (let i = waiters.length - 1; i >= 0; i--) {
                    if (waiters[i].name === entry.event) waiters.splice(i, 1)[0].resolve(entry);
                }
            }
        }
    })().catch(() => {});

    return {
        promise,
        seen,
        /** Resolves with the first matching event, past or future. */
        wait(name, timeoutMs = 90_000) {
            const existing = seen.find((e) => e.event === name);
            if (existing) return Promise.resolve(existing);
            return new Promise((resolve, reject) => {
                const waiter = {
                    name,
                    resolve: (entry) => {
                        clearTimeout(timer);
                        resolve(entry);
                    },
                };
                const timer = setTimeout(() => {
                    const i = waiters.indexOf(waiter);
                    if (i !== -1) waiters.splice(i, 1);
                    reject(new Error(`timed out waiting for SSE '${name}' event`));
                }, timeoutMs);
                waiters.push(waiter);
            });
        },
    };
}

let pdfKey = null;
let pdfSize = 0;
const aborter = new AbortController();
let events = null;

try {
    await check("start binds a loopback server", async () => {
        const url = await viewer.start();
        assert.match(url, /^http:\/\/127\.0\.0\.1:\d+\/$/);
        assert.equal(viewer.url, url);
    });

    await check("start is idempotent", async () => {
        const url = viewer.url;
        assert.equal(await viewer.start(), url);
    });

    await check("serves the viewer shell at /", async () => {
        const res = await get("/");
        assert.equal(res.status, 200);
        assert.match(res.headers.get("content-type"), /text\/html/);
        const body = await res.text();
        assert.match(body, /<iframe/i);
        assert.match(body, /app\.js/);
    });

    await check("serves the stylesheet and script", async () => {
        for (const [asset, type] of [
            ["/app.css", /text\/css/],
            ["/app.js", /javascript/],
        ]) {
            const res = await get(asset);
            assert.equal(res.status, 200, `${asset} -> ${res.status}`);
            assert.match(res.headers.get("content-type"), type);
        }
    });

    await check("static routes cannot escape the ui directory", async () => {
        const res = await get("/..%2f..%2fextension.mjs");
        assert.equal(res.status, 404);
    });

    await check("state before a document is open", async () => {
        const state = await (await get("/api/state")).json();
        assert.equal(state.status, "idle");
        assert.equal(state.doc, null);
        assert.equal(state.pdfUrl, null);
    });

    await check("outline before a document is open is a typed 400", async () => {
        const res = await get("/api/outline");
        assert.equal(res.status, 400);
        const body = await res.json();
        assert.equal(body.error.code, "not_open");
    });

    await check("browse finds the fixture in the workspace", async () => {
        const body = await (await get("/api/browse")).json();
        assert.equal(body.workspacePath, workRoot);
        const hit = body.workspaceDocs.find((d) => d.name === "report.docx");
        assert.ok(hit, `report.docx not in ${JSON.stringify(body.workspaceDocs)}`);
        assert.equal(hit.relative, path.join("docs", "report.docx"));
    });

    await check("open a document over the API", async () => {
        events = listen(aborter.signal);
        await events.wait("state", 10_000); // initial snapshot on connect

        const res = await get("/api/open", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ path: fixture }),
        });
        assert.equal(res.status, 200);
        const { doc, state } = await res.json();
        assert.equal(doc.path, normalizeDocPath(fixture));
        assert.ok(doc.pageCount > 1, `expected several pages, got ${doc.pageCount}`);
        assert.equal(state.status, "ready");
        assert.equal(state.pdfUrl, `/pdf/${doc.key}.pdf`);
        pdfKey = doc.key;
        process.stderr.write(`       ${doc.name}: ${doc.pageCount} pages, key ${doc.key}\n`);
    });

    await check("opening a missing file is a typed 400", async () => {
        const res = await get("/api/open", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ path: path.join(workRoot, "nope.docx") }),
        });
        assert.equal(res.status, 400);
        assert.equal((await res.json()).error.code, "file_not_found");
    });

    await check("a failed open leaves the previous document intact", async () => {
        const state = await (await get("/api/state")).json();
        assert.equal(state.doc.key, pdfKey);
        // status reflects the failure, but the rendered document is still served
        assert.equal(state.error.code, "file_not_found");
    });

    await check("serves the rendered PDF", async () => {
        const res = await get(`/pdf/${pdfKey}.pdf`);
        assert.equal(res.status, 200);
        assert.equal(res.headers.get("content-type"), "application/pdf");
        assert.equal(res.headers.get("accept-ranges"), "bytes");
        const buf = Buffer.from(await res.arrayBuffer());
        assert.equal(buf.subarray(0, 5).toString("latin1"), "%PDF-");
        assert.equal(buf.length, Number(res.headers.get("content-length")));
        pdfSize = buf.length;
        assert.ok(pdfSize > 1000, `suspiciously small PDF: ${pdfSize} bytes`);
    });

    await check("honours byte-range requests", async () => {
        const res = await get(`/pdf/${pdfKey}.pdf`, { headers: { Range: "bytes=0-99" } });
        assert.equal(res.status, 206);
        assert.equal(res.headers.get("content-range"), `bytes 0-99/${pdfSize}`);
        const buf = Buffer.from(await res.arrayBuffer());
        assert.equal(buf.length, 100);
        assert.equal(buf.subarray(0, 5).toString("latin1"), "%PDF-");
    });

    await check("serves an open-ended range", async () => {
        const from = pdfSize - 50;
        const res = await get(`/pdf/${pdfKey}.pdf`, { headers: { Range: `bytes=${from}-` } });
        assert.equal(res.status, 206);
        assert.equal(res.headers.get("content-range"), `bytes ${from}-${pdfSize - 1}/${pdfSize}`);
        assert.equal(Buffer.from(await res.arrayBuffer()).length, 50);
    });

    await check("rejects an unsatisfiable range", async () => {
        const res = await get(`/pdf/${pdfKey}.pdf`, {
            headers: { Range: `bytes=${pdfSize + 10}-${pdfSize + 20}` },
        });
        assert.equal(res.status, 416);
        assert.equal(res.headers.get("content-range"), `bytes */${pdfSize}`);
    });

    await check("outline lists headings with page numbers", async () => {
        const outline = await (await get("/api/outline")).json();
        assert.ok(outline.headings.length > 0, "empty outline");
        assert.equal(outline.count, outline.headings.length);
        assert.ok(outline.headings.every((h) => typeof h.text === "string" && h.page >= 1));
        assert.ok(outline.headings.every((h) => h.level >= 1));
    });

    await check("outline honours limit", async () => {
        const outline = await (await get("/api/outline?limit=2")).json();
        assert.ok(outline.headings.length <= 2, `got ${outline.headings.length} headings`);
    });

    await check("search returns matches with pages and snippets", async () => {
        const body = await (await get("/api/search?q=ZORBLAX&limit=5")).json();
        assert.equal(body.query, "ZORBLAX");
        assert.ok(body.hits.length > 0, "the unique fixture token was not found");
        assert.ok(body.hits.every((h) => h.page >= 1 && typeof h.snippet === "string"));
    });

    await check("search for something absent returns no hits", async () => {
        const body = await (await get("/api/search?q=QQZZNOTPRESENT")).json();
        assert.deepEqual(body.hits, []);
    });

    await check("a canvas search for a non-ASCII term finds it", async () => {
        // Issue #40, and not redundant with the tool-side tests: the canvas
        // action boundary and the tool boundary have already been measured
        // behaving differently from one another (integer schema bounds are
        // enforced on one and not the other, issue #28), so neither is inferable
        // from the other even though one line fixes both.
        //
        // The symptom this guards is the quietest of the three the defect
        // caused, and the one a user meets first: `count: 0` with no error, no
        // warning, indistinguishable from the term genuinely not being there.
        //
        // Built from codepoints, so the assertion does not rest on how this
        // file's bytes are decoded.
        const strasse = `Stra${String.fromCharCode(0x00df)}e`;
        const body = await (await get(`/api/search?q=${encodeURIComponent(strasse)}&limit=5`)).json();
        // One response, both directions. `query` is echoed from the host's
        // decoding of what we sent (inbound); the snippet comes from the
        // document over COM (outbound). Broken, this response carried a correct
        // sz in the snippet and a corrupted one in the query -- same character,
        // same request, one right and one wrong.
        assert.equal(body.query, strasse, `the query came back as ${codepoints(body.query)}, sent ${codepoints(strasse)}`);
        assert.ok(body.count > 0, `no hits for ${codepoints(strasse)}, which the fixture contains`);
        assert.ok(
            body.hits.some((h) => h.snippet.includes(strasse)),
            `hits found but no snippet contains the term: ${body.hits.map((h) => codepoints(h.snippet)).join(" | ")}`,
        );
    });

    await check("an empty search query is a typed 400", async () => {
        const res = await get("/api/search?q=");
        assert.equal(res.status, 400);
        assert.equal((await res.json()).error.code, "invalid_query");
    });

    await check("page position round-trips", async () => {
        const res = await get("/api/page", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ page: 2 }),
        });
        assert.equal((await res.json()).lastPage, 2);
        assert.equal((await (await get("/api/state")).json()).lastPage, 2);
    });

    await check("goToPage broadcasts to the iframe", async () => {
        viewer.goToPage(3);
        const entry = await events.wait("goto", 5_000);
        assert.equal(entry.data.page, 3);
    });

    await check("refresh on an unchanged file reports no change", async () => {
        const res = await get("/api/refresh", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({}),
        });
        assert.equal((await res.json()).changed, false);
    });

    await check("unknown api routes 404", async () => {
        assert.equal((await get("/api/nonsense")).status, 404);
    });

    await check("rewriting the document triggers an automatic reload", async () => {
        process.stderr.write("       regenerating the fixture with more chapters...\n");
        await makeFixture(fixture, 6);
        const entry = await events.wait("reloaded", 120_000);
        assert.notEqual(entry.data.doc.key, pdfKey, "render key should change");
        assert.ok(
            entry.data.doc.pageCount > 0,
            `expected a page count, got ${entry.data.doc.pageCount}`,
        );
        assert.equal(entry.data.restorePage, 3, "should restore the page we navigated to");
        assert.equal(entry.data.pdfUrl, `/pdf/${entry.data.doc.key}.pdf`);
        process.stderr.write(
            `       reloaded: ${entry.data.doc.pageCount} pages, key ${entry.data.doc.key}\n`,
        );
        pdfKey = entry.data.doc.key;
    });

    await check("the re-rendered PDF is served under the new key", async () => {
        const res = await get(`/pdf/${pdfKey}.pdf`);
        assert.equal(res.status, 200);
        const buf = Buffer.from(await res.arrayBuffer());
        assert.equal(buf.subarray(0, 5).toString("latin1"), "%PDF-");
        assert.notEqual(buf.length, pdfSize, "expected a different render");
    });

    await check("recents include the opened document", async () => {
        const body = await (await get("/api/browse")).json();
        assert.ok(
            body.recents.some((r) => normalizeDocPath(r.path) === normalizeDocPath(fixture)),
            "fixture missing from recents",
        );
    });

    await check("close shuts the server down", async () => {
        const url = viewer.url;
        aborter.abort();
        await viewer.close();
        assert.equal(viewer.url, null);
        await assert.rejects(fetch(url), /fetch failed|ECONNREFUSED/);
    });
} finally {
    aborter.abort();
    await viewer.close().catch(() => {});
    await cache.dispose().catch(() => {});
}

await check("no new WINWORD.EXE is left behind", () => assertNoLeakedWord(pidsBefore));

await rm(workRoot, { recursive: true, force: true }).catch(() => {});

const failed = results.filter((r) => !r.ok);
process.stderr.write(`\n${results.length - failed.length}/${results.length} passed\n`);
process.exit(failed.length ? 1 : 0);
