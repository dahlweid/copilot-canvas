// Smoke test for the Word bridge. Run with:
//   node .github/extensions/word-canvas/test/host-smoke.mjs
//
// Requires Word. Generates its own fixture, exercises every host command, and
// asserts that no WINWORD.EXE is left behind.

import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { fileURLToPath } from "node:url";
import { mkdtemp, rm, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import assert from "node:assert/strict";

import { WordHost } from "../src/word-host.mjs";

const execFileAsync = promisify(execFile);
const HERE = path.dirname(fileURLToPath(import.meta.url));

async function wordPids() {
    const { stdout } = await execFileAsync("powershell.exe", [
        "-NoProfile",
        "-NonInteractive",
        "-Command",
        "@(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id) -join ','",
    ]);
    return stdout
        .trim()
        .split(",")
        .filter(Boolean)
        .map((n) => Number(n));
}

const results = [];
function check(name, fn) {
    return (async () => {
        try {
            await fn();
            results.push({ name, ok: true });
            process.stderr.write(`  ok   ${name}\n`);
        } catch (err) {
            results.push({ name, ok: false, err });
            process.stderr.write(`  FAIL ${name}\n       ${err.message}\n`);
        }
    })();
}

const workRoot = await mkdtemp(path.join(tmpdir(), "word-canvas-test-"));
const fixture = path.join(workRoot, "fixture.docx");
const pidsBefore = await wordPids();

process.stderr.write("generating fixture...\n");
await execFileAsync("powershell.exe", [
    "-NoProfile",
    "-NonInteractive",
    "-ExecutionPolicy",
    "Bypass",
    "-File",
    path.join(HERE, "make-fixture.ps1"),
    "-Out",
    fixture,
]);

const host = new WordHost({
    log: (m) => process.stderr.write(`[host] ${m}\n`),
    // Deliberately stable across runs: a host that crashed in a previous run
    // records its Word pid here, and the next run reaps it.
    pidDir: path.join(tmpdir(), "word-canvas-test-pids"),
});
const docId = "smoke";
let opened;

try {
    await check("ping starts Word and reports a version", async () => {
        const t0 = Date.now();
        const res = await host.ping();
        process.stderr.write(`       word ${res.wordVersion}, owned=${res.owned}, ${Date.now() - t0}ms\n`);
        assert.equal(res.ready, true);
        assert.ok(res.wordVersion, "expected a Word version");
    });

    await check("open returns metadata and a multi-page count", async () => {
        opened = await host.openDocument({ docId, path: fixture, workDir: path.join(workRoot, "work") });
        process.stderr.write(`       ${opened.name}: ${opened.pageCount} pages, ${opened.wordCount} words\n`);
        assert.ok(opened.pageCount > 1, `expected a multi-page fixture, got ${opened.pageCount}`);
        assert.equal(opened.title, "Word Canvas Fixture");
        assert.notEqual(path.resolve(opened.workingCopy), path.resolve(fixture), "must not open the original file");
    });

    await check("the original file is never locked or modified", async () => {
        const before = await stat(fixture);
        // If Word held the original open, this rename would fail on Windows.
        const probe = `${fixture}.probe`;
        const { renameSync } = await import("node:fs");
        renameSync(fixture, probe);
        renameSync(probe, fixture);
        const after = await stat(fixture);
        assert.equal(before.mtimeMs, after.mtimeMs, "original mtime changed");
        assert.equal(before.size, after.size, "original size changed");
    });

    await check("export produces a real PDF", async () => {
        const out = path.join(workRoot, "out.pdf");
        const t0 = Date.now();
        const res = await host.exportPdf({ docId, out });
        process.stderr.write(`       ${res.pageCount} pages, ${res.sizeBytes} bytes, ${Date.now() - t0}ms\n`);
        assert.ok(res.sizeBytes > 1000, "PDF looks too small");
        const { readFileSync } = await import("node:fs");
        assert.equal(readFileSync(out).subarray(0, 5).toString("latin1"), "%PDF-", "not a PDF");
    });

    await check("export honours a page range", async () => {
        const out = path.join(workRoot, "range.pdf");
        const full = (await stat(path.join(workRoot, "out.pdf"))).size;
        const res = await host.exportPdf({ docId, out, from: 1, to: 1 });
        assert.ok(res.sizeBytes < full, "single-page export should be smaller than the full export");
    });

    await check("outline finds headings with levels and page numbers", async () => {
        const res = await host.outline({ docId });
        process.stderr.write(`       ${res.count} headings, first: ${JSON.stringify(res.headings[0])}\n`);
        assert.ok(res.count >= 10, `expected >= 10 headings, got ${res.count}`);
        assert.ok(Array.isArray(res.headings), "headings must be an array");
        const h1 = res.headings.filter((h) => h.level === 1);
        const h2 = res.headings.filter((h) => h.level === 2);
        assert.ok(h1.length >= 4, `expected >= 4 level-1 headings, got ${h1.length}`);
        assert.ok(h2.length >= 6, `expected >= 6 level-2 headings, got ${h2.length}`);
        assert.ok(
            res.headings.every((h) => h.page >= 1),
            "every heading needs a page number",
        );
        // The whole point of the outline is jumping to the right page, so page
        // numbers must actually vary across a multi-page document.
        const pages = new Set(res.headings.map((h) => h.page));
        assert.ok(pages.size > 1, `headings all resolved to one page: ${[...pages].join(",")}`);
        // Document order, not style order.
        const starts = res.headings.map((h) => h.start);
        assert.deepEqual(starts, [...starts].sort((a, b) => a - b), "headings must be in document order");
        assert.equal(res.headings[0].text, "Chapter 1");
    });

    await check("search returns hits with page numbers and snippets", async () => {
        const res = await host.search({ docId, query: "ZORBLAX" });
        assert.equal(res.count, 1, `expected exactly one hit, got ${res.count}`);
        assert.ok(res.hits[0].page >= 1, "hit needs a page number");
        assert.match(res.hits[0].snippet, /ZORBLAX/);
        assert.doesNotMatch(res.hits[0].snippet, /[\r\n]/, "snippet must be single-line");
    });

    await check("search with no hits returns an empty array", async () => {
        const res = await host.search({ docId, query: "definitely-not-in-the-document-xyzzy" });
        assert.equal(res.count, 0);
        assert.ok(Array.isArray(res.hits), `hits must stay an array, got ${JSON.stringify(res.hits)}`);
    });

    await check("text extracts the whole document", async () => {
        const res = await host.text({ docId });
        assert.match(res.text, /Chapter 1/);
        assert.match(res.text, /ZORBLAX/);
        assert.equal(res.fromPage, 1);
    });

    await check("text extracts a single page", async () => {
        const whole = await host.text({ docId });
        const page1 = await host.text({ docId, fromPage: 1, toPage: 1 });
        assert.ok(page1.text.length > 0, "page 1 should not be empty");
        assert.ok(page1.text.length < whole.text.length, "one page must be shorter than the whole document");
    });

    await check("text rejects a page past the end", async () => {
        await assert.rejects(() => host.text({ docId, fromPage: 9999 }), /beyond the end/i);
    });

    await check("non-ASCII survives the stdio round trip", async () => {
        const res = await host.search({ docId, query: "Muenchen" });
        assert.ok(res.count > 0, "expected to find the umlaut-check sentence");
    });

    await check("info reports document properties", async () => {
        const res = await host.info({ docId });
        assert.equal(res.title, "Word Canvas Fixture");
        assert.ok(res.pageCount > 1);
        assert.ok(res.paragraphs > 10);
    });

    await check("a missing file produces a clean error, not a hang", async () => {
        await assert.rejects(
            () => host.openDocument({ docId: "missing", path: path.join(workRoot, "nope.docx"), workDir: workRoot }),
            /File not found/i,
        );
    });

    await check("an unknown command is reported, not swallowed", async () => {
        await assert.rejects(() => host.request("frobnicate", {}), /Unknown command/i);
    });

    await check("commands recover after the host process dies", async () => {
        // Simulates a Word crash / stuck-dialog teardown: the bridge must
        // transparently restart and replay the open.
        await host.request("__force_kill_for_test", {}).catch(() => {});
        const { execFileSync } = await import("node:child_process");
        const pids = await wordPids();
        for (const pid of pids.filter((p) => !pidsBefore.includes(p))) {
            try {
                execFileSync("powershell.exe", [
                    "-NoProfile",
                    "-NonInteractive",
                    "-Command",
                    `Stop-Process -Id ${pid} -Force -ErrorAction SilentlyContinue`,
                ]);
            } catch {
                /* ignore */
            }
        }
        const res = await host.info({ docId });
        assert.ok(res.pageCount > 1, "expected the bridge to recover and reopen the document");
    });

    await check("close releases the document", async () => {
        const res = await host.closeDocument({ docId });
        assert.equal(res.closed, true);
    });
} finally {
    await host.dispose();
}

await check("no WINWORD.EXE is left behind", async () => {
    await new Promise((r) => setTimeout(r, 1500));
    const after = await wordPids();
    const leaked = after.filter((p) => !pidsBefore.includes(p));
    assert.deepEqual(leaked, [], `leaked Word processes: ${leaked.join(", ")}`);
});

await rm(workRoot, { recursive: true, force: true }).catch(() => {});

const failed = results.filter((r) => !r.ok);
process.stderr.write(`\n${results.length - failed.length}/${results.length} passed\n`);
process.exit(failed.length === 0 ? 0 : 1);
