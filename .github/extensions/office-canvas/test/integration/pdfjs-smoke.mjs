// The pdf.js viewer against a real Word export.
//
// Run: node .github/extensions/office-canvas/test/integration/pdfjs-smoke.mjs
//
// ## What this can and cannot reach
//
// The vendored `pdf.min.mjs` is pdf.js's *browser* build. Node cannot import it
// -- measured: it reaches for `DOMMatrix` at module scope and throws before any
// of its exports exist -- so nothing here renders a page or extracts text. That
// is not a gap this file can close by trying harder; the legacy build that would
// run under Node is a second copy of pdf.js and the budget has no room for one.
//
// So the split is deliberate:
//
//   * everything that is a decision -- the split, the reassembly, the locator,
//     which pages get marked -- is a pure function and lives in `test/unit/`,
//     where it runs on `ubuntu-latest` with no Office at all;
//   * rendering and locating **against real pdf.js** were verified by loading
//     the running instance in a browser and measuring the DOM: one box on the
//     changed line at dx 0 with the span's width, exactly one page marked, the
//     banner and marker cleared by Dismiss and restored by the next edit;
//   * what is left for this file is the part only an installed Word can answer:
//     that a real export is a real PDF, that the vendored assets are reachable
//     from the same instance that serves it, and that a record built from a real
//     `edit_document` result describes that export rather than something else.
//
// Nothing here asserts that pdf.js renders. If it passes and the viewer is
// blank, that is consistent -- read the browser measurement, not this file.

import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { execFile } from "node:child_process";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { promisify } from "node:util";
import { fileURLToPath } from "node:url";

import { changeRecordFrom } from "../../src/change-record.mjs";
import { RenderCache } from "../../src/render-cache.mjs";
import { ViewerInstance } from "../../src/server.mjs";
import { locateText, normalizeText } from "../../src/ui/locate-text.mjs";
import { planChangeMarks } from "../../src/ui/change-plan.mjs";
import { assertNoLeakedWord, newWordPids, wordPids } from "./word-pids.mjs";

const execFileAsync = promisify(execFile);
const HERE = path.dirname(fileURLToPath(import.meta.url));
const VENDOR = path.join(HERE, "..", "..", "src", "ui", "vendor");
// Deliberately not a word the fixture generator could produce on its own.
const MARKER = "Zebrafisch Quastenflosser Nebelkraehe";
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

const workRoot = await mkdtemp(path.join(tmpdir(), "word-pdfjs-test-"));
const fixture = path.join(workRoot, "report.docx");
const pidsBefore = await wordPids();

process.stderr.write("generating fixture...\n");
await makeFixture(fixture, 3);

const cache = new RenderCache({
    cacheRoot: path.join(workRoot, "artifacts"),
    log: (m) => process.stderr.write(`[cache] ${m}\n`),
});
const viewer = new ViewerInstance({
    cache,
    instanceId: "pdfjs-smoke-1",
    workspacePath: workRoot,
    log: (m) => process.stderr.write(`[viewer] ${m}\n`),
});

const get = (route, init) => fetch(new URL(route, viewer.url), init);

/** The manifest entry for a vendored file, by name -- never restated here. */
const fileEntry = (name) => {
    const entry = manifest.files.find((f) => f.name === name);
    assert.ok(entry, `no manifest entry for ${name}`);
    return entry;
};

let manifest = null;
let state = null;

try {
    await viewer.start();
    manifest = JSON.parse(await readFile(path.join(VENDOR, "pdfjs.manifest.json"), "utf8"));

    await check("a real Word export is served as a PDF", async () => {
        await viewer.openDocument(fixture);
        state = viewer.state;
        assert.equal(state.status, "ready", state.error?.message ?? "");
        assert.ok(state.doc.pageCount >= 1, `pageCount ${state.doc.pageCount}`);

        const res = await get(state.pdfUrl);
        assert.equal(res.status, 200);
        assert.match(res.headers.get("content-type"), /application\/pdf/);
        const body = Buffer.from(await res.arrayBuffer());
        assert.equal(body.subarray(0, 5).toString("latin1"), "%PDF-");
        assert.equal(Number(res.headers.get("content-length")), body.length);
        // Serving a length is what lets pdf.js decide it needs no ranges at all.
        assert.equal(res.headers.get("accept-ranges"), "bytes");
    });

    await check("the vendored worker reassembles byte-exact off the live instance", async () => {
        // The unit suite proves the parts on disk reassemble. This proves the
        // running server hands the browser those same bytes -- a different
        // claim, and the one the viewer actually depends on.
        const res = await get("/vendor/pdf.worker.min.mjs");
        assert.equal(res.status, 200);
        assert.match(res.headers.get("content-type"), /javascript/);
        const body = Buffer.from(await res.arrayBuffer());
        assert.equal(body.length, manifest.worker.bytes);
        assert.equal(createHash("sha256").update(body).digest("hex"), manifest.worker.sha256);
    });

    await check("every other vendored file is served beside the document", async () => {
        for (const entry of manifest.files) {
            const res = await get(`/vendor/${entry.name}`);
            assert.equal(res.status, 200, `${entry.name} -> ${res.status}`);
            const body = Buffer.from(await res.arrayBuffer());
            assert.equal(body.length, entry.bytes, entry.name);
            assert.equal(createHash("sha256").update(body).digest("hex"), entry.sha256, entry.name);
        }
        // Derived, so a vendoring bump that adds a file cannot leave this test
        // quietly checking the old set.
        assert.ok(fileEntry("pdf.min.mjs"), "the main build left the manifest");
    });

    await check("a record from a real edit describes the export that follows it", async () => {
        const read = await cache.readStructure(fixture);
        // A body paragraph well past the first, so `page` is not trivially 1.
        const target = read.paragraphs.filter((p) => p.headingLevel === null && p.text.length > 40).at(-1);
        assert.ok(target, "fixture had no body paragraph to edit");

        const result = await cache.editDocument(
            fixture,
            { op: "replace_text", address: target.address, text: `${MARKER}: die Zeile, auf die der Datensatz zeigt.` },
            { revisionToken: read.revisionToken },
        );

        const record = changeRecordFrom(result);
        assert.ok(record, "no record was built from a real edit");
        assert.equal(record.locatable, true);
        assert.equal(record.op, "replace_text");
        // The record's text is the post-edit read's, not the text we asked for:
        // Word is free to have changed it, and the overlay searches for what is
        // on the page rather than what was requested.
        assert.equal(record.text, normalizeText(result.paragraph.text));
        assert.ok(record.text.includes(MARKER), record.text);

        await viewer.refresh({ force: true, change: record });
        const after = viewer.state;
        assert.equal(after.status, "ready", after.error?.message ?? "");
        assert.ok(
            record.page >= 1 && record.page <= after.doc.pageCount,
            `record page ${record.page} outside 1..${after.doc.pageCount}`,
        );
        assert.notEqual(after.doc.key, state.doc.key, "the render key did not move across a real edit");
        // The record has to survive the stamping, or the overlay never sees it.
        assert.equal(after.change?.text, record.text, "the record did not reach the viewer state");
    });

    await check("the marking plan finds a real record on the page it names", async () => {
        // The one thing here that is not a pdf.js measurement: the locator and
        // the plan are fed the record from the edit above and a page carrying
        // the text Word actually wrote. It cannot prove pdf.js reports that text
        // -- only a browser can -- but it does prove the two halves agree on a
        // record that came out of Word rather than out of a fixture.
        const read = await cache.readStructure(fixture);
        const edited = read.paragraphs.find((p) => p.text.includes(MARKER));
        assert.ok(edited, "the edited paragraph was not found on re-read");

        const record = { op: "replace_text", page: 2, text: normalizeText(edited.text), locatable: true };
        const items = [{ str: "unrelated" }, { str: edited.text }];
        assert.equal(locateText(items, record.text).status, "located");

        const marks = planChangeMarks(record, [
            { number: 1, items: [{ str: "unrelated" }] },
            { number: 2, items },
        ]);
        assert.deepEqual(
            marks.map((m) => m.number),
            [2],
        );
        assert.ok(marks[0].found, "the page was marked without a match behind it");
    });
} finally {
    await viewer.close().catch(() => {});
    // Word holds the working copy open until the host is disposed; unlinking the
    // tree first fails with EBUSY, and the failure is the teardown's, not the
    // test's. Dispose, assert nothing leaked, then delete.
    await cache.dispose().catch(() => {});
}

const teardownAt = Date.now();

await check("no Word process was left behind", async () => {
    // Advisory, and the message the shared helper prints is stronger than what
    // it can know: "these started during this test, so they are ours" is exactly
    // the inference issue #25 says is unsound -- every WINWORD.EXE is parented to
    // the DCOM launcher, so a PID that appeared during the run may belong to any
    // session on the machine. Measured on this very run: three pids reported,
    // one of which started *after* teardown had begun, and two more appeared
    // while the report was being printed.
    //
    // So the failure is annotated with start times rather than suppressed. A
    // process that started after teardown cannot be ours; one that started
    // during the run and outlived it is a question worth answering by hand.
    try {
        await assertNoLeakedWord(pidsBefore);
    } catch (err) {
        const { stdout } = await execFileAsync("powershell.exe", [
            "-NoProfile",
            "-NonInteractive",
            "-Command",
            "Get-Process -Name WINWORD -ErrorAction SilentlyContinue | " +
                "ForEach-Object { '{0} {1:o}' -f $_.Id, $_.StartTime }",
        ]);
        const started = new Map(
            stdout
                .trim()
                .split(/\r?\n/)
                .filter(Boolean)
                .map((line) => line.split(" "))
                .map(([pid, iso]) => [Number(pid), Date.parse(iso)]),
        );
        const suspects = await newWordPids(pidsBefore);
        const detail = suspects
            .map((pid) => {
                const at = started.get(pid);
                if (at === undefined) return `  ${pid}: already exited`;
                const after = at > teardownAt ? " -- started AFTER teardown, cannot be ours" : "";
                return `  ${pid}: started ${new Date(at).toISOString()}${after}`;
            })
            .join("\n");
        err.message += `\n\nStart times (issue #25 -- attribution is unsound here):\n${detail}`;
        throw err;
    }
});

await rm(workRoot, { recursive: true, force: true }).catch(() => {});

const failed = results.filter((r) => !r.ok);
process.stderr.write(`\n${results.length - failed.length}/${results.length} checks passed\n`);
for (const f of failed) process.stderr.write(`\n${f.name}:\n${f.err.stack}\n`);
process.exit(failed.length ? 1 : 0);
