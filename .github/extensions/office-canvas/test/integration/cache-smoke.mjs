// Smoke test for the document service / PDF cache.
//
// Run: node .github/extensions/office-canvas/test/integration/cache-smoke.mjs

import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { mkdtemp, readFile, rm, stat, utimes } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { promisify } from "node:util";
import { fileURLToPath } from "node:url";
import { RenderCache, DocumentError, normalizeDocPath } from "../../src/render-cache.mjs";
import { assertNoLeakedWord, ownedWordLedger, wordPids } from "./word-pids.mjs";

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

const workRoot = await mkdtemp(path.join(tmpdir(), "word-cache-test-"));
const fixture = path.join(workRoot, "doc.docx");
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

const cache = new RenderCache({
    cacheRoot: path.join(workRoot, "artifacts"),
    log: (m) => process.stderr.write(`[cache] ${m}\n`),
});
// A leak assertion without a ledger can say a Word survived but not whose it
// was, and the two answers call for opposite responses -- our teardown versus
// another session's Word. Now that the host attributes by window handle rather
// than by differencing, what it reports here is provable, so the split in the
// failure message is worth something.
const ledger = ownedWordLedger().watch(cache.host);

let firstKey = null;

try {
    await check("open reports metadata for a real document", async () => {
        const info = await cache.open(fixture);
        assert.equal(info.path, normalizeDocPath(fixture));
        assert.ok(info.pageCount > 1, `expected multiple pages, got ${info.pageCount}`);
        assert.ok(info.docId.length >= 8);
        assert.ok(info.key);
        firstKey = info.key;
        process.stderr.write(`       ${info.name}: ${info.pageCount} pages, key ${info.key}\n`);
    });

    await check("open is idempotent for an unchanged file", async () => {
        const again = await cache.open(fixture);
        assert.equal(again.key, firstKey);
    });

    await check("a relative path resolves to the same document", async () => {
        const relative = path.relative(process.cwd(), fixture);
        const info = await cache.open(relative);
        assert.equal(info.key, firstKey);
        assert.equal(cache.openCount, 1, "the same file must not open twice");
    });

    let pdfFile = null;
    await check("pdf exports a renderable file", async () => {
        const t0 = Date.now();
        const res = await cache.pdf(fixture);
        pdfFile = res.file;
        const bytes = await readFile(res.file);
        assert.equal(bytes.subarray(0, 5).toString("latin1"), "%PDF-");
        assert.ok(bytes.length > 5000, `pdf looks too small: ${bytes.length} bytes`);
        assert.ok(res.pageCount > 1);
        process.stderr.write(`       ${bytes.length} bytes in ${Date.now() - t0}ms\n`);
    });

    await check("a cached pdf is served without re-exporting", async () => {
        const before = await stat(pdfFile);
        const t0 = Date.now();
        const res = await cache.pdf(fixture);
        const elapsed = Date.now() - t0;
        const after = await stat(pdfFile);
        assert.equal(res.file, pdfFile);
        assert.equal(after.mtimeMs, before.mtimeMs, "the cached pdf was rewritten");
        assert.ok(elapsed < 1500, `cache hit took ${elapsed}ms`);
    });

    await check("concurrent pdf requests share a single export", async () => {
        await rm(pdfFile, { force: true });
        const [a, b, c] = await Promise.all([cache.pdf(fixture), cache.pdf(fixture), cache.pdf(fixture)]);
        assert.equal(a.file, b.file);
        assert.equal(b.file, c.file);
        const bytes = await readFile(a.file);
        assert.equal(bytes.subarray(0, 5).toString("latin1"), "%PDF-");
    });

    await check("refresh reports no change when the file is untouched", async () => {
        const res = await cache.refresh(fixture);
        assert.equal(res.changed, false);
        assert.equal(res.key, firstKey);
    });

    await check("editing the file produces a new key and a fresh render", async () => {
        const stalePdf = pdfFile;
        // Rewrite the fixture with different content so both mtime and size move.
        await execFileAsync("powershell.exe", [
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            path.join(HERE, "make-fixture.ps1"),
            "-Out",
            fixture,
            "-Chapters",
            "4",
        ]);
        const res = await cache.refresh(fixture);
        assert.equal(res.changed, true, "refresh did not notice the edit");
        assert.notEqual(res.key, firstKey);

        const rendered = await cache.pdf(fixture);
        assert.notEqual(rendered.file, stalePdf);
        const bytes = await readFile(rendered.file);
        assert.equal(bytes.subarray(0, 5).toString("latin1"), "%PDF-");
        assert.ok(rendered.pageCount > 1);
        process.stderr.write(`       re-rendered: ${rendered.pageCount} pages, key ${res.key}\n`);
    });

    await check("stale renders are pruned", async () => {
        const { readdir } = await import("node:fs/promises");
        const info = await cache.open(fixture);
        const dir = path.join(workRoot, "artifacts", "pdf", info.docId);
        const entries = (await readdir(dir)).filter((n) => n.endsWith(".pdf"));
        assert.deepEqual(entries, [`${info.key}.pdf`], `unexpected cache contents: ${entries.join(", ")}`);
    });

    await check("a touch that changes nothing else still renders correctly", async () => {
        const when = new Date(Date.now() + 1000);
        await utimes(fixture, when, when);
        const res = await cache.refresh(fixture);
        assert.equal(res.changed, true, "mtime is part of the cache key");
        const rendered = await cache.pdf(fixture);
        const bytes = await readFile(rendered.file);
        assert.equal(bytes.subarray(0, 5).toString("latin1"), "%PDF-");
    });

    await check("outline, search and text all work through the cache", async () => {
        const outline = await cache.outline(fixture, { limit: 50 });
        assert.ok(Array.isArray(outline.headings) && outline.headings.length > 0);

        const hits = await cache.search(fixture, "ZORBLAX", { limit: 5 });
        assert.ok(hits.hits.length > 0, "the unique fixture token was not found");
        assert.ok(hits.hits[0].page >= 1);

        const text = await cache.text(fixture, { fromPage: 1, toPage: 1 });
        assert.ok(text.text.length > 50);
    });

    await check("a missing file is a typed error, not a crash", async () => {
        await assert.rejects(
            () => cache.open(path.join(workRoot, "nope.docx")),
            (err) => err instanceof DocumentError && err.code === "file_not_found",
        );
    });

    await check("a non-Word file is rejected before Word ever sees it", async () => {
        await assert.rejects(
            () => cache.open(path.join(workRoot, "notes.txt")),
            (err) => err instanceof DocumentError && err.code === "unsupported_type",
        );
    });

    await check("querying a document that is not open is a typed error", async () => {
        await assert.rejects(
            () => cache.outline(path.join(workRoot, "other.docx")),
            (err) => err instanceof DocumentError && err.code === "not_open",
        );
    });

    await check("close forgets the document and cleans its working copy", async () => {
        const info = await cache.open(fixture);
        const workDir = path.join(workRoot, "artifacts", "work", info.docId);
        await cache.close(fixture);
        assert.equal(cache.openCount, 0);
        const { existsSync } = await import("node:fs");
        assert.equal(existsSync(workDir), false, "the working copy was left behind");
    });
} finally {
    await cache.dispose();
}

await check("no new WINWORD.EXE is left behind", () => assertNoLeakedWord(pidsBefore, { ledger }));

await rm(workRoot, { recursive: true, force: true }).catch(() => {});

const failed = results.filter((r) => !r.ok);
process.stderr.write(`\n${results.length - failed.length}/${results.length} passed\n`);
process.exit(failed.length === 0 ? 0 : 1);
