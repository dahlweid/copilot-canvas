// Smoke test for read_document against real Word.
//
// Run: node .github/extensions/office-canvas/test/integration/read-smoke.mjs
//
// Requires Word. Generates its own fixture (with duplicate paragraph text, so
// the occurrence index is actually exercised), reads it through the same path
// the tool uses, and ends by asserting it left no Word process behind.
//
// What this test is here to prove that the unit tests cannot:
//   * `Range.WordOpenXML` really does return a Flat OPC package with the styles
//     part in it, which is what makes localized styles resolvable in one call.
//   * The style ids Word writes on this machine really are localized.
//   * Reading does not lock or modify the original.

import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { mkdtemp, readFile, rm, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { promisify } from "node:util";
import { fileURLToPath } from "node:url";

import { RenderCache } from "../../src/render-cache.mjs";
import { fileRevisionToken, tokensMatch } from "../../src/revision-token.mjs";
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

const makeFixture = (out, { chapters = 2, duplicates = true } = {}) =>
    execFileAsync("powershell.exe", [
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
        ...(duplicates ? ["-Duplicates"] : []),
    ]);

const workRoot = await mkdtemp(path.join(tmpdir(), "word-read-test-"));
const fixture = path.join(workRoot, "docs", "structured.docx");
const pidsBefore = await wordPids();

process.stderr.write("generating fixture...\n");
await makeFixture(fixture);

const cache = new RenderCache({
    cacheRoot: path.join(workRoot, "artifacts"),
    log: (m) => process.stderr.write(`[cache] ${m}\n`),
});

let firstRead = null;

try {
    await check("a structure map and a revision token come back together", async () => {
        const started = Date.now();
        firstRead = await cache.readStructure(fixture);
        process.stderr.write(`  [read took ${Date.now() - started}ms for ${firstRead.paragraphCount} paragraphs]\n`);

        assert.ok(firstRead.paragraphCount > 30, `expected a substantial document, got ${firstRead.paragraphCount}`);
        assert.equal(firstRead.returned, firstRead.paragraphCount);
        assert.match(firstRead.revisionToken, /^sha256:[0-9a-f]{16}$/);
        assert.equal(firstRead.name, "structured.docx");
        assert.equal(firstRead.title, "Word Canvas Fixture");
    });

    await check("WordOpenXML carried the styles part, so styles are resolvable", async () => {
        // If this fails, style resolution silently degrades to w:outlineLvl only
        // and the whole styles-part strategy needs revisiting.
        assert.ok(firstRead.styleCount > 0, "no styles part in the returned package");
        process.stderr.write(`  [${firstRead.styleCount} styles, ${firstRead.markupBytes} bytes of markup]\n`);
    });

    await check("headings resolve to the right level despite a localized style id", async () => {
        const headings = firstRead.paragraphs.filter((p) => p.headingLevel !== null);
        assert.ok(headings.length >= 4, `expected headings, got ${headings.length}`);

        const chapter = headings.find((p) => p.text === "Chapter 1");
        assert.ok(chapter, "the level 1 heading was not found");
        assert.equal(chapter.headingLevel, 1);

        const section = headings.find((p) => p.text === "Section 1.1");
        assert.ok(section, "the level 2 heading was not found");
        assert.equal(section.headingLevel, 2);
        assert.deepEqual(section.headingPath, ["Chapter 1"]);

        // The measured surprise: Word does not write `Überschrift1`. It mints
        // the id from the localized name "Überschrift 1" and drops the
        // non-ASCII character, so the id in the file is `berschrift1`. Neither
        // the English name nor the obvious German spelling matches it, which is
        // why the id is only ever carried.
        process.stderr.write(`  [heading 1 style id on this machine: ${JSON.stringify(chapter.styleId)}]\n`);
        assert.ok(chapter.styleId, "no w:pStyle on a heading paragraph");
        assert.equal(chapter.styleName, "heading 1", "w:name should hold the canonical built-in name");
        assert.notEqual(chapter.styleId, chapter.styleName, "the id is not the canonical name");
    });

    await check("a body paragraph carries the full heading path", async () => {
        const body = firstRead.paragraphs.find((p) => p.text.startsWith("Paragraph 1.1.1"));
        assert.ok(body, "the first body paragraph was not found");
        assert.equal(body.headingLevel, null);
        assert.deepEqual(body.headingPath, ["Chapter 1", "Section 1.1"]);
    });

    await check("duplicate paragraph text is disambiguated by the occurrence index", async () => {
        // The case demo.docx never tested: identical text, same heading path.
        const duplicates = firstRead.paragraphs.filter((p) => p.text.startsWith("DUPLICATE LINE:"));
        assert.ok(duplicates.length >= 6, `expected repeated paragraphs, got ${duplicates.length}`);

        const underOne = duplicates.filter((p) => p.headingPath[0] === "Twin Chapter 1");
        assert.equal(underOne.length, 3);
        assert.deepEqual(
            underOne.map((p) => p.occurrence),
            [1, 2, 3],
            "the occurrence index did not increment for identical text",
        );
        assert.equal(new Set(underOne.map((p) => p.address)).size, 3, "identical paragraphs shared an address");
    });

    await check("every address in the document is unique", async () => {
        const addresses = firstRead.paragraphs.map((p) => p.address);
        const unique = new Set(addresses);
        assert.equal(unique.size, addresses.length, `${addresses.length - unique.size} addresses collided`);
    });

    await check("the same text under a different heading is a different address", async () => {
        const [one] = firstRead.paragraphs.filter(
            (p) => p.text.startsWith("DUPLICATE LINE:") && p.headingPath[0] === "Twin Chapter 1",
        );
        const [two] = firstRead.paragraphs.filter(
            (p) => p.text.startsWith("DUPLICATE LINE:") && p.headingPath[0] === "Twin Chapter 2",
        );
        assert.equal(one.text, two.text);
        assert.equal(one.occurrence, two.occurrence, "the occurrence counter is scoped to the heading path");
        assert.notEqual(one.address, two.address);
    });

    await check("a second read of an untouched document gives the same token and addresses", async () => {
        const again = await cache.readStructure(fixture);
        assert.ok(tokensMatch(again.revisionToken, firstRead.revisionToken), "the token moved without an edit");
        assert.deepEqual(
            again.paragraphs.map((p) => p.address),
            firstRead.paragraphs.map((p) => p.address),
        );
    });

    await check("the revision token matches the file's own bytes", async () => {
        assert.equal(await fileRevisionToken(fixture), firstRead.revisionToken);
    });

    await check("reading neither locks nor modifies the original", async () => {
        // Word opens a copy, so the original is never held: a read is free to
        // overlap with a script regenerating the document.
        assert.equal(firstRead.writable, true, "the original was reported as locked during a read");

        const info = await stat(fixture);
        assert.equal(info.size, firstRead.sizeBytes);
        // Writable in the strongest sense: we can open it for exclusive write.
        const handle = await readFile(fixture);
        await writeFile(fixture, handle);
        assert.equal(await fileRevisionToken(fixture), firstRead.revisionToken, "rewriting identical bytes moved the token");
    });

    await check("paging returns a window without moving an address", async () => {
        const page = await cache.readStructure(fixture, { limit: 5, offset: 10 });
        assert.equal(page.paragraphCount, firstRead.paragraphCount);
        assert.equal(page.returned, 5);
        assert.equal(page.truncated, true);
        assert.deepEqual(
            page.paragraphs.map((p) => p.address),
            firstRead.paragraphs.slice(10, 15).map((p) => p.address),
        );
    });

    await check("the same content authored in a separate Word run gives the same addresses", async () => {
        // The decisive test that addresses are derived from content and nothing
        // else. Word stamps per-session revision identifiers (w:rsidR and
        // friends) into the markup, so a second run produces a byte-different
        // file for identical content. Addresses must not notice; the revision
        // token must.
        const twin = path.join(workRoot, "docs", "twin.docx");
        await makeFixture(twin);
        const other = await cache.readStructure(twin);

        assert.deepEqual(
            other.paragraphs.map((p) => p.address),
            firstRead.paragraphs.map((p) => p.address),
            "identical content produced different addresses across two Word runs",
        );
        assert.ok(
            !tokensMatch(other.revisionToken, firstRead.revisionToken),
            "two separately authored files should not share a revision token",
        );
        process.stderr.write(`  [twin: same ${other.paragraphCount} addresses, token ${other.revisionToken}]\n`);
    });

    await check("regenerating the document moves the revision token", async () => {
        // The behaviour optimistic concurrency depends on: an edit citing the
        // old token must be refusable after this.
        await makeFixture(fixture, { chapters: 3 });
        const after = await cache.readStructure(fixture);
        assert.ok(!tokensMatch(after.revisionToken, firstRead.revisionToken), "the token survived a regeneration");
        assert.ok(after.paragraphCount > firstRead.paragraphCount);
    });

    await check("a missing file is reported without starting Word", async () => {
        await assert.rejects(() => cache.readStructure(path.join(workRoot, "nope.docx")), /file_not_found|No such file/i);
    });

    await check("reading needs no canvas open", async () => {
        // The tools are the product; the canvas is a display surface. Nothing
        // above ever opened one.
        assert.equal(cache.openCount, 0, "a structure read left a document open");
    });
} finally {
    await cache.dispose().catch(() => {});
}

await check("no new WINWORD.EXE is left behind", () => assertNoLeakedWord(pidsBefore));

await rm(workRoot, { recursive: true, force: true }).catch(() => {});

const failed = results.filter((r) => !r.ok);
process.stderr.write(`\n${results.length - failed.length}/${results.length} passed\n`);
process.exit(failed.length === 0 ? 0 : 1);
