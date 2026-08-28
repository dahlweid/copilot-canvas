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
import { execFile, spawn } from "node:child_process";
import { existsSync } from "node:fs";
import { mkdtemp, readFile, rm, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { promisify } from "node:util";
import { fileURLToPath } from "node:url";

import { RenderCache } from "../../src/render-cache.mjs";
import { fileRevisionToken, tokensMatch } from "../../src/revision-token.mjs";
import { assertNoLeakedWord, newWordPids, wordPids } from "./word-pids.mjs";

// Hold `target` open from a second process with a given FileShare mode, and do
// not return until the handle is provably taken.
//
// The handshake is the point. A fixed sleep before reading looks like it waits
// for the holder, but it only waits for the clock: on a loaded machine
// PowerShell can still be starting when the sleep expires, the read then runs
// against an unheld file and **succeeds without any contention**. The test goes
// green having exercised nothing. Timings on this machine are typical-case and
// never bounds, so a sleep can never establish this.
//
// The holder writes its ready file only after `[IO.File]::Open` returns, so the
// file's existence *is* evidence the handle is held; and it never closes the
// handle, so "still running" continues to mean "still holding". Callers assert
// `exitCode === null` after the work to close the other end of the bracket.
//
// Measured, by making the holder signal ready *before* opening: the exclusive
// (`None`) test goes red -- "a read of an exclusively held file should not
// succeed" -- while the `ReadWrite` test stays **green**. The two are
// asymmetric. A test asserting a refusal detects its own false-ready; a test
// asserting success cannot, because an unheld read succeeds for the wrong
// reason and looks identical. So the Word-like-holder check rests entirely on
// this handshake, which is why it may not be relaxed back to a sleep.
async function holdFile(target, share, readyPath) {
    const holder = spawn(
        "powershell.exe",
        [
            "-NoProfile",
            "-NonInteractive",
            "-Command",
            "$fs=[IO.File]::Open($env:PROBE_PATH,'Open','ReadWrite',$env:PROBE_SHARE);" +
                "Set-Content -LiteralPath $env:PROBE_READY -Value 'held';" +
                "while ($true) { Start-Sleep -Seconds 5 }",
        ],
        { env: { ...process.env, PROBE_PATH: target, PROBE_SHARE: share, PROBE_READY: readyPath } },
    );

    const deadline = Date.now() + 60_000;
    while (!existsSync(readyPath)) {
        if (holder.exitCode !== null) {
            holder.kill();
            throw new Error(`the ${share} holder exited before taking the handle`);
        }
        if (Date.now() > deadline) {
            holder.kill();
            throw new Error(`the ${share} holder never signalled ready`);
        }
        await new Promise((resolve) => setTimeout(resolve, 100));
    }
    return holder;
}

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
    await check("a missing file is reported without starting Word", async () => {
        // Deliberately the first check of the run: nothing has warmed Word yet,
        // so the pid assertion below can actually fail. Run after any successful
        // read it could not -- a warm Word would be reused and no new pid would
        // appear however late the existence check happened.
        //
        // The property is real on this path: DocumentReader.read() stats the
        // file and throws before #fetchMarkup, which is what starts the bridge.
        // Delete that stat and this goes red.
        // Narrow on purpose: "No such file" is the JS-side ReadError, so this
        // also pins *which layer* rejected. The host's own message is "File not
        // found", which this deliberately does not match -- if the request ever
        // reaches the host, that is the defect, not an alternative success.
        await assert.rejects(() => cache.readStructure(path.join(workRoot, "nope.docx")), /file_not_found|No such file/i);

        const appeared = await newWordPids(pidsBefore);
        assert.deepEqual(
            appeared,
            [],
            `a doomed read started Word (pid ${appeared.join(", ")}). ` +
                "If another session was driving Word during this window, re-run before believing this.",
        );
    });

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

        // Writable in the strongest sense: nobody holds a conflicting handle,
        // so we can open it denying all sharing. Node's own fs cannot show
        // this -- it always opens with FILE_SHARE_READ|WRITE|DELETE, so a
        // readFile/writeFile round trip succeeds even against a file Word has
        // open, and would prove nothing.
        //
        // The path travels in the environment rather than interpolated into
        // the command: a Windows profile may contain an apostrophe (O'Brien),
        // which would close the PowerShell string and fail the test for a
        // quoting reason that has nothing to do with what it asserts.
        await execFileAsync(
            "powershell.exe",
            [
                "-NoProfile",
                "-NonInteractive",
                "-Command",
                "[IO.File]::Open($env:PROBE_PATH, 'Open', 'ReadWrite', 'None').Close()",
            ],
            { env: { ...process.env, PROBE_PATH: fixture } },
        );

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

    await check("an exclusively locked original is reported as locked, not as a generic failure", async () => {
        // The one path units cannot cover: a real sharing violation from a real
        // second process. Measured behaviour this depends on -- an exclusive
        // hold gives EBUSY, whereas Word's own lock (a write handle granting
        // FileShare::ReadWrite) does not fail a ReadWrite-granting reader at
        // all -- is why `file_locked` means "stricter than Word", not "open in
        // Word". The companion check below asserts that other half.
        //
        // Safe despite ADR 0005: the revision token is read before Word is
        // started, so this fails fast on the filesystem and never reaches the
        // `Documents.Open` that would hang.
        const holder = await holdFile(fixture, "None", path.join(workRoot, "ready-none"));
        try {
            const err = await cache.readStructure(fixture).then(
                () => null,
                (e) => e,
            );
            assert.equal(holder.exitCode, null, "the holder released during the read, so nothing was contended");
            assert.ok(err, "a read of an exclusively held file should not succeed");
            assert.equal(err.code, "file_locked", `expected file_locked, got ${err.code}: ${err.message}`);
            assert.doesNotMatch(
                err.message,
                /permission/i,
                "a sharing violation must not be reported as a permissions problem",
            );
        } finally {
            holder.kill();
            await new Promise((resolve) => setTimeout(resolve, 500));
        }
    });

    await check("a read succeeds against a Word-like holder, so 'open in Word' is never file_locked", async () => {
        // The load-bearing half of the `file_locked` contract, and until now
        // the untested one. Every message and every ADR asserts that a document
        // the user has open in Word reads fine; nothing failed if that stopped
        // being true.
        //
        // The holder models Word exactly: a handle with **write** access,
        // granting ReadWrite. That distinction is the point. This repo
        // documented Word as taking `FileShare::Read` for months -- same
        // conclusion, wrong mechanism -- and the wrong one predicts that *any*
        // reader succeeds. It does not: a reader granting only `Read` refuses
        // to let anyone else write, conflicts with Word's write handle, and
        // gets a sharing violation on a file `Copy-Item` copies fine. So this
        // check is also the guard against someone "hardening" our copy to a
        // narrower share mode: that change breaks reads of open documents and
        // nothing else would notice.
        //
        // The holder is taken through `holdFile`, which does not return until the
        // handle is provably held -- a fixed sleep would let the read run against
        // an unheld file on a loaded machine and pass having contended nothing.
        // The `exitCode` assertion after the read closes the other end: the
        // holder never releases voluntarily, so still-running means still-holding.
        const holder = await holdFile(fixture, "ReadWrite", path.join(workRoot, "ready-rw"));
        try {
            const result = await cache.readStructure(fixture);
            assert.equal(holder.exitCode, null, "the holder released during the read, so nothing was contended");
            assert.ok(result.paragraphCount > 0, "a read of a Word-held document returned nothing");
            assert.ok(result.revisionToken, "a read of a Word-held document produced no revision token");
        } finally {
            holder.kill();
            await new Promise((resolve) => setTimeout(resolve, 500));
        }
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
