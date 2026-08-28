// Smoke test for edit_document and revert_document against real Word.
//
// Run: node .github/extensions/office-canvas/test/integration/edit-smoke.mjs
//
// Requires Word. This is the suite that exercises the things no Office-free
// test can reach, and every one of them was a surprise when first measured:
//
//   * An edit reaches the user's own file, in place, and the lock is gone
//     afterwards (ADR 0005). The budget is one warm operation, ~228 ms.
//   * A mark-of-the-web document can be edited *without* stripping its zone
//     marker, via the Protected View route (ADR 0007). Measured: a plain
//     `Documents.Open` on such a file hangs forever rather than failing.
//   * A locked document is refused cleanly, with no `~$` stub left behind --
//     the failure mode that needed two processes killed by hand.
//   * The address -> `Document.Paragraphs` join survives a table, which a
//     document-order walk does not.
//   * Text goes in verbatim: no autocorrect, no smart quotes.

import assert from "node:assert/strict";
import { execFile, spawn } from "node:child_process";
import { access, mkdtemp, readdir, readFile, rm, writeFile } from "node:fs/promises";
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
        process.stderr.write(`  FAIL ${name}\n       ${err.stack ?? err.message}\n`);
    }
}

const powershell = (script) =>
    execFileAsync("powershell.exe", ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", script]);

const makeFixture = (out, { chapters = 2, duplicates = true, table = true } = {}) =>
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
        ...(table ? ["-Table"] : []),
    ]);

const exists = (p) =>
    access(p).then(
        () => true,
        () => false,
    );

/** The `~$name.docx` owner file Word leaves behind when a document is held. */
const ownerFileFor = (docPath) => path.join(path.dirname(docPath), `~$${path.basename(docPath)}`);

const workRoot = await mkdtemp(path.join(tmpdir(), "word-edit-test-"));
const docs = path.join(workRoot, "docs");
const fixture = path.join(docs, "editable.docx");
const pidsBefore = await wordPids();

process.stderr.write("generating fixture...\n");
await makeFixture(fixture);

const cache = new RenderCache({
    cacheRoot: path.join(workRoot, "artifacts"),
    log: (m) => process.stderr.write(`[cache] ${m}\n`),
});

/** Re-reads and returns the paragraph at an address, or null. */
const paragraphAt = async (docPath, address) => {
    const map = await cache.readStructure(docPath);
    return map.paragraphs.find((p) => p.address === address) ?? null;
};

try {
    let read = await cache.readStructure(fixture);

    await check("replacing a paragraph's text changes that paragraph and nothing else", async () => {
        const before = read.paragraphs.find((p) => p.headingLevel === null && p.text.length > 20);
        assert.ok(before, "no body paragraph in the fixture");

        const untouched = read.paragraphs.filter((p) => p.address !== before.address).map((p) => p.address);

        const result = await cache.editDocument(
            fixture,
            { op: "replace_text", address: before.address, text: "Rewritten by the edit smoke test." },
            { revisionToken: read.revisionToken },
        );

        assert.equal(result.paragraph.text, "Rewritten by the edit smoke test.");
        assert.equal(result.document.paragraphCount, read.paragraphCount, "the paragraph count moved");
        assert.ok(!tokensMatch(result.document.revisionToken, read.revisionToken), "the token did not move");

        // Every other address survived, which is what makes an address a usable
        // coordinate rather than a one-shot ticket.
        const after = new Set(result.document.paragraphs.map((p) => p.address));
        const lost = untouched.filter((a) => !after.has(a));
        assert.deepEqual(lost, [], `rewriting one paragraph invalidated ${lost.length} unrelated addresses`);

        process.stderr.write(
            `  [lock held ${result.timings.lockHeldMs}ms: open ${result.timings.openMs}, edit ${result.timings.editMs}, ` +
                `save ${result.timings.saveMs}, release ${result.timings.releaseMs}]\n`,
        );
        read = result.document;
    });

    await check("the edit is in the user's own file, not a copy", async () => {
        // ADR 0005 in one assertion: the bytes on disk changed, and the change
        // is the text we asked for.
        const bytes = await readFile(fixture);
        assert.ok(bytes.length > 0);
        assert.equal(await fileRevisionToken(fixture), read.revisionToken);
        assert.ok(
            read.paragraphs.some((p) => p.text === "Rewritten by the edit smoke test."),
            "a fresh read of the original does not show the edit",
        );
    });

    await check("the lock is released when the operation ends", async () => {
        // Not assumed: `Close()` returning is not proof the file is free, in the
        // same way `Quit()` returns ~120 ms before Word's process exits.
        // Anything else here would be a re-open into the unbounded hang.
        const handle = await readFile(fixture);
        await writeFile(fixture, handle);
        assert.equal(await fileRevisionToken(fixture), read.revisionToken, "rewriting identical bytes moved the token");
        assert.equal(await exists(ownerFileFor(fixture)), false, "Word left a ~$ owner file behind");
    });

    await check("text goes in verbatim -- no autocorrect, no smart quotes", async () => {
        // `Range.Text` is not typing, so Word's autocorrect should not fire.
        // Asserted rather than assumed: an agent that writes code samples into a
        // document cannot have "(c)" silently become a copyright sign.
        const verbatim = 'straight "quotes", (c), teh, i, 1/2 and -- a dash';
        const target = read.paragraphs.find((p) => p.headingLevel === null);
        const result = await cache.editDocument(
            fixture,
            { op: "replace_text", address: target.address, text: verbatim },
            { revisionToken: read.revisionToken },
        );
        assert.equal(result.paragraph.text, verbatim);
        read = result.document;
    });

    await check("a paragraph can be inserted, and inherits a sensible style", async () => {
        const heading = read.paragraphs.find((p) => p.headingLevel === 1);
        assert.ok(heading, "no level 1 heading in the fixture");

        const result = await cache.editDocument(
            fixture,
            { op: "insert_paragraph_after", address: heading.address, text: "Inserted under the first heading." },
            { revisionToken: read.revisionToken },
        );

        assert.equal(result.document.paragraphCount, read.paragraphCount + 1);
        assert.equal(result.paragraph.text, "Inserted under the first heading.");
        // Word's own rule for pressing Enter at the end of a heading: the next
        // paragraph is body text, not another heading.
        assert.equal(result.paragraph.headingLevel, null, "the inserted paragraph inherited the heading style");
        assert.deepEqual(result.paragraph.headingPath, [heading.text]);
        read = result.document;
    });

    await check("a heading level can be set without naming a style", async () => {
        // The localization trap, closed: `Range.Style` accepts neither the style
        // id Word writes ("berschrift1") nor the English name ("Heading 1") --
        // both throw here. Only the numeric wd* constant works, so only a
        // numeric level is ever sent.
        const body = read.paragraphs.find((p) => p.text === "Inserted under the first heading.");
        const result = await cache.editDocument(
            fixture,
            { op: "set_heading_level", address: body.address, headingLevel: 2 },
            { revisionToken: read.revisionToken },
        );
        assert.equal(result.paragraph.headingLevel, 2);
        process.stderr.write(`  [level 2 heading has styleId '${result.paragraph.styleId}']\n`);
        read = result.document;
    });

    await check("a paragraph can be deleted", async () => {
        const victim = read.paragraphs.find((p) => p.text === "Inserted under the first heading.");
        const result = await cache.editDocument(
            fixture,
            { op: "delete_paragraph", address: victim.address },
            { revisionToken: read.revisionToken },
        );
        assert.equal(result.document.paragraphCount, read.paragraphCount - 1);
        assert.equal(await paragraphAt(fixture, victim.address), null, "the deleted paragraph is still addressable");
        read = result.document;
    });

    await check("an edit inside a table hits the right paragraph", async () => {
        // The join that a document-order walk gets wrong: Word counts an extra
        // paragraph for every table row's end-of-row mark, so anything after a
        // table is offset. Measured on a 2x2-table fixture: map 8, Word 10.
        const cell = read.paragraphs.find((p) => p.inTable);
        assert.ok(cell, "no table paragraph in the fixture");

        const result = await cache.editDocument(
            fixture,
            { op: "replace_text", address: cell.address, text: "cell rewritten" },
            { revisionToken: read.revisionToken },
        );
        assert.equal(result.paragraph.text, "cell rewritten");
        assert.equal(result.paragraph.inTable, true, "the edit landed outside the table");
        read = result.document;
    });

    await check("revert restores the previous state and steps backwards", async () => {
        const beforeRevert = read.revisionToken;
        const result = await cache.revertDocument(fixture);

        assert.ok(!tokensMatch(result.document.revisionToken, beforeRevert), "revert did not change the file");
        assert.ok(
            !result.document.paragraphs.some((p) => p.text === "cell rewritten"),
            "the reverted edit is still in the document",
        );
        assert.equal(result.restored.op, "replace_text");

        // Consumed, so a second revert goes further back rather than redoing.
        const second = await cache.revertDocument(fixture);
        assert.equal(second.snapshotsRemaining, result.snapshotsRemaining - 1);
        read = second.document;
    });

    await check("a stale revision token is refused before Word is touched", async () => {
        await assert.rejects(
            () =>
                cache.editDocument(
                    fixture,
                    { op: "replace_text", address: read.paragraphs[0].address, text: "nope" },
                    { revisionToken: "sha256:0000000000000000" },
                ),
            (err) => err.code === "stale_revision_token",
        );
    });

    await check("an edit with no revision token is refused", async () => {
        // Null-safe by construction: a missing token must never compare equal.
        await assert.rejects(
            () => cache.editDocument(fixture, { op: "delete_paragraph", address: read.paragraphs[0].address }, {}),
            (err) => err.code === "stale_revision_token",
        );
    });

    await check("an address that no longer exists is refused, not guessed", async () => {
        const token = await fileRevisionToken(fixture);
        await assert.rejects(
            () => cache.editDocument(fixture, { op: "replace_text", address: "p:ffffffffffff", text: "nope" }, { revisionToken: token }),
            (err) => err.code === "address_not_found",
        );
    });

    await check("a locked document is refused cleanly and leaves no ~$ stub", async () => {
        // The measured disaster this guards: a second Word opening a held file
        // hangs indefinitely, with DisplayAlerts already off, and both processes
        // have to be killed by hand. The pre-flight write-handle probe is what
        // stops us ever making that call.
        // FileShare::Read, not None, because that is the lock Word itself takes:
        // a document open in Word can still be copied, which is why layer 1's
        // copy-based read keeps working while the original is held. (With
        // FileShare::None even the read fails, with a raw EBUSY from copyFile --
        // a gap in layer 1's error contract, reported separately.)
        const holder = spawn(
            "powershell.exe",
            [
                "-NoProfile",
                "-NonInteractive",
                "-Command",
                `$s=[IO.File]::Open('${fixture.replace(/'/g, "''")}',[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::Read); Start-Sleep -Seconds 25; $s.Close()`,
            ],
            { stdio: "ignore" },
        );
        try {
            // Give the holder time to actually take the handle.
            for (let i = 0; i < 50; i++) {
                const map = await cache.readStructure(fixture).catch(() => null);
                if (map && map.writable === false) break;
                await new Promise((r) => setTimeout(r, 200));
            }

            const locked = await cache.readStructure(fixture);
            assert.equal(locked.writable, false, "the holder never took an exclusive handle");

            const started = Date.now();
            await assert.rejects(
                () =>
                    cache.editDocument(
                        fixture,
                        { op: "replace_text", address: locked.paragraphs[0].address, text: "nope" },
                        { revisionToken: locked.revisionToken },
                    ),
                (err) => err.code === "document_locked",
            );
            const elapsed = Date.now() - started;
            assert.ok(elapsed < 30_000, `refusing a locked document took ${elapsed}ms; it should be immediate`);
            assert.equal(await exists(ownerFileFor(fixture)), false, "a ~$ owner file appeared");
            process.stderr.write(`  [locked document refused in ${elapsed}ms]\n`);
        } finally {
            holder.kill();
        }
    });

    await check("a mark-of-the-web document is edited without stripping its zone marker", async () => {
        // ADR 0007. `Documents.Open` on a marked file does not fail -- it hangs,
        // exactly as a locked file does, which is why the check is an alternate
        // data stream read and never a trial open. The Protected View route is
        // the automation equivalent of clicking "Enable Editing"; the cost is a
        // per-file trust record under HKCU, and the marker itself survives.
        const marked = path.join(docs, "downloaded.docx");
        await makeFixture(marked, { chapters: 1, duplicates: false, table: false });
        await powershell(
            `Set-Content -LiteralPath '${marked.replace(/'/g, "''")}' -Stream Zone.Identifier -Value "[ZoneTransfer]\`r\`nZoneId=3"`,
        );

        const zoneBefore = (
            await powershell(`Get-Content -LiteralPath '${marked.replace(/'/g, "''")}' -Stream Zone.Identifier -Raw`)
        ).stdout;
        assert.match(zoneBefore, /ZoneId=3/, "the fixture is not actually marked");

        const map = await cache.readStructure(marked);
        const target = map.paragraphs.find((p) => p.headingLevel === null);

        const started = Date.now();
        const result = await cache.editDocument(
            marked,
            { op: "replace_text", address: target.address, text: "Edited through Protected View." },
            { revisionToken: map.revisionToken },
        );
        process.stderr.write(`  [protected-view edit took ${Date.now() - started}ms]\n`);

        assert.equal(result.markOfTheWeb, true, "the mark of the web was not detected");
        assert.equal(result.protectedView, true, "the direct open path was used on a marked file");
        assert.equal(result.paragraph.text, "Edited through Protected View.");

        const zoneAfter = (
            await powershell(`Get-Content -LiteralPath '${marked.replace(/'/g, "''")}' -Stream Zone.Identifier -Raw`)
        ).stdout;
        assert.match(zoneAfter, /ZoneId=3/, "editing stripped the file's mark of the web");
    });

    await check("editing needs no canvas open", async () => {
        assert.equal(cache.openCount, 0, "an edit left a document open");
    });

    await check("snapshots are capped rather than growing without bound", async () => {
        const dir = path.join(workRoot, "artifacts", "snapshots");
        const perDoc = await readdir(dir);
        assert.ok(perDoc.length >= 1, "no snapshots were taken");
        for (const sub of perDoc) {
            const files = await readdir(path.join(dir, sub));
            assert.ok(files.length <= 40, `${files.length} snapshot files for one document`);
        }
    });
} finally {
    await cache.dispose().catch(() => {});
}

await check("no new WINWORD.EXE is left behind", () => assertNoLeakedWord(pidsBefore));

await rm(workRoot, { recursive: true, force: true }).catch(() => {});

const failed = results.filter((r) => !r.ok);
process.stderr.write(`\n${results.length - failed.length}/${results.length} passed\n`);
process.exit(failed.length === 0 ? 0 : 1);
