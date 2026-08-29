// Smoke test for edit_document and revert_document against real Word.
//
// Run: node .github/extensions/office-canvas/test/integration/edit-smoke.mjs
//
// Requires Word. This is the suite that exercises the things no Office-free
// test can reach, and every one of them was a surprise when first measured:
//
//   * An edit reaches the user's own file, in place, and the lock is gone
//     afterwards (ADR 0005). Measured warm: the lock is held ~300-500 ms, of
//     which open/edit/save is only ~230 -- the rest is post-save COM reporting
//     that happens while the document is still open. Typical-case figures on a
//     quiet machine, not bounds; the assertions here are relational for that
//     reason, since a threshold would only measure the runner.
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
import { access, copyFile, mkdtemp, readdir, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { promisify } from "node:util";
import { fileURLToPath } from "node:url";

import { RenderCache } from "../../src/render-cache.mjs";
import { fileRevisionToken, tokensMatch } from "../../src/revision-token.mjs";
import { asToolError } from "../../src/tool-error.mjs";
import { codepoints, documentPlainText, documentXml } from "./docx-zip.mjs";
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

// Paths reach the script through the environment, never interpolated into it.
// Escaping was tried first and is the wrong fix: the danger set is per-parser
// and barely overlaps -- an apostrophe breaks a PowerShell single-quoted
// literal and is harmless to cmd, while `&` and `^` break cmd and are harmless
// here. Doubling apostrophes therefore looks correct right up until the path
// reaches a different parser. Removing the interpolation removes the whole
// class, including the next parser's set, which we do not know yet.
const powershell = (script, vars = {}) =>
    execFileAsync("powershell.exe", ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", script], {
        env: { ...process.env, ...vars },
    });

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

/**
 * Holds an open handle on a file until the returned handle is released.
 *
 * Both brackets of the hold are guarded, because a contention test is only
 * meaningful while the contention exists and neither end of that window is
 * self-evident:
 *
 *   * **Acquisition.** The holder writes its ready file only after
 *     `[IO.File]::Open` has returned, so waiting on that file waits on the
 *     handle rather than on the process having been spawned. Sleeping a fixed
 *     interval instead would be a guess about scheduler latency.
 *   * **Release.** The holder never closes the handle and never times out; it
 *     blocks until it is killed. A time-boxed hold makes the test's outcome
 *     depend on how fast the machine is, which this repo treats as never a
 *     bound. `release()` asserts the process was still alive when the work
 *     finished, so a holder that died early fails loudly instead of quietly
 *     turning the test into one that measured nothing.
 *
 * A mutation check tells you which of your assertions this actually protects:
 * signal ready *before* opening, and a test asserting a **refusal** fails,
 * because the refusal stops firing. A test asserting a **success** under
 * contention still passes, because an uncontended success is identical on every
 * observable. Only the refusing kind detects its own false ready.
 */
async function holdFile(target, share) {
    const readyPath = `${target}.holder-ready`;
    await rm(readyPath, { force: true });

    const child = spawn(
        "powershell.exe",
        [
            "-NoProfile",
            "-NonInteractive",
            "-Command",
            "$s=[IO.File]::Open($env:HOLD_DOC,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::" +
                `${share});` +
                // Signalled only once the handle is genuinely held.
                "Set-Content -LiteralPath $env:HOLD_READY -Value 'held';" +
                // Never closes, never lapses: the handle outlives the work by
                // construction, and the kill in release() is the only way out.
                "while ($true) { Start-Sleep -Seconds 3600 }",
        ],
        { stdio: "ignore", env: { ...process.env, HOLD_DOC: target, HOLD_READY: readyPath } },
    );

    let exitedEarly = null;
    child.once("exit", (code) => {
        exitedEarly = code;
    });

    for (let i = 0; i < 100; i++) {
        if (await exists(readyPath)) break;
        if (exitedEarly !== null) {
            throw new Error(`the holder exited with ${exitedEarly} before taking a handle on ${target}`);
        }
        await new Promise((r) => setTimeout(r, 50));
    }
    if (!(await exists(readyPath))) {
        child.kill();
        throw new Error(`the holder never signalled that it had opened ${target}`);
    }

    return {
        /**
         * Asserts the handle was still held. Call this *before* examining the
         * result of the contended operation: if the holder died, that is the
         * explanation, and reporting it first beats a confusing "expected a
         * rejection" from an operation that correctly succeeded against a file
         * nobody was holding any more.
         */
        assertStillHeld() {
            assert.ok(
                child.exitCode === null && child.signalCode === null,
                `the holder released ${path.basename(target)} before the test finished, so nothing was contended`,
            );
        },
        kill() {
            child.kill();
        },
    };
}

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

        // ADR 0005's central claim is that the lock window is short, so the
        // number cited as evidence for it must measure the window and not the
        // command. It did not: `lockHeldMs` was the whole-command stopwatch,
        // started before the pre-flight write probe and stopped after the
        // post-close release poll, so it overstated the window -- in the
        // flattering direction, which is the one nobody checks.
        //
        // These are relational, not thresholds: the machine's speed is not the
        // property under test, and a bound would only measure the runner.
        const t = result.timings;
        assert.ok(
            t.lockHeldMs < t.totalMs,
            `lockHeldMs (${t.lockHeldMs}ms) must exclude the pre-flight and the release poll that totalMs (${t.totalMs}ms) includes`,
        );
        assert.ok(
            t.lockHeldMs >= t.editMs + t.saveMs,
            `lockHeldMs (${t.lockHeldMs}ms) must cover the edit and save it contains (${t.editMs}+${t.saveMs}ms)`,
        );
        assert.ok(
            t.openMs <= t.lockHeldMs,
            `openMs (${t.openMs}ms) happens inside the lock window (${t.lockHeldMs}ms)`,
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
        // same way `Quit()` returns seconds before Word's process exits
        // (measured 3-28 ms to return, 3039-3702 ms to exit --
        // probe-quit-exit-gap.ps1).
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
        // The token is required now, and each revert mints a new one, so a
        // second revert has to carry the token the first one returned. That is
        // the contract working, not ceremony: it is what stops a revert
        // overwriting a document someone changed in between.
        const beforeRevert = read.revisionToken;
        const result = await cache.revertDocument(fixture, { revisionToken: beforeRevert });

        assert.ok(!tokensMatch(result.document.revisionToken, beforeRevert), "revert did not change the file");
        assert.ok(
            !result.document.paragraphs.some((p) => p.text === "cell rewritten"),
            "the reverted edit is still in the document",
        );
        assert.equal(result.restored.op, "replace_text");

        // Consumed, so a second revert goes further back rather than redoing.
        const second = await cache.revertDocument(fixture, { revisionToken: result.document.revisionToken });
        assert.equal(second.snapshotsRemaining, result.snapshotsRemaining - 1);
        read = second.document;
    });

    await check("a restore that cannot run is typed, and says the document is intact", async () => {
        // The recovery path's own failure mode. Reaching it needs contention on
        // the *snapshot*, not on the document: the writable pre-flight at the
        // top of revert() refuses a held document before `revertToLatest` runs,
        // so holding the document would test the pre-flight and never this.
        //
        // Asserted through `asToolError`, because that is the boundary the agent
        // stands at, and an error that is correct one layer beneath it is a
        // defect this repo has already shipped once.
        const root = path.join(workRoot, "artifacts", "snapshots");
        const snapshots = [];
        for (const sub of await readdir(root)) {
            const dir = path.join(root, sub);
            for (const name of await readdir(dir)) {
                if (!name.endsWith(".snapshot")) continue;
                const manifest = JSON.parse(await readFile(path.join(dir, `${name}.json`), "utf8"));
                if (manifest.documentPath === fixture) snapshots.push(path.join(dir, name));
            }
        }
        snapshots.sort();
        const snapshot = snapshots.at(-1);
        assert.ok(snapshot, "no snapshot of the fixture to contend for");

        const holder = await holdFile(snapshot, "None");
        const before = await fileRevisionToken(fixture);
        let thrown = null;
        try {
            await cache.revertDocument(fixture, { revisionToken: before });
        } catch (err) {
            thrown = err;
        } finally {
            holder.assertStillHeld();
            holder.kill();
        }

        assert.ok(thrown, "a revert whose snapshot could not be read reported success");
        const wire = asToolError(thrown);
        assert.equal(wire.code, "revert_failed", `unexpected code ${wire.code}: ${wire.message}`);
        assert.equal(wire.data?.step, "copy");
        assert.equal(wire.data?.errno, "EBUSY", `unexpected errno ${wire.data?.errno}`);
        assert.equal(wire.data?.documentUnchanged, true);
        assert.equal(wire.data?.snapshotRetained, true);

        // The two claims on `data`, checked rather than trusted -- they are the
        // whole reason this failure is safe to retry.
        assert.ok(tokensMatch(before, await fileRevisionToken(fixture)), "the document changed despite the failure");
        assert.ok(await exists(snapshot), "the failed revert discarded the snapshot it could not read");

        // Nothing was consumed, so the document is still revertable afterwards.
        // `kill()` returns before the handle is gone, so wait on the handle
        // rather than on the signal -- the release end of a hold is no more
        // self-evident than the acquisition end.
        for (let i = 0; i < 100; i++) {
            try {
                await readFile(snapshot);
                break;
            } catch {
                await new Promise((r) => setTimeout(r, 50));
            }
        }
        const recovered = await cache.revertDocument(fixture, { revisionToken: before });
        read = recovered.document;
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
        // The holder grants FileShare::ReadWrite because that is what Word itself
        // grants: probed, Word holds a *write* handle and shares ReadWrite. It is
        // not FileShare::Read, and that distinction is not cosmetic here -- of the
        // share modes a holder can take, ReadWrite is the most permissive, so it
        // is the hardest case for a refusal to fire. A test holding a stricter
        // mode would pass while proving strictly less.
        //
        // It fires anyway because Test-FileWritable opens with FileShare::None,
        // which conflicts with any existing handle regardless of what that handle
        // shares. So the refusal covers the case that actually happens in
        // production -- the user has the document open in Word -- and this is the
        // one place read and edit genuinely diverge: layer 1's copy-based read
        // sails through this same lock, because the copy is not the original and
        // an edit has to be.
        const holder = await holdFile(fixture, "ReadWrite");
        try {
            // The ready file proves a handle exists; this proves it is the
            // handle that matters. `writable` is the product's own observable,
            // the same field the edit path pre-flights on, so asserting it here
            // means the test contends with exactly what production contends
            // with rather than with something merely correlated.
            const locked = await cache.readStructure(fixture);
            assert.equal(locked.writable, false, "the holder never took a handle the pre-flight can see");

            const started = Date.now();
            let refusal = null;
            try {
                await cache.editDocument(
                    fixture,
                    { op: "replace_text", address: locked.paragraphs[0].address, text: "nope" },
                    { revisionToken: locked.revisionToken },
                );
            } catch (err) {
                refusal = err;
            }
            const elapsed = Date.now() - started;

            holder.assertStillHeld();
            assert.ok(refusal, "editing a locked document was not refused");
            assert.equal(refusal.code, "file_locked", `expected file_locked, got ${refusal.code}`);
            // A few hundred milliseconds in practice, and the spread is wide
            // enough across runs that quoting a tight range would just be a
            // number to go stale: the write-handle pre-flight is ~4 ms of it
            // and the rest is the read that precedes it. The bound is
            // deliberately nowhere near it. Every timing in this repo is
            // typical-case and never a bound, and Word's shutdown has already
            // been measured contending on per-user state when another session
            // drives it -- so a tight bound here would fail for load rather
            // than for regression. What must not regress is the *shape*: a
            // refusal that never opens Word, rather than the indefinite hang a
            // trial open gives on a held file. Three seconds is far below that
            // hang and far above the noise.
            assert.ok(elapsed < 3_000, `refusing a locked document took ${elapsed}ms; it should not involve Word`);
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
            `Set-Content -LiteralPath $env:SMOKE_DOC -Stream Zone.Identifier -Value "[ZoneTransfer]\`r\`nZoneId=3"`,
            { SMOKE_DOC: marked },
        );

        const zoneBefore = (
            await powershell(`Get-Content -LiteralPath $env:SMOKE_DOC -Stream Zone.Identifier -Raw`, { SMOKE_DOC: marked })
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

        // The Protected View route opens a second COM object -- a
        // ProtectedViewWindow -- that the ordinary path never creates, and that
        // window holds the user's file. Nothing asserted it was gone, so a
        // retained handle here would have surfaced only as the next edit hanging
        // on its open. It is checked through `lockReleased`, which the host sets
        // by polling `Test-FileWritable`: that grants `FileShare::None` and
        // requests write, so it conflicts with any handle whatever it shares.
        assert.equal(result.lockReleased, true, "the Protected View path left the file held after the edit");

        const zoneAfter = (
            await powershell(`Get-Content -LiteralPath $env:SMOKE_DOC -Stream Zone.Identifier -Raw`, { SMOKE_DOC: marked })
        ).stdout;
        assert.match(zoneAfter, /ZoneId=3/, "editing stripped the file's mark of the web");
    });

    await check("paragraphs carrying Word's layout marks are editable", async () => {
        // Regression for a defect that made ordinary documents silently
        // unusable. The map renders `w:br` as "\n" and `w:noBreakHyphen` as "-";
        // Word's Range.Text renders them as \u000B and \u001E. The host's
        // normalizer used to *delete* those control characters, so the two texts
        // never matched and the pre-mutation check rejected every edit.
        //
        // What made it worse than a refusal: the rejection is reported as "the
        // file changed between the read and the edit. Read it again and use the
        // new address." That is false, and the advice is a loop -- re-reading
        // mints the identical address, so an agent following it never stops.
        // Soft breaks are common in headings and postal addresses.
        const marks = path.join(docs, "layout-marks.docx");
        await execFileAsync("powershell.exe", [
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            path.join(HERE, "inject-layout-marks.ps1"),
            "-Source",
            fixture,
            "-Out",
            marks,
        ]);

        const map = await cache.readStructure(marks);
        let token = map.revisionToken;

        for (const [marker, expectedMapText] of [
            ["MARKbr", "MARKbr abc\ndef"],
            ["MARKnbh", "MARKnbh e-mail"],
            ["MARKtab", "MARKtab a\tb"],
            ["MARKsoft", "MARKsoft softhyphen"],
        ]) {
            const target = map.paragraphs.find((p) => p.text.startsWith(marker));
            assert.ok(target, `no ${marker} paragraph in the fixture`);
            assert.equal(target.text, expectedMapText, `the map renders ${marker} differently than expected`);

            const result = await cache.editDocument(
                marks,
                { op: "replace_text", address: target.address, text: `${marker} rewritten` },
                { revisionToken: token },
            );
            assert.equal(result.paragraph.text, `${marker} rewritten`, `${marker} was not rewritten`);
            token = result.document.revisionToken;
        }
    });

    // Issue #40. Three checks, deliberately not one: the write direction, the
    // lockout the write direction creates, and a paragraph Word itself authored.
    // They fail for different reasons and a single round trip would prove only
    // the first of them.
    //
    // Every umlaut below is built from codepoints rather than typed. A literal
    // would make the assertion depend on how *this* file's bytes are decoded,
    // which is a member of the class of question under test -- the probe that
    // established this defect broke on exactly that before it was rewritten this
    // way.
    const UML = String.fromCharCode(0x00e4, 0x00f6, 0x00fc, 0x00df, 0x00c4, 0x00d6, 0x00dc);
    // 0xC3, the UTF-8 lead byte of every Latin-1 supplement character, decoded
    // as the OEM codepage. It is the signature of one crossing of the corrupting
    // boundary. The codepage is deliberately *not* named here: this ran under
    // 850 (measured with `chcp`), while an earlier version of this comment
    // asserted 437. The mark is identical either way, because both map 0xC3 to
    // U+251C -- which is exactly why naming one was a claim nothing here had
    // established, and why the assertion below does not depend on it.
    // Note the inverse test would be wrong: mojibake *contains* U+00E4, so
    // "no U+00E4 present" does not detect corruption.
    const MOJIBAKE_MARK = String.fromCharCode(0x251c);
    const umlautDoc = path.join(docs, "umlauts.docx");
    // Copied here rather than inside the first check, so the third does not
    // silently depend on the first having run. The second depends on the first
    // deliberately and says so; the third must be able to fail on its own.
    await copyFile(fixture, umlautDoc);

    await check("non-ASCII text an agent writes reaches the disk intact", async () => {
        // Asserted against the bytes in the zip, not through our own reader. A
        // reader that corrupted symmetrically would cancel the defect out and
        // report a green test over a wrecked file; the user's file is the thing
        // that has to be right, so that is what is read.
        const map = await cache.readStructure(umlautDoc);
        // An ASCII target on purpose. Editing a paragraph that already contains
        // non-ASCII fails for the *next* check's reason, which would make this
        // one red for a property it is not testing.
        const target = map.paragraphs.find((p) => p.text.includes("ZORBLAX"));
        assert.ok(target, "no ZORBLAX paragraph in the fixture");
        assert.ok(!/[^\u0000-\u007f]/.test(target.text), "this check needs an ASCII target");

        const written = `UMLWRITE ${UML}`;
        await cache.editDocument(umlautDoc, { op: "replace_text", address: target.address, text: written }, { revisionToken: map.revisionToken });

        const onDisk = documentPlainText(await documentXml(umlautDoc));
        // `documentPlainText` concatenates runs with no separator, on purpose --
        // Word splits a run wherever it likes -- so the diagnostic slices a
        // fixed window rather than looking for a line that does not exist.
        const at = onDisk.indexOf("UMLWRITE");
        const found = at === -1 ? "<no UMLWRITE paragraph at all>" : onDisk.slice(at, at + written.length + 8);
        assert.ok(onDisk.includes(written), `the text did not survive the trip to disk.\n  sent:  ${codepoints(written)}\n  found: ${codepoints(found)}`);
        assert.ok(!onDisk.includes(MOJIBAKE_MARK), "the document contains U+251C, the signature of a UTF-8 sequence decoded as OEM");
    });

    await check("a paragraph is still editable after non-ASCII has been written to it", async () => {
        // The half a round trip cannot reach. Corruption here is not symmetric
        // and does not cancel: reads are faithful, so an agent gets the on-disk
        // mojibake back exactly, and sending it in as `expectedText` re-decodes
        // the previous result's UTF-8 bytes a second time and expands it again.
        // The pre-mutation check can then never match, for any number of
        // retries, and the paragraph is locked for good.
        //
        // This runs after the check above deliberately: the text it edits is the
        // text that came back from a read of what we ourselves wrote, which is
        // the exact sequence a user performs.
        const map = await cache.readStructure(umlautDoc);
        const target = map.paragraphs.find((p) => p.text.startsWith("UMLWRITE"));
        assert.ok(target, "the previous check left no UMLWRITE paragraph to edit");

        const result = await cache.editDocument(
            umlautDoc,
            { op: "replace_text", address: target.address, text: "UMLWRITE rewritten" },
            { revisionToken: map.revisionToken },
        );
        assert.equal(result.paragraph.text, "UMLWRITE rewritten", "the second edit did not take");
    });

    await check("a paragraph Word itself wrote with umlauts is editable", async () => {
        // The other origin. Above, the non-ASCII arrived through our own write
        // path; here it was authored by Word, so the bytes on disk are correct
        // UTF-8 and only the inbound crossing is at fault. A user opening their
        // own German document meets this one without our write path ever having
        // run.
        const map = await cache.readStructure(umlautDoc);
        const target = map.paragraphs.find((p) => p.text.includes("UMLAUTMARKER"));
        assert.ok(target, "no UMLAUTMARKER paragraph in the fixture");
        assert.match(target.text, /[^\u0000-\u007f]/, "the UMLAUTMARKER paragraph is not actually non-ASCII -- the fixture is fake again");
        assert.ok(!target.text.includes(MOJIBAKE_MARK), "the fixture itself is already mojibake; make-fixture.ps1 is at fault, not the edit path");

        const result = await cache.editDocument(
            umlautDoc,
            { op: "replace_text", address: target.address, text: "UMLAUTMARKER rewritten" },
            { revisionToken: map.revisionToken },
        );
        assert.equal(result.paragraph.text, "UMLAUTMARKER rewritten", "a Word-authored non-ASCII paragraph could not be edited");
    });

    await check("deleting the only paragraph in a table cell is refused, not silently reported as done", async () => {
        // A cell must retain at least one paragraph, so Word declines rather
        // than throwing. Nothing downstream noticed, because the delete branch
        // reported a hard-coded paragraph index and no count was compared.
        const map = await cache.readStructure(fixture);
        const cell = map.paragraphs.find((p) => p.inTable && p.text.trim().length > 0);
        assert.ok(cell, "no table paragraph in the fixture");

        const countBefore = map.paragraphCount;
        await assert.rejects(
            () => cache.editDocument(fixture, { op: "delete_paragraph", address: cell.address }, { revisionToken: map.revisionToken }),
            (err) => {
                assert.equal(err.code, "edit_failed", `unexpected code ${err.code}`);
                assert.match(err.message, /had no effect|cannot be deleted|at least one/i);
                return true;
            },
        );

        const after = await cache.readStructure(fixture);
        assert.equal(after.paragraphCount, countBefore, "the refused delete changed the document anyway");
    });

    await check("revert refuses to run without a revision token", async () => {
        // A revert overwrites the document with older bytes and keeps no copy of
        // what it replaced. Left optional, the sequence "agent edits, user works
        // in Word, agent reverts" destroys the user's work with nothing to
        // recover it from -- while `edit_document` refuses that exact situation.
        await assert.rejects(
            () => cache.revertDocument(fixture, {}),
            (err) => err.code === "stale_revision_token",
        );
        await assert.rejects(
            () => cache.revertDocument(fixture, { revisionToken: "sha256:0000000000000000" }),
            (err) => err.code === "stale_revision_token",
        );
    });

    await check("editing needs no canvas open", async () => {
        assert.equal(cache.openCount, 0, "an edit left a document open");
    });

    await check("snapshot files pair up and stay within the retention cap", async () => {
        // The previous form of this asserted `<= 40` after roughly five edits,
        // which no possible behaviour could have violated. Retention itself is
        // proved on disk by the unit test; what is worth checking here is that
        // real edits leave a coherent set -- one manifest per snapshot, nothing
        // orphaned by the pruning.
        const dir = path.join(workRoot, "artifacts", "snapshots");
        const perDoc = await readdir(dir);
        assert.ok(perDoc.length >= 1, "no snapshots were taken");
        for (const sub of perDoc) {
            const files = await readdir(path.join(dir, sub));
            const snaps = files.filter((f) => f.endsWith(".snapshot"));
            const manifests = files.filter((f) => f.endsWith(".snapshot.json"));
            assert.equal(
                snaps.length,
                manifests.length,
                `${sub}: ${snaps.length} snapshots but ${manifests.length} manifests`,
            );
            assert.ok(snaps.length <= 20, `${snaps.length} snapshots retained for one document, cap is 20`);
            // Names must sort in the order they were taken, so revert pops the
            // newest. Same-millisecond ties are broken by a monotonic counter.
            const sorted = [...snaps].sort();
            assert.deepEqual(snaps.slice().sort(), sorted);
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
