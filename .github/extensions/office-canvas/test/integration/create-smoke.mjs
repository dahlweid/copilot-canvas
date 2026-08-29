// Smoke test for create_document against real Word.
//
// Run: node .github/extensions/office-canvas/test/integration/create-smoke.mjs
//
// Requires Word. Everything here is a claim no Office-free test can reach, and
// each one cost a probe to establish:
//
//   * A spec goes in and a real .docx comes out, whose structure map is the one
//     read_document returns -- so the addresses create hands back are usable by
//     edit_document immediately, with no intervening read.
//   * Headings and list styles survive the save/reopen round trip. They are
//     applied by numeric constant, never by name: Word's UI here is German, and
//     `Range.Style = 'Heading 1'`, `= 'Überschrift 1'` and the OOXML id
//     `= 'berschrift1'` all throw InvalidCastException. Measured in
//     spikes/isolation/probes/probe-authoring-save.ps1.
//   * Text goes in verbatim. Autocorrect rewrites inserted text with no error
//     raised, so an unsuppressed instance produces a document that says
//     something other than what was asked for. Only three baits are asserted --
//     see the comment on VERBATIM_BAITS, which is the part of this file most
//     likely to be "improved" into a vacuous pass.
//   * Word does not overwrite. The refusal is checked against a file that
//     exists, because the interesting failure is a silent overwrite.
//   * No WINWORD.EXE is left behind.

import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { mkdtemp, readFile, rm, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { promisify } from "node:util";

import { RenderCache } from "../../src/render-cache.mjs";
import { WordHost } from "../../src/word/word-host.mjs";
import { assertNoLeakedWord, wordPids } from "./word-pids.mjs";

const execFileAsync = promisify(execFile);
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

/**
 * The only three autocorrect substitutions with a demonstrated positive control
 * on this machine.
 *
 * This is a deliberate, measured restriction and not laziness. Arm H of
 * spikes/isolation/probes/probe-autocorrect.ps1 forces autocorrect to run
 * (`Content.AutoFormat()`) over six bait strings and records which are rewritten.
 * Measured, on this Word, with the exact codepoints it produced:
 *
 *   She said "hello" and left.  ->  She said \u201Ehello\u201C and left.
 *   width -- height             ->  width\u2014height
 *   (c) 2026                    ->  \u00A9 2026
 *
 * The other three baits -- a lowercase sentence start, a doubled initial capital,
 * and the classic `teh` typo -- come back **unchanged even with autocorrect fully
 * on**, because this Word's list is German: 402 entries, and `teh` is confirmed
 * `<no entry>`.
 *
 * So an assertion that `teh` survives, or that a lowercase `i` stays lowercase,
 * passes on a machine where autocorrect is switched on and doing its worst. It
 * would be a test whose coverage depends on a property of the machine rather
 * than on the code -- the exact shape this repo has twice mistaken for a pass.
 * Adding a bait here requires arm H to rewrite it first.
 *
 * Two corrections that came out of re-running arm H, both of which had been
 * carried from what an *English* Word does rather than measured on this one:
 *
 *   - The opening quote here is \u201E (low-9), not \u201C. This list previously
 *     named \u201C and \u201D. \u201D is not produced by this Word at all, and
 *     \u201E -- the one character a German Word actually opens with -- was
 *     absent. The assertion would still have failed on the closing \u201C, so
 *     this was a narrowed net rather than a hole, but it was a net woven to the
 *     wrong locale while the file above claimed every bait was verified live.
 *   - `--` becomes an em dash \u2014, not an en dash, and it eats the spaces on
 *     either side. The doc above said en dash.
 *
 * \u2013 and \u201D are retained below and are explicitly NOT measured here:
 * they are what other locales produce, kept so the assertion does not silently
 * narrow to this machine. Every other codepoint in this list was observed.
 *
 * WHAT A GREEN HERE DOES AND DOES NOT MEAN. This is a forward guard, not a live
 * regression detector, and the difference matters because the test's name reads
 * like the latter. Arms A, B, E, F and G of probe-autocorrect.ps1 insert this
 * bait through Range.Text and through Selection.TypeText, whole and character by
 * character, with every autocorrect setting ON, and rewrite 0 of 6. The only arm
 * that rewrites anything is an explicit Content.AutoFormat(), which this host
 * never calls. So no insertion path the product uses can trigger a substitution
 * today, and a green here is not evidence that suppression works -- the evidence
 * for that is the separate settings read-back check. What this guards is a
 * *future* insertion path that does trigger autocorrect.
 *
 * THE ANCHOR IS LOAD-BEARING, AND ITS ABSENCE MADE THE GUARD UNABLE TO GUARD.
 * Each bait paragraph opens with an anchor and is located by it. That is not
 * cosmetic. This test previously located the paragraph by its own first twelve
 * characters, and for all three baits that prefix spanned the substitution site:
 *
 *   'He said "hel'   contains the quote that becomes \u201E
 *   'A dash -- li'   contains the -- that becomes \u2014
 *   'Copyright (c'   contains the (c) that becomes \u00A9
 *
 * So a *successful* substitution destroyed the locator. The paragraph was not
 * found, and the only assertion that could fire was "the bait paragraph is
 * missing" -- which is equally consistent with an #40 stdin corruption, a paging
 * bug, or a document that was never authored. The two assertions that name the
 * defect were both unreachable for the case they exist to catch, on any Word in
 * any locale, independently of the arm results above. Measured by replaying the
 * find-then-assert against a substituted string; both orderings reported
 * "missing" and neither reported a substitution.
 *
 * An anchor must therefore be invariant under every substitution being detected.
 * ANCHOR_SHAPE enforces that mechanically rather than by inspection: letters and
 * spaces only, so it cannot contain a quote, a dash, or a parenthesised (c),
 * (r), (e) or (tm). Adding a bait whose anchor does not match is a hard failure.
 *
 * This is a convention rather than an invention, and the precedent is stricter
 * than what is done here. edit-smoke.mjs:601 locates by MARKbr, MARKnbh, MARKtab
 * and MARKsoft -- markers deliberately free of the character each case is about,
 * so "MARKnbh e-mail" carries the hyphen under test in the payload and none in
 * the locator. read-smoke.mjs does the same with "DUPLICATE LINE:". This file
 * was the outlier in locating by the payload itself.
 *
 * The other locators here -- "Quarterly Report", "First point", "End of report."
 * -- carry no trigger sequence, so the rule above does not convict them, but
 * that is luck rather than construction: nothing stopped a heading being given
 * bait-like text. A substitution in one of those would be caught and named by
 * the bait check below, which is the one place now built for it.
 *
 * mustNotContain is checked BEFORE the equality assertion, so a substitution is
 * reported as "smart quotes appeared in: ..." rather than as the generic "Word
 * rewrote the text". Equality still runs afterwards as the catch-all for a
 * rewrite these lists do not name. Reversing that order does not break the test,
 * it just downgrades every message -- which is how the specific assertions came
 * to be dead without anything going red.
 *
 * Do not try to mutation-check these by putting a curly quote in the bait text.
 * That was tried: it turns the suite red, but with "the bait paragraph is
 * missing", because the character is corrupted crossing stdin (issue #40) and
 * the paragraph is never found. The mutant scores as killed and the kill is
 * attributed to an assertion that never ran. Mutate the mustNotContain list
 * instead -- forbid a character the bait certainly has.
 */
const ANCHOR_SHAPE = /^[A-Za-z][A-Za-z ]*\.$/;

const VERBATIM_BAITS = [
    {
        anchor: "Bait one.",
        text: `He said "hello" to her.`,
        mustNotContain: ["\u201e", "\u201c", "\u201d"],
        what: "smart quotes",
    },
    {
        anchor: "Bait two.",
        text: "A dash -- like this.",
        mustNotContain: ["\u2014", "\u2013"],
        what: "an em or en dash",
    },
    {
        anchor: "Bait three.",
        text: "Copyright (c) 2024.",
        mustNotContain: ["\u00a9"],
        what: "a copyright sign",
    },
];

const baitParagraph = (b) => `${b.anchor} ${b.text}`;

const spec = {
    blocks: [
        { kind: "heading", level: 1, text: "Quarterly Report" },
        { kind: "paragraph", text: "This document was authored by create_document." },
        { kind: "heading", level: 2, text: "Findings" },
        ...VERBATIM_BAITS.map((b) => ({ kind: "paragraph", text: baitParagraph(b) })),
        { kind: "list", ordered: false, items: ["First point", "Second point"] },
        { kind: "list", ordered: true, items: ["Step one", "Step two", "Step three"] },
        {
            kind: "table",
            rows: [
                ["Region", "Revenue"],
                ["North", "120"],
                ["South", "95"],
            ],
        },
        { kind: "paragraph", text: "End of report." },
    ],
};

const workRoot = await mkdtemp(path.join(tmpdir(), "word-create-test-"));
const docs = path.join(workRoot, "docs");
const pidsBefore = await wordPids();

const cache = new RenderCache({
    cacheRoot: path.join(workRoot, "artifacts"),
    log: (m) => process.stderr.write(`[cache] ${m}\n`),
});

// Paths reach PowerShell as discrete argv elements or through the environment,
// never interpolated into a command string. See the same note in edit-smoke.mjs:
// the dangerous character sets are per-parser and barely overlap, so escaping
// per site is the mistake and removing the parser is the fix.
const readStyleIds = async (docPath) => {
    const { stdout } = await execFileAsync("powershell.exe", [
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-Command",
        "Add-Type -AssemblyName WindowsBase;" +
            "$p=[IO.Packaging.Package]::Open($env:PROBE_DOC,'Open','Read');" +
            "$u=[Uri]::new('/word/document.xml',[UriKind]::Relative);" +
            "$x=[xml](New-Object IO.StreamReader($p.GetPart($u).GetStream())).ReadToEnd();" +
            "$p.Close();" +
            "($x.document.body.p | ForEach-Object { $_.pPr.pStyle.val }) -join '|'",
    ], { env: { ...process.env, PROBE_DOC: docPath } });
    return stdout.trim();
};

try {
    await rm(docs, { recursive: true, force: true });
    await execFileAsync("powershell.exe", [
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-Command",
        "New-Item -ItemType Directory -Force -Path $env:PROBE_DIR | Out-Null",
    ], { env: { ...process.env, PROBE_DIR: docs } });

    const target = path.join(docs, "report.docx");
    let created;

    await check("a spec becomes a document on disk", async () => {
        created = await cache.createDocument(target, spec);
        const info = await stat(target);
        assert.ok(info.size > 0, "the file is empty");
        assert.equal(created.created.path, target);
        assert.ok(created.revisionToken, "no revision token was returned");
    });

    await check("autocorrect was suppressed on the instance that authored it", () => {
        // Asserted rather than assumed, and `suppressed` is a READ-BACK: the host
        // records the prior values, writes false, then reads all five and
        // compares. It used to mean "five assignments did not throw", which is a
        // write reporting itself as a read and could not have gone red if Word
        // had accepted an assignment without applying it.
        //
        // RETRACTED, and this comment is the reason the retraction matters. It
        // used to say the settings are per-*process* -- switched off on instance
        // A, a second process reads them still on -- and concluded suppression
        // cannot reach the user's Word. The observation reproduces; the
        // conclusion is false. That probe read B **while A was still alive**, and
        // a concurrent reader sees the pre-write value, so it cannot tell
        // isolation from persistence-with-lag. Re-measured
        // sequentially (set, QUIT, read a fresh instance) all five come back
        // changed: they persist for the user.
        //
        // So the host now captures and restores around the authoring call, and
        // the assertion below has two halves. `restored` is the one carrying the
        // user-safety claim -- a suppression proven and a restoration unproven is
        // the same defect one step later.
        //
        // What this check CANNOT see, stated so it is not mistaken for full
        // cover: this machine's found state is already all-false, which is also
        // what suppression writes, so prior == target and a skipped write leaves
        // the right value behind. Both halves were proven falsifiable by
        // mutating to a *wrong* value rather than a no-op -- `not_applied
        // (CorrectInitialCaps)` and `not_restored (ReplaceText)` -- but the case
        // the restore actually exists for is a user whose autocorrect is ON, and
        // that case cannot occur here. It is manufactured in
        // spikes/isolation/probes/probe-autocorrect-restore.mjs, which sets all
        // five ON, runs this tool, and reads a fresh instance afterwards.
        // Both messages name the settings the host reported, not just the reason.
        // A red that says only `not_applied` proves the check can fail; it does
        // not prove it failed for the thing it is watching, and those are
        // different claims. The list is taken from the report rather than
        // restated here, so it cannot drift from what the host actually checks.
        const naming = (label, why, settings) =>
            `${label}: ${why ?? "no reason given"}${settings?.length ? ` (${settings.join(", ")})` : ""}`;

        assert.equal(
            created.autoCorrect.suppressed,
            true,
            naming("not suppressed", created.autoCorrect.reason, created.autoCorrect.settings),
        );
        assert.equal(
            created.autoCorrect.restored,
            true,
            naming("not restored", created.autoCorrect.restoreReason, created.autoCorrect.restoreSettings),
        );

        // The capture is what makes a restore possible at all, so its absence is
        // a distinct failure from a restore that ran and did not take.
        assert.ok(created.autoCorrect.prior, "no prior values were captured");
        assert.equal(
            Object.keys(created.autoCorrect.prior).length,
            5,
            `captured ${Object.keys(created.autoCorrect.prior).length} prior values, expected all 5`,
        );
    });

    await check("text goes in verbatim, with no autocorrect substitution", async () => {
        // This check CANNOT discriminate, and the name reads as though it can.
        // Measured (probe-autocorrect.ps1 arms A, B, E, F, G): with every
        // setting ON, these baits are rewritten 0 of 6 through every insertion
        // path this host uses. So it is green whether or not suppression
        // applied, and it is a forward guard against a future insertion path,
        // never evidence that autocorrect was off. The evidence for that is the
        // settings read-back check above, which is the acceptance criterion's
        // real test. Full reasoning in the file header.
        const map = await cache.readStructure(target, { limit: 0 });
        const texts = map.paragraphs.map((p) => p.text);
        for (const bait of VERBATIM_BAITS) {
            assert.ok(
                ANCHOR_SHAPE.test(bait.anchor),
                `anchor is not substitution-proof, so a rewrite would read as a missing paragraph: ${bait.anchor}`,
            );
            const found = texts.find((t) => t.startsWith(bait.anchor));
            assert.ok(found !== undefined, `the bait paragraph is missing: ${bait.anchor}`);
            for (const ch of bait.mustNotContain) {
                assert.ok(!found.includes(ch), `${bait.what} appeared in: ${found}`);
            }
            assert.equal(found, baitParagraph(bait), `Word rewrote the text`);
        }
    });

    await check("the structure map create returns is the one read_document returns", async () => {
        // If these diverge, an address handed back by create is not an address
        // edit_document can use, and the caller has to re-read for no reason.
        const fresh = await cache.readStructure(target, { limit: 0 });
        assert.deepEqual(
            created.document.paragraphs.map((p) => [p.address, p.text, p.headingLevel]),
            fresh.paragraphs.map((p) => [p.address, p.text, p.headingLevel]),
        );
    });

    await check("headings survive the round trip at the level they were asked for", async () => {
        const map = await cache.readStructure(target, { limit: 0 });
        const h1 = map.paragraphs.find((p) => p.text === "Quarterly Report");
        const h2 = map.paragraphs.find((p) => p.text === "Findings");
        assert.equal(h1?.headingLevel, 1, "the level 1 heading did not survive");
        assert.equal(h2?.headingLevel, 2, "the level 2 heading did not survive");
    });

    await check("style ids are Word's own, never constructed", async () => {
        // Word mints a style id from the *localized* name with non-ASCII
        // dropped, so a German Word writes `berschrift1` -- not `Heading1` and
        // not `Überschrift1`. This asserts what is in the file rather than what
        // we think it should be: the point is that nothing in our code ever
        // builds or matches one of these.
        const ids = await readStyleIds(target);
        assert.ok(ids.length > 0, "no styled paragraphs found in document.xml");
        assert.doesNotMatch(ids, /Heading\d/, `an English style id appeared, so something constructed one: ${ids}`);
    });

    await check("list blocks come back styled as lists, not as body text", async () => {
        const map = await cache.readStructure(target, { limit: 0 });
        for (const text of ["First point", "Step one"]) {
            const item = map.paragraphs.find((p) => p.text === text);
            assert.ok(item, `list item is missing: ${text}`);
            assert.notEqual(item.headingLevel, 1, `${text} came back as a heading`);
            assert.ok(item.styleId, `${text} carries no style id`);
        }
        const bullet = map.paragraphs.find((p) => p.text === "First point");
        const numbered = map.paragraphs.find((p) => p.text === "Step one");
        assert.notEqual(
            bullet.styleId,
            numbered.styleId,
            "the bulleted and numbered lists got the same style, so `ordered` did nothing",
        );
    });

    await check("the table survives, and its cell text is trimmed correctly", async () => {
        // Inside a table a cell paragraph's Range.Text ends with \r plus chr(7)
        // -- two characters in the string, but `End - Start` counts the
        // end-of-cell mark as one position. A trim derived from string length
        // eats the last character of every cell. Measured: an empty cell reads
        // length 2, span 1, char codes 13 and 7.
        const map = await cache.readStructure(target, { limit: 0 });
        for (const cell of ["Region", "Revenue", "North", "120", "South", "95"]) {
            assert.ok(
                map.paragraphs.some((p) => p.text === cell),
                `cell text is missing or mistrimmed: expected exactly "${cell}"`,
            );
        }
        assert.equal(created.tableCount, 1, "the table did not reach the document");
    });

    await check("every block in the spec reached the document", async () => {
        // What the caller actually asked for: nothing silently dropped. This
        // replaces a check against a predicted paragraph count, which was wrong
        // by construction — the prediction counted paragraphs in Word's COM
        // coordinate system (a row-end mark per table row) while the map is
        // OOXML-derived, where row-end marks are not paragraphs. Measured here:
        // the prediction said 18, the map says 14, and the document was correct.
        //
        // It is also the assertion that would have caught the real bug this
        // suite found: `Paragraphs.Add()` returns the *previously* last
        // paragraph, so an early version overwrote each block with the next and
        // lost "Quarterly Report" entirely, raising nothing.
        const map = await cache.readStructure(target, { limit: 0 });
        const texts = new Set(map.paragraphs.map((p) => p.text));
        for (const block of spec.blocks) {
            const wanted =
                block.kind === "list"
                    ? block.items
                    : block.kind === "table"
                      ? block.rows.flat()
                      : [block.text];
            for (const text of wanted) {
                assert.ok(texts.has(text), `block text is missing from the document: "${text}"`);
            }
        }
        // And nothing the spec did not ask for. A stray empty paragraph is what
        // an unconsumed after-table paragraph looks like from out here.
        assert.equal(
            map.paragraphs.filter((p) => p.text === "").length,
            0,
            "the document carries a paragraph the spec never asked for",
        );
    });

    await check("an address minted by create is immediately usable by edit_document", async () => {
        // The whole point of returning a structure map: no intervening read.
        const target2 = created.document.paragraphs.find((p) => p.text === "End of report.");
        assert.ok(target2, "could not find the paragraph to edit");
        const after = await cache.editDocument(
            target,
            { op: "replace_text", address: target2.address, text: "Edited in place." },
            { revisionToken: created.revisionToken },
        );
        assert.ok(
            after.document.paragraphs.some((p) => p.text === "Edited in place."),
            "the edit did not land",
        );
    });

    await check("an existing file is refused, not overwritten", async () => {
        const before = await readFile(target);
        await assert.rejects(
            () => cache.createDocument(target, spec),
            (err) => {
                assert.equal(err.code, "file_exists");
                // The message must point at the tool that *can* do this safely,
                // and must not name a cause nobody measured.
                assert.match(err.message, /edit_document/);
                return true;
            },
        );
        assert.deepEqual(await readFile(target), before, "the refused create still changed the file");
    });

    await check("the host refuses an existing file on its own, not only the JS preflight", async () => {
        // The check above cannot see the host at all. `createDocument` refuses in
        // JavaScript before the host is asked anything, so it passes whether the
        // host would overwrite or not -- and the mutant for it removes that same
        // JS preflight, so the mutation check inherited the blind spot. The
        // guarantee in the tool description is "will not overwrite", and until
        // this arm nothing tested the layer that actually holds the file handle.
        //
        // Calling `host.create` directly is the point, not a shortcut: it is the
        // only way to reach `Cmd-Create` with the preflight out of the way.
        //
        // What this pins is the refusal, not *which* of the host's two checks
        // produced it. With the file present before the call, the early check at
        // the head of `Cmd-Create` fires and the one before `SaveAs2` is
        // unreachable. The late check is verified by a positive control instead:
        // with the early check deleted, this arm still reports `file_exists`,
        // and reports `created` if the late check is also deleted. That is a
        // mutation result, recorded here because it is not something this
        // assertion can show on its own.
        const host = new WordHost({ log: () => {} });
        try {
            const before = await readFile(target);
            const out = await host.create({ path: target, blocks: [{ kind: "paragraph", text: "Overwrite." }] });

            assert.equal(out.status, "file_exists", `the host reported ${out.status} for a path that already exists`);
            assert.deepEqual(await readFile(target), before, "the host overwrote a document that already existed");
        } finally {
            await host.dispose().catch(() => {});
        }
    });

    await check("a path in a folder that does not exist is refused before Word is started", async () => {
        await assert.rejects(
            () => cache.createDocument(path.join(docs, "nope", "x.docx"), spec),
            (err) => {
                assert.equal(err.code, "directory_not_found");
                return true;
            },
        );
    });

    await check("a type this tool does not author is refused", async () => {
        await assert.rejects(
            () => cache.createDocument(path.join(docs, "other.rtf"), spec),
            (err) => {
                assert.equal(err.code, "unsupported_type");
                return true;
            },
        );
    });

    await check("a COM failure reports the type Word threw, not a PowerShell type", async () => {
        // `create_failed` is deliberately unreachable through the public API --
        // validating the spec is what makes every COM failure impossible -- so
        // this drives the host directly with a table the validator refuses, and
        // Word rejects at `Tables.Add`.
        //
        // What it pins is that `exception` carries the type of the thing that
        // actually failed. It cannot fail for a missing unwrap of PowerShell's
        // MethodInvocationException, and does not claim to: the unwrap is a
        // guard, and a mutation check confirmed that removing it changes nothing
        // here, because COM calls arrive unwrapped. It *can* fail -- verified by
        // mutation -- if the reported type becomes a PowerShell artefact such as
        // the ErrorRecord, which is the shape of the regression that would put a
        // caller back to reading the message. And the message is why the field
        // exists at all: this failure's detail arrives in German.
        const host = new WordHost({ log: () => {} });
        try {
            const out = await host.create({
                path: path.join(docs, "never-written.docx"),
                blocks: [{ kind: "table", headerRow: false, rows: [] }],
            });
            assert.equal(out.status, "create_failed", `expected a COM failure, got ${out.status}`);
            assert.equal(
                out.exception,
                "System.Runtime.InteropServices.COMException",
                `reported ${out.exception}, which is not the type the COM call threw`,
            );
            assert.equal(out.leftBehind, false, "a failed create left a file behind");
        } finally {
            await host.dispose().catch(() => {});
        }
    });

    await check("quitting Word actually calls Quit, rather than relying on the process dying", async () => {
        // The defect this pins was invisible to every black-box signal.
        // `Application.Quit` declares its parameters as `VARIANT*`, so a value
        // argument fails to bind -- "Argument: 1 muss
        // System.Management.Automation.PSReference sein" -- and it threw on every
        // shutdown of every run. Word exited anyway, because killing the host
        // process releases the last COM reference, so the quit RPC succeeded, the
        // leak assertion passed and the census was clean while the Quit never ran.
        //
        // Measured by instrumenting all 11 swallowing catches around
        // Quit/Close/Release and re-running this suite: `Quit` was the only site
        // throwing, at every occurrence, and with `[ref]` the log is empty.
        //
        // Two things it deliberately does not claim. It is not a
        // value-versus-variable problem -- that site passed a variable. And it is
        // not a general VARIANT* rule: in the same process on the same run,
        // `$doc.Close($WD_DO_NOT_SAVE_CHANGES)` never threw at any of its four
        // sites, so neither call shape may be changed on the other's evidence.
        //
        // The assertion is only possible because the catch reports instead of
        // swallowing. That was the actual fix; the `[ref]` was the easy part.
        //
        // The `owned` guard below is not decoration. `Initialize-Word` attaches
        // to a Word the user already had running when one is available, and
        // `Stop-Word` never quits an instance it did not start -- so on an
        // attached instance `Quit` is skipped and `quitError` is null for a
        // reason that has nothing to do with the binding. Without the guard this
        // check would pass vacuously, on a machine where 8-18 WINWORD.EXE are
        // routinely alive from other sessions. It fails loudly instead, because
        // "could not measure" and "the defect is gone" must not look alike.
        const host = new WordHost({ log: () => {} });
        try {
            const ready = await host.request("ping", {});
            assert.equal(ready.owned, true, "attached to an existing Word, so the Quit below is skipped and this check cannot measure anything");
            const out = await host.request("quit", {});
            assert.equal(out.stopped, true, "the host did not report a stop");
            assert.equal(out.quitError, null, `Application.Quit threw and was swallowed: ${out.quitError}`);
        } finally {
            await host.dispose().catch(() => {});
        }
    });

    await check("creating needs no canvas open", () => {
        assert.equal(cache.openCount, 0, "a create left a document open");
    });

    await check("a file created next to one Word holds open still works", async () => {
        // The case the lock model is about. Word holds `report.docx` only
        // transiently here, so this is the weaker half; what it does prove is
        // that authoring into a directory containing a Word-held document is not
        // itself refused.
        const sibling = path.join(docs, "sibling.docx");
        const out = await cache.createDocument(sibling, { blocks: [{ kind: "paragraph", text: "Sibling." }] });
        assert.ok(out.revisionToken);
    });
} finally {
    await cache.dispose().catch(() => {});
}

await check("no new WINWORD.EXE is left behind", () => assertNoLeakedWord(pidsBefore));

await rm(workRoot, { recursive: true, force: true }).catch(() => {});

const failed = results.filter((r) => !r.ok);
process.stderr.write(`\n${results.length - failed.length}/${results.length} passed\n`);
process.exit(failed.length === 0 ? 0 : 1);
