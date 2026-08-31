// Autocorrect must not be suppressed. Nothing in the host may write those five
// settings, and nothing above it may report on them.
//
// THE FACT THIS GUARDS
//
// `Application.AutoCorrect.ReplaceText`, `.CorrectSentenceCaps`,
// `.CorrectInitialCaps`, `Application.Options.AutoFormatAsYouTypeReplaceQuotes`
// and `.AutoFormatAsYouTypeReplaceSymbols` **persist for the user**. They were
// once documented as per-process, and that was wrong: the probe that established
// it read a second instance while the writing instance was still alive, and a
// concurrent reader sees the pre-write value, so isolation and persistence are
// the same observation from there. Quitting the writer first separates them.
// Re-measured in `spikes/isolation/probes/probe-autocorrect-concurrency.mjs`
// arm 1: seed all five ON, a writer writes OFF and quits, a FRESH instance
// started afterwards reads all five OFF.
//
// So writing them is writing the user's own Word. The host used to do it around
// authoring and put them back in a `finally`, which worked but is a promise
// about shared state that cannot be fully kept: arm 4 of the same probe measured
// that a user's Word alive at the same time and exiting AFTER us overwrites our
// restore with its own stale copy.
//
// What removed the hazard rather than managing it is that the suppression buys
// nothing. `spikes/isolation/probes/probe-autocorrect-necessity.ps1` inserted 8
// baits -- 4 of them proven live in this machine's own 402-entry GERMAN
// replacement list -- through the host's exact insertion path with all five
// settings ON, and all 8 came back verbatim, while `Content.AutoFormat()` in the
// same run rewrote 3 of them, so the instrument can see a rewrite. Autocorrect
// and autoformat-as-you-type are TYPING features and this host never types:
// every character goes in through `Set-ParagraphText` -> `Range.Text`.
//
// And the argument that settles it is about this codebase: `Cmd-Edit` has never
// suppressed anything, writes through the identical `Range.Text` assignment, and
// `test/integration/edit-smoke.mjs` already asserts its text lands verbatim.
//
// WHY A SOURCE ASSERTION
//
// The behavioural half is covered where it can be: `create-smoke.mjs` asserts a
// live bait goes in verbatim, and the two probes above need a real Word so
// neither runs in CI. The failure this file exists to catch is a later edit --
// someone reading the missing suppression as an omission and helpfully putting
// it back. That edit is visible in the source and needs no Office to see.
//
// FALSIFIABILITY
//
// A "this string is absent" test is the exact shape that passes vacuously when
// its detector is broken, and this repo has shipped that bug twice. So every
// detector below is *also* run against a synthetic source that does contain the
// thing, in the same test, and must report it. A regex that stops matching
// reddens the test rather than silently blessing the file.
//
// MUTANTS (each run red, with the test that caught it)
//
//   1. Gut `word-host.ps1` to a single comment line -> "the sources these
//      assertions read are non-empty and still the right files" fails. Without
//      that test the other five would all have passed on an empty corpus.
//   2. Add `$script:App.AutoCorrect.ReplaceText = $false` to `Set-ParagraphText`
//      -> "the host assigns none of the five autocorrect settings" AND "the
//      create_document description is backed by the host" both fail: the
//      description promises those settings are never touched, so reintroducing
//      the suppression makes the shipped description a lie, and the test says so.
//   3. Reintroduce an unreferenced `Restore-AutoCorrect` helper -> "the suppress
//      and restore helpers are gone" fails with `word-host.ps1 still carries
//      Restore-AutoCorrect`.
//   4. Reintroduce `autoCorrect: result.autoCorrect ?? null` in
//      `document-author.mjs` -> "no autoCorrect field crosses the tool boundary"
//      fails with `document-author.mjs still surfaces an autoCorrect field`.
//   5. Delete the "never read and never changed" sentence from create_document's
//      description -> "the create_document description is backed by the host"
//      fails. The tie runs both ways on purpose: the claim cannot be quietly
//      dropped either.
//   6. Add `$script:App.Selection.TypeText($text)` to `Set-ParagraphText` -> "the
//      host never types" fails. This is the mutant that matters most, because
//      typing is the one edit that would make CorrectSentenceCaps and
//      CorrectInitialCaps reachable again.
//
// Office-free.
//
// Run: node --test ".github/extensions/office-canvas/test/unit/*.test.mjs"

import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";

import { blankComments } from "./ps-encoding-rule.mjs";
import { REPO } from "./tracked-files.mjs";

const HOST = path.join(REPO, ".github/extensions/office-canvas/src/word/word-host.ps1");
const AUTHOR = path.join(REPO, ".github/extensions/office-canvas/src/word/document-author.mjs");
const EXTENSION = path.join(REPO, ".github/extensions/office-canvas/extension.mjs");

/**
 * The five settings, as the property names they are reached by.
 *
 * Named individually rather than matched by a `AutoCorrect\.\w+` wildcard so a
 * failure says which one came back, and so reading a setting -- which is
 * harmless -- is distinguishable from writing one, which is not.
 */
const SUPPRESSED = [
    "AutoCorrect.ReplaceText",
    "AutoCorrect.CorrectSentenceCaps",
    "AutoCorrect.CorrectInitialCaps",
    "Options.AutoFormatAsYouTypeReplaceQuotes",
    "Options.AutoFormatAsYouTypeReplaceSymbols",
];

/** Line number of `index`, 1-based, for a failure message a reader can act on. */
const lineOf = (source, index) => source.slice(0, index).split("\n").length;

/**
 * Every assignment to `<anything>.<property>` in `source`, as line numbers.
 *
 * Assignment only: `= ` with no second `=`, so a comparison (`-eq`) and a plain
 * read both pass. The property may be reached off any expression, because the
 * container is written `$script:App.AutoCorrect` in one place and could be held
 * in a local in another, and a rule that only knew the first spelling would be
 * the vacuous-detector bug this file is built to avoid.
 */
function assignments(source, property) {
    const escaped = property.replace(/\./g, "\\.");
    const re = new RegExp(`\\.${escaped}\\s*=(?!=)`, "g");
    const hits = [];
    for (const m of source.matchAll(re)) hits.push(lineOf(source, m.index));
    return hits;
}

const hostSource = blankComments(await readFile(HOST, "utf8"));
const authorSource = await readFile(AUTHOR, "utf8");
const extensionSource = await readFile(EXTENSION, "utf8");

/**
 * Anchors that must appear in each corpus, so absence can be trusted.
 *
 * Every other test here asserts a string is *missing*, which is the shape that
 * passes forever once the corpus becomes empty -- a moved file, a rename, or a
 * `blankComments` that blanks too much. `readFile` throws on a missing path, but
 * it does not throw on a file that has been emptied, truncated or gutted by the
 * comment stripper. So each corpus is proven to still contain the code the other
 * tests are looking through.
 */
const CORPUS = [
    // Chosen for what they mean, not for being long: this is the sole function
    // through which document text is written, and the sole write inside it.
    ["word-host.ps1", hostSource, ["function Set-ParagraphText", "(Get-TextRange $para).Text ="]],
    ["document-author.mjs", authorSource, ["export", "created"]],
    ["extension.mjs", extensionSource, ['name: "create_document"']],
];

test("the sources these assertions read are non-empty and still the right files", () => {
    for (const [name, source, anchors] of CORPUS) {
        assert.ok(
            source.length > 1000,
            `${name} is ${source.length} bytes -- too small to be the real file, so every ` +
                `"string is absent" assertion below would pass on nothing`,
        );
        for (const anchor of anchors) {
            assert.ok(
                source.includes(anchor),
                `${name} no longer contains ${JSON.stringify(anchor)}, so this file is searching ` +
                    `the wrong text and its silence means nothing`,
            );
        }
    }
});

test("the host assigns none of the five autocorrect settings", () => {
    for (const property of SUPPRESSED) {
        // The detector is proven able to fire before it is trusted to stay
        // silent. Without this, a typo in the property list turns the assertion
        // below into a tautology and the file into decoration.
        const bait = `        $script:App.${property} = $false\n`;
        assert.deepEqual(
            assignments(bait, property),
            [1],
            `the detector for ${property} cannot see an assignment, so its silence on the host means nothing`,
        );

        const found = assignments(hostSource, property);
        assert.deepEqual(
            found,
            [],
            `word-host.ps1:${found.join(", ")} assigns ${property}. These settings persist for the user; ` +
                `they were measured not to affect any insertion path this host uses, and the suppression was ` +
                `removed rather than restored. See the block above Initialize-Word.`,
        );
    }
});

test("the suppress and restore helpers are gone, not merely unreferenced", () => {
    // Removing the call sites while leaving the functions behind is the natural
    // half-measure, and it leaves the next reader a loaded gun: a function whose
    // comment says it is needed, sitting one call away from being needed again.
    for (const name of ["Disable-AutoCorrect", "Restore-AutoCorrect", "Suppress-AutoCorrect"]) {
        const re = new RegExp(`(?<![-\\w])${name}(?![-\\w])`, "g");
        assert.ok(re.test(`x ${name} y`), `the detector for ${name} cannot match its own name`);
        re.lastIndex = 0;
        assert.ok(!re.test(hostSource), `word-host.ps1 still carries ${name}`);
    }
});

test("no autoCorrect field crosses the tool boundary", async () => {
    // The host reported the suppression outcome as `autoCorrect` so a caller
    // could assert it. With nothing suppressed there is no outcome, and a field
    // that always says the same thing is worse than no field: it reads like
    // evidence. The replacement evidence is a bait assertion on the text itself
    // in create-smoke.
    const sources = {
        "word-host.ps1": hostSource,
        "document-author.mjs": authorSource,
        "extension.mjs": extensionSource,
    };
    const re = /(?<![-\w])autoCorrect(?![-\w])/;
    assert.ok(re.test("{ autoCorrect: 1 }"), "the detector cannot match the field it looks for");

    for (const [name, source] of Object.entries(sources)) {
        // `document-author.mjs` keeps the word "Autocorrect" in prose explaining
        // why there is nothing to report; the field is the camel-cased
        // identifier, so match that and not the English word.
        assert.ok(!re.test(source), `${name} still surfaces an autoCorrect field`);
    }
});

// The promise, verbatim, as `create_document`'s description makes it. Held here
// so the test fails when the description drifts rather than when it is deleted.
const PROMISE = "Your own Word settings are never read and never changed";

/**
 * The file with each `[...].join(" ")` description seam closed up.
 *
 * Descriptions are written as arrays of wrapped lines, so a sentence the agent
 * reads as one string is several string literals in the source and a plain
 * `includes` on the file misses it. This reproduces what `.join(" ")` does at
 * runtime, which is the text that actually ships.
 */
const unwrap = (source) => source.replace(/",\s*\n\s*"/g, " ");

test("the create_document description is backed by the host, not just asserted", () => {
    // Prove the unwrapper before trusting it: a seam it fails to close makes
    // every search below miss, which would read as "the promise is absent".
    assert.equal(
        unwrap('    "a sentence split",\n        "across two lines.",\n'),
        '    "a sentence split across two lines.",\n',
        "the description unwrapper does not reproduce .join(\" \"), so it cannot find any sentence",
    );

    // A description is the one string every agent reads before deciding what to
    // do, and it outlives the code it describes unless something ties them. This
    // is the tie: the sentence may only stand while the host earns it.
    const shipped = unwrap(extensionSource);
    const tool = shipped.indexOf('name: "create_document"');
    assert.notEqual(tool, -1, "create_document is no longer registered under that name");
    const promise = shipped.indexOf(PROMISE, tool);
    assert.notEqual(
        promise,
        -1,
        `create_document's description no longer carries "${PROMISE}". If that was deliberate, the ` +
            `host must have started touching those settings -- which the rest of this file forbids. ` +
            `Do not delete this assertion to make it pass.`,
    );

    // ...and it must be inside that tool's own description, not somewhere later
    // in the file, or the tie is to nothing.
    const nextTool = shipped.indexOf('name: "', tool + 10);
    assert.ok(
        nextTool === -1 || promise < nextTool,
        "the promise sits outside create_document's description, so it does not describe this tool",
    );

    // "Never read" is the stronger half and is not covered by the assignment
    // test above, which deliberately ignores reads. Here a read is a defect too,
    // because the description says there are none.
    const mention = /\.AutoCorrect\.|AutoFormatAsYouType/g;
    assert.ok(
        mention.test("$script:App.AutoCorrect.ReplaceText"),
        "the detector cannot see a mention, so its silence means nothing",
    );
    mention.lastIndex = 0;
    const hits = [...hostSource.matchAll(mention)].map((m) => lineOf(hostSource, m.index));
    assert.deepEqual(
        hits,
        [],
        `word-host.ps1:${hits.join(", ")} touches those settings while the description promises it ` +
            `never does. Either the code or the promise is wrong; do not leave them disagreeing.`,
    );
});

test("the host never types, so the two settings no bait could reach are unreachable", () => {
    // `probe-autocorrect-necessity.ps1` gives a positive control for three of the
    // five: ReplaceText, ReplaceQuotes and ReplaceSymbols each have a bait proven
    // to fire. CorrectSentenceCaps and CorrectInitialCaps have none -- they are
    // keystroke handlers, and nothing programmatic was found that triggers them.
    // So the probe alone can only say they did not fire, not that they could not.
    //
    // This says they could not, and it is a fact about this codebase rather than
    // about Word: every character of document text goes in through
    // `Set-ParagraphText` -> `Range.Text` assignment, and no keystroke API is
    // called anywhere. A future edit that reaches for one is the event that would
    // make those two settings live again, and it reddens here.
    for (const api of ["TypeText", "TypeParagraph", "SendKeys", "TypeBackspace"]) {
        const re = new RegExp(`(?<![-\\w])${api}(?![-\\w])`, "g");
        assert.ok(re.test(`$sel.${api}()`), `the detector for ${api} cannot match its own name`);
        re.lastIndex = 0;
        const hits = [...hostSource.matchAll(re)].map((m) => lineOf(hostSource, m.index));
        assert.deepEqual(
            hits,
            [],
            `word-host.ps1:${hits.join(", ")} calls ${api}. Typing makes autocorrect and ` +
                `autoformat-as-you-type reachable, including CorrectSentenceCaps and ` +
                `CorrectInitialCaps, which no probe here has a positive control for. ` +
                `Measure before adding a keystroke path.`,
        );
    }
});
