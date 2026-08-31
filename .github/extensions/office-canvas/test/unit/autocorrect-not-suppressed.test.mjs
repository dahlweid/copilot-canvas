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
// MUTANT (run red): add `$script:App.AutoCorrect.ReplaceText = $false` to
// `Initialize-Word`. Reports `word-host.ps1:759 assigns
// AutoCorrect.ReplaceText`.
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
        "document-author.mjs": await readFile(AUTHOR, "utf8"),
        "extension.mjs": await readFile(EXTENSION, "utf8"),
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
