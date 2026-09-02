// Every PowerShell script that reads its stdin must pin *both* console
// encodings, not one.
//
// The asymmetry is the defect. Setting `[Console]::OutputEncoding` alone leaves
// the inbound direction on the OEM codepage, so UTF-8 JSON written to the
// child's stdin is decoded as codepage 850 and every non-ASCII character an
// agent sends becomes mojibake -- while reads stay faithful, which is what makes
// it survive being looked at. That was issue #40 in `word-host.ps1`, fixed in
// PR #46. The identical shape was then found still living in
// `spikes/live-word/live-word.ps1` (#50): outbound set, inbound never.
//
// There was a second file, `word-host-encoding.test.mjs`, asserting the same
// three properties of `word-host.ps1` alone. It is folded in here (#104): with
// that host pinned into `MUST_BE_COVERED` below, every assertion it made is one
// this file already makes over a set that contains it. Only its citation test
// was unique, and it is the last test in this file.
//
// What the fold does not fold away is the reason breadth won. A per-file
// assertion covers exactly the file it names, and this defect's whole history is
// that it reappears in the next file nobody named -- the depth check could not
// see #50 at all. The rule is enforced against whatever is in the tree, so a
// host added tomorrow is covered on the day it lands rather than on the day
// someone remembers to extend a list.
//
// Three things the deleted file recorded, which are the argument for this file
// existing at all and would otherwise have gone with it:
//
// The corruption **compounds**, which is why it is not merely cosmetic. The read
// direction stays faithful, so an agent gets the on-disk mojibake back exactly,
// and sending that in re-decodes the previous result's UTF-8 bytes as OEM a
// second time (measured: `Gr├╝nchen` in, `GrÔö£ÔòØnchen` out). One transform
// applied to both sides would cancel; this expands, so a non-ASCII paragraph
// becomes permanently *uneditable* -- `expectedText` can never equal what is
// stored, for any number of retries.
//
// It survived two merged layers not because anything hid it -- a faithful read
// makes it visible end to end on first look -- but because every fixture in the
// repo was ASCII, including the line in `make-fixture.ps1` labelled "Umlaut
// check:" that contained no umlaut.
//
// A source assertion is not merely the half that runs without Word. It is the
// only half that cannot be silenced by the machine it runs on. The integration
// checks can see this defect only where the OEM codepage differs from UTF-8 --
// measured here as 850, with `chcp`. Configure a host for UTF-8 system-wide and
// an unset `InputEncoding` defaults to UTF-8: nothing is corrupted, and those
// checks go green with the fix reverted, reporting a passing suite for an
// environment that could not produce the input. This file has no such blind
// spot, because it never decodes anything.
//
// Probe code is in scope deliberately. This repo's evidence fails by returning a
// plausible wrong answer rather than by crashing, so a probe that mangles its
// own non-ASCII input reports an encoding artefact of the harness as a platform
// behaviour, and nothing distinguishes the two.
//
// Office-free.
//
// Run: node --test ".github/extensions/office-canvas/test/unit/*.test.mjs"

import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { assignsUtf8, blankComments, READS_STDIN, USES_CONSOLE, WRITES_CONSOLE } from "./ps-encoding-rule.mjs";
import { REPO, gitAvailable, trackedFiles } from "./tracked-files.mjs";

// The extension folder, resolved from this file rather than from `REPO`.
//
// `REPO` is five levels up and only means anything in a checkout. An installed
// extension folder is the extension root, so five levels up from `test/unit/`
// there lands outside the install entirely. Everything in this file that
// enumerates goes through `REPO` and skips without git; the one test that names
// a single shipped file goes through this instead, and so keeps working.
const EXTENSION = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..", "..");
const SHIPPED_HOST = path.join(EXTENSION, "src", "word", "word-host.ps1");

// The probe that measured the console-encoding behaviour every preamble below
// claims. One constant, two readers: the swept citation test and the unguarded
// per-file one. Written as a full path, not a basename, because
// `tools/check-citations.mjs` resolves basenames repo-wide and several collide
// across `spikes/isolation/` and `spikes/powerpoint/`, so a basename match would
// accept a citation naming the wrong probe.
const ENCODING_PROBE = "spikes/isolation/probes/probe-console-input-encoding.mjs";

// The one script allowed to read stdin without pinning the encoding, because
// varying that assignment across arms is the measurement it performs. The test
// below asserts it is still present, so the exemption cannot outlive the file
// and quietly start covering something else.
const VARIES_THE_ENCODING = "spikes/isolation/probes/probe-console-input-encoding.ps1";

// Named so a walk that silently stopped finding things cannot pass. Both are
// stdin-reading hosts today; if one is renamed, this reddens and asks a human
// whether the rule moved with it rather than reporting a clean sweep of two
// files it never opened.
const MUST_BE_COVERED = [
    ".github/extensions/office-canvas/src/word/word-host.ps1",
    "spikes/live-word/live-word.ps1",
];

async function trackedPowerShellFiles() {
    return trackedFiles("*.ps1");
}

/**
 * Comments blanked, offsets preserved. Imported rather than restated -- see
 * `ps-encoding-rule.mjs` for why there is exactly one copy of these.
 */
async function stdinReadingScripts() {
    const found = [];
    for (const file of await trackedPowerShellFiles()) {
        if (file === VARIES_THE_ENCODING) continue;
        const raw = await readFile(path.join(REPO, file), "utf8");
        const source = blankComments(raw);
        if (READS_STDIN.test(source)) found.push({ file, source });
    }
    return found;
}

/**
 * The part of a script above its first touch of the console.
 *
 * Takes the *raw* source and returns raw text: the citation being looked for
 * **is** a comment, and the comment-blanking the sweeps above rely on would
 * erase the thing being asserted. The offset is still taken from the blanked
 * copy, so a `#` inside a comment cannot move the boundary.
 *
 * One copy, because the swept citation test and the unguarded per-file one ask
 * the same question of two differently-rooted files, and two records of one
 * quantity are free to disagree -- the failure `tracked-files.mjs` was
 * extracted to stop.
 *
 * @param {string} rawSource
 */
function preambleOf(rawSource) {
    return rawSource.slice(0, blankComments(rawSource).search(USES_CONSOLE));
}

test("the enumeration finds the scripts it is supposed to be checking", async (t) => {
    if (!(await gitAvailable())) return t.skip("no git checkout to enumerate");

    const files = (await stdinReadingScripts()).map((s) => s.file);
    // An empty set makes every assertion below pass without reading anything --
    // the same failure as a glob that matches nothing and exits 0, which this
    // repo has already been bitten by once in CI.
    assert.ok(files.length > 0, "no stdin-reading .ps1 found; the enumeration is broken, not the tree");
    for (const expected of MUST_BE_COVERED) {
        assert.ok(
            files.includes(expected),
            `${expected} is no longer discovered as a stdin-reading script. Either it stopped reading stdin, or it ` +
                `moved and this guard is now sweeping a tree that does not contain it.`,
        );
    }

    // `SHIPPED_HOST` and `MUST_BE_COVERED[0]` are two records of one path,
    // written against different roots because they answer to different
    // conditions -- one has to survive an install, the other has to be
    // comparable against `git ls-files` output. Two records of one quantity are
    // free to disagree, which is the failure `tracked-files.mjs` was extracted
    // to stop, so they are tied together here rather than left to drift.
    //
    // Here and not at module scope: this equality is path arithmetic through
    // `REPO`, and `REPO` is meaningless in an installed folder. Asserting it
    // beside the constant would redden the one test that is supposed to run
    // there. This test already only runs where `REPO` means something.
    assert.equal(
        path.resolve(REPO, MUST_BE_COVERED[0]),
        SHIPPED_HOST,
        "the repo-relative pin and the extension-relative one name different files; one of them moved",
    );
});

test("every stdin-reading .ps1 decodes its stdin as UTF-8", async (t) => {
    if (!(await gitAvailable())) return t.skip("no git checkout to enumerate");

    const offenders = [];
    for (const { file, source } of await stdinReadingScripts()) {
        if (!assignsUtf8("InputEncoding").test(source)) offenders.push(file);
    }

    assert.deepEqual(
        offenders,
        [],
        "these scripts read [Console]::In without setting [Console]::InputEncoding, so Windows PowerShell decodes " +
            "the UTF-8 they are sent as the OEM codepage and corrupts every non-ASCII character silently " +
            "(issue #40, #50):\n  " +
            offenders.join("\n  "),
    );
});

test("every stdin-reading .ps1 encodes its stdout as UTF-8", async (t) => {
    if (!(await gitAvailable())) return t.skip("no git checkout to enumerate");

    // Its own test rather than a clause of the one above, so a regression names
    // the direction that broke. The two have different signatures -- an unset
    // InputEncoding expands `ß` into two characters, an unset OutputEncoding
    // replaces it with U+FFFD -- and a message saying only "encoding" sends the
    // reader to the wrong end of the pipe.
    const offenders = [];
    for (const { file, source } of await stdinReadingScripts()) {
        if (!assignsUtf8("OutputEncoding").test(source)) offenders.push(file);
    }

    assert.deepEqual(offenders, [], "these scripts write to the console without pinning UTF-8 on the way out:\n  " + offenders.join("\n  "));
});

test("both encodings are pinned before the console is first touched", async (t) => {
    if (!(await gitAvailable())) return t.skip("no git checkout to enumerate");

    // Not ceremony. .NET builds the reader behind `[Console]::In` on first use
    // and caches it, so an assignment made afterwards replaces the property
    // while that reader keeps decoding with the old codepage. A correct pair of
    // assignments moved below the read loop restores the defect and satisfies
    // both greps above.
    const offenders = [];
    for (const { file, source } of await stdinReadingScripts()) {
        const firstUse = source.search(USES_CONSOLE);
        const input = source.search(assignsUtf8("InputEncoding"));
        const output = source.search(assignsUtf8("OutputEncoding"));
        // `search` reports -1 when absent, and -1 is below every real offset, so
        // without these two guards a file setting neither encoding would pass
        // this test -- which is the defect it exists to catch.
        if (input < 0 || output < 0) continue; // already reported by the tests above
        if (input > firstUse) offenders.push(`${file} (InputEncoding set after the console reader is built)`);
        if (output > firstUse) offenders.push(`${file} (OutputEncoding set after the console writer is built)`);
    }

    assert.deepEqual(offenders, [], "encoding pinned too late to take effect:\n  " + offenders.join("\n  "));
});

test("every shipped .ps1 that writes to the console pins UTF-8 on the way out", async (t) => {
    if (!(await gitAvailable())) return t.skip("no git checkout to enumerate");

    // A separate rule from the stdin one because it catches a different set.
    // `word-icon.ps1` reads no stdin at all, so every assertion above passes it
    // over -- and it wrote to both stdout and stderr with nothing pinned. Its
    // stdout payload is base64 and so identical under every codepage in play,
    // which is why this was never visible: the channel was safe because of what
    // happened to be travelling on it, not because of any rule.
    //
    // The exposure is the failure path. Measured on this machine, spawning that
    // script's exact shape with the encoding unpinned, non-ASCII on stderr
    // arrives as one U+FFFD per character -- lossy, unlike the inbound
    // direction's reversible mojibake -- and `$ErrorActionPreference = 'Stop'`
    // means a German PowerShell error is precisely what a failure sends there.
    //
    // Scoped to the shipped extension rather than the whole tree. Probes are
    // held to the stdin rule above but not to this one: several deliberately
    // exercise unpinned console behaviour, and a rule that swept them would be
    // met with exemptions until it meant nothing.
    const SHIPPED = ".github/extensions/office-canvas/src/";

    const shipped = (await trackedPowerShellFiles()).filter((f) => f.startsWith(SHIPPED));
    assert.ok(shipped.length > 0, `no shipped .ps1 found under ${SHIPPED}; the enumeration is broken, not the tree`);

    const offenders = [];
    let writers = 0;
    for (const file of shipped) {
        const source = blankComments(await readFile(path.join(REPO, file), "utf8"));
        if (!WRITES_CONSOLE.test(source)) continue;
        writers += 1;
        if (!assignsUtf8("OutputEncoding").test(source)) {
            offenders.push(file);
            continue;
        }
        // Same reason as the ordering test above, and it bites harder here:
        // the early-exit diagnostic in `word-icon.ps1` writes to stderr well
        // before the payload reaches stdout, so an assignment placed between
        // the two would be too late for the only message that carries German.
        const firstUse = source.search(USES_CONSOLE);
        if (source.search(assignsUtf8("OutputEncoding")) > firstUse) {
            offenders.push(`${file} (OutputEncoding set after the console writer is built)`);
        }
    }

    assert.ok(writers > 0, "no shipped .ps1 was found writing to the console; the matcher stopped matching");
    assert.deepEqual(
        offenders,
        [],
        "these shipped scripts write to the console without pinning UTF-8 outbound, so a non-ASCII diagnostic " +
            "reaches the extension as U+FFFD and cannot be recovered (#55):\n  " + offenders.join("\n  "),
    );
});

test("every stdin-reading .ps1 cites the probe that measured the encoding rule", async (t) => {
    if (!(await gitAvailable())) return t.skip("no git checkout to enumerate");

    // Swept rather than named, and that is the whole point of the shape. #147
    // asked for "a second, guarded test covering live-word.ps1", which closes
    // today's gap and leaves tomorrow's: the next stdin-reading script gets its
    // encoding *assignment* swept automatically and its *citation* required by
    // nothing. This file's own stated principle is that the defect reappears in
    // the file nobody named.
    //
    // Note this adds no enumeration of its own. It runs over exactly the
    // population `stdinReadingScripts()` already computes, so it inherits the
    // one exemption that already exists -- `VARIES_THE_ENCODING`, the probe
    // itself, which has nothing to cite because it *is* the measurement. An
    // exemption list and a coverage list fail in opposite directions: a file
    // missing from a coverage list is silently uncovered, while a file missing
    // from an exemption list is covered. Only the second is safe to forget.
    const offenders = [];
    let checked = 0;
    for (const { file } of await stdinReadingScripts()) {
        checked += 1;
        const raw = await readFile(path.join(REPO, file), "utf8");
        if (!preambleOf(raw).includes(ENCODING_PROBE)) offenders.push(file);
    }

    // Same reason as the enumeration test above: an empty set would pass this
    // without opening a file.
    assert.ok(checked > 0, "no stdin-reading .ps1 was checked for a citation; the enumeration is broken, not the tree");
    assert.deepEqual(
        offenders,
        [],
        `these scripts pin the console encodings without citing ${ENCODING_PROBE}, the probe that measured the ` +
            `behaviour they are relying on, above the first use of the console (#147):\n  ` + offenders.join("\n  "),
    );
});

test("the probe that varies the encoding is still present", async (t) => {
    if (!(await gitAvailable())) return t.skip("no git checkout to enumerate");

    // The exemption is by path. If the file is renamed or deleted, the constant
    // above stops matching anything and silently exempts nothing -- which is
    // harmless -- but it also stops being reviewed, and the next reader has no
    // way to tell an obsolete exemption from a load-bearing one.
    const files = await trackedPowerShellFiles();
    assert.ok(
        files.includes(VARIES_THE_ENCODING),
        `${VARIES_THE_ENCODING} is exempt from the rule above but is not in the tree; the exemption is stale`,
    );
});

// Moved here from `word-host-encoding.test.mjs` when that file was folded in
// (#104). It is the one assertion of that file's four which this one did not
// already make over a set containing `word-host.ps1`.
//
// Two deliberate departures from the original, so this is a move plus two
// strengthenings rather than a move:
//
//   * the preamble boundary is `USES_CONSOLE` (`In|Out|Error`) rather than the
//     original's `In|Out`. An extra alternative can only pull a first match
//     earlier, so the boundary can only shrink -- fail-closed. Measured in
//     `word-host.ps1` it does not move at all today: `:88-89` are the encoding
//     *properties*, which `\b` excludes, the first real use is `[Console]::Out`
//     at `:142`, and the first `Error` is at `:344`. It would shrink only if a
//     stderr write were ever added above the assignments, which is the case
//     worth reddening on.
//   * the pattern is the probe's full path, not its basename, for the reason
//     given at `ENCODING_PROBE`.
//
// **What the fold costs, stated beside what it keeps.** All four tests of the
// deleted file were extension-relative and unguarded, so in an installed folder
// `main` asserts four properties of `word-host.ps1` and this file asserts one.
// The other three -- both `assignsUtf8` checks and the ordering check -- are
// absorbed into the swept tests above and now run only in a checkout. Measured
// with `PATH` cleared: the deleted file was 4 passed / 0 skipped; this file is
// 1 passed / 6 skipped, against 0 passed / 6 skipped before.
//
// That is the right trade, and the reason is that these are source assertions.
// They guard against an *edit* to `word-host.ps1`, and every edit happens in a
// checkout; an installed folder holds a copy nobody modifies. What breadth buys
// in exchange is the thing depth structurally could not -- it is why #50 was
// found in `live-word.ps1` and the depth check could not see it.
//
// **Why this test still names one file when the test above sweeps.** #147 asked
// for the citation gap on `live-word.ps1` to be closed without guarding this
// assertion, and the sweep above is that fix. This one is *not* the redundant
// half left behind: the two are rooted differently on purpose. The sweep goes
// through `REPO`, which is meaningless in an installed extension folder, so it
// sits behind `gitAvailable()` with its five neighbours. This one goes through
// `EXTENSION` (`import.meta.url`) and is the only assertion this file makes
// without a checkout. Widening it to cover `live-word.ps1` -- which is the
// tidiest-looking diff available here -- would have dragged it through `REPO`
// too and taken the install-mode pass count to zero, which is precisely what
// #147 exists to prevent. The overlap on `word-host.ps1` is the price of that
// and is cheap: one `readFile` of a file the sweep also reads.
test("the probe backing the encoding claim is cited where the encodings are set", async () => {
    // This repo's rule is that a claim about platform behaviour is backed by a
    // probe that was actually run. The citation is what makes the next reader
    // able to re-run it rather than re-derive it, and `tools/check-citations.mjs`
    // is what stops the path going stale -- but only for citations that exist.
    // That it exists at all is what this asserts.
    const rawSource = await readFile(SHIPPED_HOST, "utf8");

    assert.ok(
        preambleOf(rawSource).includes(ENCODING_PROBE),
        `the encoding preamble in src/word/word-host.ps1 must cite ${ENCODING_PROBE}, the probe that measured it`,
    );
});
