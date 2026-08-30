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
// `word-host-encoding.test.mjs` is the depth check on the one host -- ordering,
// citation, comment-blanking. This is the breadth check, and it exists because
// the depth check could not see #50 at all: a per-file assertion covers exactly
// the file it names, and the defect's whole history is that it reappears in the
// next file nobody named. The rule is enforced against whatever is in the tree,
// so a host added tomorrow is covered on the day it lands rather than on the day
// someone remembers to extend a list.
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
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { assignsUtf8, blankComments, READS_STDIN, USES_CONSOLE, WRITES_CONSOLE } from "./ps-encoding-rule.mjs";

const execFileAsync = promisify(execFile);
const HERE = path.dirname(fileURLToPath(import.meta.url));
const REPO = path.resolve(HERE, "..", "..", "..", "..", "..");

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
    const { stdout } = await execFileAsync("git", ["ls-files", "-z", "*.ps1"], {
        cwd: REPO,
        maxBuffer: 8 * 1024 * 1024,
    });
    return stdout.split("\0").filter(Boolean).map((f) => f.replace(/\\/g, "/"));
}

let available = null;
async function gitAvailable() {
    if (available !== null) return available;
    try {
        await execFileAsync("git", ["rev-parse", "--git-dir"], { cwd: REPO });
        available = true;
    } catch {
        available = false;
    }
    return available;
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
