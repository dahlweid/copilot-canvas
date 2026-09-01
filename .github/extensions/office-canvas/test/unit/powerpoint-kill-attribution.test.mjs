// A PowerPoint the probes did not start is never quit and never killed.
//
// `spikes/powerpoint/probes/_common.ps1` used to snapshot the `POWERPNT` pid
// set, attach with `New-Object -ComObject PowerPoint.Application`, and treat the
// difference as "the instance we own" -- then `Quit()` it and force-kill it.
// Both halves are measured wrong in this repo:
//
//   * `New-Object` ATTACHES. PowerPoint is single-instance, so the object the
//     probe holds is routinely the USER'S PowerPoint, unsaved decks and all
//     (spikes/powerpoint/FINDINGS.md, single-instance).
//   * Differencing a census OVER-REPORTS: `probe-init-attribution.ps1` measured
//     2 new pids for 1 instance created, and the difference is non-empty by
//     construction whenever anybody else starts PowerPoint during the window.
//
// Which inverts the guard. `if ($ctx.Owned.Count -gt 0) { Quit }` passed
// PRECISELY in the race where the probe had attached to somebody else's
// instance -- the case it was written to catch. Issue #139.
//
// The rule this pins, in three parts:
//
//   1. `Stop-Process` appears under `spikes/powerpoint/` in exactly one place:
//      the body of `Stop-VerifiedPpt`. No exemptions. Every other kill in that
//      tree was removed rather than made cleverer, so there is nothing here to
//      carve out -- and a guard with no carve-out cannot be satisfied by
//      re-adding the defect.
//   2. Nothing calls `.Quit()` on a PowerPoint obtained through COM. Not a
//      grep: a per-file taint set, because four of the five original sites
//      reached the instance through a factory or a parameter rather than
//      through a literal `New-Object`.
//   3. `Stop-VerifiedPpt` itself decides correctly, checked by executing a
//      model of its extracted decision sequence.
//
// Why a source assertion rather than a leak or liveness assertion: these are
// probes. CI cannot run them -- they need a licensed PowerPoint -- and nobody
// runs all of them. The argument in `quit-argument.test.mjs` applies unchanged.
// It is sharper here: the only way to observe this defect at runtime is to have
// somebody's presentation destroyed by it.
//
// WHAT THIS CANNOT VERIFY, stated plainly:
//
//   * It does not run PowerShell. Test 3 assumes each extracted condition means
//     what the model says it means; an unrecognised condition FAILS rather than
//     being skipped, so the model cannot silently stop describing the code.
//   * It proves nothing about behaviour against a real PowerPoint. It does not
//     show that `Quit()` reaps, that a kill succeeds, or that `Stop-VerifiedPpt`
//     ever runs at all.
//   * Test 2's taint analysis is SYNTACTIC. An application passed through a
//     function parameter (other than the close helper's, which is named below),
//     stored in an array, or returned from a helper this does not know about
//     would escape it.
//   * Against the parent commit test 3 fails at EXTRACTION -- `Stop-VerifiedPpt`
//     does not exist there -- not on a case. So test 3 has no coverage of the
//     parent's behaviour and is not claimed to; tests 1 and 2 are what fail on
//     the defect, and they fail by naming its sites.
//
// Office-free. Reads tracked files and runs `git ls-files`.
//
// Run: node --test ".github/extensions/office-canvas/test/unit/*.test.mjs"

import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";

import { blankComments } from "./ps-encoding-rule.mjs";
import { REPO, gitAvailable, trackedFiles } from "./tracked-files.mjs";

const PATHSPEC = "spikes/powerpoint/*.ps1";
const KILLER = "Stop-VerifiedPpt";
const KILLER_FILE = "spikes/powerpoint/probes/_common.ps1";

// Both spellings of the factory are recognised, and both spellings of the close
// helper. The rename from `New-OwnedPowerPoint` / `Close-OwnedPowerPoint` is
// part of the fix -- the old names asserted an ownership the code cannot
// establish -- so a guard keyed only on the new ones would go quiet on exactly
// the tree it is meant to indict.
const FACTORIES = ["New-PowerPointInstance", "New-OwnedPowerPoint"];
const CLOSE_HELPERS = ["Close-PowerPointInstance", "Close-OwnedPowerPoint"];

/** The body of a `function name(...) { ... }`, by brace matching, or null. */
function functionBody(source, name) {
    const start = source.search(new RegExp(`function\\s+${name}\\b`));
    if (start === -1) return null;
    const open = source.indexOf("{", start);
    if (open === -1) return null;
    let depth = 0;
    for (let i = open; i < source.length; i++) {
        if (source[i] === "{") depth++;
        else if (source[i] === "}" && --depth === 0) return { text: source.slice(open + 1, i), start: open + 1, end: i };
    }
    return null;
}

const lineOf = (source, index) => source.slice(0, index).split("\n").length;

// --- 1: the only Stop-Process is inside Stop-VerifiedPpt ----------------------

test("spikes/powerpoint: Stop-Process appears only inside Stop-VerifiedPpt", async (t) => {
    if (!(await gitAvailable())) return t.skip("no git checkout to enumerate");

    const files = await trackedFiles(PATHSPEC);
    // An empty list would make this pass without examining anything -- the same
    // shape as a glob that matches nothing and exits 0, which this repo has
    // measured once already.
    assert.ok(files.length > 0, `git ls-files found no files for '${PATHSPEC}'; the enumeration is broken, not the tree`);

    const offenders = [];
    for (const file of files) {
        const source = blankComments(await readFile(path.join(REPO, file), "utf8"));
        // Comments are blanked first: three of the surviving mentions of the
        // old kills in this tree are prose explaining why they were removed,
        // and a scan over raw text would report the explanation as the bug.
        const sanctioned = file === KILLER_FILE ? functionBody(source, KILLER) : null;
        for (const m of source.matchAll(/Stop-Process/g)) {
            if (sanctioned && m.index >= sanctioned.start && m.index < sanctioned.end) continue;
            offenders.push(`${file}:${lineOf(source, m.index)}`);
        }
    }

    assert.deepEqual(
        offenders,
        [],
        `Stop-Process under spikes/powerpoint/ is only sound inside ${KILLER}, which verifies a CreateProcess pid ` +
            "against the process name and the StartTime recorded at launch. A pid taken from a census difference is " +
            "not attribution: New-Object attaches to the user's PowerPoint, and the difference over-reports. See " +
            `issue #139 and ${KILLER_FILE}.\n  ` +
            offenders.join("\n  "),
    );
});

test(`${KILLER} still exists and still kills`, async (t) => {
    if (!(await gitAvailable())) return t.skip("no git checkout to enumerate");

    // Without this, test 1 passes trivially on a tree where the one sanctioned
    // kill was deleted or renamed -- "no kills anywhere" and "kills only where
    // they are verified" are different properties and only one is intended.
    const source = blankComments(await readFile(path.join(REPO, KILLER_FILE), "utf8"));
    const body = functionBody(source, KILLER);
    assert.ok(body, `${KILLER} is not in ${KILLER_FILE}; test 1's exemption now covers nothing`);
    assert.match(body.text, /Stop-Process\s+-Id\s+\$ProcessId\s+-Force/, `${KILLER} no longer kills the pid it was given`);
});

// --- 2: no .Quit() on a COM-obtained PowerPoint -------------------------------

/**
 * Variables in `source` holding a PowerPoint the code cannot prove it created.
 *
 * Seeded from a literal `New-Object -ComObject PowerPoint.Application` and from
 * either factory name, then propagated through the one aliasing form the tree
 * actually uses, `$alias = $tainted.App` (spikes/powerpoint/probes/probe-hide.ps1,
 * spikes/powerpoint/probes/probe-export.ps1 and four others). Without that
 * propagation a future `$app.Quit()` in any of them passes.
 */
function comObtained(source) {
    const tainted = new Set();
    const factory = FACTORIES.map((f) => f.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")).join("|");
    for (const m of source.matchAll(/\$(\w+)\s*=\s*New-Object\s+-ComObject\s+PowerPoint\.Application/gi)) tainted.add(m[1]);
    for (const m of source.matchAll(new RegExp(String.raw`\$(\w+)\s*=\s*(?:${factory})\b`, "gi"))) tainted.add(m[1]);
    // Fixed point, because an alias may be assigned before its source in a
    // single pass over a file that defines functions out of order.
    for (let changed = true; changed; ) {
        changed = false;
        for (const m of source.matchAll(/\$(\w+)\s*=\s*\$(\w+)\.App\s*(?:$|[\r\n])/gim)) {
            if (tainted.has(m[2]) && !tainted.has(m[1])) {
                tainted.add(m[1]);
                changed = true;
            }
        }
    }
    return tainted;
}

test("spikes/powerpoint: nothing quits a PowerPoint obtained through COM", async (t) => {
    if (!(await gitAvailable())) return t.skip("no git checkout to enumerate");

    const files = await trackedFiles(PATHSPEC);
    assert.ok(files.length > 0, `git ls-files found no files for '${PATHSPEC}'; the enumeration is broken, not the tree`);

    const offenders = [];
    let scanned = 0;
    for (const file of files) {
        const source = blankComments(await readFile(path.join(REPO, file), "utf8"));
        const tainted = comObtained(source);

        // The close helper's `$ctx` is a PARAMETER, so no assignment-based taint
        // set can ever contain it -- and `_common.ps1`'s `$ctx.App.Quit()` is
        // the site this whole issue is about. It is discriminated by which
        // FUNCTION the quit sits in, which is structural rather than inferred:
        // this helper's ctx always comes from the factory, whereas
        // Stop-IsolatedPowerPoint's comes from CreateProcess and is genuinely
        // ours. Tainting every `X.App` instead would false-positive on that one.
        const helperBodies = CLOSE_HELPERS.map((name) => functionBody(source, name)).filter(Boolean);

        if (/New-Object\s+-ComObject\s+PowerPoint\.Application/i.test(source)) {
            scanned++;
            assert.ok(
                tainted.size > 0,
                `${file} creates a PowerPoint COM object but the taint scan found no variable holding it; ` +
                    "a degenerate parse must not pass",
            );
        }

        for (const m of source.matchAll(/\$(\w+)(?:\.App)?\.Quit\s*\(/g)) {
            const inHelper = helperBodies.some((b) => m.index >= b.start && m.index < b.end);
            if (!tainted.has(m[1]) && !(inHelper && m[1].toLowerCase() === "ctx")) continue;
            offenders.push(`${file}:${lineOf(source, m.index)}  $${m[1]}...Quit()`);
        }
    }

    assert.ok(scanned > 0, "no file in the pathspec creates a PowerPoint COM object; the scan is looking at the wrong tree");

    assert.deepEqual(
        offenders,
        [],
        "Quit() on a PowerPoint obtained through New-Object closes the USER'S decks: PowerPoint is single-instance, " +
            "so New-Object attaches rather than starts. There is no signal at this layer that establishes we created " +
            "the instance -- the census difference is not one, in either direction -- so release the RCW and report " +
            "instead. See issue #139.\n  " +
            offenders.join("\n  "),
    );
});

// --- 3: Stop-VerifiedPpt decides correctly ------------------------------------

// Each condition the model understands, as a predicate over a world. An
// unlisted condition is a parse failure, not a pass.
const CONDITIONS = [
    [/^-not \$ProcessId$/, (w) => !w.processId],
    [/^-not \$p$/, (w) => !w.exists],
    [/^\$p\.ProcessName -ne 'POWERPNT'$/, (w) => w.name !== "POWERPNT"],
    [/^\$null -eq \$ExpectedStart$/, (w) => w.expectedStart === null],
    [/^\$actual -ne \$ExpectedStart$/, (w) => w.actualStart !== w.expectedStart],
];

// A `try { STMT } catch { return 'x' }` guard: the read failing IS the
// condition. Both of these are reads of a live process that can vanish or be
// protected, which is why they are expressed as exceptions rather than tests.
const READS = [
    [/^\$null = \$p\.Handle$/, (w) => !w.handleReadable],
    [/^\$actual = \$p\.StartTime$/, (w) => !w.startReadable],
];

/** The ordered guard sequence in a function body, plus its fall-through return. */
function guardsOf(body) {
    const guards = [];
    let fallthrough = null;
    for (const raw of body.split("\n")) {
        const line = raw.trim();
        if (!line) continue;

        const conditional = /^if \((.+)\)\s*\{\s*return '([\w:]+)'\s*\}$/.exec(line);
        if (conditional) {
            const [, condition, state] = conditional;
            const known = CONDITIONS.find(([pattern]) => pattern.test(condition.trim()));
            assert.ok(known, `unrecognised condition \`${condition.trim()}\` in ${KILLER} -- the model no longer describes the code, so update it deliberately`);
            guards.push({ test: known[1], state });
            continue;
        }

        const guarded = /^try \{\s*(.+?)\s*\} catch \{\s*return '([\w:]+)'\s*\}$/.exec(line);
        if (guarded) {
            const [, statement, state] = guarded;
            const known = READS.find(([pattern]) => pattern.test(statement.trim()));
            assert.ok(known, `unrecognised guarded read \`${statement.trim()}\` in ${KILLER} -- update the model deliberately`);
            guards.push({ test: known[1], state });
            continue;
        }

        const plain = /^return '([\w:]+)'$/.exec(line);
        if (plain) fallthrough = plain[1];
    }
    assert.ok(fallthrough, `${KILLER} has no fall-through return`);
    return { guards, fallthrough };
}

/** Walk the extracted sequence: the first matching guard decides. */
const evaluate = ({ guards, fallthrough }, world) => {
    for (const guard of guards) if (guard.test(world)) return guard.state;
    return fallthrough;
};

const OURS = 1_700_000_000_000;
const THEIRS = 1_600_000_000_000;
const ok = { processId: 4242, exists: true, name: "POWERPNT", handleReadable: true, startReadable: true, actualStart: OURS, expectedStart: OURS };

// Every world names the result it must produce. "Handled somehow" is not the
// property: a helper that declined EVERY case would satisfy a looser test while
// silently converting "the CreateProcess kills are kept and hardened" into
// "there are no kills" -- which is a different change, shipped under this one's
// name. The `killed` case is what makes that impossible.
const CASES = [
    ["A: no pid recorded at all", { ...ok, processId: 0 }, "declined:nopid"],
    ["B: pid already gone -- sound, and not a decline", { ...ok, exists: false }, "gone"],
    ["C: handle cannot be pinned", { ...ok, handleReadable: false }, "declined:handle"],
    ["D: pid recycled onto something that is not PowerPoint", { ...ok, name: "notepad" }, "declined:name"],
    ["E: launcher never recorded a StartTime", { ...ok, expectedStart: null }, "declined:unverified"],
    ["F: StartTime unreadable", { ...ok, startReadable: false }, "declined:unreadable"],
    ["G: pid recycled onto another POWERPNT", { ...ok, actualStart: THEIRS }, "declined:start"],
    // I pins the ORDER. Absence is sound WITHOUT a recorded StartTime -- no
    // process holds the pid, so nothing we launched runs under it. Hoist the
    // null-ExpectedStart guard above the absence check and this is the only
    // case that changes: an already-gone process comes back as a decline and
    // the caller prints a leak warning about nothing. Measured: without this
    // world the mutation passes, because every other absent case here carries a
    // recorded start time. It is reachable in the tree -- _isolated.ps1 leaves
    // StartTime null when the launch-time read fails, and that process may well
    // have exited by teardown.
    ["I: pid absent AND no recorded StartTime", { ...ok, exists: false, expectedStart: null }, "gone"],
    // H is the one that fails if the helper is made to refuse everything.
    ["H: our own process, fully verified", ok, "killed"],
];

test(`${KILLER} kills a verified pid and declines every other world`, async () => {
    const source = blankComments(await readFile(path.join(REPO, KILLER_FILE), "utf8"));
    const body = functionBody(source, KILLER);
    assert.ok(body, `${KILLER} not found in ${KILLER_FILE} -- re-point or remove this test`);
    const sequence = guardsOf(body.text);

    // Deliberately far below the real guard count: this only exists so a
    // degenerate parse reports itself instead of surfacing as confusing case
    // failures. It must NOT sit at or near the actual number, or the test
    // degenerates into counting lines and would pass any rewrite that kept the
    // count while inverting a verdict.
    assert.ok(sequence.guards.length >= 3, `only ${sequence.guards.length} guards recognised in ${KILLER}; a degenerate parse must not pass`);

    // Pins the ORDER, not just the presence: E's `declined:unverified` must sit
    // AFTER B's absence check. Hoisted above it, an absent pid -- which is sound
    // without any recorded StartTime, because no process holds it -- would come
    // back as a decline, and the caller would print a leak warning about a
    // process that is already gone.
    const failures = CASES.flatMap(([label, world, expected]) => {
        const actual = evaluate(sequence, world);
        return actual === expected ? [] : [`${label}: expected '${expected}', got '${actual}'`];
    });
    assert.deepEqual(failures, [], `${failures.length} of ${CASES.length} worlds decided wrongly\n  ${failures.join("\n  ")}`);
});
