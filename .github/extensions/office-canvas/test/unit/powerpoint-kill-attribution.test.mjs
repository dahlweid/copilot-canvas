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
//      model of its extracted decision sequence -- in which the KILL is a step
//      with a position, not a footnote, so a kill hoisted above the guards is
//      caught rather than hidden behind declines that have become dead code.
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
//     the defect, and they fail by naming its sites. Measured, by running this
//     file against `7ff59bb` in a worktree of that commit: all four tests fail,
//     test 3 and the existence test on extraction, and tests 1 and 2 between
//     them name 15 sites -- 10 kills and 5 `Quit()`s.
//
// STANDING MUTATION SET for the extraction model. Each of these is a post-kill
// re-check that is PRESENT BUT WRONG, and each passed all four tests in this
// file before `decisionsOf` was widened. Re-run them against any change to the
// extractor: a change that admits one has re-opened the hole.
//
//   1. `$ProcessId` rebound between the kill and the re-check, so the check
//      spells the verified pid and queries another one -- rejected by the
//      reassignment assertion at the top of `decisionsOf`.
//   2. The re-check hoisted ABOVE the kill and split across two physical lines
//      (`if (` then `Get-Process ...) { return 'killed:survived' }`). The first
//      line carries no `return` and the second does not start with `if (`, so
//      both slipped the catch-all and POST_KILL was never consulted -- rejected
//      by the unparsed-returning-line assertion.
//   3. A nested `function Get-Process` shadowing the real lookup so its value
//      cannot gate the state -- same assertion; the nested body carries a
//      `return`.
//   4. The pattern present only inside a string literal, with no re-check at
//      all -- same assertion.
//
// Found by review round 1 on #145 and measured against the RUNNING guard rather
// than against the regex alone, which is the distinction that made them
// visible: all four green before, all four red after, each naming the assertion
// above. The general lesson outlives the four -- an existence assertion pins
// TEXT, so it is worth exactly what the extractor's line coverage is worth.
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

    // And that it checks what the kill did. `Stop-Process -ErrorAction
    // SilentlyContinue` swallows its errors, so without this re-check 'killed'
    // is a claim about a CALL, not about a process, and a caller that reads it
    // as "reaped" hides a leak.
    //
    // Pinned HERE rather than in test 3 because test 3's model deliberately
    // cannot see it: reaching the kill is what decides 'killed' there, and what
    // the helper reports afterwards cannot terminate a process the guards would
    // have spared, so the model treats it as outside the decision. That leaves
    // presence unpinned -- measured, by deleting the re-check and watching test
    // 3 stay green. Test 3 pins the check's POSITION (it may only appear after
    // the kill); this pins its EXISTENCE. Neither implies the other. The
    // position pin is only as good as the extractor's line coverage, which is
    // why `decisionsOf` fails any unparsed line carrying a `return`.
    assert.match(
        body.text,
        /Get-Process\s+-Id\s+\$ProcessId\s+-ErrorAction\s+SilentlyContinue\)\s*\{\s*return\s+'killed:survived'/,
        `${KILLER} issues the kill but no longer verifies it, so 'killed' has gone back to being a claim about a call ` +
            "that swallowed its errors rather than about a process",
    );
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
    [/^\$name -ne 'POWERPNT'$/, (w) => w.name !== "POWERPNT"],
    [/^\$null -eq \$ExpectedStart$/, (w) => w.expectedStart === null],
    [/^\$actual -ne \$ExpectedStart$/, (w) => w.actualStart !== w.expectedStart],
];

// A `try { STMT } catch { return 'x' }` guard: the read failing IS the
// condition. All three are reads of a live process that can vanish or be
// protected, which is why they are expressed as exceptions rather than tests.
const READS = [
    [/^\$null = \$p\.Handle$/, (w) => !w.handleReadable],
    [/^\$name = \$p\.ProcessName$/, (w) => !w.nameReadable],
    [/^\$actual = \$p\.StartTime$/, (w) => !w.startReadable],
];

/**
 * Conditions that may only appear AFTER the kill: observations of what the
 * machine did, not decisions about whether to act.
 *
 * They are listed rather than ignored so their POSITION is asserted. One of
 * these appearing above the kill would be a `Get-Process` existence test
 * standing in for attribution -- precisely the reasoning #139 is about -- and
 * `decisionsOf` refuses it.
 */
const POST_KILL = [/^Get-Process -Id \$ProcessId -ErrorAction SilentlyContinue$/];

/**
 * The ordered decision sequence in a function body: guards AND the kill, in
 * source order.
 *
 * The kill is a step, not a footnote. Without it in the sequence the model
 * describes only which worlds return early, and a rewrite that hoisted the
 * `Stop-Process` to the FIRST line of the helper -- killing before a single
 * guard runs -- would satisfy every case below, because each case's expected
 * decline is still reachable further down as dead code. That mutation is the
 * single most dangerous regression this helper has, and it is caught only
 * because reaching the kill terminates the walk.
 *
 * Plain `return` lines are deliberately NOT collected. An earlier version of
 * this extractor took the LAST one as a fall-through, which models a function
 * that cannot exist: a body with two plain returns at one nesting level returns
 * at the FIRST, so the last-wins reading describes the wrong function. Nothing
 * needs them now -- reaching the kill is what decides `killed`.
 */
function decisionsOf(body) {
    // Every line after the guards names `$ProcessId`, so the guards' verdict
    // only transfers to the kill and the re-check if that name still means the
    // pid they verified. Rebinding it is invisible to the walk below, which
    // matches on TEXT: measured, inserting `$ProcessId = $OtherPid` between the
    // kill and the re-check left all four tests of this file green while the
    // re-check queried a process nothing had verified.
    assert.ok(
        !/^\s*\$ProcessId\s*=[^=]/m.test(body),
        `${KILLER} reassigns $ProcessId; every line after that names a pid the guards never verified, while still reading as though it did`,
    );

    const steps = [];
    let kills = 0;
    for (const raw of body.split("\n")) {
        const line = raw.trim();
        if (!line) continue;

        const conditional = /^if \((.+)\)\s*\{\s*return '([\w:-]+)'\s*\}$/.exec(line);
        if (conditional) {
            const [, condition, state] = conditional;
            if (POST_KILL.some((pattern) => pattern.test(condition.trim()))) {
                assert.equal(kills, 1, `\`${condition.trim()}\` in ${KILLER} is an observation of the kill's effect, but it appears BEFORE the kill, where it would be an existence test doing attribution's job (#139)`);
                steps.push({ postKill: true, state });
                continue;
            }
            const known = CONDITIONS.find(([pattern]) => pattern.test(condition.trim()));
            assert.ok(known, `unrecognised condition \`${condition.trim()}\` in ${KILLER} -- the model no longer describes the code, so update it deliberately`);
            steps.push({ test: known[1], state });
            continue;
        }

        const guarded = /^try \{\s*(.+?)\s*\} catch \{\s*return '([\w:-]+)'\s*\}$/.exec(line);
        if (guarded) {
            const [, statement, state] = guarded;
            const known = READS.find(([pattern]) => pattern.test(statement.trim()));
            assert.ok(known, `unrecognised guarded read \`${statement.trim()}\` in ${KILLER} -- update the model deliberately`);
            steps.push({ test: known[1], state });
            continue;
        }

        // A returning line the patterns above did not understand must FAIL, not
        // be skipped. On the sibling guard `declined:unreadable-name` was
        // silently dropped because the state pattern's character class had no
        // hyphen, and the model then reported that world as reaching the kill.
        // A guard the model cannot see is a guard the model reports as absent.
        //
        // The test is `contains a return`, NOT `starts with if( or try{`. An
        // earlier version asked the narrower question and could be blinded by a
        // LINE BREAK: splitting the post-kill check as `if (` + `Get-Process
        // ...) { return 'killed:survived' }` left the first line returnless and
        // the second line not starting with `if (`, so both slipped past, the
        // check never entered the sequence, and POST_KILL was never consulted --
        // measured, a re-check hoisted ABOVE the kill that way passed all four
        // tests of this file. The same widening also rejects a `return` hidden
        // in a string literal or in a nested `function` declaration.
        assert.ok(
            !/\breturn\b/.test(line) || /^return '[\w:-]+'$/.test(line),
            `unparsed returning line \`${line}\` in ${KILLER} -- it decides something the model cannot see, so extend the model deliberately`,
        );

        if (/^Stop-Process\b/.test(line)) {
            kills += 1;
            steps.push({ kill: true });
        }
    }
    // Exactly one. Two would mean a world can be killed down a path the cases
    // below never walk; zero would mean test 1's exemption covers nothing.
    assert.equal(kills, 1, `${KILLER} contains ${kills} Stop-Process calls; the sanctioned kill must be exactly one, reached only after every guard`);
    return steps;
}

/**
 * Walk the sequence: the first matching guard decides, and reaching the kill
 * decides `killed`.
 *
 * `killed` here means REACHED THE KILL. What the helper reports afterwards
 * ('killed' vs 'killed:survived') is an observation about the machine, not an
 * attribution decision, and is deliberately outside this model -- no value of it
 * can terminate a process that the guards would have spared.
 */
const evaluate = (steps, world) => {
    for (const step of steps) {
        if (step.postKill) continue;
        if (step.kill) return "killed";
        if (step.test(world)) return step.state;
    }
    return "fell-through-without-killing";
};

const OURS = 1_700_000_000_000;
const THEIRS = 1_600_000_000_000;
const ok = { processId: 4242, exists: true, name: "POWERPNT", handleReadable: true, nameReadable: true, startReadable: true, actualStart: OURS, expectedStart: OURS };

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
    // F2 is why the read is guarded at all. ProcessName on a Process whose
    // process has exited can throw, and this helper is called from a finally in
    // both of its callers -- so an uncaught throw skips a CloseDesktop and, in
    // probe-second-process.ps1, the leak census itself. Declining is the only
    // outcome that leaves the caller's cleanup running.
    ["F2: name unreadable -- must decline, not throw out of the caller's finally", { ...ok, nameReadable: false }, "declined:unreadable-name"],
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
    // J is the mirror of I for the handle pin. The pin must sit ABOVE the
    // identity reads, so that a pid recycled between the reads cannot be killed
    // under a name checked on its predecessor. Moving the pin below the name
    // check leaves this world reporting a name verdict for a process whose
    // handle was never held.
    ["J: handle unpinnable AND recycled onto another name", { ...ok, handleReadable: false, name: "notepad" }, "declined:handle"],
    // K pins the ORDER AMONG THE GUARDS: the name check must sit above both
    // verification checks, so a pid that fails identity AND verification is
    // reported by identity -- the fact about the process -- rather than by our
    // own bookkeeping. Measured: moving either the `$ExpectedStart` guard or the
    // `$p.StartTime` read above the name check flips K and nothing else, in both
    // cases 1 of 12.
    //
    // An earlier version of this comment claimed K was the only world a hoisted
    // kill would flip. That was false and this file refutes it: reaching the
    // kill terminates the walk, so a kill on the first line returns 'killed' for
    // every world, and the mutation is reported as 11 of 12 worlds decided
    // wrongly. The kill's position is pinned by the walk itself, not by K.
    ["K: identity and verification both fail -- identity must be the verdict", { ...ok, name: "notepad", expectedStart: null, startReadable: false }, "declined:name"],
    // H is the one that fails if the helper is made to refuse everything.
    ["H: our own process, fully verified", ok, "killed"],
];

test(`${KILLER} kills a verified pid and declines every other world`, async () => {
    const source = blankComments(await readFile(path.join(REPO, KILLER_FILE), "utf8"));
    const body = functionBody(source, KILLER);
    assert.ok(body, `${KILLER} not found in ${KILLER_FILE} -- re-point or remove this test`);
    const steps = decisionsOf(body.text);

    // Deliberately far below the real step count: this only exists so a
    // degenerate parse reports itself instead of surfacing as confusing case
    // failures. It must NOT sit at or near the actual number, or the test
    // degenerates into counting lines and would pass any rewrite that kept the
    // count while inverting a verdict.
    assert.ok(steps.length >= 4, `only ${steps.length} decision steps recognised in ${KILLER}; a degenerate parse must not pass`);

    const failures = CASES.flatMap(([label, world, expected]) => {
        const actual = evaluate(steps, world);
        return actual === expected ? [] : [`${label}: expected '${expected}', got '${actual}'`];
    });
    assert.deepEqual(failures, [], `${failures.length} of ${CASES.length} worlds decided wrongly\n  ${failures.join("\n  ")}`);
});
