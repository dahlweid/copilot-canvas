// `Stop-VerifiedWord` in the SHIPPED Word host decides, for one pid, whether to
// terminate a process. It is the highest-consequence function in this tree:
// wrong, it kills a stranger's Word and takes their unsaved documents with it.
//
// It had no decision test at all, while its two imitations under `spikes/` --
// which kill nothing a user owns -- have thorough ones
// (word-kill-attribution.test.mjs, powerpoint-kill-attribution.test.mjs). That
// inversion is how issue #150's defect survived: the divergence between the
// three was visible, but only the cheap two were pinned.
//
// This is NOT a copy of the sibling model, and the differences are forced:
//
//   * The shipped states are PROSE ('declined:the start time could not be
//     read, so ownership is unproven'), not the spikes' short tokens, and one
//     of them interpolates. So each step carries an ID and a CLASS, and the
//     cases below assert those rather than pinning wording that is meant to be
//     readable and will be reworded.
//   * The name check is UNGUARDED here, and that is correct rather than an
//     omission -- see the world `name unreadable` below, whose predicate
//     encodes a measurement rather than the intuitive reading.
//   * The post-kill steps must be EVALUATED, not skipped. In the spikes helpers
//     what follows the kill is an observation that cannot spare a process; here
//     it decides between 'gone', a decline, and 'killed'. That removes the
//     sibling's neatest property -- there, reaching the kill ends the walk, so
//     a kill hoisted to the top is caught by every case at once -- so the
//     kill's POSITION is asserted directly instead.
//
// Office-free: reads one tracked file.
//
// Run: node --test ".github/extensions/office-canvas/test/unit/*.test.mjs"

import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";

import { blankComments } from "./ps-encoding-rule.mjs";
import { REPO } from "./tracked-files.mjs";

const KILLER = "Stop-VerifiedWord";
const KILLER_FILE = ".github/extensions/office-canvas/src/word/word-host.ps1";

/**
 * The body of a `function name(...) { ... }`, by brace matching, or null.
 *
 * A third copy of this in the unit tree, and deliberately not extracted: the
 * two existing copies are in merged files, and moving them is a change to
 * merged work that belongs in its own commit rather than smuggled into a fix.
 */
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

const readKiller = async () => blankComments(await readFile(path.join(REPO, KILLER_FILE), "utf8"));

// --- 1: the host has exactly one terminate, and it is inside the helper -------

test(`${KILLER_FILE} terminates a process only inside ${KILLER}`, async () => {
    const source = await readKiller();
    const body = functionBody(source, KILLER);
    assert.ok(body, `${KILLER} is not in ${KILLER_FILE}`);

    // Both spellings, because they are interchangeable at the call site and only
    // one of them is what this file happens to use today. A future terminate
    // added through the other must be caught by the same test.
    const sites = [...source.matchAll(/\.Kill\s*\(|\bStop-Process\b|\btaskkill\b/g)];
    assert.equal(
        sites.length,
        1,
        `${KILLER_FILE} contains ${sites.length} process terminations; there must be exactly one, and it must be the ` +
            `verified kill inside ${KILLER}. Found: ${sites.map((m) => m[0]).join(", ")}`,
    );
    assert.ok(
        sites[0].index > body.start && sites[0].index < body.end,
        `the one termination in ${KILLER_FILE} is outside ${KILLER}, so it kills a pid nothing proved is ours`,
    );
});

// --- 2: the decision sequence -------------------------------------------------

// Each condition the model understands, as an id, a pattern, and a predicate
// over a world. An unlisted condition is a parse failure, not a pass.
const CONDITIONS = [
    ["absent", /^\$null -eq \$p$/, (w) => !w.exists],
    // The predicate is `unreadable OR not Word`, and the first disjunct is the
    // measured part rather than a convenience. `$p.ProcessName` on a process
    // that has gone does not throw into a PowerShell caller: the .NET getter
    // throws, and the property adapter converts it to $null with `$Error`
    // growing by ZERO even under 'Stop'. So an unreadable name arrives here as
    // $null, `$null -ne 'WINWORD'` is true, and the helper returns 'gone'.
    // Measured, spikes/isolation/probes/probe-processname-after-exit.ps1 arms
    // A1/A2/A3. Wrapping this read in a try/catch would add a decline state
    // that cannot be reached -- which is what both spikes helpers did, on a
    // justification that was never probed.
    ["name", /^\$p\.ProcessName -ne 'WINWORD'$/, (w) => !w.nameReadable || w.name !== "WINWORD"],
    // The pin is a CONDITION, not a guarded read, and that is the second half of
    // the same measurement. Written as `try { $null = $p.Handle } catch { return
    // ... }` it could never decline: `.Handle` on an exited process answers
    // $null without throwing and with `$Error` growing by zero, exactly as the
    // name read does -- probe arms F1/F2/F3. The value has to be bound and
    // tested, so this step exists only while the code reads it back.
    ["handle", /^\$null -eq \$pinned$/, (w) => !w.handleReadable],
    ["no-expected", /^\$null -eq \$expectedStart$/, (w) => w.expectedStart === null],
    ["unreadable-start", /^\$null -eq \$actual$/, (w) => !w.startReadable],
    ["start-mismatch", /^\$actual -ne \$expectedStart$/, (w) => w.actualStart !== w.expectedStart],
];

// Conditions that may only appear AFTER the terminate. Unlike the spikes model
// these are not observations -- they decide the outcome -- so they are walked,
// but their position is still asserted: either of them above the kill would be
// reading a `$failure` no attempt had yet been able to set.
const POST_KILL = [
    ["kill-exited", /^\$failure -is \[InvalidOperationException\]$/, (w) => w.kill === "exited"],
    ["kill-failed", /^\$null -ne \$failure$/, (w) => w.kill === "failed"],
];

// There is deliberately no table of `try { STMT } catch { return '...' }` reads.
// The helper had one -- the pin -- and it was unreachable; see the `handle`
// entry above. A new one would land on the unparsed-returning-line assertion at
// the bottom of the walk, which is where an unmodelled decision belongs.

// Non-deciding lines the model recognises and steps over. The StartTime binding
// is accepted in BOTH the guarded and the bare form, because they are
// indistinguishable at runtime -- the same adapter conversion as the name read
// yields $null either way (probe arm E, on a live but protected process). A
// test that demanded the `try` would be asserting a property the platform does
// not actually give the code.
const BINDINGS = [
    /^\$actual = try \{ \$p\.StartTime \} catch \{ \$null \}$/,
    /^\$actual = \$p\.StartTime$/,
    /^\$pinned = try \{ \$p\.Handle \} catch \{ \$null \}$/,
    /^\$pinned = \$p\.Handle$/,
    /^\$p = Get-Process -Id \$candidate -ErrorAction SilentlyContinue$/,
    /^\$failure = \$null$/,
];

/** The ordered decision sequence in `body`: guards, the kill, and what follows. */
function decisionsOf(body) {
    // Every line after the pin reads `$p`, so the guards' verdict only transfers
    // to the kill if that name still means the process they pinned. Rebinding it
    // is invisible to a walk that matches on TEXT.
    const pin = body.indexOf("$p.Handle");
    assert.ok(pin !== -1, `${KILLER} no longer pins a handle before its identity reads`);
    assert.ok(
        !/^\s*\$p\s*=[^=]/m.test(body.slice(pin)),
        `${KILLER} rebinds $p after pinning the handle; every line below that names a process the guards never verified`,
    );
    // The pin's VALUE must survive the line that takes it. Discarding it into
    // $null leaves nothing for the condition above to test, and reaching the
    // next line proves nothing: measured, probe arm F3 -- `.Handle` on an exited
    // process answers $null without throwing, so control continues and every
    // guard below then passes on an object that was never pinned (arm F4).
    const pinLine = body.split("\n").find((l) => l.includes("$p.Handle")).trim();
    assert.ok(
        !/^\$null\s*=/.test(pinLine),
        `${KILLER} discards the pinned handle into $null (\`${pinLine}\`), so no line below can tell a handle from the $null an exited process answers with`,
    );

    const steps = [];
    const killCatch = [];
    let kills = 0;
    let inKillCatch = false;

    for (const raw of body.split("\n")) {
        const line = raw.trim();
        if (!line) continue;

        // The terminate spans lines, because its catch unwraps to the root
        // exception. Consume the block rather than letting its interior lines
        // reach the returning-line check below.
        if (/\$p\.Kill\s*\(\s*\)/.test(line)) {
            kills += 1;
            steps.push({ kill: true });
            inKillCatch = !/\}\s*$/.test(line);
            continue;
        }
        if (inKillCatch) {
            if (line === "}") inKillCatch = false;
            else killCatch.push(line);
            assert.ok(!/\breturn\b/.test(line), `${KILLER}'s terminate catch returns directly (\`${line}\`), so the outcome bypasses the decisions below it`);
            continue;
        }

        const conditional = /^if \((.+?)\)\s*\{\s*return (.+?)\s*\}$/.exec(line);
        if (conditional) {
            const [, condition, quoted] = conditional;
            const state = quoted.replace(/^["']|["']$/g, "");
            const post = POST_KILL.find(([, pattern]) => pattern.test(condition.trim()));
            if (post) {
                assert.equal(kills, 1, `\`${condition.trim()}\` in ${KILLER} inspects the terminate's failure but appears BEFORE the terminate`);
                steps.push({ id: post[0], test: post[2], state });
                continue;
            }
            const known = CONDITIONS.find(([, pattern]) => pattern.test(condition.trim()));
            assert.ok(known, `unrecognised condition \`${condition.trim()}\` in ${KILLER} -- the model no longer describes the code, so update it deliberately`);
            assert.equal(kills, 0, `\`${condition.trim()}\` in ${KILLER} is a guard on whether to kill, but it appears AFTER the kill, where it can only be dead code`);
            steps.push({ id: known[0], test: known[2], state });
            continue;
        }

        const final = /^return (.+)$/.exec(line);
        if (final) {
            assert.equal(kills, 1, `${KILLER} returns unconditionally before its terminate, so no world can reach the kill`);
            steps.push({ id: "killed", test: () => true, state: final[1].replace(/^["']|["']$/g, "") });
            continue;
        }

        // A returning line the patterns above did not understand must FAIL, not
        // be skipped: a guard the model cannot see is a guard the model reports
        // as absent, and it would report the world it guards as reaching the
        // kill. The sibling model lost `declined:unreadable-name` exactly this
        // way, to a character class.
        assert.ok(
            !/\breturn\b/.test(line),
            `unparsed returning line \`${line}\` in ${KILLER} -- it decides something the model cannot see, so extend the model deliberately`,
        );
        assert.ok(
            BINDINGS.some((pattern) => pattern.test(line)),
            `unrecognised line \`${line}\` in ${KILLER} -- if it cannot decide anything, add it to BINDINGS deliberately`,
        );
    }

    assert.ok(!inKillCatch, `${KILLER}'s terminate block was never closed, so the parse above is not describing the code`);
    assert.equal(kills, 1, `${KILLER} contains ${kills} terminations; the sanctioned kill must be exactly one, reached only after every guard`);

    // The two POST_KILL steps both read `$failure`, and the walk above takes
    // that variable on trust -- it models the world's `kill` outcome, not the
    // code that records it. So the catch that records it is asserted here
    // directly. Without this, emptying the body to `catch { }` leaves every one
    // of the twelve worlds green while a failed terminate is reported 'killed'.
    const caught = killCatch.join("\n");
    assert.match(
        caught,
        /\$failure\s*=\s*\$_\.Exception/,
        `${KILLER}'s terminate catch never records the failure, so $failure stays $null and a terminate that threw is reported as 'killed'`,
    );
    // And it must unwrap. PowerShell wraps a method-call failure, so the OUTER
    // type is MethodInvocationException whatever went wrong: measured, probe
    // arms C2/C3/C4, where `outer -is [InvalidOperationException]` is False
    // while the root IS one. Testing the wrapper would send the exit race to
    // the decline branch -- which is the defect this whole helper was fixed for.
    assert.match(
        caught,
        /while \(\$null -ne \$failure\.InnerException\)[^\n]*\$failure = \$failure\.InnerException/,
        `${KILLER}'s terminate catch does not unwrap to the root exception, so the InvalidOperationException test below inspects PowerShell's wrapper and can never match`,
    );
    return steps;
}

/** The class of an outcome. Prose states are pinned by class, not by wording. */
const classOf = (state) => (state.startsWith("declined:") ? "declined" : state);

const walk = (steps, world) => {
    for (const step of steps) {
        if (step.kill) continue;
        if (step.test(world)) return { id: step.id, class: classOf(step.state) };
    }
    return { id: "fell-through", class: "fell-through" };
};

const OURS = 1_700_000_000_000;
const THEIRS = 1_600_000_000_000;
const ok = {
    exists: true,
    handleReadable: true,
    nameReadable: true,
    name: "WINWORD",
    startReadable: true,
    actualStart: OURS,
    expectedStart: OURS,
    kill: "ok",
};

// Every world names the id that must decide it AND the class that id must
// return. Both, because the defect this test exists for was a step returning
// the wrong CLASS while sitting in exactly the right place: the exit race was
// decided by the terminate, correctly, and then reported as a decline.
const CASES = [
    ["A: pid already gone", { ...ok, exists: false }, "absent", "gone"],
    ["B: handle cannot be pinned", { ...ok, handleReadable: false }, "handle", "declined"],
    ["C: pid holds something that is not Word", { ...ok, name: "notepad" }, "name", "gone"],
    // D is the world the shipped code and both spikes helpers disagree about,
    // and the model's predicate for it is the measured one. It must be decided
    // by the NAME step -- reaching a decline state for an unreadable name would
    // mean a state that cannot occur.
    ["D: name unreadable -- $null, not a throw, so the name step decides", { ...ok, nameReadable: false }, "name", "gone"],
    ["E: no start time was ever recorded", { ...ok, expectedStart: null }, "no-expected", "declined"],
    ["F: start time unreadable", { ...ok, startReadable: false }, "unreadable-start", "declined"],
    ["G: pid recycled onto another WINWORD", { ...ok, actualStart: THEIRS }, "start-mismatch", "declined"],
    // H and I pin the ORDER of the identity guards against our own bookkeeping.
    // Hoist either start-time guard above the name check and H flips: a pid
    // holding a stranger's process would be reported by our missing ledger entry
    // instead of by the fact about the process.
    ["H: not Word AND nothing recorded -- the process decides, not our ledger", { ...ok, name: "notepad", expectedStart: null, startReadable: false }, "name", "gone"],
    ["I: unpinnable AND not Word -- the pin decides", { ...ok, handleReadable: false, name: "notepad" }, "handle", "declined"],
    // J is issue #150's defect. The process exits between the last guard and the
    // terminate; every guard passes on a corpse (probe arms B1/B2/B3), so
    // control genuinely reaches the kill and the kill throws. Reported as a
    // decline, callers printed four assertions about a cause -- "a leaked Word",
    // "an unrelated process that inherited the pid", "our Word exited and the
    // pid was reused", "its identity could not be proved" -- every one of them
    // false. A pid cannot be recycled while the pinned handle is open, so the
    // only knowable fact is that it exited, and that is 'gone'.
    ["J: exited between the last guard and the terminate", { ...ok, kill: "exited" }, "kill-exited", "gone"],
    ["K: the terminate failed for some other reason", { ...ok, kill: "failed" }, "kill-failed", "declined"],
    // L is what fails if the helper is ever made to refuse everything, which
    // would otherwise satisfy every case above.
    ["L: our own process, fully verified, terminate accepted", ok, "killed", "killed"],
];

test(`${KILLER} kills only a verified pid, and reports an exited one as gone`, async () => {
    const source = await readKiller();
    const body = functionBody(source, KILLER);
    assert.ok(body, `${KILLER} not found in ${KILLER_FILE} -- re-point or remove this test`);
    const steps = decisionsOf(body.text);

    // Deliberately far below the real step count, so a degenerate parse reports
    // itself rather than surfacing as confusing case failures. It must not sit
    // at or near the actual number, or the test degenerates into counting lines.
    assert.ok(steps.length >= 5, `only ${steps.length} decision steps recognised in ${KILLER}; a degenerate parse must not pass`);

    const failures = CASES.flatMap(([label, world, id, cls]) => {
        const got = walk(steps, world);
        return got.id === id && got.class === cls ? [] : [`${label}: expected ${id}/${cls}, got ${got.id}/${got.class}`];
    });
    assert.deepEqual(failures, [], `${failures.length} of ${CASES.length} worlds decided wrongly\n  ${failures.join("\n  ")}`);
});
