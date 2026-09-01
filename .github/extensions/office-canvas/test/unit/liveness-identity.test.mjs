// A missing start time is never proof that our Word exited.
//
// The teardown in `probe-addressing.ps1` and `probe-bulk-read.ps1` records the
// owned pid's `StartTime` at attribution and re-reads it on every poll, so that
// a pid recycled onto another process is not reported as our Word surviving.
// `Get-WordStartTime` returns `$null` both when the process cannot be found and
// from its catch, so the recorded value can legitimately be null.
//
// The defect this guards against: with `$expectedStart` null and `$actual` a
// real `DateTime`, `$actual -ne $expectedStart` is `$true` in PowerShell, so the
// liveness check returned 'gone' for a Word that was running and had just been
// created. The poll loop never iterated, the verdict stayed 0, and the probe
// printed "our Word (pid N) exited." -- a clean teardown that nothing observed,
// inverted and silent. Measured directly: `(Get-Date) -ne $null` is `$true`.
//
// The rule has two halves and they are NOT symmetric, which is why this test
// pins the ORDER of the guards and not merely their presence:
//
//   * Absence is sound without a start time. No process holds the pid, so
//     nothing we created runs under it. This must still conclude 'gone'.
//   * Presence is NOT sound without one. The pid may have been recycled. This
//     must conclude 'unknown', which routes to the "teardown UNVERIFIED"
//     verdict rather than to a clean one.
//
// So the null-expectedStart guard has to sit AFTER the two absence checks. At
// the top of the function it would convert sound 'gone' results into 'unknown'
// and manufacture UNVERIFIED verdicts out of good evidence. Case E below is the
// one that fails if it is moved there.
//
// Why a source assertion rather than a behavioural one: these are probes. CI
// cannot run them -- they need a licensed Word -- and nobody runs all of them.
// The reasoning in `quit-argument.test.mjs` applies unchanged: a leak assertion
// only fires on the probe someone happened to execute. This defect was found by
// reading, and reading does not happen twice.
//
// What this does and does not prove. It extracts the ordered decision sequence
// from the shipped source and EXECUTES a model of it, so it is not a check that
// a line exists -- moving, reordering or deleting a guard changes the outcome.
// It does not run PowerShell, so it assumes each condition means what the map
// below says it means. That mapping was verified against the real interpreter
// before this was written; the semantics it turns on (`DateTime -ne $null` is
// `$true`) are recorded above. An unrecognised condition FAILS rather than
// being skipped, so the model cannot silently stop describing the code.
//
// Office-free. Reads two files and runs no subprocess.
//
// Run: node --test ".github/extensions/office-canvas/test/unit/*.test.mjs"

import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";

import { blankComments } from "./ps-encoding-rule.mjs";
import { REPO } from "./tracked-files.mjs";

const PROBES = [
    "spikes/isolation/probes/probe-addressing.ps1",
    "spikes/isolation/probes/probe-bulk-read.ps1",
];

const FUNCTION = "Get-WordLiveness";

// Each condition the model understands, as a predicate over a world. An
// unlisted condition is a parse failure, not a pass.
const CONDITIONS = [
    [/^\$null -eq \$p$/, (w) => !w.exists],
    [/^\$p\.ProcessName -ne 'WINWORD'$/, (w) => w.name !== "WINWORD"],
    [/^\$null -eq \$expectedStart$/, (w) => w.expectedStart === null],
    [/^\$null -eq \$actual$/, (w) => w.actualStart === null],
    [/^\$actual -ne \$expectedStart$/, (w) => w.actualStart !== w.expectedStart],
];

/** The body of a `function name(...) { ... }`, by brace matching. */
const functionBody = (source, name) => {
    const start = source.indexOf(`function ${name}`);
    assert.notEqual(start, -1, `${name} not found -- re-point or remove this test`);
    const open = source.indexOf("{", start);
    assert.notEqual(open, -1, `${name} has no body`);
    let depth = 0;
    for (let i = open; i < source.length; i++) {
        if (source[i] === "{") depth++;
        else if (source[i] === "}" && --depth === 0) return source.slice(open + 1, i);
    }
    assert.fail(`${name} body is unbalanced`);
};

/**
 * The ordered guard sequence: every `if (cond) { return 'state' }`, plus the
 * fall-through `return 'state'`. The catch clause is excluded -- it is the
 * exception path, not a decision about evidence.
 */
const guardsOf = (body, where) => {
    const guards = [];
    let fallthrough = null;
    for (const raw of body.split("\n")) {
        const line = raw.trim();
        if (!line || line.includes("catch")) continue;

        const conditional = /^if \((.+)\)\s*\{\s*return '(\w+)'\s*\}$/.exec(line);
        if (conditional) {
            const [, condition, state] = conditional;
            const known = CONDITIONS.find(([pattern]) => pattern.test(condition.trim()));
            assert.ok(known, `${where}: unrecognised condition \`${condition.trim()}\` in ${FUNCTION} -- the model no longer describes the code, so update it deliberately`);
            guards.push({ test: known[1], state });
            continue;
        }

        const plain = /^return '(\w+)'$/.exec(line);
        if (plain) fallthrough = plain[1];
    }
    assert.ok(fallthrough, `${where}: ${FUNCTION} has no fall-through return`);
    return { guards, fallthrough };
};

/** Walk the extracted sequence: the first matching guard decides. */
const evaluate = ({ guards, fallthrough }, world) => {
    for (const guard of guards) if (guard.test(world)) return guard.state;
    return fallthrough;
};

const RUNNING = new Date("2026-09-01T10:00:00Z").getTime();
const OTHER = new Date("2026-01-01T00:00:00Z").getTime();

const CASES = [
    // The defect. A Word we created, running, whose start time was never read.
    ["A: running WINWORD, no recorded start time", { exists: true, name: "WINWORD", actualStart: RUNNING, expectedStart: null }, "unknown"],
    ["B: running WINWORD, identity confirmed", { exists: true, name: "WINWORD", actualStart: RUNNING, expectedStart: RUNNING }, "alive"],
    ["C: pid recycled onto another WINWORD", { exists: true, name: "WINWORD", actualStart: RUNNING, expectedStart: OTHER }, "gone"],
    ["D: pid recycled onto a non-Word process", { exists: true, name: "notepad", actualStart: RUNNING, expectedStart: RUNNING }, "gone"],
    // E pins the guard ORDER: sound absence must still conclude without a start
    // time. Moving the null-expectedStart guard above the absence checks breaks
    // exactly this case and nothing else.
    ["E: pid absent, no recorded start time", { exists: false, name: null, actualStart: null, expectedStart: null }, "gone"],
    ["F: pid absent, start time recorded", { exists: false, name: null, actualStart: null, expectedStart: RUNNING }, "gone"],
];

for (const probe of PROBES) {
    test(`${probe}: a missing start time is never read as proof of exit`, async () => {
        const source = blankComments(await readFile(path.join(REPO, probe), "utf8"));
        const sequence = guardsOf(functionBody(source, FUNCTION), probe);

        // Deliberately far below the real guard count. This exists only so that
        // a degenerate parse reports itself clearly instead of surfacing as
        // confusing case failures. It must NOT be set at or near the actual
        // number of guards: a missing guard has to be caught by the behavioural
        // cases below, otherwise this test degenerates into counting lines and
        // would "pass" any rewrite that kept the count while inverting a verdict.
        assert.ok(
            sequence.guards.length >= 2,
            `${probe}: only ${sequence.guards.length} guards recognised in ${FUNCTION}; a degenerate parse must not pass`,
        );

        // Every case is evaluated, and the failures reported together. Stopping
        // at the first would hide whether a regression is narrow or total --
        // and "it fires only on the case it is meant to" is a property of this
        // test that has to be visible when it fails.
        const failures = CASES.flatMap(([label, world, expected]) => {
            const actual = evaluate(sequence, world);
            return actual === expected ? [] : [`${label}: expected '${expected}', got '${actual}'`];
        });
        assert.deepEqual(failures, [], `${probe}: ${failures.length} of ${CASES.length} cases wrong\n  ${failures.join("\n  ")}`);
    });
}
