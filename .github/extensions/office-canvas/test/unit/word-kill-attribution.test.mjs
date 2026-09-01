// A WINWORD chosen by differencing a pid census is never killed.
//
// Probes under `spikes/isolation/` snapshot the `WINWORD` pid set, create a
// Word, snapshot again, and treat the difference as "the instance we own".
// Eleven of them then force-killed that set. This repository has measured the
// inference false twice:
//
//   * `spikes/isolation/probes/probe-init-attribution.ps1` differenced 2 new
//     pids for the 1 instance it created.
//   * A census control in the same probe saw 2 strangers' WINWORDs appear in a
//     40-second window with NOTHING launched. This is a shared machine and a
//     cold `New-Object` costs seconds, so the window routinely catches
//     strangers.
//
// So the difference is an inference over a POPULATION, and what a kill needs is
// a fact about ONE PROCESS. `Stop-Process -Force` on the wrong one destroys
// somebody's unsaved document with no prompt. Issue #136.
//
// NOT the same defect as #139, and the difference is measured. There, the fix
// was "never quit a COM-obtained PowerPoint at all", because PowerPoint is
// single-instance and `New-Object` ATTACHES to the user's running application.
// `spikes/isolation/probes/probe-newobject-attach.ps1` measured that `New-Object
// Word.Application` never attached to a running Word in either arm. Word is
// multi-instance, so an instance we created IS ours and quitting it is sound.
// Only the pid census is unsound. Importing #139's conclusion here would ban a
// safe operation and is how the two issues came to be split; this file
// deliberately says nothing about `.Quit()`.
//
// The rule this pins, in four parts:
//
//   1. Every `Stop-Process` under `spikes/isolation/` is either inside
//      `Stop-VerifiedWord`, or names a pid the KERNEL handed us -- from
//      `Start-Process -PassThru` or `CreateProcess`. Admission is by the
//      PROVENANCE OF THE VARIABLE, not by which file it sits in: a filename
//      allow-list would admit a differenced kill re-added to an admitted file.
//   2. No `Stop-Process` consumes a variable that came from a census
//      difference. This is the assertion that indicts the parent commit, and it
//      fails by naming the sites. It is a per-file taint set rather than a grep,
//      because the differenced set reached the kill through `foreach ($p in
//      $leaked)` and through `$pid_ = [int]$new[0]` -- neither of which looks
//      like a census at the kill site.
//   3. `Stop-VerifiedWord` still exists and still kills, so 1 cannot be
//      satisfied by deleting the one sanctioned kill.
//   4. `Stop-VerifiedWord` itself decides correctly, checked by executing a
//      model of its extracted decision sequence.
//
// Why a source assertion rather than a leak or liveness assertion: these are
// probes. CI cannot run them -- they need a licensed Word -- and nobody runs all
// of them. The only way to observe this defect at runtime is to have somebody's
// document destroyed by it.
//
// WHAT THIS CANNOT VERIFY, stated plainly:
//
//   * It does not run PowerShell. Test 4 assumes each extracted condition means
//     what the model says; an unrecognised condition FAILS rather than being
//     skipped, so the model cannot silently stop describing the code.
//   * It proves nothing about behaviour against a real Word. It does not show
//     that a kill succeeds, that a decline is reported, or that
//     `Stop-VerifiedWord` is ever reached.
//   * Tests 1 and 2 are SYNTACTIC. A pid passed through a function parameter,
//     stored in a hashtable, or returned from a helper this does not model would
//     escape both the provenance set and the taint set. Test 1 fails CLOSED on
//     the first of those (an unrecognised target is an offender) and test 2
//     fails OPEN (an unrecognised source is untainted), which is the safe
//     direction for each.
//   * It says nothing about `.Quit()`, `.Visible`, `.WindowState` or
//     `DisplayAlerts`. Those write to an instance reached through our own RCW,
//     not through a census, and the measurement above makes that instance ours.
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

const PATHSPEC = "spikes/isolation/*.ps1";
const KILLER = "Stop-VerifiedWord";
const KILLER_FILE = "spikes/isolation/probes/_common.ps1";

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

/**
 * The logical PowerShell statement starting at `index`.
 *
 * A physical line is not enough: every `Start-Process` in this tree spans
 * several, either by backtick continuation or by an open `@(` that runs to the
 * next page. Reading one line would miss `-PassThru` on four of the five worker
 * launches and silently demote them to unadmitted, which test 1 would then
 * report as defects.
 */
function statementAt(source, index) {
    const lines = source.slice(index).split("\n");
    let depth = 0;
    const taken = [];
    for (const line of lines) {
        taken.push(line);
        for (const ch of line) {
            if (ch === "(" || ch === "[" || ch === "{") depth++;
            else if (ch === ")" || ch === "]" || ch === "}") depth--;
        }
        const continued = /`\s*$/.test(line);
        if (depth <= 0 && !continued) break;
    }
    return taken.join("\n");
}

/**
 * Variables holding a pid the KERNEL returned for a process it made for us.
 *
 * Two routes, and only two:
 *   - `$X = Start-Process ... -PassThru`, where `$X.Id` is the created process.
 *   - `$X = $pi.pid`, the PROCESS_INFORMATION out-parameter of `CreateProcess`.
 *
 * Neither consults the population, so neither can name a stranger. The hwnd
 * route in spikes/isolation/probes/probe-hide.ps1 is sound for a third and
 * stronger reason -- it reads the pid off a window reached through our own RCW
 * -- but it is deliberately NOT modelled here, because that file now kills
 * through Stop-VerifiedWord and needs no exemption. Adding an unused admission
 * rule would widen this guard for nothing.
 */
function kernelMinted(source) {
    const minted = new Set();
    for (const m of source.matchAll(/\$(\w+)\s*=\s*Start-Process\b/g)) {
        if (/-PassThru\b/.test(statementAt(source, m.index))) minted.add(m[1]);
    }
    if (/CreateProcess/.test(source)) {
        for (const m of source.matchAll(/\$(\w+)\s*=\s*\$(\w+)\.pid\b/g)) minted.add(m[1]);
    }
    return minted;
}

/**
 * Variables holding, or derived from, a census DIFFERENCE.
 *
 * Two-stage, because the difference and the census are frequently not the same
 * statement. `spikes/isolation/probes/probe-quit-exit-gap.ps1` writes
 * `$after = Get-WordPids` and then `$new = @($after | Where-Object { $before
 * -notcontains $_ })`: the differencing line names no census call at all. A
 * one-stage seed that looked for `Get-Process` beside `-notcontains` missed it
 * entirely -- measured, by mutation, not by reading. So:
 *
 *   1. CENSUS variables: anything assigned from an enumeration of WINWORD
 *      processes, directly or through this tree's `Get-WordPids` helper.
 *   2. DIFFERENCED variables: an assignment using `-notcontains` or `-notin`
 *      that mentions a census variable or performs a census inline.
 *
 * Then propagated to a fixed point through the three forms that carried a
 * differenced set to a kill in the parent commit:
 *
 *   - `foreach ($p in $leaked)`, which is how eight of the eleven reached it;
 *   - `$pid_ = [int]$new[0]`, an index into the set;
 *   - `$b = $a`, plain aliasing.
 *
 * A grep at the kill site sees `$p` and `$pid_` and cannot tell either from a
 * kernel-minted pid, which is exactly why this is a taint set.
 */
function censusDifferenced(source) {
    const CENSUS = /Get-Process[^|\r\n]*WINWORD|WINWORD[^|\r\n]*Get-Process|Get-WordPids/i;
    const mentions = (text, names) => [...names].some((n) => new RegExp(`\\$${n}\\b`).test(text));

    const census = new Set();
    for (const m of source.matchAll(/\$(\w+)\s*=/g)) {
        if (CENSUS.test(statementAt(source, m.index))) census.add(m[1]);
    }

    const tainted = new Set();
    for (const m of source.matchAll(/\$(\w+)\s*=/g)) {
        const statement = statementAt(source, m.index);
        if (!/-not(?:contains|in)\b/.test(statement)) continue;
        if (CENSUS.test(statement) || mentions(statement, census)) tainted.add(m[1]);
    }

    for (let changed = true; changed; ) {
        changed = false;
        const add = (name) => {
            if (name && !tainted.has(name)) {
                tainted.add(name);
                changed = true;
            }
        };
        for (const m of source.matchAll(/foreach\s*\(\s*\$(\w+)\s+in\s+([^)]+)\)/gi)) {
            if (mentions(m[2], tainted)) add(m[1]);
        }
        for (const m of source.matchAll(/\$(\w+)\s*=\s*([^\r\n]+)/g)) {
            if (m[1] === "_") continue;
            // Only simple derivations: an index, a cast, or a bare alias. A
            // richer expression is left untainted deliberately -- test 2 fails
            // open, and test 1 fails closed over the same sites, so an escape
            // here is still caught there unless the pid is ALSO kernel-minted.
            const simple = /^\s*(?:\[[\w.]+\])?\s*\$(\w+)\s*(?:\[[^\]]*\])?\s*$/.exec(m[2]);
            if (simple && tainted.has(simple[1])) add(m[1]);
        }
    }
    return tainted;
}

/**
 * Blank the contents of single-line string literals, leaving the quotes and the
 * line/column geometry intact.
 *
 * Needed because a `Stop-Process` inside a message -- "worker did NOT exit after
 * Stop-Process" -- is prose, not a kill, and the scanner below reads text. It
 * was reported as an unattributed kill site the first time such a message was
 * written. Failing closed on prose is the right direction to fail, but it is
 * still wrong, and a maintainer who cannot phrase a diagnostic without tripping
 * the guard will eventually phrase the guard away instead.
 */
const blankStrings = (source) =>
    source.replace(/"[^"\r\n]*"|'[^'\r\n]*'/g, (m) => m[0] + " ".repeat(m.length - 2) + m[0]);

/** Every `Stop-Process` in `source`, with the root of the variable it targets. */
function killSites(source) {
    const sites = [];
    for (const m of blankStrings(source).matchAll(/Stop-Process\b[^\r\n]*/g)) {
        const target = /-Id\s+\$(\w+)/.exec(m[0]);
        sites.push({ index: m.index, text: m[0].trim(), target: target ? target[1] : null });
    }
    return sites;
}

// --- 1: every kill names a kernel-minted pid ----------------------------------

test("spikes/isolation: every Stop-Process names a pid the kernel minted", async (t) => {
    if (!(await gitAvailable())) return t.skip("no git checkout to enumerate");

    const files = await trackedFiles(PATHSPEC);
    // An empty list would make this pass without examining anything -- the same
    // shape as a glob that matches nothing and exits 0, which this repo has
    // measured once already.
    assert.ok(files.length > 0, `git ls-files found no files for '${PATHSPEC}'; the enumeration is broken, not the tree`);

    const offenders = [];
    let launchers = 0;
    for (const file of files) {
        // Comments are blanked first: most of the surviving mentions of the old
        // kills in this tree are prose explaining why they were removed, and a
        // scan over raw text would report the explanation as the bug.
        const source = blankComments(await readFile(path.join(REPO, file), "utf8"));
        const minted = kernelMinted(source);
        const sanctioned = file === KILLER_FILE ? functionBody(source, KILLER) : null;

        if (minted.size > 0) launchers++;

        for (const site of killSites(source)) {
            if (sanctioned && site.index >= sanctioned.start && site.index < sanctioned.end) continue;
            if (site.target && minted.has(site.target)) continue;
            offenders.push(`${file}:${lineOf(source, site.index)}  ${site.text}`);
        }
    }

    // The provenance scan is what ADMITS a kill, so a scan that silently stopped
    // matching would not fail open here -- every admitted site would become an
    // offender and this test would fail loudly. What it would do is stop being
    // an assertion about provenance and become one about the absence of kills.
    // This pins that it still recognises the kernel routes somewhere in the tree.
    // Like the taint assertion in test 2, this is a fact about the tree rather
    // than a count of defects, so fixing a site cannot turn it red.
    assert.ok(
        launchers > 0,
        "no file in the pathspec was found to mint a pid from Start-Process -PassThru or CreateProcess; the " +
            "provenance scan has stopped matching, so its admissions are no longer meaningful",
    );

    assert.deepEqual(
        offenders,
        [],
        `A Stop-Process under spikes/isolation/ is only sound on a pid the kernel returned for a process it made for ` +
            `us, or inside ${KILLER}, which re-verifies such a pid against the process name and the StartTime recorded ` +
            "at launch. A pid observed appearing in a census is not attribution: differencing over-reports (2 pids for " +
            "1 instance) and strangers' WINWORDs appear in the window unprompted (2 in 40 s). See issue #136 and " +
            `${KILLER_FILE}.\n  ` +
            offenders.join("\n  "),
    );
});

// --- 2: no kill consumes a census difference ----------------------------------

test("spikes/isolation: no Stop-Process consumes a census-differenced pid", async (t) => {
    if (!(await gitAvailable())) return t.skip("no git checkout to enumerate");

    const files = await trackedFiles(PATHSPEC);
    assert.ok(files.length > 0, `git ls-files found no files for '${PATHSPEC}'; the enumeration is broken, not the tree`);

    const offenders = [];
    let differencing = 0;
    for (const file of files) {
        const source = blankComments(await readFile(path.join(REPO, file), "utf8"));
        const tainted = censusDifferenced(source);
        if (tainted.size > 0) differencing++;

        for (const site of killSites(source)) {
            if (site.target && tainted.has(site.target)) {
                offenders.push(`${file}:${lineOf(source, site.index)}  $${site.target} came from a census difference`);
            }
        }
    }

    // The taint seeding is the whole test. If it stops matching -- a rename, a
    // reshaped Where-Object -- every file yields an empty set and this passes
    // over a tree full of differenced kills. Differencing is still WIDESPREAD in
    // this tree, and correctly so: it is how a probe REPORTS what appeared. So
    // "some file differences" is a fact about the tree, not a count of defects,
    // and pinning it cannot make a future fix turn CI red.
    assert.ok(
        differencing > 0,
        "no file in the pathspec was found to difference a WINWORD census; the taint seeding has stopped matching, " +
            "so this test is passing vacuously",
    );

    assert.deepEqual(
        offenders,
        [],
        "A census difference is an inference over a population: probe-init-attribution.ps1 measured 2 new pids for 1 " +
            "instance created, and its census control saw 2 strangers' WINWORDs appear in a 40 s window with nothing " +
            "launched. Killing that set destroys another session's unsaved document. Report the survivors instead -- " +
            `see Write-CensusSurvivors in ${KILLER_FILE} -- and keep the poll, which is the instrument. Issue #136.\n  ` +
            offenders.join("\n  "),
    );
});

// --- 3: the sanctioned kill still exists --------------------------------------

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

// --- 4: Stop-VerifiedWord decides correctly -----------------------------------

// Each condition the model understands, as a predicate over a world. An
// unlisted condition is a parse failure, not a pass.
const CONDITIONS = [
    [/^-not \$ProcessId$/, (w) => !w.processId],
    [/^-not \$p$/, (w) => !w.exists],
    [/^\$name -ne 'WINWORD'$/, (w) => w.name !== "WINWORD"],
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
 */
/**
 * Conditions that may only appear AFTER the kill: observations of what the
 * machine did, not decisions about whether to act.
 *
 * They are listed rather than ignored so their POSITION is asserted. One of
 * these appearing above the kill would be a `Get-Process` existence test
 * standing in for attribution -- precisely the reasoning #136 is about -- and
 * `decisionsOf` refuses it.
 */
const POST_KILL = [/^Get-Process -Id \$ProcessId -ErrorAction SilentlyContinue$/];

function decisionsOf(body) {
    const steps = [];
    let kills = 0;
    for (const raw of body.split("\n")) {
        const line = raw.trim();
        if (!line) continue;

        const conditional = /^if \((.+)\)\s*\{\s*return '([\w:-]+)'\s*\}$/.exec(line);
        if (conditional) {
            const [, condition, state] = conditional;
            if (POST_KILL.some((pattern) => pattern.test(condition.trim()))) {
                assert.equal(kills, 1, `\`${condition.trim()}\` in ${KILLER} is an observation of the kill's effect, but it appears BEFORE the kill, where it would be an existence test doing attribution's job (#136)`);
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
        // be skipped. `declined:unreadable-name` was silently dropped once
        // because the state pattern's character class had no hyphen, and the
        // model then reported that world as reaching the kill. A guard the model
        // cannot see is a guard the model reports as absent.
        assert.ok(
            !/^(?:if\s*\(|try\s*\{)/.test(line) || !/\breturn\b/.test(line),
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
const ok = { processId: 4242, exists: true, name: "WINWORD", handleReadable: true, nameReadable: true, startReadable: true, actualStart: OURS, expectedStart: OURS };

// Every world names the result it must produce. "Handled somehow" is not the
// property: a helper that declined EVERY case would satisfy a looser test while
// silently converting "the CreateProcess kills are kept and hardened" into
// "there are no kills" -- which is a different change, shipped under this one's
// name. The `killed` case is what makes that impossible.
const CASES = [
    ["A: no pid recorded at all", { ...ok, processId: 0 }, "declined:nopid"],
    ["B: pid already gone -- sound, and not a decline", { ...ok, exists: false }, "gone"],
    ["C: handle cannot be pinned", { ...ok, handleReadable: false }, "declined:handle"],
    ["D: pid recycled onto something that is not Word", { ...ok, name: "notepad" }, "declined:name"],
    ["F2: name unreadable -- must decline, not throw out of the caller's finally", { ...ok, nameReadable: false }, "declined:unreadable-name"],
    ["E: launcher never recorded a StartTime", { ...ok, expectedStart: null }, "declined:unverified"],
    ["F: StartTime unreadable", { ...ok, startReadable: false }, "declined:unreadable"],
    ["G: pid recycled onto another WINWORD", { ...ok, actualStart: THEIRS }, "declined:start"],
    // I pins the ORDER, and it is the case #143's guard was missing. Absence is
    // sound WITHOUT a recorded StartTime -- no process holds the pid, so nothing
    // we launched runs under it. Hoist the null-ExpectedStart guard above the
    // absence check and this is the ONLY case that changes: an already-gone
    // process comes back as a decline and the caller prints a leak warning about
    // nothing. Every other absent world here carries a recorded start time, so
    // without this one sitting in the intersection the mutation passes -- which
    // is measured, not hypothetical. It is reachable in the tree:
    // Get-WordStartTime returns $null when the launch-time read fails, and that
    // process may well have exited by teardown.
    ["I: pid absent AND no recorded StartTime", { ...ok, exists: false, expectedStart: null }, "gone"],
    // J is the mirror of I for the handle pin. The pin must sit ABOVE the
    // identity reads, so that a pid recycled between the reads cannot be killed
    // under a name checked on its predecessor. Moving the pin below the name
    // check leaves this world reporting a name verdict for a process whose
    // handle was never held.
    ["J: handle unpinnable AND recycled onto another name", { ...ok, handleReadable: false, name: "notepad" }, "declined:handle"],
    // K sits in the intersection that pins the kill's POSITION. Every other
    // world here would still be decided correctly by a helper that killed on its
    // first line, because the decline it expects is reachable below the kill as
    // dead code -- the walk simply never gets there. K is a world where the
    // guards are unanimous and the correct answer is a decline, so a hoisted
    // kill turns it into 'killed' and nothing else has to change.
    ["K: nothing about this pid verifies -- a hoisted kill would take it", { ...ok, name: "notepad", expectedStart: null, startReadable: false }, "declined:name"],
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
