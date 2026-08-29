// Killing a Word process needs provable ownership; asserting a leak does not.
//
// The defect this guards against shipped for months: `killNewWord` force-killed
// every WINWORD that appeared during a run, on the stated grounds that "only a
// PID that appeared after we started can be ours". Measured false by
// spikes/isolation/probes/probe-word-ownership.ps1 -- with one concurrent Word
// started from a separate process, differencing reported 2 new pids for the 1
// instance we created. `host-smoke.mjs` calls this, so a run could
// `Stop-Process -Force` another session's deliberately parked Word.
//
// Office-free by construction: process listing and process killing are both
// injected, so what is asserted here is the *policy* -- which pids are eligible
// to be killed -- rather than any Windows behaviour. The policy is the part that
// was wrong, and a source change is what would break it again, so it belongs on
// the hosted runner.
//
// Run: node --test ".github/extensions/office-canvas/test/unit/*.test.mjs"

import test from "node:test";
import assert from "node:assert/strict";
import {
    assertNoLeakedWord,
    killOwnedWord,
    newWordPids,
    ownedWordLedger,
} from "../integration/word-pids.mjs";

/** Captures kill attempts instead of making them. Returns the pids targeted. */
function recordingExec(targeted) {
    return (_exe, args) => {
        const command = args[args.length - 1];
        const match = /Get-Process -Id (\d+)/.exec(command);
        if (match) targeted.push(Number(match[1]));
    };
}

test("the ledger records what the host reports, and nothing else", () => {
    const ledger = ownedWordLedger();
    ledger.record(101);
    ledger.record(102);
    ledger.record(101); // a restart can re-report the same pid
    assert.deepEqual(ledger.pids(), [101, 102]);
});

test("the ledger ignores the shapes a host reports when it owns nothing", () => {
    // `Cmd-Ping` returns `ownedPid: null` when Word was attached rather than
    // created. A null must not become a kill target, and `Number.isInteger`
    // rejects it along with the other non-pids.
    const ledger = ownedWordLedger();
    for (const value of [null, undefined, 0, -1, "103", 1.5, NaN]) ledger.record(value);
    assert.deepEqual(ledger.pids(), []);
});

test("a pid the host owns is killed", () => {
    const targeted = [];
    const ledger = ownedWordLedger();
    ledger.record(201);
    killOwnedWord(ledger, { exec: recordingExec(targeted) });
    assert.deepEqual(targeted, [201]);
});

test("a pid that merely appeared during the run is never killed", () => {
    // The whole point. 301 is ours; 302 and 303 appeared during the run and are
    // indistinguishable from ours by differencing -- which is exactly what the
    // old implementation used to authorize the kill.
    const targeted = [];
    const ledger = ownedWordLedger();
    ledger.record(301);

    const pidsBefore = [1];
    const currentPids = [1, 301, 302, 303];
    const appearedDuringRun = currentPids.filter((pid) => !pidsBefore.includes(pid));
    assert.deepEqual(appearedDuringRun, [301, 302, 303], "differencing cannot separate them");

    killOwnedWord(ledger, { exec: recordingExec(targeted) });
    assert.deepEqual(targeted, [301], "only the pid the host reported owning may be killed");
});

test("killing an already-dead pid is not an error", () => {
    const ledger = ownedWordLedger();
    ledger.record(401);
    const throwing = () => {
        throw new Error("No process with that Id");
    };
    assert.deepEqual(killOwnedWord(ledger, { exec: throwing }), []);
});

test("the kill is name-checked, because pids are reused", () => {
    // A recorded pid can exit and its number be reissued to something else
    // before teardown runs. Killing by number alone would end an unrelated
    // process; the command must confirm it is still a WINWORD.
    const commands = [];
    const ledger = ownedWordLedger();
    ledger.record(501);
    killOwnedWord(ledger, { exec: (_exe, args) => commands.push(args[args.length - 1]) });
    assert.match(commands[0], /ProcessName -eq 'WINWORD'/);
});

test("nothing new means no leak", async () => {
    const list = async () => [1, 2];
    await assertNoLeakedWord([1, 2], { list, timeoutMs: 0 });
});

test("an unattributed process still fails the assertion", async () => {
    // Deliberately broader than the kill. An over-broad assertion costs a loud
    // false failure; an over-broad kill destroys silently. It also catches the
    // leak attribution cannot see -- ProtectedViewWindows.Open spawns a second
    // WINWORD the bridge never holds a handle to, so it is never reported owned.
    const list = async () => [1, 999];
    await assert.rejects(
        () => assertNoLeakedWord([1], { list, timeoutMs: 0, ledger: ownedWordLedger() }),
        /999/,
    );
});

test("the failure separates owned pids from ones that merely appeared", async () => {
    const ledger = ownedWordLedger();
    ledger.record(700);
    const list = async () => [1, 700, 800];

    const err = await assertNoLeakedWord([1], { list, timeoutMs: 0, ledger }).then(
        () => null,
        (e) => e,
    );
    assert.ok(err, "expected a leak to be reported");
    assert.match(err.message, /Reported as owned by this test's host: 700/);
    assert.match(err.message, /never reported as owned: 800/);
});

test("the message no longer claims every new process is ours", async () => {
    // The old message asserted a cause the code never established, which is the
    // error-message rule this repo keeps re-learning. It sent readers looking
    // for a teardown bug when the process belonged to another session.
    const list = async () => [1, 900];
    const err = await assertNoLeakedWord([1], { list, timeoutMs: 0 }).then(
        () => null,
        (e) => e,
    );
    assert.ok(err);
    assert.doesNotMatch(err.message, /so they are ours/);
});

test("newWordPids reports only what appeared", async () => {
    assert.deepEqual(await newWordPids([1, 2], async () => [2, 3]), [3]);
});
