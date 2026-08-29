// Word process bookkeeping, shared by every integration suite.
//
// The assertion is **PID-set differencing**, not "no WINWORD.EXE exists". A
// global emptiness check is not concurrency-safe: another session, another
// worktree, or the developer's own open Word all fail it, and the failure looks
// like ours. Differencing asserts what we actually mean -- this test left
// nothing behind -- and is a strictly better assertion even when run alone.
//
// This file used to add: "the same rule makes killing safe: only a PID that
// appeared after we started can be ours". That is false, and it was load-bearing
// for a `Stop-Process -Force`. Measured by
// `spikes/isolation/probes/probe-word-ownership.ps1`: with one concurrent Word
// started from a separate process, differencing reported **2 new pids for the 1
// instance we created**. Parentage carries no signal either -- every WINWORD
// here parents to the DCOM launcher (`svchost`, pid 1684), not to its creator.
//
// Asserting and killing need different standards of evidence, and serving both
// from one rule was the mistake:
//
//   an over-broad assertion fails loudly; an over-broad kill destroys silently.
//   Only the destructive operation needs provable attribution.
//
// So the assertion keeps differencing -- over-reporting costs a visible false
// failure, and it catches leaks attribution cannot see, such as the second
// WINWORD that L2 measured `ProtectedViewWindows.Open` spawning without the
// bridge ever holding a handle to it. Killing goes through a ledger of pids the
// *host itself* reported as owned, and touches nothing else however suspicious.
//
// Attribution is not yet perfect at its source: `Initialize-Word` in
// `word-host.ps1` derives `ownedPid` by this same differencing, so a ledger
// entry inherits that. A sound route exists and is measured in
// `spikes/isolation/probes/probe-word-ownership-hwnd.ps1` -- once a document is
// open, `Application.ActiveWindow.Hwnd` plus `GetWindowThreadProcessId` yields a
// pid that is ours by construction. Wiring it into the host is issue #25's
// second half; that file belongs to an open PR and is not this one's to change.

import { execFile, execFileSync } from "node:child_process";
import { promisify } from "node:util";
import assert from "node:assert/strict";

const execFileAsync = promisify(execFile);

/** PIDs of every WINWORD.EXE currently running, ours or not. */
export async function wordPids() {
    const { stdout } = await execFileAsync("powershell.exe", [
        "-NoProfile",
        "-NonInteractive",
        "-Command",
        "@(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id) -join ','",
    ]);
    return stdout
        .trim()
        .split(",")
        .filter(Boolean)
        .map(Number);
}

/** WINWORD PIDs that were not running when the test started. */
export async function newWordPids(pidsBefore, list = wordPids) {
    return (await list()).filter((pid) => !pidsBefore.includes(pid));
}

/**
 * Collects the pids a `WordHost` reports as owned. Pass `record` as the host's
 * `onOwnedPid`.
 *
 * A host may own several pids over one run, because it restarts Word after a
 * crash, so this accumulates rather than replaces. A pid that has since exited
 * stays recorded: killing a dead pid is a no-op, while forgetting a live one is
 * a leak.
 */
export function ownedWordLedger() {
    const owned = new Set();
    return {
        record(pid) {
            if (Number.isInteger(pid) && pid > 0) owned.add(pid);
        },
        pids() {
            return [...owned];
        },
    };
}

/**
 * Asserts this test leaked no Word process.
 *
 * Word exits asynchronously after `Quit`: the call returns in ~120 ms but the
 * process lingers. Measured on this machine, Quit-to-exit was 2.7 s, 6.1 s and
 * 2.7 s over three runs on an otherwise idle machine -- so a flat settle short
 * of that is a coin toss, and a flat settle long enough to be safe wastes that
 * long on every green run. We poll instead: it returns as soon as the process
 * is actually gone, and only spends the deadline when there is something to
 * report.
 *
 * The deadline is 90 s rather than the ~6 s the idle measurements suggest,
 * because exit latency is not bounded by them. With a second session driving
 * Word concurrently, one run here had a Word survive a 30 s poll and then exit
 * on its own -- a false failure, and worse, one that would have surfaced in
 * whichever PR happened to run while another was busy. Word's shutdown touches
 * per-user state (Normal.dotm and friends) that concurrent instances contend
 * for, so the tail is long and load-dependent. Polling makes the generous
 * deadline free on green runs, so there is no reason to shave it.
 *
 * Pass `ledger` to have a failure separate the pids the host told us it owned
 * from the ones that merely appeared. Both still fail; the split exists so a
 * human reading a red run knows whether to look at our teardown or at what else
 * was running.
 */
export async function assertNoLeakedWord(
    pidsBefore,
    { timeoutMs = 90000, intervalMs = 250, ledger = null, list = wordPids } = {},
) {
    const started = Date.now();
    const deadline = started + timeoutMs;
    let leaked = await newWordPids(pidsBefore, list);
    while (leaked.length > 0 && Date.now() < deadline) {
        await new Promise((resolve) => setTimeout(resolve, intervalMs));
        leaked = await newWordPids(pidsBefore, list);
    }
    const waited = ((Date.now() - started) / 1000).toFixed(1);

    const ownedPids = ledger ? ledger.pids() : [];
    const ours = leaked.filter((pid) => ownedPids.includes(pid));
    const unattributed = leaked.filter((pid) => !ownedPids.includes(pid));

    let detail = "";
    if (ours.length > 0) {
        detail += ` Reported as owned by this test's host: ${ours.join(", ")}.`;
    }
    if (unattributed.length > 0) {
        // Deliberately not "so they are ours". That claim was wrong, and it sent
        // readers looking for a teardown bug that need not exist.
        detail +=
            ` Appeared during this test but were never reported as owned: ${unattributed.join(", ")}` +
            " -- either a Word the bridge spawned without holding a handle to, or another session's.";
    }

    assert.deepEqual(
        leaked,
        [],
        `Word processes still running ${waited}s after teardown: ${leaked.join(", ")}.${detail}`,
    );
}

/**
 * Kills the Word processes the host reported owning, for the tests that
 * deliberately simulate a crash.
 *
 * Only ledger pids are touched. A Word that appeared during this run but was
 * never reported as owned is left alone however suspicious it looks: it may be
 * another session's, and a wrong kill is silent and unrecoverable whereas a
 * missed one surfaces as a loud leak assertion.
 *
 * Returns the pids a kill was actually issued for, which is narrower than the
 * pids we looked at. Two of the three outcomes below are *not* kills, and both
 * are ordinary rather than exceptional: by the time teardown runs, the bridge's
 * own exit hook has usually reaped the pid already (`gone`), and a reaped pid's
 * number can be reissued to something unrelated (`notword`). Neither raises, so
 * the outcome has to be reported out of PowerShell and read here -- inferring it
 * from "the command did not throw" counts both as kills.
 *
 * "Issued" is the honest verb: `Process.Kill()` requests termination and returns
 * without waiting, and measured `Quit()`-to-exit here is 2.7-6.1 s and rises
 * under load. Exit is confirmed by `assertNoLeakedWord` polling to a deadline,
 * not by this function.
 */
export function killOwnedWord(ledger, { exec = execFileSync } = {}) {
    const killed = [];
    for (const pid of ledger.pids()) {
        try {
            const outcome = String(
                exec("powershell.exe", [
                    "-NoProfile",
                    "-NonInteractive",
                    "-Command",
                    // Name-checked because pids are reused.
                    `$p = Get-Process -Id ${pid} -ErrorAction SilentlyContinue; ` +
                        "if ($p -and $p.ProcessName -eq 'WINWORD') { $p.Kill(); 'killed' } " +
                        "elseif ($p) { 'notword' } else { 'gone' }",
                ]) ?? "",
            ).trim();
            if (outcome === "killed") killed.push(pid);
        } catch {
            /* already gone, or refused -- either way nothing was killed */
        }
    }
    return killed;
}
