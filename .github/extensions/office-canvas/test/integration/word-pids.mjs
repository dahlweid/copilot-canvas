// Word process bookkeeping, shared by every integration suite.
//
// The assertion is **PID-set differencing**, not "no WINWORD.EXE exists". A
// global emptiness check is not concurrency-safe: another session, another
// worktree, or the developer's own open Word all fail it, and the failure looks
// like ours. Differencing asserts what we actually mean -- this test left
// nothing behind -- and is a strictly better assertion even when run alone.
//
// The same rule makes killing safe: only a PID that appeared after we started
// can be ours, and nothing else may be touched.

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
export async function newWordPids(pidsBefore) {
    return (await wordPids()).filter((pid) => !pidsBefore.includes(pid));
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
 */
export async function assertNoLeakedWord(pidsBefore, { timeoutMs = 90000, intervalMs = 250 } = {}) {
    const started = Date.now();
    const deadline = started + timeoutMs;
    let leaked = await newWordPids(pidsBefore);
    while (leaked.length > 0 && Date.now() < deadline) {
        await new Promise((resolve) => setTimeout(resolve, intervalMs));
        leaked = await newWordPids(pidsBefore);
    }
    const waited = ((Date.now() - started) / 1000).toFixed(1);
    assert.deepEqual(
        leaked,
        [],
        `Word processes still running ${waited}s after teardown: ${leaked.join(", ")}. ` +
            `These started during this test, so they are ours.`,
    );
}

/**
 * Kills the Word processes this test started, for the tests that deliberately
 * simulate a crash. Processes that predate the test are never touched -- they
 * belong to someone else.
 */
export function killNewWord(pidsBefore, currentPids) {
    for (const pid of currentPids.filter((p) => !pidsBefore.includes(p))) {
        try {
            execFileSync("powershell.exe", [
                "-NoProfile",
                "-NonInteractive",
                "-Command",
                `Stop-Process -Id ${pid} -Force -ErrorAction SilentlyContinue`,
            ]);
        } catch {
            /* already gone */
        }
    }
}
