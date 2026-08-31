// A machine-wide mutex for the Word integration suites.
//
// ## What it is for
//
// Every suite here asserts on **PID-set differencing**: a WINWORD.EXE that was
// not in the census taken at startup is a finding. That is the right assertion
// (see the header of `word-pids.mjs` for why breadth is deliberate), but it
// reads the whole machine, so *any* concurrent producer of Word processes is
// indistinguishable from a defect in the code under test.
//
// Measured, issue #37: `read-smoke` scored 17/18, 18/18, 17/18 on an unchanged
// tree, with sibling sessions driving Word throughout. A gate that reds on most
// runs of a clean tree has stopped carrying information -- the trained response
// is to re-run until green, which is also the response to a genuine intermittent
// leak.
//
// So this removes the condition rather than narrowing the detector: hold one
// lock for the duration of a suite and two suites cannot interleave. Nothing
// about any assertion changes, and in particular the `ProtectedViewWindows`
// second-instance case that differencing exists to catch is untouched.
//
// ## What it does not fix, stated because the gap is the interesting part
//
// It serialises *lock-aware* runs. A Word started by anything else -- a probe, a
// sibling session on a branch that predates this file, the user's own Word, or
// the external producer #34 measured (two new WINWORDs in 40 s with nothing
// launched) -- is still inside the window. Serialisation is necessary here, and
// it is not sufficient. A red leak assertion is still worth reading rather than
// re-running by reflex.
//
// ## Why a lock file, and not the obvious alternatives
//
// *Not* an exclusively-opened file handle. That is the usual Windows answer --
// let the OS drop the lock when the process dies -- and it does not work from
// Node: Node opens files with `FILE_SHARE_READ|WRITE|DELETE` and offers no way
// to ask for anything narrower, which is the same fact that makes `readFile` in
// a test prove nothing about exclusivity. A handle we cannot make exclusive is
// not a lock.
//
// *Not* a named `System.Threading.Mutex`. It is the right primitive and it would
// need a PowerShell process alive for the whole suite purely to hold it, which
// adds a second process whose death is a failure mode of its own.
//
// So: `writeFile(..., { flag: "wx" })`, which is `O_CREAT | O_EXCL` and reaches
// Windows as `CREATE_NEW` -- a single atomic create-and-write. Measured with 8
// concurrent acquirers in `probe-suite-lock.mjs`: exactly one winner per round,
// and no two holds overlapped across 8 rounds.
//
// Losing the OS's automatic release is the cost, and it is paid explicitly
// below: a lock whose owner is gone is reclaimed.

import { readFile, rename, rm, stat, writeFile } from "node:fs/promises";
import { readFileSync, unlinkSync } from "node:fs";
import { hostname, tmpdir } from "node:os";
import { randomUUID } from "node:crypto";
import path from "node:path";

/**
 * Where the lock lives.
 *
 * `tmpdir()` is per-*user*, not per-machine, and that is the correct scope
 * rather than a compromise: the sessions that contend here all run as the same
 * user, and what they contend over -- Word's per-user state, `Normal.dotm` and
 * friends -- is per-user too. A path under `ProgramData` would be wider than the
 * resource and would need privileges CI does not have.
 *
 * Overridable so a probe or a unit test can exercise the real code against a
 * scratch path instead of the lock a live suite may be holding.
 */
export const wordSuiteLockPath = () =>
    process.env.OFFICE_CANVAS_WORD_LOCK_PATH || path.join(tmpdir(), "office-canvas-word-suite.lock");

/** Does this pid exist? Signal 0 tests existence without delivering anything. */
function pidAlive(pid) {
    if (!Number.isInteger(pid) || pid <= 0) return false;
    try {
        process.kill(pid, 0);
        return true;
    } catch (err) {
        // EPERM means the process exists and is not ours to signal. Reading that
        // as "gone" would reclaim a live owner's lock, so it counts as alive.
        return err.code === "EPERM";
    }
}

/** The lock's payload, or null if it is absent, empty or unparseable. */
async function readOwner(lockPath) {
    try {
        const parsed = JSON.parse(await readFile(lockPath, "utf8"));
        return parsed && typeof parsed === "object" ? parsed : null;
    } catch {
        return null;
    }
}

/**
 * May this lock be taken from its owner?
 *
 * Three reclaimable states, and one deliberately non-reclaimable one:
 *
 *   - the owner pid is gone: the run crashed, was killed, or was interrupted
 *     before its exit hook ran. This is the case the OS would have handled for
 *     a real handle and the reason this function exists at all.
 *   - the owner is alive but has held past `maxHoldMs`: the backstop for pid
 *     reuse, where a dead run's number has been reissued to something unrelated
 *     that will never release. Set well beyond any real suite -- a wrong steal
 *     here costs a false red, which is what we are trying to stop.
 *   - the payload is unreadable *and* the file has not been touched for
 *     `maxHoldMs`: a half-written or corrupted lock, which would otherwise wedge
 *     every future run forever.
 *
 * A payload that is merely unreadable *now* is not stale. `writeFile` creates
 * then writes, so a reader can catch the file empty microseconds after a
 * legitimate acquisition; treating that as abandonment would hand the lock to
 * two holders at once. It will parse on the next poll.
 */
async function reclaimable(lockPath, owner, { maxHoldMs, alive }) {
    if (!owner) {
        try {
            const info = await stat(lockPath);
            return Date.now() - info.mtimeMs > maxHoldMs;
        } catch {
            // Gone between the failed read and here -- released, not stale.
            return false;
        }
    }
    if (!alive(owner.pid)) return true;
    return Number.isFinite(owner.startedAt) && Date.now() - owner.startedAt > maxHoldMs;
}

/**
 * Takes the lock away from a dead owner, atomically enough that two waiters
 * cannot both succeed.
 *
 * `rm` then `writeFile` would let two waiters interleave -- A removes, A
 * creates, B removes *A's fresh lock*, B creates, and both run. Renaming is the
 * fix: the second rename of the same path fails with ENOENT, so exactly one
 * waiter is entitled to try the create, and the loser goes back to waiting.
 */
async function stealFrom(lockPath) {
    const aside = `${lockPath}.${process.pid}.${Date.now()}.stale`;
    try {
        await rename(lockPath, aside);
    } catch {
        return false;
    }
    await rm(aside, { force: true }).catch(() => {});
    return true;
}

/**
 * Deletes the lock, but only while it is still ours.
 *
 * Synchronous because the reliable place to call it is `process.on("exit")`:
 * every suite here ends in `process.exit()`, so an async release would be
 * scheduled and then never run. The token check is what makes a release safe
 * after a reclamation -- if someone took the lock from us, the file we are
 * looking at is theirs and must not be removed.
 */
export function releaseWordSuiteLock(lockPath, token) {
    try {
        if (JSON.parse(readFileSync(lockPath, "utf8")).token !== token) return false;
    } catch {
        return false;
    }
    try {
        unlinkSync(lockPath);
        return true;
    } catch {
        return false;
    }
}

/**
 * Acquires the Word suite lock, and arranges for it to be released when this
 * process exits.
 *
 * Call it *before* the `wordPids()` census: the census has to be taken while we
 * are already exclusive, or a neighbour's Word can appear between the two and
 * land in the differencing window after all.
 *
 * ## On the timeout, which does not fail the run
 *
 * Waiting out is reported loudly and then the suite proceeds **unlocked**. That
 * is deliberate. The failure this file exists to remove is a gate going red for
 * a reason unrelated to the tree under test; converting "a neighbour was busy
 * for 30 minutes" into a red suite reproduces exactly that, in a new place. Run
 * unlocked and the worst case is the behaviour we had before this file, with a
 * line of output saying so.
 */
export async function acquireWordSuiteLock(
    label,
    {
        lockPath = wordSuiteLockPath(),
        timeoutMs = 1_800_000,
        pollMs = 500,
        maxHoldMs = 2_700_000,
        alive = pidAlive,
        log = (m) => process.stderr.write(m),
        onExit = true,
    } = {},
) {
    const token = randomUUID();
    const startedAt = Date.now();
    const deadline = startedAt + timeoutMs;
    const payload = () =>
        JSON.stringify({ token, pid: process.pid, label, host: hostname(), startedAt: Date.now() });

    let announced = false;
    for (;;) {
        try {
            await writeFile(lockPath, payload(), { flag: "wx" });
            const waitedMs = Date.now() - startedAt;
            if (announced) log(`[word-lock] acquired after ${(waitedMs / 1000).toFixed(1)}s\n`);
            const handle = {
                held: true,
                token,
                lockPath,
                waitedMs,
                release: () => releaseWordSuiteLock(lockPath, token),
            };
            if (onExit) {
                const hook = () => releaseWordSuiteLock(lockPath, token);
                process.on("exit", hook);
                handle.detach = () => process.off("exit", hook);
            }
            return handle;
        } catch (err) {
            if (err.code !== "EEXIST") throw err;
        }

        const owner = await readOwner(lockPath);
        if (await reclaimable(lockPath, owner, { maxHoldMs, alive })) {
            const who = owner ? `pid ${owner.pid} (${owner.label})` : "an unreadable lock file";
            if (await stealFrom(lockPath)) log(`[word-lock] reclaimed the lock from ${who}\n`);
            continue;
        }

        if (!announced) {
            announced = true;
            const who = owner ? `${owner.label} (pid ${owner.pid})` : "another run";
            log(`[word-lock] ${label} is waiting for ${who} to finish with Word...\n`);
        }
        if (Date.now() >= deadline) {
            log(
                `[word-lock] WARNING: gave up waiting after ${(timeoutMs / 1000).toFixed(0)}s and is ` +
                    "running UNLOCKED. Word process assertions in this run may reflect another " +
                    "session's activity rather than this tree.\n",
            );
            return { held: false, token: null, lockPath, waitedMs: Date.now() - startedAt, release: () => false };
        }
        await new Promise((resolve) => setTimeout(resolve, pollMs));
    }
}
