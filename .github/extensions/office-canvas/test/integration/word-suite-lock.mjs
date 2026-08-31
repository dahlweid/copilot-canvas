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

import { readFile, writeFile } from "node:fs/promises";
import { readFileSync, renameSync, statSync, unlinkSync, writeFileSync } from "node:fs";
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
 * Is this owner's claim dead?
 *
 * Two independent tests, and both are needed:
 *
 *   - the owner pid is gone: the run crashed, was killed, or was interrupted
 *     before its exit hook ran. This is the case the OS would have handled for
 *     a real handle, and the reason reclamation exists at all.
 *   - the owner is alive but has held past `maxHoldMs`: the backstop for pid
 *     reuse, where a dead run's number has been reissued to something unrelated
 *     that will never release. Set well beyond any real suite -- a wrong steal
 *     here costs a false red, which is the thing being fixed.
 */
function isStaleOwner(owner, { maxHoldMs, alive }) {
    if (!alive(owner.pid)) return true;
    return Number.isFinite(owner.startedAt) && Date.now() - owner.startedAt > maxHoldMs;
}

/** Parse a lock payload, or null if it is empty or unreadable. */
function parseOwner(raw) {
    try {
        const parsed = JSON.parse(raw);
        return parsed && typeof parsed === "object" ? parsed : null;
    } catch {
        return null;
    }
}

/**
 * Re-judges the lock and takes it away from a dead owner. Returns whether the
 * caller is now entitled to create its own.
 *
 * ## Why every step of this is synchronous
 *
 * The obvious shape -- read the owner, decide it is stale, then rename -- is
 * wrong, and wrong in the direction that matters. Each `await` between the read
 * and the rename is a point where another waiter runs, and a waiter that
 * *already* judged the old owner stale will happily rename away the lock of the
 * live holder that has since replaced it, then create its own on top. Two
 * holders, which is precisely the "two suites in Word at once" this file
 * exists to prevent. Measured: with the read and the rename separated by
 * awaits, 8 waiters racing one abandoned lock produced overlapping holds
 * every time (`word-suite-lock.test.mjs`, "two waiters racing ...").
 *
 * Synchronous calls make the judge-then-take pair atomic with respect to this
 * process's event loop, so a decision cannot go stale before it is acted on.
 *
 * ## Why the direction of the rename matters
 *
 * The lock path is the rename's **source**, moved aside to a name unique to
 * this caller. Exactly one waiter can move a given file, so the losers get
 * ENOENT and go back to waiting. The inverse -- each waiter renaming its own
 * temp *onto* the lock -- looks equivalent and is not: on Windows a rename onto
 * an existing target replaces it and succeeds for every caller, so every waiter
 * would believe it had won.
 *
 * ## The residual window, stated rather than hidden
 *
 * Across processes the pair is two syscalls rather than one, so a holder that
 * acquires between our read and our rename could still be moved. That needs a
 * third party to have emptied the path in the same instant, and it is closed
 * further by verifying what the rename actually captured and putting it back if
 * it turns out to be live. What is left is a genuine but far narrower race than
 * the one above, and it fails safe: the restore uses `wx`, so it can never
 * overwrite a lock someone else legitimately created.
 */
function takeIfStale(lockPath, { maxHoldMs, alive }) {
    let raw;
    try {
        raw = readFileSync(lockPath, "utf8");
    } catch {
        // Already gone. Nothing to reclaim; the caller retries its create.
        return false;
    }

    const owner = parseOwner(raw);
    if (owner) {
        if (!isStaleOwner(owner, { maxHoldMs, alive })) return false;
    } else {
        // An unreadable payload is not by itself abandonment: `writeFile`
        // creates then writes, so a waiter can catch the file empty
        // microseconds after a legitimate acquisition. Only age settles it.
        try {
            if (Date.now() - statSync(lockPath).mtimeMs <= maxHoldMs) return false;
        } catch {
            return false;
        }
    }

    const aside = `${lockPath}.${process.pid}.${randomUUID()}.stale`;
    try {
        renameSync(lockPath, aside);
    } catch {
        return false;
    }

    // We now hold `aside` exclusively, whatever it turned out to be. If it is
    // not what we judged, put it back rather than keeping a lock we were not
    // entitled to.
    let captured;
    try {
        captured = readFileSync(aside, "utf8");
    } catch {
        captured = "";
    }
    const capturedOwner = parseOwner(captured);
    if (capturedOwner && !isStaleOwner(capturedOwner, { maxHoldMs, alive })) {
        try {
            writeFileSync(lockPath, captured, { flag: "wx" });
        } catch {
            // Someone else already holds the path. Theirs stands; ours is not
            // restored, and we are not entitled either way.
        }
        try {
            unlinkSync(aside);
        } catch {}
        return false;
    }

    try {
        unlinkSync(aside);
    } catch {}
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

        // Advisory only: used to say who we are waiting for. The decision to
        // reclaim is re-taken synchronously inside `takeIfStale`, because a
        // judgement made out here would be stale by the time it was acted on.
        const owner = await readOwner(lockPath);
        if (takeIfStale(lockPath, { maxHoldMs, alive })) {
            const who = owner ? `pid ${owner.pid} (${owner.label})` : "an unreadable lock file";
            log(`[word-lock] reclaimed the lock from ${who}\n`);
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
