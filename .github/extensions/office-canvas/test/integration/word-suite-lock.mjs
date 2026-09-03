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

import { readFile } from "node:fs/promises";
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
 * Runs `fn` with the lock file's mutation baton held, or returns `null` if the
 * baton could not be taken.
 *
 * ## Why every mutation goes through here, not just reclamation
 *
 * Three versions of this file got this wrong, each time by assuming that
 * *narrowing* the race was the same as closing it. The measurements are in
 * `probe-suite-lock.mjs` part 4, 6 processes racing one abandoned lock under CPU
 * load, counting rounds in which two processes genuinely held at once:
 *
 * | shape | rounds with a double hold |
 * | --- | --- |
 * | read, judge, rename, create -- all `await`ed | 10/25, 8/25, 13/25 |
 * | the same, made synchronous | 1/25, 1/25, 0/25 |
 * | the same, plus a baton around *reclaimers only* | 0/25, 1/25, 0/25 |
 * | this: a baton around *every* mutation | 0/25 x 10 reps, 250 rounds |
 *
 * The last row is only worth reading because the instrument that produced it
 * still goes red: re-run in the same session, under the same load, against the
 * first shape, it reported 3/25, 4/25 and 5/25 with a worst overlap of 160 ms
 * against a 150 ms hold -- a whole concurrent hold, not clock skew.
 *
 * The middle two are the instructive ones. Making the sequence synchronous
 * closes the in-process interleaving and nothing else -- another process is not
 * on this event loop, and the OS will preempt between any two instructions. A
 * baton that only reclaimers take is better and still wrong, because the
 * surviving trace does not need two reclaimers:
 *
 * 1. R reads the lock and sees owner H, which is alive, so far so good.
 * 2. R is descheduled.
 * 3. H releases the lock and **exits**.
 * 4. W, an ordinary waiter, creates the now-free lock and holds it.
 * 5. R resumes and judges H by pid -- H is dead now, so the lock looks
 *    abandoned -- and installs over W. Two holders.
 *
 * R's judgement was true when it was made and false when it was acted on, and
 * no amount of re-reading fixes that, because the re-read has the same problem.
 * What fixes it is denying step 4: if creating the lock also requires the baton,
 * W cannot appear inside R's window at all, and R's rename-away is safe to undo
 * because nobody can have taken the path in the meantime.
 *
 * The baton is held for the microseconds of a decision, never for the duration
 * of a suite; the lock file remains the long-lived thing. So contention on it is
 * negligible, and a holder dying while holding it is a microsecond target rather
 * than the minutes-long one the lock itself presents.
 */
function withBaton(lockPath, fn) {
    const baton = `${lockPath}.baton`;
    const BATON_MAX_MS = 60_000;
    const mine = JSON.stringify({ pid: process.pid, at: Date.now() });

    try {
        writeFileSync(baton, mine, { flag: "wx" });
    } catch {
        // Either a decision is in progress -- microseconds, so just come back --
        // or a process died mid-decision and left it. Only age tells them apart,
        // and the gate is 60 s against a hold measured in microseconds, so a
        // live baton is never mistaken for an abandoned one except under a
        // four-order-of-magnitude stall.
        let abandoned = false;
        try {
            abandoned = Date.now() - statSync(baton).mtimeMs > BATON_MAX_MS;
        } catch {
            return null;
        }
        if (!abandoned) return null;
        try {
            // Exactly one waiter can move a given file; the losers get ENOENT.
            renameSync(baton, `${baton}.${process.pid}.${randomUUID()}.stale`);
            writeFileSync(baton, mine, { flag: "wx" });
        } catch {
            return null;
        }
    }

    try {
        return fn();
    } finally {
        try {
            unlinkSync(baton);
        } catch {}
    }
}

/**
 * Takes the lock if it is free or abandoned. Runs under the baton, so no other
 * process can create, reclaim or otherwise mutate the lock path while it works.
 */
function takeLocked(lockPath, payload, { maxHoldMs, alive }) {
    try {
        writeFileSync(lockPath, payload, { flag: "wx" });
        return true;
    } catch (err) {
        // EEXIST is the ordinary "someone holds it". It is not the only way that
        // presents on Windows: a create against a file another process has just
        // unlinked, while a handle to it is still open, is refused with EPERM
        // until the last handle closes -- the delete-pending state. Measured by
        // `probe-suite-lock.mjs` part 4 before this branch existed: 2 of 150
        // child acquisitions under CPU load died with an uncaught EPERM here,
        // which the probe scored as a patient waiter rather than a crash until
        // it was taught to tell them apart.
        //
        // Rethrowing a genuine permission fault would be no better. The caller
        // is a poll loop with a deadline, so returning "did not take it" retries
        // a transient refusal and lets a permanent one fall through to the
        // warn-and-proceed path, which is already the designed answer for a lock
        // that cannot be used.
        if (err.code !== "EEXIST" && err.code !== "EPERM" && err.code !== "EACCES" && err.code !== "EBUSY") {
            throw err;
        }
        if (err.code !== "EEXIST") return false;
    }

    let raw;
    try {
        raw = readFileSync(lockPath, "utf8");
    } catch {
        return false;
    }

    const owner = parseOwner(raw);
    if (owner) {
        if (!isStaleOwner(owner, { maxHoldMs, alive })) return false;
    } else {
        // An unreadable payload is not by itself abandonment: a create-then-write
        // can be caught with the file empty microseconds after a legitimate
        // acquisition. Only age settles it.
        try {
            if (Date.now() - statSync(lockPath).mtimeMs <= maxHoldMs) return false;
        } catch {
            return false;
        }
    }

    // Move it aside rather than deleting it, so a misjudgement is recoverable.
    // Under the baton nobody can take the path while it is briefly empty, which
    // is what makes the undo below sound rather than merely likely.
    const aside = `${lockPath}.${process.pid}.${randomUUID()}.stale`;
    try {
        renameSync(lockPath, aside);
    } catch {
        return false;
    }

    let captured;
    try {
        captured = readFileSync(aside, "utf8");
    } catch {
        captured = "";
    }
    const capturedOwner = parseOwner(captured);
    if (capturedOwner && !isStaleOwner(capturedOwner, { maxHoldMs, alive })) {
        // Not what we judged. Put it back exactly as it was.
        try {
            renameSync(aside, lockPath);
        } catch {}
        return false;
    }

    try {
        unlinkSync(aside);
    } catch {}
    try {
        writeFileSync(lockPath, payload, { flag: "wx" });
        return true;
    } catch {
        return false;
    }
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
    const grant = () => {
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
    };

    for (;;) {
        // Every mutation of the lock path goes through the baton, including the
        // ordinary create. That is the point: a create that bypassed it could
        // land inside a reclaimer's window, which is the trace that survived
        // the two previous versions of this file.
        const took = withBaton(lockPath, () => takeLocked(lockPath, payload(), { maxHoldMs, alive }));
        if (took) return grant();

        // Advisory only: used to say who we are waiting for.
        const owner = took === null ? null : await readOwner(lockPath);

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
