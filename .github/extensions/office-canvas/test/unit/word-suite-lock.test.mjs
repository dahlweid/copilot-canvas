// The Word suite lock's *policy*, on the hosted runner.
//
// The lock exists because pid-set differencing reads the whole machine, so a
// second suite running concurrently is indistinguishable from a leak (issue
// #37, measured 17/18, 18/18, 17/18 on an unchanged tree). Its cross-process
// exclusion is a Windows filesystem property and is measured in
// spikes/isolation/probes/probe-suite-lock.mjs -- 8 concurrent acquirers, 0
// overlapping holds -- which is not something a hosted runner can attest.
//
// What *is* asserted here is the part that a source change breaks and a probe
// would not catch: when a lock may be taken from its owner, and when it may not.
// Both directions matter and they fail in opposite ways -- stealing too eagerly
// puts two suites in Word at once, which is the bug being fixed; never stealing
// wedges every future run behind a lock nobody will ever release.
//
// Office-free by construction: liveness is injected, so nothing here asks the
// operating system about a process. The files are real, because the exclusion
// primitive is `O_CREAT | O_EXCL` and a fake filesystem would assert on the
// fake instead.
//
// Run: node --test ".github/extensions/office-canvas/test/unit/*.test.mjs"

import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, rm, utimes, writeFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import { acquireWordSuiteLock, releaseWordSuiteLock, wordSuiteLockPath } from "../integration/word-suite-lock.mjs";

/** A scratch lock path, never the machine-wide one a live suite may hold. */
async function scratch() {
    const dir = await mkdtemp(path.join(tmpdir(), "word-suite-lock-test-"));
    return { lockPath: path.join(dir, "suite.lock"), cleanup: () => rm(dir, { recursive: true, force: true }) };
}

/** Options that keep a test from ever waiting on a real deadline. */
const fast = { timeoutMs: 0, pollMs: 1, onExit: false, log: () => {} };

test("an uncontended lock is acquired, and says who holds it", async () => {
    const { lockPath, cleanup } = await scratch();
    try {
        const handle = await acquireWordSuiteLock("unit", { ...fast, lockPath });
        assert.equal(handle.held, true);

        const owner = JSON.parse(await readFile(lockPath, "utf8"));
        assert.equal(owner.pid, process.pid);
        assert.equal(owner.label, "unit");
        assert.equal(owner.token, handle.token);
    } finally {
        await cleanup();
    }
});

test("releasing removes the lock, so the next run can take it", async () => {
    const { lockPath, cleanup } = await scratch();
    try {
        const first = await acquireWordSuiteLock("first", { ...fast, lockPath });
        assert.equal(first.release(), true);
        assert.equal(existsSync(lockPath), false);

        const second = await acquireWordSuiteLock("second", { ...fast, lockPath });
        assert.equal(second.held, true);
    } finally {
        await cleanup();
    }
});

test("a lock held by a living owner is not taken", async () => {
    // The assertion that has to be able to fail: if this ever returns held,
    // two suites are in Word at once and the gate is back where it started.
    const { lockPath, cleanup } = await scratch();
    try {
        await acquireWordSuiteLock("holder", { ...fast, lockPath });
        const warnings = [];
        const second = await acquireWordSuiteLock("waiter", {
            ...fast,
            lockPath,
            alive: () => true,
            log: (m) => warnings.push(m),
        });
        assert.equal(second.held, false);
        assert.equal(existsSync(lockPath), true, "the waiter deleted the holder's lock");
        assert.match(warnings.join(""), /running UNLOCKED/);
    } finally {
        await cleanup();
    }
});

test("a lock whose owner is gone is reclaimed", async () => {
    // The crashed, killed or Ctrl-C'd run. Nothing releases the file, and
    // without this every later run waits out its deadline and runs unlocked.
    const { lockPath, cleanup } = await scratch();
    try {
        await writeFile(
            lockPath,
            JSON.stringify({ token: "someone-else", pid: 4242, label: "crashed", startedAt: Date.now() }),
        );
        const handle = await acquireWordSuiteLock("survivor", { ...fast, lockPath, alive: () => false });
        assert.equal(handle.held, true);
        assert.equal(JSON.parse(await readFile(lockPath, "utf8")).label, "survivor");
    } finally {
        await cleanup();
    }
});

test("a living owner that has held past the cap is reclaimed anyway", async () => {
    // The pid-reuse backstop: a dead run's number reissued to something
    // unrelated looks alive forever, so age alone has to be able to break it.
    const { lockPath, cleanup } = await scratch();
    try {
        await writeFile(
            lockPath,
            JSON.stringify({ token: "stale", pid: 4242, label: "ancient", startedAt: Date.now() - 60_000 }),
        );
        const handle = await acquireWordSuiteLock("survivor", {
            ...fast,
            lockPath,
            alive: () => true,
            maxHoldMs: 1_000,
        });
        assert.equal(handle.held, true);
    } finally {
        await cleanup();
    }
});

test("two waiters racing for the same abandoned lock, and only one wins", async () => {
    // The reclamation race, which the *direction* of the rename decides.
    //
    // `stealFrom` renames the stale lock AWAY, to a name unique to the caller:
    // the path being raced on is the rename's SOURCE, so exactly one waiter
    // moves it and the losers get ENOENT. The inverse -- each waiter renaming
    // its own temp ONTO the lock -- looks equivalent and is not: on Windows a
    // rename onto an existing target replaces it and SUCCEEDS for every
    // caller, so every waiter would believe it had won.
    //
    // Nothing above pins that, which is why this exists: every other test here
    // passes with the steal written either way round.
    const { lockPath, cleanup } = await scratch();
    const ABANDONED = 999_999_998;
    try {
        await writeFile(
            lockPath,
            JSON.stringify({ token: "gone", pid: ABANDONED, label: "abandoned", startedAt: Date.now() }),
        );

        let holding = 0;
        let overlapped = false;
        const contend = async (n) => {
            const handle = await acquireWordSuiteLock(`waiter-${n}`, {
                lockPath,
                timeoutMs: 10_000,
                pollMs: 1,
                onExit: false,
                log: () => {},
                // Only the seeded owner is dead. A waiter that has won holds a
                // lock stamped with this process's pid, and must read as alive
                // or the others would reclaim it out from under it.
                alive: (pid) => pid !== ABANDONED,
            });
            if (!handle.held) return false;
            if (++holding > 1) overlapped = true;
            await new Promise((r) => setTimeout(r, 5));
            const released = handle.release();
            holding--;
            return released;
        };

        const outcomes = await Promise.all([1, 2, 3, 4, 5, 6, 7, 8].map(contend));

        assert.equal(overlapped, false, "two waiters held the abandoned lock at the same time");
        // Every waiter must also eventually get in. A steal that cannot be
        // completed wedges them all instead of double-booking them, which is
        // the other way the wrong rename direction shows up.
        assert.deepEqual(outcomes, [true, true, true, true, true, true, true, true]);
        assert.equal(existsSync(lockPath), false);
    } finally {
        await cleanup();
    }
});

test("a freshly created, not-yet-written lock is not mistaken for abandoned", async () => {
    // `writeFile` creates then writes, so a waiter can read the file empty    // microseconds after a legitimate acquisition. Treating unparseable as
    // abandoned would hand the lock to two holders at once.
    const { lockPath, cleanup } = await scratch();
    try {
        await writeFile(lockPath, "");
        const handle = await acquireWordSuiteLock("waiter", { ...fast, lockPath, maxHoldMs: 600_000 });
        assert.equal(handle.held, false);
        assert.equal(existsSync(lockPath), true);
    } finally {
        await cleanup();
    }
});

test("a lock file that has stayed unreadable past the cap is reclaimed", async () => {
    // The other half of the case above: a genuinely truncated or corrupted
    // lock must not wedge the gate forever.
    const { lockPath, cleanup } = await scratch();
    try {
        await writeFile(lockPath, "not json");
        const stale = new Date(Date.now() - 60_000);
        await utimes(lockPath, stale, stale);
        const handle = await acquireWordSuiteLock("survivor", { ...fast, lockPath, maxHoldMs: 1_000 });
        assert.equal(handle.held, true, "the stale unreadable lock was not reclaimed");
        assert.equal(JSON.parse(await readFile(lockPath, "utf8")).label, "survivor");
    } finally {
        await cleanup();
    }
});

test("release only removes a lock that is still ours", async () => {
    // After a reclamation the file belongs to someone else. Deleting it on the
    // way out would be the same class of mistake as killing an unattributed
    // Word: silent, and destructive to a run that did nothing wrong.
    const { lockPath, cleanup } = await scratch();
    try {
        const handle = await acquireWordSuiteLock("original", { ...fast, lockPath });
        await writeFile(
            lockPath,
            JSON.stringify({ token: "a-different-run", pid: process.pid, label: "usurper", startedAt: Date.now() }),
        );
        assert.equal(handle.release(), false);
        assert.equal(existsSync(lockPath), true);
        assert.equal(releaseWordSuiteLock(lockPath, "a-different-run"), true);
    } finally {
        await cleanup();
    }
});

test("releasing an absent lock is a no-op rather than a throw", async () => {
    const { lockPath, cleanup } = await scratch();
    try {
        assert.equal(releaseWordSuiteLock(lockPath, "whatever"), false);
    } finally {
        await cleanup();
    }
});

test("the lock is released when the suite process exits", async () => {
    // Every suite ends in `process.exit()`, so an async release would be
    // scheduled and never run. The hook is the only thing that covers a suite
    // exiting non-zero, which is exactly when a stale lock is worst.
    const { lockPath, cleanup } = await scratch();
    const before = process.listenerCount("exit");
    try {
        const handle = await acquireWordSuiteLock("exiting", { ...fast, lockPath, onExit: true });
        assert.equal(process.listenerCount("exit"), before + 1);
        handle.detach();
        assert.equal(process.listenerCount("exit"), before);
    } finally {
        await cleanup();
    }
});

test("the default lock path is one shared location, not a per-worktree one", async () => {
    // A lock under the repo would be per-worktree, and sibling worktrees are
    // precisely the sessions that contend. It has to be outside the tree.
    const previous = process.env.OFFICE_CANVAS_WORD_LOCK_PATH;
    delete process.env.OFFICE_CANVAS_WORD_LOCK_PATH;
    try {
        const resolved = wordSuiteLockPath();
        assert.equal(resolved, path.join(tmpdir(), "office-canvas-word-suite.lock"));
        assert.equal(resolved.startsWith(process.cwd()), false);
    } finally {
        if (previous === undefined) delete process.env.OFFICE_CANVAS_WORD_LOCK_PATH;
        else process.env.OFFICE_CANVAS_WORD_LOCK_PATH = previous;
    }
});
