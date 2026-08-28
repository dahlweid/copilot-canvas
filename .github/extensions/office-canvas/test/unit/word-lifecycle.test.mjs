// When Word is allowed to go away. Office-free.
//
// The race this guards against is not hypothetical: the agent issues tool calls
// in parallel, and a structure read can run for minutes because it may include
// a cold Word start. The first call to finish used to arm the shutdown timer
// while the second was still running, and the timer then disposed the host
// underneath it -- surfacing as "Word is unavailable" on a machine where Word
// was fine.
//
// Run: node --test ".github/extensions/office-canvas/test/unit/*.test.mjs"

import test from "node:test";
import assert from "node:assert/strict";

import { createIdleShutdown } from "../../src/word-lifecycle.mjs";

/**
 * A controllable clock. Real timers would make these tests either slow or
 * flaky, and the thing under test is ordering, not duration.
 */
function fakeTimers() {
    let next = 1;
    const scheduled = new Map();
    return {
        setTimer(fn, ms) {
            const id = next++;
            scheduled.set(id, { fn, ms });
            return id;
        },
        clearTimer(id) {
            scheduled.delete(id);
        },
        get pending() {
            return scheduled.size;
        },
        /** Fires every timer currently armed, as the event loop would. */
        fire() {
            const due = [...scheduled.entries()];
            scheduled.clear();
            for (const [, entry] of due) entry.fn();
            return due.length;
        },
    };
}

function harness({ displaying = false } = {}) {
    const timers = fakeTimers();
    const state = { disposals: 0, displaying };
    const lifecycle = createIdleShutdown({
        idleMs: 60_000,
        isDisplaying: () => state.displaying,
        dispose: () => {
            state.disposals += 1;
        },
        setTimer: timers.setTimer,
        clearTimer: timers.clearTimer,
    });
    return { lifecycle, timers, state };
}

test("work that finishes with nothing else running arms the shutdown", async () => {
    const { lifecycle, timers, state } = harness();
    await lifecycle.run(async () => "done");
    assert.equal(timers.pending, 1);
    timers.fire();
    assert.equal(state.disposals, 1);
});

test("a finished call does not shut Word down under one still running", async () => {
    const { lifecycle, timers, state } = harness();

    let releaseSlow;
    const slowStarted = { value: false };
    const slow = lifecycle.run(async () => {
        slowStarted.value = true;
        await new Promise((resolve) => {
            releaseSlow = resolve;
        });
    });

    await lifecycle.run(async () => "fast");

    assert.ok(slowStarted.value);
    assert.equal(timers.pending, 0, "the fast call must not arm a timer while the slow one runs");
    // Even if a timer were somehow armed, firing it must not dispose.
    timers.fire();
    assert.equal(state.disposals, 0);

    releaseSlow();
    await slow;
    assert.equal(timers.pending, 1, "the last call out arms the timer");
    timers.fire();
    assert.equal(state.disposals, 1);
});

test("work starting while the timer is armed cancels it", async () => {
    const { lifecycle, timers, state } = harness();
    await lifecycle.run(async () => "first");
    assert.equal(timers.pending, 1);

    let release;
    const second = lifecycle.run(async () => {
        await new Promise((resolve) => {
            release = resolve;
        });
    });
    assert.equal(timers.pending, 0, "starting work must cancel the armed shutdown");

    release();
    await second;
    timers.fire();
    assert.equal(state.disposals, 1);
});

test("a canvas displaying a document keeps Word alive", async () => {
    const { lifecycle, timers, state } = harness({ displaying: true });
    await lifecycle.run(async () => "done");
    assert.equal(timers.pending, 0);
    assert.equal(state.disposals, 0);
});

test("a canvas opened while the timer is armed prevents the disposal", async () => {
    const { lifecycle, timers, state } = harness();
    await lifecycle.run(async () => "done");
    assert.equal(timers.pending, 1);

    state.displaying = true; // a canvas opened in the meantime
    timers.fire();
    assert.equal(state.disposals, 0, "the timer must re-check on the way out, not only on the way in");
});

test("the counter unwinds even when the work throws", async () => {
    const { lifecycle, timers } = harness();
    await assert.rejects(
        () =>
            lifecycle.run(async () => {
                throw new Error("read failed");
            }),
        /read failed/,
    );
    assert.equal(lifecycle.inFlight, 0);
    assert.equal(timers.pending, 1, "a failed call still releases Word");
});

test("busy reports whether anything is using Word", async () => {
    const { lifecycle } = harness();
    assert.equal(lifecycle.busy, false);

    let release;
    const running = lifecycle.run(async () => {
        assert.equal(lifecycle.busy, true);
        await new Promise((resolve) => {
            release = resolve;
        });
    });
    assert.equal(lifecycle.busy, true);
    release();
    await running;
    assert.equal(lifecycle.busy, false);
});

test("nested work keeps Word alive until the outermost call finishes", async () => {
    // An edit that reads first is the obvious shape; the inner call finishing
    // must not release Word out from under the outer one.
    const { lifecycle, timers, state } = harness();
    await lifecycle.run(async () => {
        await lifecycle.run(async () => "inner");
        assert.equal(timers.pending, 0, "the inner call must not arm a timer");
        assert.equal(state.disposals, 0);
    });
    assert.equal(timers.pending, 1);
});

test("a failed disposal does not escape as an unhandled rejection", async () => {
    const timers = fakeTimers();
    const lifecycle = createIdleShutdown({
        idleMs: 60_000,
        isDisplaying: () => false,
        dispose: async () => {
            throw new Error("Word was already gone");
        },
        setTimer: timers.setTimer,
        clearTimer: timers.clearTimer,
    });
    await lifecycle.run(async () => "done");
    timers.fire();
    // Give the rejected disposal a turn to land; the process-level
    // unhandledRejection handler would fail the run if it escaped.
    await new Promise((resolve) => setTimeout(resolve, 10));
});
