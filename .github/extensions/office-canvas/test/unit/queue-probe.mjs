// Two primitives for testing a promise queue, shared rather than copied.
//
// Both exist to replace a sleep. A queue test has to answer "did this operation
// run yet?", and a fixed interval answers it on slack: the assertion passes
// because the timer was long enough, not because the ordering held. `deferred`
// pins the moment an operation is *inside* the gate, and `turns` bounds "did
// not run" to a countable number of microtask turns rather than a duration.

/** A promise with its resolver exposed, so a fake can gate on it. */
export function deferred() {
    let resolve;
    const promise = new Promise((r) => {
        resolve = r;
    });
    return { promise, resolve };
}

/**
 * Yields `count` times, so anything already queued has had a chance to run.
 *
 * Used to give a "still queued" claim a bounded meaning: an operation that has
 * not started after this many turns is being held by the queue, not merely
 * slower than the assertion.
 */
export const turns = async (count) => {
    for (let i = 0; i < count; i++) await new Promise((resolve) => setImmediate(resolve));
};
