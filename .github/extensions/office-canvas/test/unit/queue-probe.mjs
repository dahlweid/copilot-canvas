// Two primitives for testing a promise queue, shared rather than copied.
//
// `deferred` pins the moment an operation is *inside* a gate, which is the one
// thing here that is exact: a fake that resolves it has demonstrably been
// entered. `turns` is weaker and is documented as such below -- it replaces a
// sleep, but it does not replace an ordering assertion.

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
 * **This bounds a "did not run" claim; it does not prove one.** `setImmediate`
 * yields to the check phase, so pending I/O callbacks do get to run, but twenty
 * iterations is still a window rather than a guarantee. An operation can be
 * absent because the queue is holding it *or* because its own asynchronous work
 * -- a `stat`, a file write -- has not finished yet, and this cannot tell those
 * apart.
 *
 * So a negative assertion after `turns` is a tripwire, not the proof. What
 * carries the weight in these tests is the settled order of the whole
 * lifecycle, asserted once everything has resolved, which is not timing-bounded
 * at all.
 *
 * Stated here because the original of this helper claimed more. Extracted from
 * `outline-interleave.test.mjs`, where every use was inside a `Promise.race` --
 * a timeout, which asserts nothing -- it acquired a doc-comment during
 * extraction saying an operation not started after these turns "is being held by
 * the queue, not merely slower than the assertion". The mechanism does not
 * deliver that, and migrating a helper is exactly when such a sentence gets
 * invented.
 */
export const turns = async (count) => {
    for (let i = 0; i < count; i++) await new Promise((resolve) => setImmediate(resolve));
};
