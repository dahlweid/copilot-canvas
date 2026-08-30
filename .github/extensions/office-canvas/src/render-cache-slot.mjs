// The single live RenderCache, and the rule that a disposal is never visible
// to a later caller.
//
// The extension keeps one RenderCache -- and therefore one hidden Word -- and
// throws it away whenever nothing needs Word any more: the last canvas closes,
// or the idle timer fires. The next caller is supposed to get a fresh one.
//
// It did not, and the reason is that disposal takes real time while the slot
// holding the cache was cleared only afterwards. `WordHost.dispose` sends `quit`
// under a 20 s timeout (`word-host.mjs:473`) and then waits up to a further 5 s
// for the child to exit (`:486`), so the window is up to ~25 s wide -- a ceiling
// derived from those two constants, not measured. For that whole window the slot
// still pointed at a cache that was on its way out, and every caller arriving
// inside it was handed that cache.
//
// What such a caller gets is *not* one failure but two, and the difference is
// worth keeping straight because only the second is the sentence from #61:
//
//   - In the ~5 s tail, `#disposed` is set, so `#ensureStarted` answers with
//     `The Word host has been shut down.`
//   - In the ~20 s before that, `#disposed` is still false -- it is set at `:483`
//     between the quit and the exit wait, and cannot be set earlier because
//     `#send` refuses to run against a disposed host and would reject the very
//     quit being sent. The caller's command is written to a host that is inside
//     `Stop-Word` and not reading, so it is never answered; when the child exits,
//     `#onExit` rejects it through `#rejectAll` with the same `word_unavailable`
//     code but a different message: `The Word host exited (code N, signal S).`
//     It does not respawn Word, because `dispose()` clears `#openArgs` on entry,
//     which makes `request()`'s replay path unreachable.
//
// That second reading is derived from those sources rather than measured; the
// unit tests here exercise the tail, which is the one that produced #61.
//
// For a tool call either is a bad half-minute. What made it *stick* was the
// panel: `ViewerInstance` captured the reference it was constructed with, and
// `open` reuses an existing instance for the same id by design ("rehydrate,
// reload, focus"), so a panel created inside the window never consulted the slot
// again and stayed dead for the life of the process. Measured downstream: the
// surface came back only on `extensions_reload`, which clears it because the
// instance map dies with the process. That is the "works for one session and is
// then dead" shape reported in #61.
//
// So the swap is synchronous and the disposal is awaited afterwards. There is
// no instant at which `get()` can return a cache that something has already
// begun to dispose.

/**
 * @param {object} options
 * @param {() => any} options.create  Builds a fresh RenderCache. Called lazily.
 * @param {(message: string) => void} [options.log]
 */
export function createRenderCacheSlot({ create, log = () => {} }) {
    let current = null;
    /** Disposals still running, so a caller that needs Word *gone* can wait. */
    const retiring = new Set();

    async function retire(doomed) {
        try {
            await doomed.dispose();
        } catch (err) {
            // Teardown is best-effort, but a failure here means a Word we may
            // not have ended -- worth a line rather than a silent swallow.
            log(`render cache disposal failed: ${err?.message ?? err}`);
        }
    }

    return {
        /** The live cache, built on first use. Never one being disposed. */
        get() {
            if (!current) current = create();
            return current;
        },

        /**
         * The live cache, or null when none has been built.
         *
         * For callers that want to observe Word without starting it -- reporting
         * `wordVersion`, closing a document that is already open, reaping on
         * process exit. Building a cache for any of those would be answering a
         * question by creating the thing being asked about.
         */
        peek() {
            return current;
        },

        /**
         * Retires the current cache. The slot is empty the instant this is
         * called, so a caller arriving during the teardown builds a fresh one
         * rather than joining the one on its way out.
         *
         * Resolves once every disposal started so far has finished, which is
         * what process shutdown needs: returning while a `quit` is still in
         * flight would leave a hidden Word behind.
         */
        async dispose() {
            const doomed = current;
            current = null;
            if (doomed) {
                const task = retire(doomed).finally(() => retiring.delete(task));
                retiring.add(task);
            }
            await Promise.all([...retiring]);
        },

        /** True while any retired cache is still shutting Word down. */
        get retiring() {
            return retiring.size > 0;
        },
    };
}
