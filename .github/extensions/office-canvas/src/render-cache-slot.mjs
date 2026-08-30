// The single live RenderCache, and the rule that a disposal is never visible
// to a later caller.
//
// The extension keeps one RenderCache -- and therefore one hidden Word -- and
// throws it away whenever nothing needs Word any more: the last canvas closes,
// or the idle timer fires. The next caller is supposed to get a fresh one.
//
// It did not, and the reason is that disposal takes real time while the slot
// holding the cache was cleared only afterwards. `WordHost.dispose` sends
// `quit` under a 20 s timeout and then waits up to a further 5 s for the child
// to exit, so the window is up to ~25 s wide, and for that whole window the
// slot still pointed at a cache whose host had already set `#disposed`. Every
// caller arriving inside it was handed that host, and `#ensureStarted` answers
// a disposed host with `The Word host has been shut down.`
//
// For a tool call that is a bad half-minute. For a canvas it is permanent:
// `ViewerInstance` captured the reference it was constructed with, so a panel
// opened inside the window held the dead host for the rest of the session, with
// nothing in the code able to give it another one. That is the "works for one
// session and is then dead" shape reported in #61.
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
