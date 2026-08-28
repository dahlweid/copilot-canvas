// When the hidden Word instance is allowed to go away.
//
// Two things can keep Word alive: a canvas displaying a document, and a tool
// call in progress. The canvas half was always tracked. The tool half was not,
// and the omission was a real defect rather than a tidy-up: a tool call can run
// for minutes (a structure read is bounded by the cold-start timeout), the
// agent issues tool calls in parallel, and the first to finish would arm a
// shutdown timer that fired underneath the ones still running. The victim saw
// its host disposed mid-call and reported that Word was unavailable -- a
// message that is both wrong and impossible to act on.
//
// So: an explicit in-flight counter, and every path that needs Word wrapped in
// `run`. Layers building further tools wrap their handler and inherit the
// lifecycle instead of each re-deriving the rule in its own `finally`.

/**
 * @param {object} options
 * @param {number} options.idleMs        How long to stay warm with nothing to do.
 * @param {() => boolean} options.isDisplaying  Whether a canvas currently owns Word.
 * @param {() => any} options.dispose    Releases Word. May be async; never called
 *                                       while work is in flight or a canvas is open.
 * @param {(message: string) => void} [options.log]
 * @param {(fn: () => void, ms: number) => any} [options.setTimer]  Injectable for tests.
 * @param {(handle: any) => void} [options.clearTimer]
 */
export function createIdleShutdown({
    idleMs,
    isDisplaying,
    dispose,
    log = () => {},
    setTimer = setTimeout,
    clearTimer = clearTimeout,
}) {
    let inFlight = 0;
    let timer = null;

    function cancel() {
        if (timer === null) return;
        clearTimer(timer);
        timer = null;
    }

    function schedule() {
        cancel();
        if (isDisplaying()) return; // a canvas is displaying something; it owns Word
        if (inFlight > 0) return; // something is still using Word
        timer = setTimer(() => {
            timer = null;
            // Re-checked on the way out as well as on the way in: work can start
            // while the timer is armed, and disposing under it is the whole bug.
            if (isDisplaying() || inFlight > 0) return;
            log("no canvas open and no tool activity, shutting Word down");
            void (async () => {
                try {
                    await dispose();
                } catch {
                    /* teardown is best-effort */
                }
            })();
        }, idleMs);
        // Never a reason to hold the process open.
        timer?.unref?.();
    }

    return {
        /** Runs `fn` with Word pinned alive for its whole duration. */
        async run(fn) {
            inFlight += 1;
            cancel();
            try {
                return await fn();
            } finally {
                inFlight -= 1;
                schedule();
            }
        },
        cancel,
        schedule,
        /** True while any wrapped work is running. */
        get busy() {
            return inFlight > 0;
        },
        get inFlight() {
            return inFlight;
        },
    };
}
