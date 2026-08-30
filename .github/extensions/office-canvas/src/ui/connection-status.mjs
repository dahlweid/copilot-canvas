// What a viewer panel is allowed to claim about its own connection.
//
// Separate from `app.js` for the same reason `change-wording.mjs` is, and the
// reason is measured rather than assumed: `app.js` cannot be imported under
// Node at all. It reaches `pdf-view.mjs`, which imports pdf.js from an absolute
// `/vendor/` URL, and module resolution fails there before any DOM access --
// `Cannot find module 'C:\vendor\pdf.min.mjs'`. Logic left inside `app.js` has
// no reachable assertion, so the decision lives here where a test can drive it.
//
// The defect this exists to end (#66): `EventSource` reconnects on a schedule of
// its own and never stops, so `readyState` sits at CONNECTING forever, so a
// status keyed on CONNECTING says "Reconnecting…" forever. Observed live -- a
// panel held that message for roughly 90 minutes against port 57318 while the
// extension process was listening on no sockets at all. The message was not
// merely unhelpful, it asserted a recovery that no code was capable of
// performing.

/**
 * `EventSource.CLOSED`, spelled numerically.
 *
 * Not a style choice: the global `EventSource` does not exist under Node, so
 * naming it here would make this module unimportable by its own test. The
 * numeric values are fixed by the specification's `readyState` enum.
 */
const CLOSED = 2;

/** How long a drop may last before the panel says anything at all. */
export const GRACE_MS = 1_500;

/**
 * How long a drop may last before the panel stops claiming a recovery.
 *
 * No measurement pins this number and none could: it separates a transient from
 * a permanent loss, and the two are indistinguishable from inside the iframe at
 * the instant they begin. What it has to be is longer than any real reconnect
 * and shorter than a reader's patience.
 *
 * It is also what keeps a *normal* panel close quiet. `ViewerInstance.close()`
 * ends every SSE client, which is correct behaviour and not a fault, and it
 * happens because the panel is going away -- so its document is destroyed and
 * nothing renders a status into it. Ten seconds is a margin wide enough that
 * this does not have to be timed precisely. Stated as reasoning, not as a
 * measurement: the ordering of the host's teardown against `onClose` has not
 * been instrumented.
 */
export const DEADLINE_MS = 10_000;

/** Said while a recovery is still possible. */
export const RECONNECTING = "Reconnecting…";

/**
 * Said once it is not.
 *
 * Every clause is something this code established. *Connection lost* was
 * observed. *Stopped updating and cannot resume* is made true by the `close()`
 * that runs immediately before this is shown -- it reports what the panel did to
 * itself rather than predicting anything about a server it cannot see.
 *
 * It names no cause, because from inside the iframe none was distinguished, and
 * it names **no remedy**, because none is established. "Re-open the viewer" was
 * proposed and withdrawn twice on #66: a re-opened canvas produces a healthy
 * server, which is server-side evidence that says nothing about whether a
 * webview attached to it. That this panel is finished is known; that any
 * particular action produces a working one is not, and only the first belongs in
 * front of a reader.
 */
export const LOST = "Connection lost — this panel has stopped updating and cannot resume.";

/**
 * Watches one `EventSource` and keeps the status line honest about it.
 *
 * `now` and `schedule` are injected so a test owns the clock; the defect is a
 * state reached only by waiting, and a test that actually waited for it would be
 * measuring the runner's patience.
 *
 * @param {{readyState: number, close: () => void, onopen: unknown, onerror: unknown}} source
 * @param {{setStatus: (text: string|null, opts?: object) => void, now?: () => number,
 *          schedule?: (fn: () => void, ms: number) => unknown}} deps
 */
export function monitorConnection(source, { setStatus, now = () => Date.now(), schedule = setTimeout } = {}) {
    /** When the current outage began, or null while connected. */
    let downSince = null;
    /** Bumped on every state change, so a timer from a healed outage is inert. */
    let epoch = 0;
    /** Set once the panel has given up. Terminal: nothing moves it back. */
    let finished = false;
    /** Whether the status line currently belongs to this monitor. */
    let owned = false;

    const say = (text, options) => {
        owned = true;
        setStatus(text, options);
    };

    /**
     * Clears the status, but only if this monitor is what put something there.
     *
     * A reconnect must not wipe a message another part of the UI is showing --
     * "Reloaded from disk." and the picker's errors both live on the same line.
     */
    const clear = () => {
        if (!owned) return;
        owned = false;
        setStatus(null);
    };

    const giveUp = () => {
        if (finished) return;
        finished = true;
        // Closing first is what makes the wording true. Reversed, the panel
        // would be asserting that it has stopped updating while still holding a
        // source that could deliver an event a moment later.
        try {
            source.close();
        } catch {
            /* already gone; the claim holds either way */
        }
        say(LOST, { error: true });
    };

    const opened = () => {
        if (finished) return;
        downSince = null;
        epoch += 1;
        clear();
    };

    const failed = () => {
        if (finished) return;

        // A source the browser has closed will not retry, so there is nothing to
        // wait for and the grace period would only delay the truth.
        if (source.readyState === CLOSED) return giveUp();

        // Errors arrive once per failed attempt. Only the first one starts the
        // clock; the rest are the same outage continuing.
        if (downSince !== null) return;
        downSince = now();

        const mine = (epoch += 1);
        const current = () => !finished && epoch === mine && downSince !== null;

        schedule(() => {
            if (current()) say(RECONNECTING, { busy: true });
        }, GRACE_MS);

        schedule(() => {
            if (current()) giveUp();
        }, DEADLINE_MS);
    };

    source.onopen = opened;
    source.onerror = failed;

    return { opened, failed };
}
