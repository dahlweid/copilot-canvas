// What the panel is allowed to say about its own connection.
//
// The defect (#66) is a status that can never stop being wrong: `EventSource`
// retries forever, so a message keyed on `readyState === CONNECTING` is
// permanent, and the message claimed a recovery was in progress. A panel held it
// for ~90 minutes against a port nothing was listening on.
//
// So the assertion that matters here is a *negative over unbounded time* --
// there is no elapsed time and no number of failures at which the panel is still
// claiming to reconnect. Driving that with a virtual clock is not a shortcut
// around a slow test; a test that really waited would be measuring the runner.
//
// The fake source below is the thing
// `spikes/viewer-connection/probes/probe-dead-server.mjs` exists to validate: it
// drives the real module with a real `EventSource` against a really-closed
// server, so the shape assumed here is checked against an implementation rather
// than against the specification as remembered.
//
// Office-free.

import { test } from "node:test";
import assert from "node:assert/strict";

import {
    monitorConnection,
    GRACE_MS,
    DEADLINE_MS,
    RECONNECTING,
    LOST,
} from "../../src/ui/connection-status.mjs";

const CONNECTING = 0;
const OPEN = 1;
const CLOSED = 2;

/**
 * A source and a clock the test drives by hand.
 *
 * `advance` fires due timers in time order, and timers scheduled *by* a timer
 * are picked up in the same sweep -- otherwise a two-stage escalation could
 * never be observed and the harness would be hiding the state under test.
 */
function harness({ readyState = CONNECTING } = {}) {
    let clock = 0;
    let seq = 0;
    const timers = [];
    const statuses = [];
    let closes = 0;

    const source = {
        readyState,
        close() {
            closes += 1;
            this.readyState = CLOSED;
        },
        onopen: null,
        onerror: null,
    };

    const api = monitorConnection(source, {
        setStatus: (text, options = {}) => statuses.push({ text, ...options }),
        now: () => clock,
        schedule: (fn, ms) => timers.push({ at: clock + ms, seq: seq++, fn }),
    });

    return {
        source,
        api,
        statuses,
        get closes() {
            return closes;
        },
        get last() {
            return statuses.at(-1);
        },
        advance(ms) {
            const until = clock + ms;
            for (;;) {
                const due = timers
                    .filter((t) => t.at <= until)
                    .sort((a, b) => a.at - b.at || a.seq - b.seq)[0];
                if (!due) break;
                timers.splice(timers.indexOf(due), 1);
                clock = due.at;
                due.fn();
            }
            clock = until;
        },
    };
}

test("a drop shorter than the grace period says nothing", () => {
    const h = harness();
    h.source.onerror();

    // Two observations, and the first is the one that carries weight. Asserting
    // only at `GRACE_MS - 1` derives the test's timing from the constant it is
    // testing, so it passes at *every* value of that constant including zero --
    // measured: setting `GRACE_MS = 0` left this test green. Advancing by 0
    // instead asks the behavioural question directly: is the panel silent at the
    // instant the first attempt fails?
    h.advance(0);
    assert.deepEqual(h.statuses, [], "the panel announced the very first failed attempt");

    h.advance(GRACE_MS - 1);
    assert.deepEqual(h.statuses, [], "the panel announced a blip the reader would never have noticed");
});

test("a drop that outlasts the grace period says it is reconnecting", () => {
    const h = harness();
    h.source.onerror();
    h.advance(GRACE_MS);

    assert.deepEqual(h.last, { text: RECONNECTING, busy: true });
});

test("no elapsed time past the deadline leaves the panel claiming to reconnect", () => {
    // THE acceptance test for #66. Not "it eventually stops" -- that a message
    // is permanent is the defect, so the assertion has to quantify over time
    // rather than sample it. An hour, with errors still arriving at the cadence
    // a browser retries at, is as unbounded as a bounded test gets.
    const h = harness();
    h.source.onerror();

    h.advance(DEADLINE_MS);
    assert.equal(h.last.text, LOST, "the panel was still claiming a recovery at the deadline");

    for (let elapsed = DEADLINE_MS; elapsed < 3_600_000; elapsed += 3_000) {
        h.advance(3_000);
        h.source.onerror();
        assert.notEqual(h.last.text, RECONNECTING, `still claiming a recovery after ${elapsed + 3_000}ms`);
        assert.notEqual(h.last.busy, true, `still showing a spinner after ${elapsed + 3_000}ms`);
    }

    assert.equal(h.last.text, LOST);
});

test("the panel closes the source before it says it has stopped updating", () => {
    // The wording is a report, not a prediction: "cannot resume" is true because
    // of this call. Reversed, the panel would claim to be finished while still
    // holding a source that could deliver an event a moment later.
    const h = harness();
    h.source.onerror();
    h.advance(DEADLINE_MS);

    assert.equal(h.closes, 1, "gave up without closing the source");
    assert.equal(h.source.readyState, CLOSED);
    assert.equal(h.last.text, LOST);
    assert.equal(h.statuses.filter((s) => s.text === LOST).length, 1, "said it more than once");
});

test("a source the browser has already closed is terminal at once", () => {
    // A CLOSED source will not retry, so there is nothing the grace period is
    // waiting for. This is the branch the old handler had no case for at all.
    const h = harness({ readyState: CLOSED });
    h.source.onerror();

    assert.equal(h.last.text, LOST);
    assert.equal(h.statuses.some((s) => s.text === RECONNECTING), false, "promised a recovery of a dead source");
});

test("the terminal line neither names a cause nor promises a remedy", () => {
    // From inside the iframe, a dropped stream does not distinguish a stopped
    // server from a broken network, and the one candidate remedy -- re-opening
    // the canvas -- was proposed and withdrawn twice on #66 for want of any
    // instrument that could confirm a webview had attached. Both halves would be
    // the very defect being fixed: a surface asserting what nothing established.
    assert.doesNotMatch(LOST, /server|network|crash|restart|Word|extension/i, "names a cause it cannot know");
    assert.doesNotMatch(LOST, /re-?open|reopen|refresh|reload|try again|retry/i, "promises a remedy nobody verified");
});

test("the terminal status is an error and carries no spinner", () => {
    const h = harness();
    h.source.onerror();
    h.advance(DEADLINE_MS);

    assert.equal(h.last.error, true, "a dead panel is not a neutral state");
    assert.notEqual(h.last.busy, true, "a spinner is an assertion that something is still happening");
});

test("a reconnect inside the deadline clears the message and re-arms the grace", () => {
    const h = harness();

    h.source.onerror();
    h.advance(GRACE_MS);
    assert.equal(h.last.text, RECONNECTING);

    h.source.readyState = OPEN;
    h.source.onopen();
    assert.equal(h.last.text, null, "the panel kept complaining after it reconnected");

    // The second outage must get its own full deadline. Sharing the first one's
    // clock would give up almost immediately on a panel that is fine.
    h.source.readyState = CONNECTING;
    h.source.onerror();
    h.advance(DEADLINE_MS - 1);
    assert.equal(h.last.text, RECONNECTING, "the second outage inherited the first one's clock");
    assert.equal(h.closes, 0);
});

test("a stale timer from a healed outage cannot speak", () => {
    // The escalation is armed when an outage starts, so a drop that heals leaves
    // two timers behind. Without the epoch guard they fire into a healthy panel.
    const h = harness();

    h.source.onerror();
    h.source.readyState = OPEN;
    h.source.onopen();
    h.advance(DEADLINE_MS * 2);

    assert.equal(h.closes, 0, "a healed connection was closed by a timer belonging to a past outage");
    assert.equal(h.statuses.filter((s) => s.text === RECONNECTING || s.text === LOST).length, 0);
});

test("a first connection does not touch a status line it never wrote", () => {
    // `onopen` fires on the initial connect too, and the status line is shared
    // with "Reloaded from disk." and the picker's errors. Clearing it
    // unconditionally would wipe somebody else's message.
    const h = harness();
    h.source.readyState = OPEN;
    h.source.onopen();

    assert.deepEqual(h.statuses, []);
});

test("nothing speaks after the panel has given up", () => {
    // Terminal means terminal. A late `onopen` -- from a source closed while an
    // attempt was in flight -- must not resurrect a panel that has stopped.
    const h = harness();
    h.source.onerror();
    h.advance(DEADLINE_MS);
    const after = h.statuses.length;

    h.source.onopen();
    h.source.onerror();
    h.advance(DEADLINE_MS * 10);

    assert.equal(h.statuses.length, after, "the panel spoke again after saying it could not resume");
    assert.equal(h.closes, 1);
});
