// Does the panel actually *show* that it has stopped, when its server dies?
//
// #72 made `connection-status.mjs` give up honestly, and
// `connection-status.test.mjs` drives that state machine hard -- against a fake
// source and a fake clock. What neither it nor the probe reaches is the panel
// itself: whether the committed `app.js`, wired to the committed markup, puts
// that message in front of a reader. Until now the only thing asserting the
// wiring was a **regex** over `app.js`'s source in `ui-contract.test.mjs`, which
// says the monitor is imported and called and cannot say that calling it renders
// anything. That file states the limit itself: it "cannot fail on a monitor that
// is imported, called, and then never reached at runtime".
//
// That is the verification gap #72 and #75 both had to declare open, and #77 is
// what makes it closable: a reload kills the render server outright, with no
// Word, no document and no race to lose, so "the server is gone and is not
// coming back" is reproducible on demand rather than waited for. That is the
// condition reproduced here -- an `EventSource` that fails and never reopens is
// exactly what a dead server looks like from inside the iframe, which is the
// only vantage the panel has.
//
// **What this is NOT.** This is Node, not Chromium. `ui-harness.mjs` stands in
// `fetch`, `EventSource`, `IntersectionObserver` and the DOM, so the reconnect
// *scheduling* of a real `EventSource` is not exercised here -- that is what
// `spikes/viewer-connection/probes/probe-dead-server.mjs` measures, against
// undici's implementation and a really-closed `ViewerInstance`. What is new here
// and nowhere else is the last hop: monitor -> `app.js` -> `setStatus` -> the
// `#status` element the markup ships. No webview is claimed.
//
// Office-free.

import { test, mock } from "node:test";
import assert from "node:assert/strict";

import { loadApp } from "./ui-harness.mjs";
import { GRACE_MS, DEADLINE_MS, RECONNECTING, LOST } from "../../src/ui/connection-status.mjs";

/** `EventSource.CONNECTING`, spelled numerically for the same reason the module does. */
const CONNECTING = 0;

/**
 * Loads the panel with its clock under test control.
 *
 * The timers have to be mocked **before** `app.js` is imported. `app.js` calls
 * `monitorConnection(source, { setStatus })` without a `schedule`, so the
 * module's `schedule = setTimeout` default resolves to whatever the global is at
 * the moment of that call -- and that call happens during the import. Enabling
 * the mock afterwards would leave the panel holding the real `setTimeout` and
 * the test waiting ten actual seconds for a message that arrives after it has
 * already asserted.
 *
 * Only `setTimeout` is mocked: `ui-harness`'s `settle()` drains the module graph
 * with `setImmediate`, and mocking that would deadlock the load.
 */
async function loadPanel() {
    mock.timers.enable({ apis: ["setTimeout"] });
    const app = await loadApp({ state: { status: "idle", doc: null } });
    return {
        ...app,
        status: () => app.el("status"),
        restore() {
            app.restore();
            mock.timers.reset();
        },
    };
}

test("the panel renders the terminal message into its own status element when the connection dies", async () => {
    const app = await loadPanel();
    try {
        const source = app.source;
        assert.ok(source, "app.js opened no EventSource, so nothing monitors the connection");

        // The panel is quiet until something goes wrong. Asserted so a status
        // element that says LOST from the very start -- which would pass every
        // assertion below -- is caught here instead.
        assert.equal(app.status().hidden, true, "the panel was already showing a status before any failure");

        source.readyState = CONNECTING;
        source.onerror();

        mock.timers.tick(GRACE_MS);
        assert.equal(
            app.status().textContent,
            RECONNECTING,
            "the panel said nothing while a recovery was still possible",
        );

        mock.timers.tick(DEADLINE_MS - GRACE_MS);
        assert.equal(
            app.status().textContent,
            LOST,
            "the panel never told the reader it had stopped -- the gap #72 and #75 left open",
        );
        assert.equal(app.status().hidden, false, "the message was rendered into a hidden element");
        assert.ok(app.status().classList.contains("error"), "the terminal message was not styled as an error");
    } finally {
        app.restore();
    }
});

test("giving up closes the panel's own source, which is what makes the wording true", async () => {
    const app = await loadPanel();
    try {
        const source = app.source;
        source.readyState = CONNECTING;
        source.onerror();

        mock.timers.tick(DEADLINE_MS - 1);
        assert.equal(source.closed, 0, "the panel dropped its source before the deadline");

        mock.timers.tick(1);
        // "cannot resume" is a claim about this panel, and it is only true
        // because the source is shut before the sentence is shown. A panel still
        // holding an open source could be handed an event a moment later.
        assert.equal(source.closed, 1, "the panel claims it cannot resume while still holding an open source");
    } finally {
        app.restore();
    }
});

test("no amount of further failure moves the panel off the terminal message", async () => {
    const app = await loadPanel();
    try {
        const source = app.source;
        source.readyState = CONNECTING;
        source.onerror();
        mock.timers.tick(DEADLINE_MS);
        assert.equal(app.status().textContent, LOST);

        // #66 in the shape it would come back: errors keep arriving from a
        // source the browser is still retrying, and each one is another chance
        // for the panel to start claiming a recovery again. An hour of them.
        for (let elapsed = 0; elapsed < 3_600_000; elapsed += 3_000) {
            source.onerror();
            mock.timers.tick(3_000);
            assert.notEqual(
                app.status().textContent,
                RECONNECTING,
                `the panel went back to claiming a recovery ${elapsed / 1000}s after giving up`,
            );
        }

        assert.equal(app.status().textContent, LOST, "the panel stopped saying it had stopped");
        assert.equal(source.closed, 1, "the panel closed its source more than once");
    } finally {
        app.restore();
    }
});

test("a blip shorter than the grace period is never mentioned to the reader", async () => {
    const app = await loadPanel();
    try {
        const source = app.source;
        source.readyState = CONNECTING;
        source.onerror();

        // Healed with a millisecond of the grace period still to run, so the
        // timer scheduled by the failure is still pending and now belongs to an
        // outage that is over. It has to fire and find that out.
        mock.timers.tick(GRACE_MS - 1);
        source.onopen();
        mock.timers.tick(DEADLINE_MS * 2);

        // Added because a mutant survived without it: dropping the `current()`
        // guard on the grace timer left every assertion in this file green while
        // the panel announced "Reconnecting…" over a connection that was already
        // back. A momentary drop is the common case, so this is the message a
        // reader would see most often.
        assert.equal(
            app.status().hidden,
            true,
            "the panel announced a reconnection attempt for an outage that had already healed",
        );
        assert.equal(source.closed, 0, "the panel closed a source that never stopped working");
    } finally {
        app.restore();
    }
});

test("a connection that recovers inside the deadline leaves no message behind", async () => {
    const app = await loadPanel();
    try {
        const source = app.source;
        source.readyState = CONNECTING;
        source.onerror();
        mock.timers.tick(GRACE_MS);
        assert.equal(app.status().textContent, RECONNECTING);

        source.onopen();

        // The negative case matters as much as the terminal one: a panel that
        // latched "Reconnecting…" after the connection came back would be the
        // same class of lie as #66, pointing the other way.
        assert.equal(app.status().hidden, true, "the panel kept a stale connection message after reconnecting");

        mock.timers.tick(DEADLINE_MS * 2);
        assert.equal(app.status().hidden, true, "a timer from the healed outage fired and said the panel was lost");
        assert.equal(source.closed, 0, "the panel closed a source that had recovered");
    } finally {
        app.restore();
    }
});
