// What a closed ViewerInstance is allowed to go on doing (#81).
//
// ## The defect
//
// `close()` tore down the server, the SSE clients and the watcher, and left the
// `#autoRefresh` retry loop running. The loop had nothing it *could* consult --
// there was no disposal flag in the file -- and `close()` does not null
// `this.doc`, so `refresh()`'s `not_open` guard did not stop it either.
//
// #80 measured that continuation with a **fake** `RenderCache` and reported it
// inert. It was, for that cache: a fake cannot open a file, so the consequence
// that matters was invisible to it by construction. Measured again against a
// *real* cache in `spikes/viewer-disposal/probes/probe-close-during-retry.mjs`:
// after `close()` returned, **2 `cache.refresh` calls and 1
// `WordHost.openDocument`** landed, followed by a 924 ms PDF export. A disposed
// panel was driving Word.
//
// ## What is guaranteed, stated narrowly
//
// **At most one already-started `cache.refresh` may complete after `close()`
// returns; nothing new is started and no result is acted on.** Not "close
// cancels in-flight work" -- that would be false. The leaf of this work is a
// PowerShell host driving Word over COM, where nothing observes a signal and
// `Documents.Open` on a locked file hangs rather than failing. Awaiting it in
// `close()` would trade a discarded render for a teardown that can block
// indefinitely.
//
// The tests below are therefore split along that boundary: one closes during
// the *backoff* (nothing in flight), one closes with a refresh *inside* the
// cache (in flight, result must be dropped). A suite with only the first goes
// green against a disposal check placed at the top of the retry loop, which is
// the version someone would plausibly write.
//
// Office-free: the cache is a stub, and the document is a real temp file so the
// committed `FileWatcher` is the one under test. Word is never started.
// Run: node --test ".github/extensions/office-canvas/test/unit/*.test.mjs"

import test, { after, before, beforeEach } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm, stat, utimes, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { ViewerClosedError, ViewerInstance } from "../../src/server.mjs";
import { FileWatcher } from "../../src/watcher.mjs";

const BUDGET_MS = 15_000;

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

function deferred() {
    let resolve;
    const promise = new Promise((r) => {
        resolve = r;
    });
    return { promise, resolve };
}

async function waitFor(predicate, label, budgetMs = BUDGET_MS) {
    const deadline = Date.now() + budgetMs;
    while (Date.now() < deadline) {
        if (predicate()) return;
        await sleep(20);
    }
    assert.fail(`timed out after ${budgetMs}ms waiting for ${label}`);
}

/**
 * A RenderCache stand-in that can fail to order and can be held open.
 *
 * `gate` is what makes the in-flight case testable at all: without it there is
 * no way to be inside `cache.refresh` at the moment `close()` lands, and that
 * is precisely the window a disposal check at the top of the loop leaves open.
 *
 * `pdfCalls` is counted separately from `refreshCalls` because they are
 * different crossings. In the real cache `refresh` opens the document in Word
 * and `pdf` exports it -- so a test that only counts refreshes cannot see the
 * second, more expensive half being performed for a panel that no longer
 * exists.
 */
class StubCache {
    wordVersion = "stub";
    refreshCalls = 0;
    pdfCalls = 0;
    inFlight = 0;
    failures = [];
    /** Set to a promise to hold the next `refresh` inside the cache. */
    gate = null;
    #version = 0;

    #info(docPath) {
        return {
            path: docPath,
            key: `${path.basename(docPath)}-${this.#version}`,
            name: path.basename(docPath),
            pageCount: 1,
        };
    }

    async open(docPath) {
        return this.#info(docPath);
    }

    async pdf(docPath) {
        this.pdfCalls += 1;
        return { file: `${docPath}.pdf`, key: this.#info(docPath).key };
    }

    async refresh(docPath) {
        this.refreshCalls += 1;
        this.inFlight += 1;
        try {
            if (this.gate) {
                const gate = this.gate;
                this.gate = null;
                await gate;
            }
            const code = this.failures.shift();
            if (code) {
                const err = new Error(`${code} raised by the test`);
                err.code = code;
                throw err;
            }
            this.#version += 1;
            return { ...this.#info(docPath), changed: true };
        } finally {
            this.inFlight -= 1;
        }
    }
}

let home;
let dir;
let docPath;
let cache;
let logLines;
const instances = [];
const dirs = [];

before(async () => {
    home = await mkdtemp(path.join(tmpdir(), "office-canvas-disposal-home-"));
    process.env.COPILOT_HOME = home;
});

beforeEach(async () => {
    while (instances.length) await instances.pop().close();
    dir = await mkdtemp(path.join(tmpdir(), "office-canvas-disposal-"));
    dirs.push(dir);
    docPath = path.join(dir, "report.docx");
    await writeFile(docPath, "original");
    cache = new StubCache();
    logLines = [];
});

after(async () => {
    while (instances.length) await instances.pop().close();
    for (const used of dirs) await rm(used, { recursive: true, force: true });
    if (home) await rm(home, { recursive: true, force: true });
});

async function openViewer({ autoRefreshDelaysMs = [200, 200, 200] } = {}) {
    const instance = new ViewerInstance({
        cache,
        instanceId: "disposal",
        workspacePath: null,
        autoRefreshDelaysMs,
        log: (line) => logLines.push(line),
    });
    instances.push(instance);
    await instance.openDocument(docPath);
    assert.equal(instance.status, "ready", "precondition: the document opened");
    return instance;
}

/**
 * Moves the file's mtime so the watcher reports a change.
 *
 * `utimes` rather than a write: the watcher fingerprints `mtimeMs|size`, so this
 * is a change by the only definition it has, and it leaves the bytes alone.
 */
async function touch() {
    const info = await stat(docPath);
    const moved = new Date(info.mtimeMs + 5000);
    await utimes(docPath, moved, moved);
}

/** Leaves the viewer sitting in an auto-refresh backoff, mid-retry. */
async function armRetryLoop(viewer) {
    // More failures than the backoff has attempts, so a loop that survives
    // `close()` has something left to consume and shows up in the count.
    cache.failures = ["file_locked", "file_locked", "file_locked", "file_locked"];
    await touch();
    await waitFor(() => cache.refreshCalls >= 1, "the watcher to deliver and the first attempt to fail");
}

test("a viewer closed mid-retry starts no further refresh", async () => {
    const viewer = await openViewer();
    await armRetryLoop(viewer);

    await viewer.close();
    const callsAtClose = cache.refreshCalls;

    // 1200ms against a remaining backoff of 200+200+200. The window is derived
    // from the delays this viewer was configured with rather than guessed, so
    // "nothing happened" is a claim about an interval that would certainly have
    // contained the three attempts had the loop survived -- not a claim resting
    // on the machine being slow enough to hide them.
    await sleep(1200);

    assert.equal(
        cache.refreshCalls,
        callsAtClose,
        `${cache.refreshCalls - callsAtClose} refresh(es) landed on a disposed viewer`,
    );
    assert.ok(cache.failures.length > 0, "precondition: the loop had failures left to consume");
});

test("closing mid-retry is not recorded as a failed refresh", async () => {
    // A cancellation dressed as a failure puts the panel into `status: "error"`
    // during its own teardown, so anything reading the retry state afterwards
    // sees an error that never happened.
    const viewer = await openViewer();
    await armRetryLoop(viewer);

    await viewer.close();
    await sleep(1200);

    assert.notEqual(viewer.status, "error", `teardown left the viewer in error: ${JSON.stringify(viewer.error)}`);
    assert.equal(viewer.error, null);
});

test("close() ends the backoff rather than waiting it out, and refuses the change", async () => {
    // Two properties in one observation, because they share the same evidence:
    // the watcher logs "change was not consumed" only when `#autoRefresh`
    // *rejects*, so the line appearing at all is the #75 rollback still being
    // signalled, and the line appearing *early* is the sleep being cancelled
    // rather than run out.
    const viewer = await openViewer({ autoRefreshDelaysMs: [4000, 4000, 4000] });
    await armRetryLoop(viewer);

    const startedAt = Date.now();
    await viewer.close();
    await waitFor(
        () => logLines.some((line) => line.includes("not consumed")),
        "the watcher to be told the change was not consumed",
        3000,
    );
    const elapsed = Date.now() - startedAt;

    // 1500 against a pending 4000. A `close()` that only set a flag would have
    // to wait the timer out and could not land inside this.
    assert.ok(elapsed < 1500, `close() waited out the backoff: the refusal arrived ${elapsed}ms later`);
});

test("a change delivered during teardown is refused, and refused for being closed", async () => {
    // The delivery-path half of the invariant. `#autoRefresh` is the watcher's
    // `onChange`, and `#deliver` (watcher.mjs) marks the change seen *before*
    // awaiting `onChange` -- it captures `seenBefore` and calls `#markSeen` on
    // the new fingerprint first, rolling back only if it throws. So a disposal guard
    // that returns cleanly here does not merely skip a render -- it leaves the
    // change recorded as consumed, with nothing left to re-fire on. That is not
    // #67-like; reached through close() instead of through a lock, it is #67.
    //
    // The message is asserted, not just the line, because a refusal logged for
    // some other reason would satisfy the shape of this test while proving
    // nothing about disposal.
    const viewer = await openViewer();
    await armRetryLoop(viewer);

    await viewer.close();
    await waitFor(
        () => logLines.some((line) => line.includes("not consumed")),
        "the watcher to be told the change was not consumed",
        3000,
    );

    const refusal = logLines.find((line) => line.includes("not consumed"));
    assert.match(refusal, /closed while a refresh was in flight/);
    assert.match(refusal, /stays pending/);
});

test("a delivery refused during teardown still rolls the change back, though the watcher is closed", async () => {
    // The other half, and the one the log line above cannot establish: in
    // `#deliver`'s `catch` (watcher.mjs) the `change was not consumed ... it
    // stays pending` log sits *outside* the `if` that rolls back, while the
    // `#markSeen(seenBefore)` restore sits *inside* it -- so a
    // refusal being logged does not by itself mean `lastSeen` was restored. The
    // rollback is conditional on `filePath` and on the seen-generation, and
    // teardown closes the watcher while the delivery is still in flight -- a
    // case #75's "a change the consumer refuses is not recorded as seen" does
    // not cover, since nothing there is being disposed mid-delivery.
    //
    // Driven through the real settle path rather than by calling the private
    // `#deliver`, so the conditions are the ones production actually presents.
    const entered = deferred();
    let onChangeCalls = 0;
    const watcher = new FileWatcher({
        log: (line) => logLines.push(line),
        onChange: async () => {
            onChangeCalls += 1;
            // Exactly what ViewerInstance.close() does to its watcher, at
            // exactly the moment it can do it: with a delivery already in
            // flight.
            watcher.close();
            entered.resolve();
            throw new ViewerClosedError();
        },
    });

    await watcher.watch(docPath);
    await watcher.acknowledge();
    const seenBefore = watcher.lastSeen;
    assert.ok(seenBefore, "precondition: the watcher recorded a starting fingerprint");

    await touch();
    await entered.promise;
    await waitFor(() => logLines.some((line) => line.includes("not consumed")), "the refusal to be logged", 3000);

    assert.equal(onChangeCalls, 1);
    assert.equal(
        watcher.lastSeen,
        seenBefore,
        "the change was left marked as seen, so nothing would ever re-fire it",
    );
});

test("a change whose delivery begins after close() is refused, not silently consumed", async () => {
    // The guard at the top of `#autoRefresh` -- reachable, and until this test
    // existed, unreached by the suite. Every other test here closes during the
    // *backoff*, so refusal comes from `refresh()`'s own entry check and the
    // delivery-path guard is never the one that fires. A mutant turning that
    // guard into a clean `return` therefore passed six other tests.
    //
    // The race is won by construction rather than by sleeping on a guess, but
    // note which timer is being held open. Churning from the start would keep
    // resetting the 300ms *debounce* -- `#schedule` clears and re-arms it on
    // every event -- so `#settle` would never begin, and `close()` would then
    // cancel the pending timer outright and nothing would ever be delivered.
    // (That is not hypothetical: it is what the first version of this test did,
    // and it failed.) The touches therefore start only once the settle pass is
    // already running, where `#schedule` merely sets `#pendingWhileSettling`.
    // From there the pass cannot complete, because it needs SETTLE_STABLE_ROUNDS
    // (2) consecutive 200ms polls in which mtime and size do not move.
    const viewer = await openViewer();
    await touch();

    // 400ms: past the 300ms debounce, so the settle pass has begun; before the
    // ~700ms at which the earliest possible delivery could occur.
    await sleep(400);
    let churn = setInterval(() => void touch(), 100);

    // 1200ms in, the pass is pinned open and delivery certainly has not
    // happened. `close()` lands squarely inside it.
    await sleep(800);
    await viewer.close();

    await sleep(200);
    clearInterval(churn);
    churn = null;

    await waitFor(
        () => logLines.some((line) => line.includes("not consumed")),
        "the delivery to be refused by the closed viewer",
        5000,
    );

    // Reports its own precondition. If the timing had slipped and `close()` had
    // landed after delivery began, the refusal would have come from the retry
    // loop instead and this test would be quietly re-testing an earlier one.
    assert.equal(cache.refreshCalls, 0, "a refresh was started, so this is not the delivery-path guard being tested");
    const refusal = logLines.find((line) => line.includes("not consumed"));
    assert.match(refusal, /closed while a refresh was in flight/);
});

test("a refresh already inside the cache when close() lands is not acted on", async () => {
    // The case a disposal check at the top of the retry loop cannot see: the
    // check has already passed, the call is inside Word, and `close()` happens
    // while it is there. Everything after the result -- the PDF export, the doc
    // assignment, the broadcast -- would then run for a panel that is gone.
    const viewer = await openViewer();
    const gate = deferred();
    cache.gate = gate.promise;

    const pending = viewer.refresh({ force: true });
    await waitFor(() => cache.inFlight > 0, "cache.refresh to be entered");

    const pdfCallsAtClose = cache.pdfCalls;
    const docAtClose = viewer.doc;

    await viewer.close();
    gate.resolve();

    await assert.rejects(pending, (err) => {
        assert.equal(err.code, "viewer_closed");
        return true;
    });
    assert.equal(cache.pdfCalls, pdfCallsAtClose, "a disposed viewer re-rendered the document");
    assert.equal(viewer.doc, docAtClose, "a disposed viewer replaced the document state it was torn down with");
});

test("a closed viewer refuses a refresh outright, rather than falling through not_open", async () => {
    // `close()` deliberately does not null `this.doc` (out of scope in #81), so
    // without an explicit check a closed viewer refreshes normally -- which is
    // the route #80 measured and attributed to the `not_open` guard it never
    // reached.
    const viewer = await openViewer();
    await viewer.close();

    assert.ok(viewer.doc, "precondition: close() left this.doc set, so not_open cannot be what refuses");
    const callsBefore = cache.refreshCalls;
    await assert.rejects(
        () => viewer.refresh({ force: true }),
        (err) => {
            assert.equal(err.code, "viewer_closed");
            return true;
        },
    );
    assert.equal(cache.refreshCalls, callsBefore, "a closed viewer reached the cache");
});
