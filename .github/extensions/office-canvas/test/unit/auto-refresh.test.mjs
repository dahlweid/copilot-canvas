// Auto-refresh when the document changes on disk, and what happens when that
// refresh fails (#67).
//
// ## The defect
//
// "Open in Word -> change it -> Save -> wait for updates in the canvas" and the
// preview never reloaded. Auto-refresh is built and works: the watcher fires,
// `#autoRefresh` re-renders, the new state is pushed. What it did not survive
// was *one* failing attempt. `#autoRefresh` caught, set `status: "error"`, and
// stopped -- no retry, no re-arm -- and the watcher had already advanced its
// fingerprint past the change before dispatching it, so there was nothing left
// to re-fire on either. A panel was measured serving `file_locked` roughly three
// minutes after the save that produced it: a stuck state, not a passing one.
//
// **What made that attempt fail is not established.** Three panels were watching
// the same document at the time and contending on it, so "Word's post-save flush
// held the file" is plausible and unproven. These tests therefore assert what
// the viewer does with a failure, and never which failure Word produces.
//
// ## What layer these observe
//
// A `ViewerInstance`, server side. They assert that the refresh retries and that
// `state` ends up current -- nothing about a browser. That distinction is real
// and not pedantic: #66 measured that a viewer panel whose `EventSource` origin
// has died retries forever and never recovers, so it now closes its own source
// at a deadline instead. Recovered state reaches a client whose stream is still
// open, and any client connecting afterwards, since `/events` writes the current
// state on connect -- but it cannot reach one that has stopped listening. These
// tests do not, and cannot from here, say the browser repainted.
//
// Office-free: a fake `RenderCache` supplies the failures, and the file on disk
// is a real temp file so the committed `FileWatcher` is the one under test. Word
// is never started.
// Run: node --test ".github/extensions/office-canvas/test/unit/*.test.mjs"

import test, { after, before, beforeEach } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { ViewerInstance } from "../../src/server.mjs";

// The watcher needs debounce + ~600 ms to settle, and then the retries run.
const BUDGET_MS = 15_000;

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function waitFor(predicate, label, budgetMs = BUDGET_MS) {
    const deadline = Date.now() + budgetMs;
    while (Date.now() < deadline) {
        if (predicate()) return;
        await sleep(25);
    }
    assert.fail(`timed out after ${budgetMs}ms waiting for ${label}`);
}

/**
 * A RenderCache stand-in whose `refresh` fails to order.
 *
 * `failures` is a queue of codes; each `refresh` consumes one and throws a
 * `DocumentError`-shaped error for it. An empty queue means the file is
 * readable again, which is the whole point -- the condition these tests are
 * about is one that clears.
 */
class FakeCache {
    wordVersion = "fake";
    refreshCalls = 0;
    failures = [];
    /** What `refresh` reports once it stops failing. */
    nextChanged = true;
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
        return { file: `${docPath}.pdf`, key: this.#info(docPath).key };
    }
    async refresh(docPath) {
        this.refreshCalls += 1;
        const code = this.failures.shift();
        if (code) {
            const err = new Error(`${code} raised by the test`);
            err.code = code;
            throw err;
        }
        this.#version += 1;
        return { ...this.#info(docPath), changed: this.nextChanged };
    }
}

let home;
let dir;
let docPath;
let cache;
const instances = [];
const dirs = [];

before(async () => {
    // `addRecent` writes under $COPILOT_HOME on every open; keep it out of the
    // developer's real one.
    home = await mkdtemp(path.join(tmpdir(), "office-canvas-autorefresh-home-"));
    process.env.COPILOT_HOME = home;
});

beforeEach(async () => {
    // Close first: a watcher left running from the previous test would keep
    // delivering into a shared fake and make these tests read each other's
    // refresh counts.
    while (instances.length) await instances.pop().close();
    dir = await mkdtemp(path.join(tmpdir(), "office-canvas-autorefresh-"));
    dirs.push(dir);
    docPath = path.join(dir, "report.docx");
    // Real bytes: the viewer watches the file it opened. Word never sees it.
    await writeFile(docPath, "original");
    cache = new FakeCache();
});

after(async () => {
    while (instances.length) await instances.pop().close();
    for (const used of dirs) await rm(used, { recursive: true, force: true });
    if (home) await rm(home, { recursive: true, force: true });
});

/** A viewer with the retry backoff collapsed, so the test is about the loop. */
async function openViewer({ autoRefreshDelaysMs = [5, 5, 5] } = {}) {
    const instance = new ViewerInstance({
        cache,
        instanceId: "auto-refresh",
        workspacePath: null,
        autoRefreshDelaysMs,
    });
    instances.push(instance);
    await instance.openDocument(docPath);
    assert.equal(instance.status, "ready", "precondition: the document opened");
    return instance;
}

test("a refresh that fails once is retried, not latched", async () => {
    const viewer = await openViewer();
    cache.failures = ["file_locked"];

    await writeFile(docPath, "the user saved from Word");

    await waitFor(() => cache.refreshCalls >= 2, "a second attempt");
    await waitFor(() => viewer.status === "ready", "the viewer to recover");
    assert.equal(viewer.error, null, "a recovered viewer must not still be carrying the error it recovered from");
    assert.equal(cache.failures.length, 0, "precondition: the scripted failure was the one that was retried");
});

test("a failure that cannot clear is reported without burning the retries", async () => {
    // An ACL is exactly as true in five seconds as it is now -- `document-
    // reader.mjs` says so of `permission_denied` in as many words -- so retrying
    // one is only a slower way to tell the user. More failures are queued than
    // the backoff has attempts, so a retrying viewer would consume them and be
    // visible in the count.
    const viewer = await openViewer();
    cache.failures = ["permission_denied", "permission_denied", "permission_denied", "permission_denied"];

    await writeFile(docPath, "the user saved from Word");

    await waitFor(() => viewer.status === "error", "the failure to be reported");
    const callsWhenReported = cache.refreshCalls;
    assert.equal(viewer.error.code, "permission_denied");
    assert.equal(callsWhenReported, 1, "a permanent condition was retried");
});

test("a viewer that gave up refreshes again on the next save, even if nothing re-rendered", async () => {
    // The third part. A bounded retry still ends somewhere, and when it does the
    // viewer must be re-armed rather than dead: the watcher gets the refusal, puts
    // its fingerprint back, and the next write reports the change again.
    //
    // The recovering refresh reports `changed: false` on purpose. That is what a
    // settle landing on an event which moved no bytes produces, and `refresh`
    // clears the error only on the branch that actually re-renders -- so without
    // forcing, the viewer would return from a *successful* refresh still reading
    // `error` over a current image.
    const viewer = await openViewer();
    cache.failures = ["file_locked", "file_locked", "file_locked", "file_locked"];

    await writeFile(docPath, "the save that lands during a lock");
    await waitFor(() => viewer.status === "error", "the viewer to give up");
    assert.equal(viewer.error.code, "file_locked");
    assert.ok(
        cache.refreshCalls >= 4,
        `the backoff must attempt more than once before giving up, got ${cache.refreshCalls}`,
    );
    assert.equal(cache.failures.length, 0, "precondition: every scripted failure was consumed");

    const callsBefore = cache.refreshCalls;
    cache.nextChanged = false;
    await writeFile(docPath, "the next save, with the lock gone");

    await waitFor(() => viewer.status === "ready", "the later save to clear the error");
    assert.equal(viewer.error, null);
    assert.ok(cache.refreshCalls > callsBefore, "the recovery must have gone through a real refresh");
});

test("our own save is still not echoed back in as a change", async () => {
    // `acknowledge()` is what stops a refresh we performed from being reported
    // to us as a change and refreshing again forever. Retrying and re-arming
    // must not defeat it: this writes the file and refreshes it the way an edit
    // does, then waits out the window in which an echo would arrive.
    const viewer = await openViewer();

    await writeFile(docPath, "an edit we made ourselves");
    await viewer.refresh({ force: true });
    const callsAfterOurSave = cache.refreshCalls;

    await sleep(2500); // debounce + settle + a wide margin
    assert.equal(cache.refreshCalls, callsAfterOurSave, "our own write came back in as a change");
    assert.equal(viewer.status, "ready");
});
