// When the change overlay is shown, and -- mostly -- when it is taken away.
//
// Office-free: `ViewerInstance` is driven with a fake `RenderCache`, which is
// the seam the render key comes through. Every test here is about the *lifetime*
// of a record rather than its content; a record that outlives the render it
// describes is a highlight sitting over text that has moved, which is the
// address rule (ADR 0006) failing at the display instead of at the tool.

import { test, after, beforeEach } from "node:test";
import assert from "node:assert/strict";
import path from "node:path";

import { ViewerInstance } from "../../src/server.mjs";

const DOC_A = path.join(process.cwd(), "a.docx");
const DOC_B = path.join(process.cwd(), "b.docx");

/**
 * A RenderCache stand-in. `key` is the only field the overlay's lifetime turns
 * on, so it is the only one the tests move.
 */
class FakeCache {
    wordVersion = "fake";
    keys = new Map();
    refreshCalls = 0;
    /** One-shot gates; each `refresh` call consumes the first still queued. */
    gates = [];

    #info(docPath) {
        if (!this.keys.has(docPath)) this.keys.set(docPath, `${path.basename(docPath)}-1`);
        return { path: docPath, key: this.keys.get(docPath), name: path.basename(docPath), pageCount: 1 };
    }

    /** Simulates a write: the render key moves, as it does on any mtime change. */
    touch(docPath) {
        const next = `${path.basename(docPath)}-${Date.now()}-${Math.random()}`;
        this.keys.set(docPath, next);
        return next;
    }

    /**
     * Holds the next `refresh` open until the returned function is called, so a
     * second edit can arrive while the first render is still in flight.
     */
    blockNext() {
        let release;
        this.gates.push(new Promise((resolve) => {
            release = resolve;
        }));
        return release;
    }

    async open(docPath) {
        return this.#info(docPath);
    }
    /**
     * Shaped like the real one, which returns `{ file }` -- `#servePdf`
     * destructures it. Returning nothing here made the destructure throw a
     * `TypeError` that the route reports as `render_failed`, a code naming a
     * cause nothing had determined.
     */
    async pdf(docPath) {
        return { file: `${docPath}.pdf`, key: this.keys.get(docPath) ?? null };
    }
    async refresh(docPath) {
        this.refreshCalls += 1;
        // Read *before* the gate. A render reports the file as it stood when it
        // began, not as it stands when it finishes -- resolving the key after the
        // wait would make a stale render unconstructible, which is the one state
        // these tests need to be able to build.
        const snapshot = this.#info(docPath);
        const gate = this.gates.shift();
        if (gate) await gate;
        // Defaults to a render, which is what most tests want. Set `changed` to
        // false to exercise the no-op path: `cache.refresh` reporting that the
        // file has not moved is what makes `ViewerInstance.refresh` return
        // early *without* advancing `this.doc`.
        const changed = this.nextChanged ?? true;
        return { ...snapshot, changed };
    }
}

/** Spins the event loop until `predicate` holds, or gives up and returns false. */
async function until(predicate, limit = 200) {
    for (let i = 0; i < limit; i++) {
        if (predicate()) return true;
        await new Promise((resolve) => setImmediate(resolve));
    }
    return predicate();
}

const record = (over = {}) => ({
    op: "replace_text",
    page: 2,
    text: "The replacement text.",
    locatable: true,
    at: "2024-01-01T00:00:00.000Z",
    ...over,
});

const instances = [];
let cache;

beforeEach(() => {
    cache = new FakeCache();
});
after(async () => {
    for (const instance of instances) await instance.close();
});

function viewer() {
    const instance = new ViewerInstance({ cache, instanceId: "test", workspacePath: null });
    instances.push(instance);
    return instance;
}

test("a change survives the refresh that produced it", async () => {
    const viewerInstance = viewer();
    await viewerInstance.openDocument(DOC_A);
    cache.touch(DOC_A);
    await viewerInstance.refresh({ force: true, change: record() });
    assert.deepEqual(viewerInstance.state.change?.text, "The replacement text.");
});

test("a change is dropped once the document is rendered again", async () => {
    // The discriminating case for the whole design: the record is not merely
    // stored, it is tied to one render. A later render is a different image and
    // the box's coordinates no longer describe it.
    const viewerInstance = viewer();
    await viewerInstance.openDocument(DOC_A);
    await viewerInstance.refresh({ force: true, change: record() });
    assert.ok(viewerInstance.state.change, "precondition: the record is showing");

    cache.touch(DOC_A);
    await viewerInstance.refresh({ force: true });
    assert.equal(viewerInstance.state.change, null);
});

test("a render that changes nothing keeps the overlay up", async () => {
    // ...and the other half of the discriminator. If the key were ignored the
    // test above would pass by the record being cleared on every refresh, which
    // would also take the overlay away the instant anything touched the file.
    const viewerInstance = viewer();
    await viewerInstance.openDocument(DOC_A);
    await viewerInstance.refresh({ force: true, change: record() });
    const key = viewerInstance.doc.key;

    await viewerInstance.refresh({ force: true });
    assert.equal(viewerInstance.doc.key, key, "precondition: the render key did not move");
    assert.ok(viewerInstance.state.change, "an unchanged render must not take the overlay down");
});

test("an edit that joins an older render waits for one that can contain it", async () => {
    // Production shape: every caller carrying a record forces (extension.mjs's
    // refreshCanvasesFor). Two rapid edits -- A lands, its render starts, B lands
    // and joins it. The join reports `changed: true`, but that image is post-A and
    // pre-B, so settling for it would stamp B's overlay onto A's page. `changed`
    // proves a render happened, not which one.
    const viewerInstance = viewer();
    await viewerInstance.openDocument(DOC_A);

    const release = cache.blockNext();
    const callsBefore = cache.refreshCalls;

    cache.touch(DOC_A); // edit A lands on disk
    const firstRefresh = viewerInstance.refresh({ force: true, change: record({ text: "Edit A." }) });
    assert.ok(
        await until(() => cache.refreshCalls > callsBefore),
        "precondition: the first refresh reached the cache and is being held there",
    );

    const freshKey = cache.touch(DOC_A); // edit B lands while A's render is in flight
    const secondRefresh = viewerInstance.refresh({ force: true, change: record({ text: "Edit B." }) });

    release();
    await Promise.all([firstRefresh, secondRefresh]);

    assert.equal(viewerInstance.state.change?.text, "Edit B.", "precondition: B's record is the one showing");
    assert.equal(
        viewerInstance.doc.key,
        freshKey,
        "B's overlay was published against a render that began before B existed",
    );
});

test("a render that completes after a newer edit does not stamp that edit", async () => {
    // The same pairing reached from the opposite side, and it needs its own test
    // because it is a different line of code. `#stampChange` stamps whatever
    // `this.change` currently holds -- so when A's render finishes, the record it
    // finds there is B's, overwritten by the caller that joined it. Stamping then
    // ties B to A's image without B's caller ever having settled for it.
    //
    // Unforced, so nothing re-renders afterwards to paper over the stamp: the only
    // safe outcome is that no overlay is published at all.
    const viewerInstance = viewer();
    await viewerInstance.openDocument(DOC_A);

    const release = cache.blockNext();
    const callsBefore = cache.refreshCalls;

    cache.touch(DOC_A);
    const firstRefresh = viewerInstance.refresh({ change: record({ text: "Edit A." }) });
    assert.ok(
        await until(() => cache.refreshCalls > callsBefore),
        "precondition: the first refresh reached the cache and is being held there",
    );

    cache.touch(DOC_A);
    const secondRefresh = viewerInstance.refresh({ change: record({ text: "Edit B." }) });

    release();
    await Promise.all([firstRefresh, secondRefresh]);

    assert.equal(
        cache.refreshCalls,
        callsBefore + 1,
        "precondition: the second edit joined the in-flight render rather than starting its own",
    );
    assert.equal(
        viewerInstance.state.change,
        null,
        "an overlay was published against a render that cannot contain the edit it describes",
    );
});

test("a forced refresh with no record still refuses the no-op it joined", async () => {
    // The other half of the join condition. The test above covers a caller
    // carrying a record; this one carries none, so only `joined.changed` stands
    // between it and settling for a refresh that rendered nothing. A watcher echo
    // that fires before Word has finished saving is exactly that no-op, and the
    // canvas's own force-refresh button is exactly this caller.
    const viewerInstance = viewer();
    await viewerInstance.openDocument(DOC_A);

    const release = cache.blockNext();
    const callsBefore = cache.refreshCalls;
    cache.nextChanged = false; // the in-flight refresh finds a file that has not moved

    const first = viewerInstance.refresh({});
    assert.ok(
        await until(() => cache.refreshCalls > callsBefore),
        "precondition: the first refresh reached the cache and is being held there",
    );
    const second = viewerInstance.refresh({ force: true });

    release();
    const [, secondResult] = await Promise.all([first, second]);

    assert.equal(cache.refreshCalls, callsBefore + 2, "the forced caller settled for a render that did not happen");
    assert.equal(secondResult.changed, true);
});

test("a second edit supersedes the first rather than accumulating", async () => {
    const viewerInstance = viewer();
    await viewerInstance.openDocument(DOC_A);
    await viewerInstance.refresh({ force: true, change: record({ text: "First." }) });
    cache.touch(DOC_A);
    await viewerInstance.refresh({ force: true, change: record({ text: "Second." }) });
    assert.equal(viewerInstance.state.change.text, "Second.");
});

test("an explicit null clears the overlay, as a revert does", async () => {
    const viewerInstance = viewer();
    await viewerInstance.openDocument(DOC_A);
    await viewerInstance.refresh({ force: true, change: record() });
    cache.touch(DOC_A);
    await viewerInstance.refresh({ force: true, change: null });
    assert.equal(viewerInstance.state.change, null);
});

test("an omitted change leaves the record alone, rather than clearing it", async () => {
    // `undefined` and `null` must not mean the same thing. Auto-refresh passes
    // neither and must not disturb an overlay; a revert passes null and must.
    const viewerInstance = viewer();
    await viewerInstance.openDocument(DOC_A);
    await viewerInstance.refresh({ force: true, change: record() });
    await viewerInstance.refresh({ force: true });
    assert.ok(viewerInstance.state.change, "a refresh with no opinion must not clear the record");
});

test("opening another document takes the overlay with it", async () => {
    const viewerInstance = viewer();
    await viewerInstance.openDocument(DOC_A);
    await viewerInstance.refresh({ force: true, change: record() });
    await viewerInstance.openDocument(DOC_B);
    assert.equal(viewerInstance.state.change, null);
    assert.equal(viewerInstance.change, null, "the record is dropped, not merely hidden by the key check");
});

test("reopening the same document also drops a stale overlay", async () => {
    // The key would not necessarily catch this one: reopening an untouched file
    // yields the same key, so without the explicit clear the overlay would come
    // back for an edit the user has since scrolled away from.
    const viewerInstance = viewer();
    await viewerInstance.openDocument(DOC_A);
    await viewerInstance.refresh({ force: true, change: record() });
    const key = viewerInstance.doc.key;
    await viewerInstance.openDocument(DOC_A);
    assert.equal(viewerInstance.doc.key, key, "precondition: the key did not move");
    assert.equal(viewerInstance.state.change, null);
});

test("a forced refresh does not settle for a no-op it joined", async () => {
    // The discriminating case the test above cannot reach, because its cache
    // always reports `changed: true`. A watcher echo can fire *before* Word has
    // finished saving: that refresh finds nothing moved and returns early
    // WITHOUT re-rendering and without advancing `this.doc`. The edit's own
    // forced refresh then joins it. Stamping the record against the joined
    // result would publish this edit's text over the *pre-edit* render -- the
    // hazard the comment above `refresh` names, arriving by the other door.
    const viewerInstance = viewer();
    await viewerInstance.openDocument(DOC_A);
    const staleKey = viewerInstance.doc.key;

    let release;
    const gate = new Promise((resolve) => {
        release = resolve;
    });
    const slowRefresh = cache.refresh.bind(cache);
    cache.refresh = async (docPath) => {
        await gate;
        return slowRefresh(docPath);
    };

    cache.nextChanged = false; // the echo sees a file that has not moved yet
    const first = viewerInstance.refresh({});
    await Promise.resolve();
    const second = viewerInstance.refresh({ force: true, change: record({ text: "Landed during a no-op." }) });

    const freshKey = cache.touch(DOC_A); // the save lands while both are in flight
    release();
    await Promise.all([first, second]);

    assert.equal(cache.refreshCalls, 2, "the forced caller must refresh for real rather than return the no-op");
    assert.equal(viewerInstance.doc.key, freshKey, "the render must advance past the pre-edit key");
    assert.notEqual(viewerInstance.doc.key, staleKey);
    assert.equal(
        viewerInstance.state.change?.text,
        "Landed during a no-op.",
        "and the record is stamped against that render, not the stale one",
    );
});

test("a record arriving during an in-flight refresh is not swallowed", async () => {
    // Two refreshes overlap all the time: the file watcher fires as an echo of
    // our own save, just as the edit's own refresh is starting. The second caller
    // must not lose its record to that overlap -- if it silently returned the
    // joined render the overlay would simply never appear for that edit.
    //
    // It does not join, and that is the point: a render already in flight began
    // before this record existed, so it cannot be shown to contain the edit.
    // A forced caller therefore pays for a second render rather than settling.
    const viewerInstance = viewer();
    await viewerInstance.openDocument(DOC_A);

    let release;
    const gate = new Promise((resolve) => {
        release = resolve;
    });
    const slowRefresh = cache.refresh.bind(cache);
    cache.refresh = async (docPath) => {
        await gate;
        return slowRefresh(docPath);
    };

    const first = viewerInstance.refresh({ force: true });
    await Promise.resolve();
    const second = viewerInstance.refresh({ force: true, change: record({ text: "Landed during." }) });
    release();
    await Promise.all([first, second]);

    assert.equal(cache.refreshCalls, 2, "the forced caller must render for real rather than settle for the join");
    assert.equal(viewerInstance.state.change?.text, "Landed during.");
});

test("an unlocatable record still reaches the viewer, so a page marker can be shown", async () => {
    const viewerInstance = viewer();
    await viewerInstance.openDocument(DOC_A);
    cache.touch(DOC_A);
    await viewerInstance.refresh({
        force: true,
        change: record({ op: "delete_paragraph", locatable: false, text: null }),
    });
    const change = viewerInstance.state.change;
    assert.equal(change.locatable, false);
    assert.equal(change.page, 2);
});

test("the state a client first sees carries the same verdict as a later push", async () => {
    // The SSE stream opens with a `state` frame built from the same getter, so a
    // canvas that connects after an edit must see the overlay too -- otherwise
    // it appears only for whoever was already listening.
    const viewerInstance = viewer();
    await viewerInstance.openDocument(DOC_A);
    cache.touch(DOC_A);
    await viewerInstance.refresh({ force: true, change: record() });

    const base = await viewerInstance.start();
    const res = await fetch(new URL("/api/state", base));
    const body = await res.json();
    assert.equal(body.change?.text, "The replacement text.");
});
