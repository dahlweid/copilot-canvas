// Recovering the Word surface after it has been shut down (#61). Office-free.
//
// ## The defect
//
// The extension keeps one RenderCache, and therefore one hidden Word, and
// throws it away whenever nothing needs Word: the last canvas closes, or the
// idle timer fires. `getCache()` builds a fresh one when the slot is empty, so
// on paper the next caller recovers.
//
// It did not, because the slot was emptied *after* the teardown finished
// instead of before it started:
//
//     await cache?.dispose().catch(() => {});   // seconds
//     cache = null;                             // only now
//
// `WordHost.dispose` sends `quit` under a 20 s timeout and then waits up to 5 s
// more for the child to exit, and it sets `#disposed` inside that span, not at
// the end of it. So for up to ~25 s the slot handed every caller a host that
// was already answering `The Word host has been shut down.` -- which is exactly
// the sentence the canvas showed and, once #45 discarded it, the bare
// `Tool execution failed` the tool caller got.
//
// For a tool call that is a bad half-minute. For a canvas it was permanent:
// `ViewerInstance` kept the reference it was constructed with, so a panel
// opened inside the window held the dead host for the rest of the session.
// That is the "three canvases open, all dead, a retry does not heal it" shape.
//
// The idle timer was *not* the trigger. It is gated on `instances.size > 0` and
// canvases were open; the disposal that matters here is the one `onClose` runs
// when the last panel goes, and the panels that came back afterwards.
//
// ## How this reaches the real wiring
//
// The bug is in extension.mjs, which imports the SDK -- unresolvable on disk --
// and is therefore not importable from a test as it stands. `module.registerHooks`
// supplies a stand-in for the SDK transport and swaps `RenderCache` for one that
// never starts Word, and nothing else: every canvas handler, the slot, the
// lifecycle and `ViewerInstance` are the committed code.
//
// Mutations used to confirm these can go red (results in the PR body):
//   A. src/render-cache-slot.mjs, `dispose()` -> await the teardown *before*
//      clearing `current`, which is what `extension.mjs` did on main @ 2625a9e.
//      Measured: 3 pass, 5 fail. Both headline tests red, with the canvas one
//      failing on the verbatim reported sentence.
//   B. src/server.mjs, `ViewerInstance` -> resolve the cache once in the
//      constructor and capture it, which is what `cache: getCache()` did.
//      Measured: 7 pass, 1 fail -- "a panel resolves the live cache rather than
//      the one it was constructed with".

import test, { after, before } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { registerHooks } from "node:module";

import { ViewerInstance } from "../../src/server.mjs";
import { createRenderCacheSlot } from "../../src/render-cache-slot.mjs";

const SDK_SPECIFIER = "@github/copilot-sdk/extension";
const stubSdkUrl = new URL("./stub-sdk.mjs", import.meta.url).href;
const stubCacheUrl = new URL("./stub-render-cache.mjs", import.meta.url).href;

registerHooks({
    resolve(specifier, context, nextResolve) {
        if (specifier === SDK_SPECIFIER) return { url: stubSdkUrl, shortCircuit: true };
        const resolved = nextResolve(specifier, context);
        // Only extension.mjs's own import is swapped. src/server.mjs imports the
        // same module for `DocumentError` and `normalizeDocPath`, and the stub
        // re-exports them from the real one -- redirecting those too would be a
        // cycle.
        if (
            resolved.url.endsWith("/src/render-cache.mjs") &&
            String(context.parentURL ?? "").endsWith("/extension.mjs")
        ) {
            return { ...resolved, url: stubCacheUrl, shortCircuit: true };
        }
        return resolved;
    },
});

let home;
let docPath;
let canvas;
let stub;

before(async () => {
    home = await mkdtemp(path.join(tmpdir(), "office-canvas-recovery-"));
    // artifactsRoot() is read at cache construction; keep it out of the user's.
    process.env.COPILOT_HOME = home;
    docPath = path.join(home, "report.docx");
    // Real bytes: ViewerInstance watches the file it opened. Word never sees it.
    await writeFile(docPath, "not a real docx");

    stub = await import(stubCacheUrl);
    const sdk = await import(stubSdkUrl);
    await import("../../extension.mjs");
    canvas = sdk.joined.canvases[0];
    assert.ok(canvas?.open && canvas?.onClose, "the extension must have registered the word-doc canvas");
});

const openPanels = new Set();

after(async () => {
    for (const instanceId of openPanels) await canvas.onClose({ instanceId }).catch(() => {});
    if (home) await rm(home, { recursive: true, force: true });
});

async function openPanel(instanceId, input = { path: docPath }) {
    // Registered before the await: `open` binds the panel's server and records
    // the instance before it opens the document, so a rejection still leaves a
    // panel that `after()` has to close or the run never exits.
    openPanels.add(instanceId);
    return canvas.open({ instanceId, input });
}

function closePanel(instanceId) {
    openPanels.delete(instanceId);
    return canvas.onClose({ instanceId });
}

/** Closes every panel, so the next test starts with the slot genuinely empty. */
async function drainPanels() {
    for (const instanceId of [...openPanels]) await closePanel(instanceId);
}

/** Polls until `predicate` holds. Bounded so a failure fails rather than hangs. */
async function until(predicate, what, timeoutMs = 5000) {
    const deadline = Date.now() + timeoutMs;
    while (!predicate()) {
        if (Date.now() > deadline) assert.fail(`timed out waiting for ${what}`);
        await new Promise((resolve) => setImmediate(resolve));
    }
}

test("a canvas opened while Word is still shutting down is not handed the dead host", async () => {
    await drainPanels();
    stub.resetCaches();

    await openPanel("panel-a");
    assert.equal(stub.built.length, 1, "one cache, built on first use");
    const first = stub.built[0];
    assert.deepEqual(first.used, ["open", "pdf"]);

    // The last panel closes, so the extension shuts Word down -- and that takes
    // seconds, which is the whole point.
    const release = stub.holdDisposals();
    const closing = closePanel("panel-a");
    try {
        await until(() => first.disposed, "the teardown to reach the host");

        // The assertion is only worth anything if the cache being retired would
        // in fact produce the reported symptom. It does:
        await assert.rejects(() => first.open(docPath), /^Error: The Word host has been shut down\.$/);

        // Now the panel comes back, inside the teardown window. On main this
        // threw `The Word host has been shut down.` and the panel stayed dead
        // for the rest of the session.
        await openPanel("panel-b");

        const second = stub.built.at(-1);
        assert.notEqual(second, first, "the second panel must not be handed the cache being disposed");
        assert.equal(second.disposed, false);
        assert.deepEqual(second.used, ["open", "pdf"], "the replacement served the open, so it is really live");

        release();
        await closing;

        assert.equal(second.disposed, false, "finishing the teardown must not take the cache that replaced it");
        assert.equal(stub.built.length, 2, "exactly one replacement, not one per call");
    } finally {
        // A failure must fail, not hang: the gate is holding a teardown that
        // `after()` and every later test are waiting behind.
        release();
        await closing.catch(() => {});
    }
});

test("a tool call arriving during the teardown gets a live host too", async () => {
    await drainPanels();
    stub.resetCaches();

    // A panel is what makes the extension build a cache without Word; the tool
    // path and the canvas path share the same slot, so this is that slot seen
    // from the other side.
    await openPanel("panel-c");
    const first = stub.built[0];

    const release = stub.holdDisposals();
    const closing = closePanel("panel-c");
    try {
        await until(() => first.disposed, "the teardown to reach the host");

        const readDocument = (await import(stubSdkUrl)).joined.tools.find((t) => t.name === "read_document");
        assert.ok(readDocument, "read_document must be registered");

        const result = await readDocument.handler({ path: docPath });
        assert.deepEqual(result, { paragraphs: [], paragraphCount: 0, truncated: false });
        assert.equal(stub.built.length, 2, "the read must have been served by a fresh cache");
        assert.equal(stub.built[1].disposed, false);
        assert.deepEqual(stub.built[1].used, ["readStructure"]);
    } finally {
        release();
        await closing.catch(() => {});
    }
});

test("a panel resolves the live cache rather than the one it was constructed with", async () => {
    // The other half of the defect, at its own level: even with the slot fixed,
    // a panel that captured a reference would keep a disposed host forever,
    // because nothing in the code ever replaced it.
    stub.resetCaches();
    const slot = createRenderCacheSlot({ create: () => new stub.RenderCache({}) });

    const instance = new ViewerInstance({ cache: () => slot.get(), instanceId: "p", workspacePath: null });
    await instance.start();
    try {
        await instance.openDocument(docPath);
        const first = stub.built[0];
        assert.deepEqual(first.used, ["open", "pdf"]);

        await slot.dispose();
        assert.equal(first.disposed, true);

        await instance.openDocument(docPath);
        assert.equal(stub.built.length, 2);
        assert.equal(stub.built[1].disposed, false);
        assert.deepEqual(stub.built[1].used, ["open", "pdf"], "the panel used the replacement, not what it captured");
    } finally {
        await instance.close();
    }
});

test("a plain RenderCache still works as a constructor argument", async () => {
    // The resolver is what the extension passes; the existing tests and the
    // integration smokes pass a cache directly, and both forms have to work.
    stub.resetCaches();
    const cache = new stub.RenderCache({});
    const instance = new ViewerInstance({ cache, instanceId: "p", workspacePath: null });
    assert.equal(instance.cache, cache);
});

// --- the slot on its own ---------------------------------------------------

test("the slot is empty the instant a disposal starts, not when it finishes", async () => {
    let release;
    const gate = new Promise((resolve) => {
        release = resolve;
    });
    let built = 0;
    const slot = createRenderCacheSlot({
        create: () => ({ id: ++built, disposed: false, async dispose() {
            this.disposed = true;
            await gate;
        } }),
    });

    const first = slot.get();
    const disposing = slot.dispose();
    assert.equal(slot.peek(), null, "peek must not report a cache that is being torn down");

    const second = slot.get();
    assert.notEqual(second, first);
    assert.equal(second.disposed, false);

    release();
    await disposing;
    assert.equal(first.disposed, true);
    assert.equal(second.disposed, false, "the disposal must not follow the slot to its new occupant");
});

test("disposing waits for a teardown already in flight", async () => {
    // Process shutdown needs this: returning while `quit` is still in flight
    // leaves a hidden Word behind.
    let release;
    const gate = new Promise((resolve) => {
        release = resolve;
    });
    let finished = false;
    const slot = createRenderCacheSlot({
        create: () => ({
            async dispose() {
                await gate;
                finished = true;
            },
        }),
    });

    slot.get();
    const first = slot.dispose();
    const second = slot.dispose(); // e.g. the last onClose racing process shutdown

    assert.equal(slot.retiring, true);
    release();
    await Promise.all([first, second]);
    assert.equal(finished, true);
    assert.equal(slot.retiring, false);
});

test("a disposal is attempted once per cache, however many callers ask", async () => {
    let disposals = 0;
    const slot = createRenderCacheSlot({
        create: () => ({
            async dispose() {
                disposals += 1;
            },
        }),
    });

    slot.get();
    await Promise.all([slot.dispose(), slot.dispose(), slot.dispose()]);
    assert.equal(disposals, 1, "an empty slot has nothing to dispose");
});

test("a failing teardown is reported and does not escape", async () => {
    const lines = [];
    const slot = createRenderCacheSlot({
        create: () => ({
            async dispose() {
                throw new Error("Word was already gone");
            },
        }),
        log: (m) => lines.push(m),
    });

    slot.get();
    await slot.dispose(); // must not reject
    assert.match(lines.join("\n"), /Word was already gone/);
    assert.notEqual(slot.get(), null, "and the slot still hands out a fresh cache afterwards");
});
