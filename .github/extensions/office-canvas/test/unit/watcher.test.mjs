// File watching: the mechanism behind auto-refresh when a script regenerates a
// document.
//
// Office-free — this only ever touches the filesystem. Timings are generous and
// waits are poll-based rather than fixed sleeps, because CI machines are slower
// and noisier than a developer's.
// Run: node --test ".github/extensions/office-canvas/test/unit/*.test.mjs"

import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { FileWatcher } from "../../src/watcher.mjs";

// The watcher polls at 200 ms and wants 2 stable rounds, so a settled change
// needs debounce + ~600 ms before it is reported.
const SETTLE_BUDGET_MS = 4000;
const QUIET_MS = 1500;

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function waitFor(predicate, budgetMs, label) {
    const deadline = Date.now() + budgetMs;
    while (Date.now() < deadline) {
        if (predicate()) return;
        await sleep(50);
    }
    throw new Error(`timed out after ${budgetMs}ms waiting for ${label}`);
}

async function withWatcher(fn) {
    const dir = await mkdtemp(path.join(tmpdir(), "office-canvas-watch-"));
    const file = path.join(dir, "doc.docx");
    await writeFile(file, "original");
    const changes = [];
    // Queued refusals: each delivery consumes one and throws it, which is how a
    // consumer says "I did not take this change".
    const failures = [];
    const watcher = new FileWatcher({
        debounceMs: 50,
        log: () => {},
        onChange: (info) => {
            changes.push(info);
            const failure = failures.shift();
            if (failure) throw failure;
        },
    });
    await watcher.watch(file);
    try {
        await fn({ dir, file, changes, watcher, failures });
    } finally {
        watcher.close();
        await rm(dir, { recursive: true, force: true });
    }
}

test("a change to the watched file is reported", async () => {
    await withWatcher(async ({ file, changes }) => {
        await writeFile(file, "changed once");
        await waitFor(() => changes.length > 0, SETTLE_BUDGET_MS, "a change");
        assert.equal(changes[0].path, file);
    });
});

test("a multi-step write is reported once, not once per write", async () => {
    // A generating script writes in several passes. Re-rendering on each one
    // would open a half-written document in Word.
    await withWatcher(async ({ file, changes }) => {
        for (const chunk of ["part 1", "part 1 + 2", "part 1 + 2 + 3"]) {
            await writeFile(file, chunk);
            await sleep(60);
        }
        await waitFor(() => changes.length > 0, SETTLE_BUDGET_MS, "a change");
        await sleep(QUIET_MS);
        assert.equal(changes.length, 1, `expected exactly one change, got ${changes.length}`);
    });
});

test("Word's ~$ owner file never counts as a change", async () => {
    // Word drops ~$doc.docx beside the document and churns it constantly.
    // Treating that as a change would re-render in a loop.
    await withWatcher(async ({ dir, changes }) => {
        for (let i = 0; i < 3; i++) {
            await writeFile(path.join(dir, "~$doc.docx"), `owner ${i}`);
            await sleep(60);
        }
        await sleep(QUIET_MS);
        assert.equal(changes.length, 0, "the owner file must not trigger a re-render");
    });
});

test("an unrelated file in the same directory is ignored", async () => {
    // The watch is on the directory, so every sibling write arrives here.
    await withWatcher(async ({ dir, changes }) => {
        await writeFile(path.join(dir, "other.docx"), "unrelated");
        await writeFile(path.join(dir, "notes.txt"), "unrelated");
        await sleep(QUIET_MS);
        assert.equal(changes.length, 0);
    });
});

test("close stops delivery", async () => {
    await withWatcher(async ({ file, changes, watcher }) => {
        watcher.close();
        await writeFile(file, "after close");
        await sleep(QUIET_MS);
        assert.equal(changes.length, 0);
    });
});

test("watching a missing file does not throw", async () => {
    // The document can vanish between selection and watch; that is the
    // renderer's error to report, not a crash here.
    const dir = await mkdtemp(path.join(tmpdir(), "office-canvas-watch-"));
    const watcher = new FileWatcher({ onChange: () => {} });
    try {
        await watcher.watch(path.join(dir, "does-not-exist.docx"));
        assert.equal(watcher.lastSeen, null);
    } finally {
        watcher.close();
        await rm(dir, { recursive: true, force: true });
    }
});

test("watching a file in a missing directory does not throw", async () => {
    const watcher = new FileWatcher({ onChange: () => {}, log: () => {} });
    try {
        await watcher.watch(path.join(tmpdir(), "office-canvas-no-such-dir", "doc.docx"));
    } finally {
        watcher.close();
    }
});

// --- delivery: a change is consumed only once it has been taken (#67) --------

test("a change the consumer refuses is not recorded as seen", async () => {
    // The watcher used to advance `lastSeen` *before* dispatching, so a refresh
    // that threw consumed the edit that triggered it. The file then matched the
    // fingerprint the watcher believed it had already handled, and that edit
    // could never fire again -- which is what pinned the panel in `error` even
    // after the file became readable.
    await withWatcher(async ({ file, changes, watcher, failures }) => {
        const seenBefore = watcher.lastSeen;
        assert.notEqual(seenBefore, null, "precondition: the original file was fingerprinted");

        failures.push(new Error("locked at the moment the settle landed"));
        await writeFile(file, "changed");
        await waitFor(() => changes.length > 0, SETTLE_BUDGET_MS, "a delivery");
        await waitFor(() => watcher.lastSeen === seenBefore, SETTLE_BUDGET_MS, "the change to stay pending");
    });
});

test("a change our own save acknowledged is not rolled back over", async () => {
    // `acknowledge()` shares `lastSeen`: it records what *we* just wrote so the
    // echo of our own save does not come back in as a change and loop. A failed
    // delivery must therefore not restore the older value blindly -- by then the
    // field belongs to that save, not to this delivery.
    let watcher;
    let acknowledged = null;
    const dir = await mkdtemp(path.join(tmpdir(), "office-canvas-watch-"));
    const file = path.join(dir, "doc.docx");
    await writeFile(file, "original");
    watcher = new FileWatcher({
        debounceMs: 50,
        log: () => {},
        onChange: async () => {
            await watcher.acknowledge();
            acknowledged = watcher.lastSeen;
            throw new Error("the render failed after our save had been acknowledged");
        },
    });
    await watcher.watch(file);
    const seenBefore = watcher.lastSeen;
    try {
        await writeFile(file, "changed");
        await waitFor(() => acknowledged !== null, SETTLE_BUDGET_MS, "the delivery to acknowledge and fail");
        assert.notEqual(acknowledged, seenBefore, "precondition: acknowledge moved the fingerprint");
        await sleep(200);
        assert.equal(
            watcher.lastSeen,
            acknowledged,
            "the rollback overwrote an acknowledgement, so our own write will echo back in",
        );
    } finally {
        watcher.close();
        await rm(dir, { recursive: true, force: true });
    }
});

test("a slow consumer is never handed two changes at once", async () => {
    // Delivery is awaited, so `#settling` covers the consume as well as the
    // settle. That is what keeps the watcher from starting a second re-render
    // while the first is still running -- the overlapping-refresh hazard that
    // `viewer-state.test.mjs` pins from the other side.
    const dir = await mkdtemp(path.join(tmpdir(), "office-canvas-watch-"));
    const file = path.join(dir, "doc.docx");
    await writeFile(file, "original");

    let release;
    const gate = new Promise((resolve) => {
        release = resolve;
    });
    let inFlight = 0;
    let maxInFlight = 0;
    let delivered = 0;
    const watcher = new FileWatcher({
        debounceMs: 50,
        log: () => {},
        onChange: async () => {
            inFlight += 1;
            maxInFlight = Math.max(maxInFlight, inFlight);
            await gate;
            inFlight -= 1;
            delivered += 1;
        },
    });
    await watcher.watch(file);
    try {
        await writeFile(file, "one");
        await waitFor(() => inFlight === 1, SETTLE_BUDGET_MS, "the first delivery to start");

        // More writes land while the consumer is still busy with the first.
        for (const chunk of ["two", "three"]) {
            await writeFile(file, chunk);
            await sleep(60);
        }
        // Long enough for a second settle to have run and dispatched: the
        // watcher needs debounce + ~600 ms, so this is well over twice the
        // window in which the overlap would appear.
        await sleep(QUIET_MS);
        assert.equal(maxInFlight, 1, "a second change was dispatched while the first was still in flight");

        release();
        await waitFor(() => delivered >= 2, SETTLE_BUDGET_MS, "the coalesced second delivery");
        assert.equal(maxInFlight, 1, "the second delivery must wait for the first, not join it");
    } finally {
        release();
        watcher.close();
        await rm(dir, { recursive: true, force: true });
    }
});
