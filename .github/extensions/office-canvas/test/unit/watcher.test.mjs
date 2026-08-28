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
    const watcher = new FileWatcher({ onChange: (info) => changes.push(info), debounceMs: 50 });
    await watcher.watch(file);
    try {
        await fn({ dir, file, changes, watcher });
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
