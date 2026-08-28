// Office-free tests for the snapshot store.
//
// Snapshots are the only mechanism behind revert -- Word's in-process undo dies
// when the document is closed at the end of every operation -- so retention and
// ordering are load-bearing, not housekeeping.

import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import {
    compactStamp,
    latestSnapshot,
    listSnapshots,
    MAX_SNAPSHOTS_PER_DOCUMENT,
    parseSnapshotName,
    pruneSnapshots,
    revertToLatest,
    snapshotDirFor,
    snapshotName,
    SnapshotError,
    takeSnapshot,
} from "../../src/word/snapshots.mjs";

const withTemp = async (fn) => {
    const dir = await mkdtemp(path.join(tmpdir(), "snapshots-test-"));
    try {
        return await fn(dir);
    } finally {
        await rm(dir, { recursive: true, force: true }).catch(() => {});
    }
};

test("names sort oldest-first as plain strings", () => {
    const early = snapshotName({ takenAt: new Date("2026-01-02T03:04:05.006Z"), op: "replace_text", token: "sha256:aa" });
    const late = snapshotName({ takenAt: new Date("2026-01-02T03:04:05.007Z"), op: "replace_text", token: "sha256:aa" });
    assert.ok(early < late, `${early} should sort before ${late}`);
    assert.match(early, /^20260102T030405006Z-replace_text-aa\.snapshot$/);
});

test("the timestamp keeps millisecond resolution", () => {
    assert.equal(compactStamp(new Date("2026-12-31T23:59:59.999Z")), "20261231T235959999Z");
});

test("names round-trip through the parser", () => {
    const name = snapshotName({
        takenAt: new Date("2026-08-28T18:30:12.123Z"),
        op: "insert_paragraph_after",
        token: "sha256:0123456789abcdef",
        nonce: "k9x2",
    });
    assert.deepEqual(parseSnapshotName(name), {
        name,
        stamp: "20260828T183012123Z",
        op: "insert_paragraph_after",
        token: "0123456789abcdef",
        nonce: "k9x2",
    });
});

test("the parser ignores anything that is not a snapshot", () => {
    for (const name of ["notes.txt", "20260828T183012123Z-replace_text-aa.snapshot.json", "random.snapshot"]) {
        assert.equal(parseSnapshotName(name), null, `${name} should not parse as a snapshot`);
    }
});

test("retention keeps the newest and discards the rest", () => {
    const names = Array.from({ length: 25 }, (_, i) =>
        snapshotName({ takenAt: new Date(Date.UTC(2026, 0, 1, 0, 0, i)), op: "replace_text", token: "sha256:aa" }),
    );
    const { keep, discard } = pruneSnapshots(names, 20);
    assert.equal(keep.length, 20);
    assert.equal(discard.length, 5);
    // Newest first, and the discarded ones are strictly older than the kept.
    assert.ok(keep[0].name > keep[19].name);
    assert.ok(discard.every((d) => d.name < keep[19].name));
});

test("retention ignores unrelated files in the directory", () => {
    const { keep, discard } = pruneSnapshots(["readme.md", "20260101T000000000Z-replace_text-aa.snapshot"], 20);
    assert.equal(keep.length, 1);
    assert.equal(discard.length, 0);
});

test("two documents get separate directories, and case does not split one", () => {
    const a = snapshotDirFor("/root", "C:\\Docs\\Report.docx");
    const b = snapshotDirFor("/root", "c:\\docs\\report.docx");
    const c = snapshotDirFor("/root", "C:\\Docs\\Other.docx");
    assert.equal(a, b, "the same document in different case must share a history");
    assert.notEqual(a, c);
});

test("a snapshot copies the bytes and records a manifest", async () => {
    await withTemp(async (dir) => {
        const doc = path.join(dir, "report.docx");
        await writeFile(doc, "original bytes");

        const snap = await takeSnapshot({
            root: dir,
            docPath: doc,
            op: "replace_text",
            token: "sha256:deadbeef",
            description: "replace the text of p:abc",
        });

        assert.equal(await readFile(snap.file, "utf8"), "original bytes");
        const manifest = JSON.parse(await readFile(`${snap.file}.json`, "utf8"));
        assert.equal(manifest.op, "replace_text");
        assert.equal(manifest.revisionToken, "sha256:deadbeef");
        assert.equal(manifest.description, "replace the text of p:abc");
        assert.equal(manifest.bytes, "original bytes".length);
    });
});

test("revert restores the newest snapshot and consumes it", async () => {
    await withTemp(async (dir) => {
        const doc = path.join(dir, "report.docx");

        await writeFile(doc, "v1");
        await takeSnapshot({ root: dir, docPath: doc, op: "replace_text", token: "sha256:1", nonce: "a" });
        await writeFile(doc, "v2");
        await takeSnapshot({ root: dir, docPath: doc, op: "replace_text", token: "sha256:2", nonce: "b" });
        await writeFile(doc, "v3");

        assert.equal((await latestSnapshot({ root: dir, docPath: doc })).revisionToken, "sha256:2");

        const first = await revertToLatest({ root: dir, docPath: doc });
        assert.equal(await readFile(doc, "utf8"), "v2");
        assert.equal(first.remaining, 1);

        // Consumed, so a second revert steps further back rather than toggling
        // between the last two states.
        const second = await revertToLatest({ root: dir, docPath: doc });
        assert.equal(await readFile(doc, "utf8"), "v1");
        assert.equal(second.remaining, 0);

        await assert.rejects(
            () => revertToLatest({ root: dir, docPath: doc }),
            (err) => err instanceof SnapshotError && err.code === "no_snapshot",
        );
    });
});

test("history is capped and keeps the newest", async () => {
    await withTemp(async (dir) => {
        const doc = path.join(dir, "report.docx");
        const total = MAX_SNAPSHOTS_PER_DOCUMENT + 5;
        for (let i = 0; i < total; i++) {
            await writeFile(doc, `v${i}`);
            await takeSnapshot({
                root: dir,
                docPath: doc,
                op: "replace_text",
                token: `sha256:${i}`,
                nonce: String(i).padStart(4, "0"),
            });
        }

        const history = await listSnapshots({ root: dir, docPath: doc });
        assert.equal(history.length, MAX_SNAPSHOTS_PER_DOCUMENT);

        // The most recent edit is still undoable; the oldest are gone.
        assert.equal(history[0].revisionToken, `sha256:${total - 1}`);
        assert.equal(history.at(-1).revisionToken, `sha256:${total - MAX_SNAPSHOTS_PER_DOCUMENT}`);
    });
});

test("listing a document with no history is empty rather than an error", async () => {
    await withTemp(async (dir) => {
        assert.deepEqual(await listSnapshots({ root: dir, docPath: path.join(dir, "never-edited.docx") }), []);
        assert.equal(await latestSnapshot({ root: dir, docPath: path.join(dir, "never-edited.docx") }), null);
    });
});
