// Office-free tests for the edit orchestration.
//
// `DocumentEditor` takes its reader and its Word host as constructor arguments,
// so the parts that decide *whether the user's document is recoverable* can be
// tested without Word -- including the failure modes Word itself will not
// reproduce on demand, like a host that throws mid-save.
//
// These exist because the failure they cover is silent: a lost snapshot looks
// exactly like a successful edit until someone tries to revert.

import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readdir, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { DocumentEditor } from "../../src/word/document-editor.mjs";
import { fileRevisionToken } from "../../src/revision-token.mjs";
import { snapshotDirFor } from "../../src/word/snapshots.mjs";

const withTemp = async (fn) => {
    const dir = await mkdtemp(path.join(tmpdir(), "editor-test-"));
    try {
        return await fn(dir);
    } finally {
        await rm(dir, { recursive: true, force: true }).catch(() => {});
    }
};

/**
 * A reader that reports one addressable paragraph, whose token always matches
 * the file on disk. Enough to get `edit()` past its pre-flight and into the
 * host call, which is what these tests are about.
 */
const stubReader = (docPath, { address = "p:0123456789ab", text = "the paragraph" } = {}) => ({
    read: async () => ({
        path: docPath,
        name: path.basename(docPath),
        revisionToken: await fileRevisionToken(docPath),
        writable: true,
        paragraphCount: 1,
        paragraphs: [{ address, text, wordIndex: 1, index: 0, headingLevel: null, styleId: null, inTable: false }],
    }),
});

const snapshotFilesIn = async (root, docPath) =>
    (await readdir(snapshotDirFor(root, docPath)).catch(() => [])).filter((f) => f.endsWith(".snapshot"));

const intent = { op: "replace_text", address: "p:0123456789ab", text: "rewritten" };

test("a host that throws without touching the file discards the snapshot", async () => {
    // The document is provably unchanged, so the snapshot is a duplicate of what
    // is already on disk. Keeping it would spend a revert step undoing nothing.
    await withTemp(async (dir) => {
        const doc = path.join(dir, "report.docx");
        await writeFile(doc, "ORIGINAL");
        const root = path.join(dir, "artifacts");

        const editor = new DocumentEditor({
            reader: stubReader(doc),
            host: {
                edit: async () => {
                    throw Object.assign(new Error("Word did not respond"), { code: "word_timeout" });
                },
            },
            snapshotRoot: root,
        });

        const err = await editor
            .edit(doc, intent, { revisionToken: await fileRevisionToken(doc) })
            .then(() => null, (e) => e);

        assert.ok(err, "the failure was swallowed");
        assert.equal(err.code, "word_timeout");
        assert.equal(err.documentUnchanged, true);
        assert.equal(await readFile(doc, "utf8"), "ORIGINAL");
        assert.deepEqual(await snapshotFilesIn(root, doc), [], "a snapshot was kept for an edit that did nothing");
    });
});

test("a host that throws AFTER changing the file rolls the document back", async () => {
    // The defect this covers: the snapshot used to be deleted here. The host
    // threw, so the outcome is unknown -- which is exactly when the snapshot is
    // the only copy of the pre-edit bytes. The reachable sequence is a request
    // timing out and the PowerShell child being killed with `taskkill /F /T`
    // while Word is mid-Save on the user's original, leaving a half-written
    // document and a revert that answers `no_snapshot`.
    await withTemp(async (dir) => {
        const doc = path.join(dir, "report.docx");
        await writeFile(doc, "ORIGINAL");
        const root = path.join(dir, "artifacts");

        const editor = new DocumentEditor({
            reader: stubReader(doc),
            host: {
                edit: async () => {
                    await writeFile(doc, "HALF-WRITTEN");
                    throw Object.assign(new Error("Word did not respond"), { code: "word_timeout" });
                },
            },
            snapshotRoot: root,
        });

        const err = await editor
            .edit(doc, intent, { revisionToken: await fileRevisionToken(doc) })
            .then(() => null, (e) => e);

        assert.ok(err, "the failure was swallowed");
        assert.equal(err.rolledBack, true, "the document was not rolled back");
        assert.equal(await readFile(doc, "utf8"), "ORIGINAL", "the half-written bytes survived");
    });
});

test("when the rollback itself fails, the snapshot is kept and named", async () => {
    // The last line of defence. If we cannot restore the bytes we must at least
    // not delete them, and must say where they are.
    await withTemp(async (dir) => {
        const doc = path.join(dir, "report.docx");
        await writeFile(doc, "ORIGINAL");
        const root = path.join(dir, "artifacts");

        const editor = new DocumentEditor({
            reader: stubReader(doc),
            host: {
                edit: async () => {
                    await writeFile(doc, "HALF-WRITTEN");
                    // Removing the snapshot directory makes the restore fail the
                    // way a vanished or unreadable snapshot would.
                    await rm(snapshotDirFor(root, doc), { recursive: true, force: true });
                    throw Object.assign(new Error("Word did not respond"), { code: "word_timeout" });
                },
            },
            snapshotRoot: root,
        });

        const err = await editor
            .edit(doc, intent, { revisionToken: await fileRevisionToken(doc) })
            .then(() => null, (e) => e);

        assert.ok(err, "the failure was swallowed");
        assert.equal(err.rolledBack, false);
        assert.match(err.snapshot ?? "", /\.snapshot$/, "the error does not name the snapshot to recover from");

        // ...and again through the tool boundary's filter. `asToolError` in
        // extension.mjs forwards only `code`, `message` and `data`, so a fact
        // recorded solely as a top-level property is visible to this test and
        // invisible to the agent. The assertions above passed while the caller
        // learned nothing but "word_timeout"; these are the ones that matter.
        assert.match(
            err.data?.snapshot ?? "",
            /\.snapshot$/,
            "the snapshot name is not on `data`, so it is dropped at the tool boundary",
        );
        assert.equal(err.data?.rolledBack, false);
        assert.match(
            err.message,
            /may be partially written/,
            "the message does not tell the agent the document may be half-written",
        );
        assert.ok(err.message.includes(err.snapshot), "the message does not name the snapshot");
    });
});

test("a recovered edit tells the caller what happened to the document", async () => {
    // The three recovery outcomes are only useful if they reach the agent, and
    // the agent reads the message. A bare host error would say the call timed
    // out and leave the state of the user's document unstated.
    await withTemp(async (dir) => {
        const doc = path.join(dir, "report.docx");
        await writeFile(doc, "ORIGINAL");
        const root = path.join(dir, "artifacts");

        const editor = new DocumentEditor({
            reader: stubReader(doc),
            host: {
                edit: async () => {
                    await writeFile(doc, "HALF-WRITTEN");
                    throw Object.assign(new Error("Word did not respond"), { code: "word_timeout" });
                },
            },
            snapshotRoot: root,
        });

        const err = await editor
            .edit(doc, intent, { revisionToken: await fileRevisionToken(doc) })
            .then(() => null, (e) => e);

        assert.ok(err, "the failure was swallowed");
        assert.equal(err.data?.rolledBack, true, "the rollback is not reported on `data`");
        assert.match(err.message, /rolled back/, "the message does not say the document was rolled back");
        assert.equal(await readFile(doc, "utf8"), "ORIGINAL", "the document was not actually restored");
    });
});

test("an untouched document after a host throw is reported as needing no recovery", async () => {
    await withTemp(async (dir) => {
        const doc = path.join(dir, "report.docx");
        await writeFile(doc, "ORIGINAL");
        const root = path.join(dir, "artifacts");

        const editor = new DocumentEditor({
            reader: stubReader(doc),
            host: {
                edit: async () => {
                    throw Object.assign(new Error("Word did not respond"), { code: "word_timeout" });
                },
            },
            snapshotRoot: root,
        });

        const err = await editor
            .edit(doc, intent, { revisionToken: await fileRevisionToken(doc) })
            .then(() => null, (e) => e);

        assert.ok(err, "the failure was swallowed");
        assert.equal(err.data?.documentUnchanged, true, "the untouched outcome is not reported on `data`");
        assert.match(err.message, /unchanged/, "the message does not say the document was untouched");
    });
});

test("two concurrent edits of one document do not both act on the same snapshot", async () => {
    // A revision token authorizes an edit but cannot serialize one. Two calls
    // minted from a single read both hold a valid token and both pass every
    // pre-mutation check, because neither changes the other's paragraph text.
    // Unserialized, both apply while both snapshots hold the pre-first-edit
    // bytes -- so one revert silently undoes two edits and reports one.
    await withTemp(async (dir) => {
        const doc = path.join(dir, "report.docx");
        await writeFile(doc, "v0");
        const root = path.join(dir, "artifacts");

        let inFlight = 0;
        let overlapped = false;
        const order = [];

        const editor = new DocumentEditor({
            reader: stubReader(doc),
            host: {
                edit: async ({ text }) => {
                    inFlight += 1;
                    if (inFlight > 1) overlapped = true;
                    await new Promise((r) => setTimeout(r, 20));
                    await writeFile(doc, text);
                    order.push(text);
                    inFlight -= 1;
                    return { status: "edited", wordIndex: 1, paragraphCount: 1, page: 1, openMs: 1, editMs: 1, saveMs: 1, releaseMs: 0, totalMs: 3, protectedView: false, markOfTheWeb: false, released: true };
                },
            },
            snapshotRoot: root,
        });

        const token = await fileRevisionToken(doc);
        const results = await Promise.allSettled([
            editor.edit(doc, { ...intent, text: "A" }, { revisionToken: token }),
            editor.edit(doc, { ...intent, text: "B" }, { revisionToken: token }),
        ]);

        assert.equal(overlapped, false, "two edits of one document ran concurrently");

        // The second must be refused rather than queued-and-applied: its token
        // describes bytes the first has already replaced. That is the whole
        // point of the token, and serialization is what makes it enforceable.
        const rejected = results.filter((r) => r.status === "rejected");
        assert.equal(rejected.length, 1, "both concurrent edits were accepted");
        assert.equal(rejected[0].reason.code, "stale_revision_token");
        assert.equal(order.length, 1, "the refused edit still reached Word");
    });
});

test("edits to different documents are not serialized against each other", async () => {
    // The lock is per document. Serializing globally would make every edit wait
    // on every other one for no benefit -- they contend for Word, not for bytes.
    await withTemp(async (dir) => {
        const a = path.join(dir, "a.docx");
        const b = path.join(dir, "b.docx");
        await writeFile(a, "v0");
        await writeFile(b, "v0");

        let concurrent = 0;
        let sawBothAtOnce = false;
        const makeEditor = (doc) =>
            new DocumentEditor({
                reader: stubReader(doc),
                host: {
                    edit: async ({ path: p, text }) => {
                        concurrent += 1;
                        if (concurrent > 1) sawBothAtOnce = true;
                        await new Promise((r) => setTimeout(r, 20));
                        await writeFile(p, text);
                        concurrent -= 1;
                        return { status: "edited", wordIndex: 1, paragraphCount: 1, page: 1, openMs: 1, editMs: 1, saveMs: 1, releaseMs: 0, totalMs: 3, protectedView: false, markOfTheWeb: false, released: true };
                    },
                },
                snapshotRoot: path.join(dir, "artifacts"),
            });

        // One editor instance, two documents -- the shape the extension uses.
        const editor = makeEditor(a);
        // Reader is bound to `a`, so drive `b` through its own editor; what is
        // under test is that the *lock* does not couple them.
        const editorB = makeEditor(b);

        await Promise.all([
            editor.edit(a, { ...intent, text: "A" }, { revisionToken: await fileRevisionToken(a) }),
            editorB.edit(b, { ...intent, text: "B" }, { revisionToken: await fileRevisionToken(b) }),
        ]);

        assert.equal(sawBothAtOnce, true, "two different documents were serialized against each other");
    });
});

test("revert refuses without a token, and refuses a stale one", async () => {
    await withTemp(async (dir) => {
        const doc = path.join(dir, "report.docx");
        await writeFile(doc, "v1");

        const editor = new DocumentEditor({
            reader: stubReader(doc),
            host: { edit: async () => assert.fail("revert must not reach Word") },
            snapshotRoot: path.join(dir, "artifacts"),
        });

        for (const options of [{}, { revisionToken: null }, { revisionToken: "sha256:0000" }]) {
            const err = await editor.revert(doc, options).then(() => null, (e) => e);
            assert.ok(err, `revert was allowed with ${JSON.stringify(options)}`);
            assert.equal(err.code, "stale_revision_token");
        }
    });
});
