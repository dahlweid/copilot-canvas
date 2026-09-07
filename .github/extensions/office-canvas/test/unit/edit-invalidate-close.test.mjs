// An edit runs to completion through the document queue, and the invalidation
// that follows it closes Word's copy at a point nothing else can interleave.
//
// `editDocument` and `revertDocument` are the queued operations that end by
// *tearing down* the document they just acted on: each calls
// `#invalidateState`, which closes the host-side document so the canvas cannot
// keep showing pre-edit bytes. That makes them the repo's highest-risk shape --
// a teardown at the end of a unit of work -- and until this file existed no
// unit test reached it. Measured at the commit this was written against: an
// `appendFileSync` probe placed at the top of `#invalidateState
// (render-cache.mjs)` produced no output across a full 712-test run, and two
// lines from this file alone once these tests existed, so the run was
// instrumented correctly and the seam really was unreached.
// `outline-interleave.test.mjs` calls `editDocument`, but only against a
// *closing* document, where the refusal returns before anything is enqueued.
//
// The fake is a Word host, not a fake `DocumentEditor`. The real
// `DocumentEditor` and `DocumentReader` run on top of it, so the queued unit of
// work here is the one the product runs -- read, edit, re-read, invalidate --
// rather than a stand-in whose duration is the only thing it models.

import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, mkdir, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { RenderCache } from "../../src/render-cache.mjs";
import { flatOpc, paragraph } from "./word-fixtures.mjs";
import { deferred, turns } from "./queue-probe.mjs";

const BODY = "The paragraph as it stood before the edit.";

/**
 * A Word host that serves a two-paragraph document and writes to the file when
 * asked to edit it.
 *
 * `structure` is what `DocumentReader` calls, so the addresses the editor
 * resolves against are minted by the real structure map rather than invented
 * here -- an invented address would prove the queue ordering while quietly
 * skipping the address resolution that makes the edit real.
 */
const fakeHost = ({ docPath, calls, gates = {} }) => ({
    async openDocument({ docId, workDir }) {
        calls.push("openDocument");
        await mkdir(workDir, { recursive: true });
        return { docId, pageCount: 1, wordCount: 7, sizeBytes: 32, modifiedIso: "2026-01-01T00:00:00.000Z" };
    },
    async closeDocument() {
        // Two markers, not one, and the distinction is the point. An async
        // function body runs synchronously up to its first `await`, and the
        // awaited *expression* is evaluated before the suspension -- so a single
        // marker pushed here is recorded even by a caller that never waits for
        // the returned promise. Dropping the `await` in front of
        // `#invalidateState (render-cache.mjs)`, or the one in front of
        // `host.closeDocument` inside it, was measured leaving a one-marker
        // version of this test green. Releasing the document queue while a close
        // is still in flight against Word is this repo's named failure class, so
        // a test that cannot see it is the wrong test.
        calls.push("closeDocument:entered");
        if (gates.closeEntered) gates.closeEntered.resolve();
        if (gates.releaseClose) await gates.releaseClose.promise;
        calls.push("closeDocument:returned");
        return { closed: true };
    },
    async structure({ workDir, out }) {
        calls.push("structure");
        await mkdir(workDir, { recursive: true });
        await writeFile(
            out,
            flatOpc([paragraph("Chapter one", { styleId: "berschrift1" }), paragraph(BODY)]),
        );
        return { writable: true, name: path.basename(docPath), sizeBytes: 32 };
    },
    async edit({ wordIndex }) {
        calls.push("edit");
        if (gates.editEntered) gates.editEntered.resolve();
        if (gates.releaseEdit) await gates.releaseEdit.promise;
        // A real edit changes the bytes, which is what makes the document's
        // cache key stale and the invalidation necessary.
        await writeFile(docPath, `EDITED at ${wordIndex}`);
        return { status: "edited", wordIndex, page: 2, released: true };
    },
    async outlineMarkup({ out }) {
        calls.push("outlineMarkup");
        await writeFile(out, flatOpc([paragraph("Chapter one", { styleId: "berschrift1" })]));
        if (gates.markupEntered) gates.markupEntered.resolve();
        if (gates.releaseMarkup) await gates.releaseMarkup.promise;
    },
    async outlinePositions({ wordIndices }) {
        calls.push("outlinePositions");
        return { positions: wordIndices.map((wordIndex) => ({ wordIndex, start: 0, page: 1 })) };
    },
});

/** Opens the fixture and returns the address and token an edit needs. */
const openAndAddress = async (cache, docPath) => {
    await cache.open(docPath);
    const structure = await cache.readStructure(docPath);
    const target = structure.paragraphs.find((p) => p.text === BODY);
    assert.ok(target, "the fixture did not produce the paragraph the edit addresses");
    return { address: target.address, wordIndex: target.wordIndex, revisionToken: structure.revisionToken };
};

const withFixture = async (name, run) => {
    const root = await mkdtemp(path.join(os.tmpdir(), "edit-invalidate-"));
    const docPath = path.join(root, name);
    await writeFile(docPath, "ORIGINAL");
    const cache = new RenderCache({
        cacheRoot: path.join(root, "cache"),
        snapshotRoot: path.join(root, "snapshots"),
    });
    try {
        await run({ cache, docPath });
    } finally {
        await rm(root, { recursive: true, force: true }).catch(() => {});
    }
};

test("an edit holds the document queue until its invalidation has closed the document", async () => {
    // The hold is inside `host.edit`, not on a timer, and that is the whole
    // design of this test. The obvious shape -- start an edit behind a held
    // outline and assert the edit "has not run yet" -- passes on slack: the real
    // editor does a stat, a hash, a snapshot and a full read before it reaches
    // Word, which outlasts any turn count, so a version of `editDocument` that
    // skipped the queue entirely still looked correctly queued. Measured: with
    // `editDocument` calling its operation directly instead of through
    // `#enqueueDocumentOperation (render-cache.mjs)`, that assertion stayed
    // green.
    //
    // Holding *inside* the edit inverts it. The window stays open for as long as
    // the test likes, so a competing operation that is not queued has unbounded
    // opportunity to run, and "it did not run" stops being a statement about
    // timing.
    //
    // The title says *has closed*, and the assertions have to mean it. An
    // earlier version of this test recorded a single `closeDocument` marker,
    // which an async fake pushes synchronously -- before its first `await` --
    // so it was recorded even by a caller that never waited for the close to
    // finish. Measured: with `await this.#invalidateState(state)` in
    // `editDocument (render-cache.mjs)` reduced to a bare call, that version
    // stayed green. It observed when the teardown *started* while its title
    // claimed it had *finished*, which is the defect #178 is about, reproduced
    // in the test written to close it. The gate inside `closeDocument` and the
    // `closeDocument:returned` marker are what fix that.
    await withFixture("queued.docx", async ({ cache, docPath }) => {
        const calls = [];
        const editEntered = deferred();
        const releaseEdit = deferred();
        const closeEntered = deferred();
        const releaseClose = deferred();
        let edit;
        let outline;

        cache.host = fakeHost({
            docPath,
            calls,
            gates: { editEntered, releaseEdit, closeEntered, releaseClose },
        });

        try {
            const { address, revisionToken } = await openAndAddress(cache, docPath);

            edit = cache.editDocument(docPath, { op: "replace_text", address, text: "rewritten" }, { revisionToken });
            await editEntered.promise;

            // `outline` reaches `host.outlineMarkup` in a couple of turns once it
            // is running -- it joins a path and issues the command -- so if it is
            // not queued behind the edit it runs here, while the edit is pinned.
            outline = cache.outline(docPath);
            await turns(20);

            assert.ok(
                !calls.includes("outlineMarkup"),
                "an outline ran against the document while an edit was mid-flight inside Word",
            );

            // Let the edit finish and pin the invalidation instead, inside the
            // close. This is the window the queue must still be holding: the
            // bytes on disk have already changed and Word's copy is mid-teardown.
            releaseEdit.resolve();
            await closeEntered.promise;
            await turns(20);

            assert.ok(
                !calls.includes("outlineMarkup"),
                "an outline ran while the invalidation's closeDocument was still in flight against Word",
            );

            releaseClose.resolve();
            const result = await edit;
            await outline;

            assert.equal(result.applied.op, "replace_text", "the edit did not run to completion");

            // The load-bearing assertion: a settled order, with nothing timing-
            // bounded about it. `closeDocument:returned` must precede
            // `outlineMarkup`, which is what distinguishes a queue released after
            // the teardown *finished* from one released when it merely started.
            //
            // The reads are filtered out rather than asserted -- a read count is
            // `document-editor.mjs`'s business and would make this test red for
            // something it is not about.
            assert.deepEqual(
                calls.filter((call) => call !== "structure" && call !== "openDocument"),
                ["edit", "closeDocument:entered", "closeDocument:returned", "outlineMarkup", "outlinePositions"],
                "the outline handshake interleaved with the edit or its invalidation",
            );
        } finally {
            releaseEdit.resolve();
            releaseClose.resolve();
            await Promise.allSettled([edit, outline].filter(Boolean));
        }
    });
});

test("a reopen cannot land between an edit and the invalidation that closes it", async () => {
    // The failure this excludes is the repo's named one: tearing down Word while
    // work is in flight, seen from the other side. Between `host.edit` returning
    // and `#invalidateState` closing the document there is a window in which the
    // file on disk has already changed. A `refresh` admitted into that window
    // reopens the document from the new bytes and *then* the invalidation closes
    // what it just opened -- leaving the canvas holding a closed docId, with no
    // error anywhere to say so.
    await withFixture("reopen.docx", async ({ cache, docPath }) => {
        const calls = [];
        const editEntered = deferred();
        const releaseEdit = deferred();
        const closeEntered = deferred();
        const releaseClose = deferred();
        let edit;
        let refresh;

        cache.host = fakeHost({
            docPath,
            calls,
            gates: { editEntered, releaseEdit, closeEntered, releaseClose },
        });

        try {
            const { address, revisionToken } = await openAndAddress(cache, docPath);

            edit = cache.editDocument(docPath, { op: "replace_text", address, text: "rewritten" }, { revisionToken });
            await editEntered.promise;

            const opensBefore = calls.filter((call) => call === "openDocument").length;
            refresh = cache.refresh(docPath);
            await turns(20);

            assert.equal(
                calls.filter((call) => call === "openDocument").length,
                opensBefore,
                "a refresh reopened the document while an edit was still in flight",
            );

            // The close is gated for the same reason it is in the test above,
            // and not merely for symmetry: ungated, this fake's two markers are
            // pushed with no suspension between them, so the settled order below
            // would hold whatever the caller did with the close's promise. The
            // gate is what makes the reopen have to *wait* for a teardown that is
            // genuinely still running.
            releaseEdit.resolve();
            await closeEntered.promise;
            await turns(20);

            assert.equal(
                calls.filter((call) => call === "openDocument").length,
                opensBefore,
                "a refresh reopened the document while its close was still in flight against Word",
            );

            releaseClose.resolve();
            await edit;
            const refreshed = await refresh;

            const lifecycle = calls.filter((call) => call !== "structure");
            assert.deepEqual(
                lifecycle,
                [
                    "openDocument",
                    "edit",
                    "closeDocument:entered",
                    "closeDocument:returned",
                    "openDocument",
                ],
                "the reopen landed inside the edit's invalidation instead of queueing behind it",
            );
            assert.equal(refreshed.changed, true, "the refresh did not observe the edited bytes");
        } finally {
            releaseEdit.resolve();
            releaseClose.resolve();
            await Promise.allSettled([edit, refresh].filter(Boolean));
        }
    });
});
