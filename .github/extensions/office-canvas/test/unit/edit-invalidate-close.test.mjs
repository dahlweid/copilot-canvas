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
        calls.push("closeDocument");
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
    await withFixture("queued.docx", async ({ cache, docPath }) => {
        const calls = [];
        const editEntered = deferred();
        const releaseEdit = deferred();
        let edit;
        let outline;

        cache.host = fakeHost({ docPath, calls, gates: { editEntered, releaseEdit } });

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

            releaseEdit.resolve();
            const result = await edit;
            await outline;

            assert.equal(result.applied.op, "replace_text", "the edit did not run to completion");

            // The claim is the ordering of the units of work, not how many reads
            // the editor needs, so the reads are filtered out rather than
            // asserted -- a read count is `document-editor.mjs`'s business and
            // would make this test red for something it is not about.
            assert.deepEqual(
                calls.filter((call) => call !== "structure" && call !== "openDocument"),
                ["edit", "closeDocument", "outlineMarkup", "outlinePositions"],
                "the outline handshake interleaved with the edit or its invalidation",
            );
        } finally {
            releaseEdit.resolve();
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
        let edit;
        let refresh;

        cache.host = fakeHost({ docPath, calls, gates: { editEntered, releaseEdit } });

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

            releaseEdit.resolve();
            await edit;
            const refreshed = await refresh;

            const lifecycle = calls.filter((call) => call === "edit" || call === "closeDocument" || call === "openDocument");
            assert.deepEqual(
                lifecycle,
                ["openDocument", "edit", "closeDocument", "openDocument"],
                "the reopen landed inside the edit's invalidation instead of queueing behind it",
            );
            assert.equal(refreshed.changed, true, "the refresh did not observe the edited bytes");
        } finally {
            releaseEdit.resolve();
            await Promise.allSettled([edit, refresh].filter(Boolean));
        }
    });
});
