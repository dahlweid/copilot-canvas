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
//
// A note on the failure this file kept walking into, because it is a class and
// not an incident. The seam test here once carried a title asserting that a
// close had *finished* over an assertion that could only see it *start* -- the
// prose a human reads and the thing the machine checks had come apart, and
// nothing in between complains when they do. The same shape bit the pull
// request that added the mutation-anchor gate from the other side: its body
// said in plain
// English that it did *not* close its tracking issue, while GitHub's
// linked-issue parser, which does not read negation, matched the keyword beside
// the reference and registered it as closing that issue. Prose was the part
// everybody read; the metadata was the part that would have acted. Neither case
// is catchable by reading the diff, because in both the sentence is true of the
// intent and false of the mechanism. The defence used here is to make the
// assertion name the mechanism -- `closeDocument:returned`, not `closeDocument`
// -- so that a title claiming completion has something underneath it that can
// only be satisfied by completion.
//
// What this file does *not* cover, recorded because its absence used to be
// hidden by a test that implied otherwise. Nothing here pairs a `refresh` with
// an in-flight edit. A second test once did, and was removed rather than
// repaired: `open (render-cache.mjs)` drops a stale working copy itself when the
// fingerprint has changed, and an edit changes it, so the reopen brought its own
// `closeDocument` and filled the same slot in the asserted order. Measured:
// removing the invalidation from `editDocument (render-cache.mjs)` entirely left
// that test green. It could not attribute the close it observed, so the
// assertion under the name was doing less work than the name implied -- the same
// shape as the two cases above. The pairing is tracked in #186; the choke point
// that serialises it is exercised by the edit-versus-outline test below, so what
// is missing is an assertion, not a guard.

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
 * Asserts that every structure read landed before the teardown began.
 *
 * `DocumentReader` is the only thing in `src/` that calls `host.structure`, and
 * `edit (document-editor.mjs)` re-reads the document after saving it, so a
 * `structure` marker recorded after `closeDocument:entered` would be a read
 * issued against a document Word has already been told to close.
 *
 * The two preconditions are not ceremony. `lastIndexOf` and `indexOf` both
 * return -1 for an absent marker, and -1 loses to any real index, so the bare
 * comparison would pass green on a run where no read happened at all -- an
 * assertion whose message claims a read was correctly ordered, satisfied by
 * there being no read. The missing-teardown case fails the other way: it would
 * report a read-ordering defect when the real defect is that nothing tore down.
 * Each of the three failures has to name its own cause, so each is asserted
 * separately.
 */
const assertReadsPrecedeTeardown = (calls) => {
    const lastRead = calls.lastIndexOf("structure");
    const teardownStart = calls.indexOf("closeDocument:entered");

    assert.notEqual(lastRead, -1, "no structure read was recorded at all, so the ordering below would pass vacuously");
    assert.notEqual(teardownStart, -1, "no teardown was recorded at all, so read-versus-teardown order was never exercised");
    assert.ok(
        teardownStart > lastRead,
        "a structure read ran against a document whose teardown had already started",
    );
};

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

test("the read-before-teardown assertion fails separately for each of its three causes", () => {
    // This exists because the assertion it covers cannot be driven red from
    // production code. Three mutations were tried and all three are preempted:
    // dropping the `await` on the post-edit re-read in
    // `edit (document-editor.mjs)` makes `after` a promise and dies on a
    // `TypeError` first; dropping the `await` on the editor call in
    // `editDocument (render-cache.mjs)` trips the settled-order assertion; and
    // reordering the invalidation before the edit deadlocks against the gates
    // and reports only `test timed out`. A timeout is a red test, but it names
    // no cause, so none of those runs demonstrates that this assertion works.
    //
    // So the assertion is verified here instead, directly. If someone later
    // "simplifies" the helper to the bare index comparison, the first case below
    // goes red -- which is the whole point of it not being the bare comparison.
    assert.throws(
        () => assertReadsPrecedeTeardown(["openDocument", "closeDocument:entered"]),
        /pass vacuously/,
        "with no read recorded, the bare comparison would have passed on -1 and claimed the ordering was checked",
    );

    // Captured rather than matched, because a regex matcher fails the wrong way
    // here. Remove the teardown-present guard and the helper falls through to
    // the ordering comparison, where -1 loses to any real index -- so it still
    // throws, and `assert.throws` re-raises what it caught when the pattern does
    // not match. The reader would then be shown "a structure read ran against a
    // document whose teardown had already started" for a run in which nothing
    // tore down at all: a true-looking message naming a cause that did not
    // occur. That is this file's own subject pointed at its own diagnostics, so
    // the failure output has to name the cause as exactly as the assertion does.
    let missingTeardown = null;
    try {
        assertReadsPrecedeTeardown(["openDocument", "structure"]);
    } catch (error) {
        missingTeardown = error;
    }

    assert.ok(missingTeardown, "with no teardown recorded, the helper must refuse rather than pass vacuously");
    assert.match(
        missingTeardown.message,
        /never exercised/,
        `no teardown was recorded, so read-versus-teardown order was never exercised; the failure must say that rather than report a read-ordering defect (received: ${missingTeardown.message})`,
    );

    assert.throws(
        () => assertReadsPrecedeTeardown(["closeDocument:entered", "structure"]),
        /teardown had already started/,
        "a read after the teardown began is the defect this assertion is for",
    );

    assert.doesNotThrow(
        () => assertReadsPrecedeTeardown(["structure", "closeDocument:entered", "closeDocument:returned"]),
        "the correct order must not be reported as a defect",
    );
});

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
            // The reads are filtered out of the comparison because a read *count*
            // is `document-editor.mjs`'s business and would make this test red for
            // something it is not about. Filtering drops their *position* too,
            // though, and position here is real information -- so it is asserted
            // just above rather than discarded with the count.
            assertReadsPrecedeTeardown(calls);
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
