// Office-free tests for the create_document orchestration.
//
// `DocumentAuthor` takes its reader and its Word host as constructor arguments,
// so the parts that decide *what the caller is told about their file* can be
// tested without Word — including the states Word will not reproduce on demand:
// a host that claims success having written nothing, and a document that exists
// but cannot be read back.
//
// Every failure here is silent by nature. "Created" with no file on disk, or
// "failed" for a document that is sitting there, are both indistinguishable from
// success to an agent that does not check — and an agent told a create failed
// will retry it.

import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, rm, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { DocumentAuthor } from "../../src/word/document-author.mjs";
import { asToolError } from "../../src/tool-error.mjs";
import { fileRevisionToken } from "../../src/revision-token.mjs";

const withTemp = async (fn) => {
    const dir = await mkdtemp(path.join(tmpdir(), "author-test-"));
    try {
        return await fn(dir);
    } finally {
        await rm(dir, { recursive: true, force: true }).catch(() => {});
    }
};

/**
 * The tool boundary, applied to an error.
 *
 * `asToolError` forwards only `code`, `message` and `data`, so anything recorded
 * as a top-level property on an error is visible to a test and invisible to the
 * agent. A test that asserts one layer beneath the strip will happily confirm a
 * fact no caller can observe — which has now happened twice in this stack, once
 * in the fix for it.
 *
 * This file used to model that projection by hand, because `asToolError` lived
 * in `extension.mjs` and that cannot be imported (it calls `joinSession()` at
 * module scope). Layer 2 extracted it to `src/tool-error.mjs`, so the real
 * function is used here instead. A model that has to be pinned to its original
 * is a second copy of it; importing removes the copy.
 */
const throughToolBoundary = asToolError;

const here = path.dirname(fileURLToPath(import.meta.url));

test("the tool boundary drops top-level properties and keeps data", () => {
    // The property this whole file depends on, asserted against the real
    // function rather than assumed of it. If `asToolError` ever started
    // forwarding a fourth field, every assertion below would keep passing while
    // testing a boundary that no longer exists.
    const err = Object.assign(new Error("nope"), {
        code: "create_failed",
        exception: "System.Runtime.InteropServices.COMException",
        data: { detail: "kept" },
    });

    const wrapped = throughToolBoundary(err);
    assert.equal(wrapped.code, "create_failed");
    assert.deepEqual(wrapped.data, { detail: "kept" });
    assert.equal(wrapped.exception, undefined, "a top-level property survived the boundary");
});

const spec = { blocks: [{ kind: "heading", level: 1, text: "Title" }, { kind: "paragraph", text: "Body." }] };

const stubReader = (docPath) => ({
    read: async () => ({
        path: docPath,
        name: path.basename(docPath),
        revisionToken: await fileRevisionToken(docPath),
        writable: true,
        paragraphCount: 2,
        paragraphs: [
            { address: "p:aaaaaaaaaaaa", text: "Title", wordIndex: 1, index: 0, headingLevel: 1 },
            { address: "p:bbbbbbbbbbbb", text: "Body.", wordIndex: 2, index: 1, headingLevel: null },
        ],
    }),
});

/** A host that behaves: writes the file, reports it created. */
const goodHost = (extra = {}) => ({
    create: async ({ path: docPath }) => {
        await writeFile(docPath, "AUTHORED");
        return {
            status: "created",
            paragraphCount: 2,
            tableCount: 0,
            autoCorrect: { suppressed: true },
            buildMs: 40,
            saveMs: 120,
            releaseMs: 0,
            released: true,
            totalMs: 200,
            ...extra,
        };
    },
});

const failed = (author, docPath, s = spec) => author.create(docPath, s).then(() => null, (e) => e);

test("a document is authored and described back to the caller", async () => {
    await withTemp(async (dir) => {
        const doc = path.join(dir, "report.docx");
        const author = new DocumentAuthor({ reader: stubReader(doc), host: goodHost() });

        const result = await author.create(doc, spec);

        assert.equal(result.created.blocks, 2);
        assert.equal(result.document.paragraphs.length, 2);
        // The token is the one read_document would mint, so the caller can edit
        // immediately without reading first.
        assert.equal(result.revisionToken, await fileRevisionToken(doc));
        assert.equal(result.lockReleased, true);
    });
});

test("autocorrect suppression is reported, not assumed", async () => {
    // Autocorrect rewrites inserted text and raises nothing, so "we switch it
    // off" is only a claim until the result says so. This is the property a
    // smoke test can assert without reading a comment.
    await withTemp(async (dir) => {
        const doc = path.join(dir, "report.docx");

        const on = new DocumentAuthor({ reader: stubReader(doc), host: goodHost() });
        assert.deepEqual((await on.create(doc, spec)).autoCorrect, { suppressed: true, reason: null });

        await rm(doc, { force: true });

        // An instance we attached to rather than started is the user's own Word,
        // and its settings are per-process — measured, probe-autocorrect.ps1
        // arm C — so changing them changes what the user is looking at. The host
        // declines, and says so rather than reporting a suppression it did not
        // perform.
        const off = new DocumentAuthor({
            reader: stubReader(doc),
            host: goodHost({ autoCorrect: { suppressed: false, reason: "attached_instance" } }),
        });
        assert.deepEqual((await off.create(doc, spec)).autoCorrect, {
            suppressed: false,
            reason: "attached_instance",
        });
    });
});

test("an existing file is refused and its bytes survive", async () => {
    await withTemp(async (dir) => {
        const doc = path.join(dir, "report.docx");
        await writeFile(doc, "THE USER'S WORK");

        let hostCalled = false;
        const author = new DocumentAuthor({
            reader: stubReader(doc),
            host: {
                create: async () => {
                    hostCalled = true;
                    throw new Error("unreachable");
                },
            },
        });

        const err = await failed(author, doc);
        assert.ok(err, "an existing document was overwritten");
        assert.equal(err.code, "file_exists");
        assert.equal(hostCalled, false, "Word was started for a call that could be refused on the filesystem");
        assert.equal(await readFile(doc, "utf8"), "THE USER'S WORK");
        // The refusal has to point somewhere, or an agent that wanted to change
        // the document has nowhere to go and will try the same call again.
        assert.match(err.message, /edit_document/);
    });
});

test("only extensions SaveAs2's format matches are creatable", async () => {
    await withTemp(async (dir) => {
        const author = new DocumentAuthor({ reader: stubReader(""), host: goodHost() });
        for (const name of ["report.rtf", "report.doc", "report.dotx", "report"]) {
            const err = await failed(author, path.join(dir, name));
            assert.ok(err, `${name} was accepted`);
            assert.equal(err.code, "unsupported_type", `${name}: ${err.message}`);
        }
    });
});

test("an invalid spec is refused before the host is called", async () => {
    await withTemp(async (dir) => {
        const doc = path.join(dir, "report.docx");
        let hostCalled = false;
        const author = new DocumentAuthor({
            reader: stubReader(doc),
            host: {
                create: async () => {
                    hostCalled = true;
                    return { status: "created" };
                },
            },
        });

        const err = await failed(author, doc, { blocks: [{ kind: "paragraph", text: "a\nb" }] });
        assert.equal(err.code, "invalid_text");
        assert.equal(hostCalled, false, "Word was started for a spec that could never be authored");
        await assert.rejects(stat(doc), "a file was created for a rejected spec");
    });
});

test("a host failure names no cause, and carries the exception type through the boundary", async () => {
    // An error code can be right while its message asserts a reason the code
    // never distinguished. All that was established here is that a COM call
    // raised: the type goes on `data` for a caller that wants to discriminate,
    // and the prose says what was observed.
    await withTemp(async (dir) => {
        const doc = path.join(dir, "report.docx");
        const author = new DocumentAuthor({
            reader: stubReader(doc),
            host: {
                create: async () => ({
                    status: "create_failed",
                    exception: "System.Runtime.InteropServices.COMException",
                    detail: "Der Befehl ist nicht verfügbar.",
                }),
            },
        });

        const raw = await failed(author, doc);
        const err = throughToolBoundary(raw);

        assert.equal(err.code, "create_failed");
        assert.equal(err.data?.exception, "System.Runtime.InteropServices.COMException");
        // Whatever the message says, it must not say why -- nothing here
        // distinguished a locked folder from a bad path from a Word that fell
        // over, and this repo has shipped a correct code with a message that
        // asserted the wrong reason more than once.
        assert.doesNotMatch(err.message, /lock|locked|permission|read-only|close Word|open in Word/i);
    });
});

test("'created' with no file on disk is reported as a failure", async () => {
    // SaveAs2 returning is not evidence that a file exists. Without this check
    // the caller is handed a revision token for nothing and an agent proceeds to
    // edit a document that was never written.
    await withTemp(async (dir) => {
        const doc = path.join(dir, "report.docx");
        const author = new DocumentAuthor({
            reader: stubReader(doc),
            host: { create: async () => ({ status: "created", released: true }) },
        });

        const err = await failed(author, doc);
        assert.equal(err.code, "create_not_persisted");
    });
});

test("a document that cannot be read back is still reported as created", async () => {
    // The file exists. Reporting this as a failed create would make an agent
    // retry, and the retry would be refused for a file the agent believes it
    // never made -- so the error has to say the opposite of what its code
    // usually implies, in words.
    await withTemp(async (dir) => {
        const doc = path.join(dir, "report.docx");
        const author = new DocumentAuthor({
            reader: {
                read: async () => {
                    throw Object.assign(new Error("Word did not respond"), { code: "word_timeout" });
                },
            },
            host: goodHost(),
        });

        const raw = await failed(author, doc);
        const err = throughToolBoundary(raw);

        assert.equal(err.code, "document_unreadable");
        assert.equal(err.data?.created, true, "the created flag is dropped at the tool boundary");
        assert.equal(err.data?.cause, "word_timeout");
        assert.match(err.message, /Do not repeat the call/);
        assert.equal(await readFile(doc, "utf8"), "AUTHORED", "the created document was deleted");
    });
});

test("a document still held after Close is flagged rather than passed off as clean", async () => {
    await withTemp(async (dir) => {
        const doc = path.join(dir, "report.docx");
        const lines = [];
        const author = new DocumentAuthor({
            reader: stubReader(doc),
            host: goodHost({ released: false, releaseMs: 5000 }),
            log: (m) => lines.push(m),
        });

        const result = await author.create(doc, spec);
        assert.equal(result.lockReleased, false);
        assert.ok(
            lines.some((l) => l.includes("still held")),
            "a document Word never released was not mentioned anywhere",
        );
    });
});
