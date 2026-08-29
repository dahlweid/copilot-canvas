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
import { mkdir, mkdtemp, readFile, rm, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { CreateError, DocumentAuthor, creatableList } from "../../src/word/document-author.mjs";
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

test("a flat detail on CreateError reaches the agent", () => {
    // The other half of the boundary contract, and the half that used to be a
    // convention rather than a check.
    //
    // `CreateError` once did `Object.assign` alone, so a detail reached `data`
    // only if the call site nested it there by hand. All four sites did, so
    // nothing was broken -- but the natural shape, the one every `EditError`
    // site uses and the one a future create site will reach for, put the fact
    // somewhere no caller could see.
    //
    // Asserting the boundary's *output* is what makes the mirror the thing
    // under test. A test reading `err.leftBehind` instead would pass with the
    // mirror deleted, because the assign puts it there either way: that is the
    // below-the-boundary trap this file's header describes, and this is the
    // assertion that is not subject to it.
    const err = new CreateError("create_failed", "nope", { leftBehind: true });

    assert.deepEqual(throughToolBoundary(err).data, { leftBehind: true });
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
    //
    // Asserted as a WHOLE-SHAPE deep-equal on purpose. These fields are what the
    // agent sees, and the defect this pins is a field being added host-side and
    // silently dropped on the way out: `restored` was written into the host
    // report and lost at three object literals in document-author.mjs, which is
    // the same class as asToolError forwarding only code/message/data. A
    // property-by-property assertion cannot see an omission; this can.
    await withTemp(async (dir) => {
        const doc = path.join(dir, "report.docx");

        const on = new DocumentAuthor({ reader: stubReader(doc), host: goodHost() });
        assert.deepEqual((await on.create(doc, spec)).autoCorrect, {
            suppressed: true,
            reason: null,
            restored: false,
            restoreReason: null,
            settings: null,
            restoreSettings: null,
            prior: null,
        });

        await rm(doc, { force: true });

        // An instance we attached to rather than started is the user's own Word.
        // The original reason given here was that the settings are per-process,
        // citing probe-autocorrect.ps1 arm C. That is RETRACTED: arm C read a
        // second instance while the first was still alive, and a concurrent
        // reader sees the pre-write value, so it cannot tell
        // isolation from persistence-with-lag. Measured sequentially they
        // persist for the user, which makes declining on an attached instance
        // MORE important rather than less -- the change would outlive us.
        const off = new DocumentAuthor({
            reader: stubReader(doc),
            host: goodHost({ autoCorrect: { suppressed: false, reason: "attached_instance" } }),
        });
        assert.deepEqual((await off.create(doc, spec)).autoCorrect, {
            suppressed: false,
            reason: "attached_instance",
            restored: false,
            restoreReason: null,
            settings: null,
            restoreSettings: null,
            prior: null,
        });

        await rm(doc, { force: true });

        // Both arms above leave six of the seven fields at their default -- five
        // `null` and `restored: false` -- so each is asserted only in the
        // direction the helper would produce anyway if it hardcoded it. A
        // `restored: false` or `settings: null` written into reportedAutoCorrect
        // survives both of them, which is the suppression/restoration asymmetry
        // one level down: `suppressed` is pinned in both directions by the two
        // arms above and nothing else was.
        //
        // So this arm carries a value in EVERY field, all distinct from the
        // defaults. Any field the helper stops forwarding changes here.
        const full = new DocumentAuthor({
            reader: stubReader(doc),
            host: goodHost({
                autoCorrect: {
                    suppressed: true,
                    reason: "verified",
                    restored: true,
                    restoreReason: "verified",
                    settings: ["ReplaceText"],
                    restoreSettings: ["ReplaceText"],
                    prior: { ReplaceText: true },
                },
            }),
        });
        assert.deepEqual((await full.create(doc, spec)).autoCorrect, {
            suppressed: true,
            reason: "verified",
            restored: true,
            restoreReason: "verified",
            settings: ["ReplaceText"],
            restoreSettings: ["ReplaceText"],
            prior: { ReplaceText: true },
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

            // Asserted through the boundary, because the boundary is where the
            // agent reads this. A refusal is only actionable if it says which
            // path was refused -- an agent that authored several in one turn
            // cannot otherwise tell which call to change. Reviewed finding:
            // the message used to open with the bare extension (".rtf cannot
            // be created"), which reads as though the extension were the thing
            // being created and never named the file.
            const seen = throughToolBoundary(err);
            const full = path.join(dir, name);
            assert.ok(seen.message.includes(full), `${name}: does not name the path -- ${seen.message}`);
            for (const ok of creatableList().split(", ")) {
                assert.ok(seen.message.includes(ok), `${name}: does not offer ${ok} -- ${seen.message}`);
            }
            const ext = path.extname(name);
            if (ext) {
                assert.ok(
                    seen.message.includes(ext),
                    `${name}: does not name the rejected type -- ${seen.message}`,
                );
            }
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

test("a failed create reports whether it actually cleaned up after itself", async () => {
    // The host's cleanup is a best-effort `Remove-Item` inside a swallowed
    // `catch`, so "no document was written" was a claim it could not make. It now
    // looks at the disk and reports what it saw. Both directions are asserted,
    // because a message that is right by accident in one case is the failure mode
    // being closed here.
    await withTemp(async (dir) => {
        const doc = path.join(dir, "report.docx");
        const hostSaying = (leftBehind) => ({
            create: async () => ({ status: "create_failed", exception: "System.Exception", leftBehind }),
        });

        const cleaned = throughToolBoundary(
            await failed(new DocumentAuthor({ reader: stubReader(doc), host: hostSaying(false) }), doc),
        );
        assert.equal(cleaned.data?.leftBehind, false);
        assert.match(cleaned.message, /No document was left behind/);
        assert.doesNotMatch(cleaned.message, /partial/i);

        const partial = throughToolBoundary(
            await failed(new DocumentAuthor({ reader: stubReader(doc), host: hostSaying(true) }), doc),
        );
        assert.equal(partial.data?.leftBehind, true);
        assert.match(partial.message, /partial document was left/i);
        // And it names the path, because "a partial document exists somewhere"
        // is not something a caller can act on.
        assert.ok(partial.message.includes(doc));
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
        // The one errno that licenses the message's claim of absence.
        assert.equal(throughToolBoundary(err).data?.cause, "ENOENT");
        assert.equal(throughToolBoundary(err).data?.created, false);
    });
});

test("a directory is not refused as a file that already exists", async () => {
    // `stat` resolves for any filesystem entry, so a preflight that asked only
    // whether it resolved refused a directory with `file_exists` -- advising the
    // caller to use edit_document on a directory.
    //
    // The wrong answer is the smaller half. The refusal runs *upstream* of the
    // host, whose `path_is_directory` classifier is correct, so the bad answer
    // also prevented the good one. This asserts the host was reached at all,
    // because that is the part a message-only fix would leave broken.
    await withTemp(async (dir) => {
        const asDirectory = path.join(dir, "report.docx");
        await mkdir(asDirectory);

        let hostSaw = null;
        const author = new DocumentAuthor({
            reader: stubReader(asDirectory),
            host: {
                create: async ({ path: p }) => {
                    hostSaw = p;
                    return { status: "path_is_directory", path: p };
                },
            },
        });

        const err = throughToolBoundary(await failed(author, asDirectory));

        assert.equal(hostSaw, asDirectory, "the preflight refused before the host could classify");
        assert.equal(err.code, "invalid_path");
        assert.match(err.message, /is a directory/i);
        assert.doesNotMatch(err.message, /edit_document/i, "a directory was described as an existing document");
    });
});

test("a file that cannot be read back is not reported as one that was never written", async () => {
    // `.catch(() => null)` collapsed every verification failure into "no file",
    // and the message then asserted absence -- a cause a failed read cannot
    // establish. The two want opposite responses: "nothing on disk" invites
    // another attempt, "could not read it back" must not, because re-authoring
    // would overwrite a document that already exists.
    //
    // The failing read is real, not stubbed: the path is a directory, so the
    // read stream fails `EISDIR`. Measured on this machine rather than assumed,
    // because the whole point of the split is which errno actually arrives.
    await withTemp(async (dir) => {
        const asDirectory = path.join(dir, "report.docx");
        await mkdir(asDirectory);

        const author = new DocumentAuthor({
            reader: stubReader(asDirectory),
            host: {
                create: async () => ({
                    status: "created",
                    released: true,
                    autoCorrect: { suppressed: false, reason: "attached_instance" },
                }),
            },
        });

        const err = throughToolBoundary(await failed(author, asDirectory));

        assert.equal(err.code, "create_unverified");
        assert.equal(err.data?.cause, "EISDIR");
        assert.equal(err.data?.created, null, "a read that only failed was reported as proof of absence");
        assert.match(err.message, /Do not create it again/);
        assert.doesNotMatch(err.message, /there is no file at that path/);
        // Asserted through the boundary, because this is a path where a document
        // may exist and the description tells callers to check the field. Whole
        // shape again: an error path is where a field is most likely to be
        // dropped, since the success path is the one people look at.
        assert.deepEqual(err.data?.autoCorrect, {
            suppressed: false,
            reason: "attached_instance",
            restored: false,
            restoreReason: null,
            settings: null,
            restoreSettings: null,
            prior: null,
        });
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

test("the autocorrect outcome survives to the caller on the failure that still authored a document", async () => {
    // The tool description tells callers to check `autoCorrect` to learn whether
    // their text was written verbatim. That has to be true on the paths where
    // there is a document to ask about -- and this is the only failure where
    // there is one. Everywhere else the create did not happen and the question
    // is meaningless.
    //
    // Asserted *through* `asToolError`, which forwards only code, message and
    // data. A check one layer beneath that strip would happily confirm a
    // property no caller can observe; that has bitten this stack twice.
    //
    // The unsuppressed case specifically, because it is the one that carries
    // information: `suppressed: true` is the default assumption a caller would
    // make anyway.
    await withTemp(async (dir) => {
        const doc = path.join(dir, "report.docx");
        const author = new DocumentAuthor({
            reader: {
                read: async () => {
                    throw Object.assign(new Error("Word did not respond"), { code: "word_timeout" });
                },
            },
            host: goodHost({ autoCorrect: { suppressed: false, reason: "attached_instance" } }),
        });

        const err = throughToolBoundary(await failed(author, doc));

        assert.equal(err.code, "document_unreadable");
        assert.equal(err.data?.autoCorrect?.suppressed, false, "the autocorrect outcome is dropped at the tool boundary");
        assert.equal(err.data?.autoCorrect?.reason, "attached_instance");
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
