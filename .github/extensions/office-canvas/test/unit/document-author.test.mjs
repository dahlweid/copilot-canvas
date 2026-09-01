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
import { toolFailure } from "../../src/tool-error.mjs";
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
 * The tool boundary, applied to an error — the *real* one, measured.
 *
 * This helper used to be `asToolError`, and the comment here used to say that
 * `code`, `message` and `data` were "what the agent sees". Measured (#45), none
 * of the three reached the agent at all: a thrown error crossed as the bare
 * string `Tool execution failed`. So every assertion made through it was
 * confirming a fact no caller could observe — precisely the failure the comment
 * warned about, committed by the warning itself.
 *
 * `toolFailure` is the boundary that now exists, and the one thing that crosses
 * it is `textResultForLlm`. Projecting to that string is what makes an
 * assertion here mean something on the far side.
 */
const throughToolBoundary = (err) => toolFailure("create_document", err).textResultForLlm;

const here = path.dirname(fileURLToPath(import.meta.url));

test("the tool boundary carries the code and the data bag, and drops top-level properties", () => {
    // The property this whole file depends on, asserted against the real
    // function rather than assumed of it. If `toolFailure` ever stopped
    // rendering `data` into the text, every assertion below would keep passing
    // while testing a boundary that no longer carries anything.
    const err = Object.assign(new Error("nope"), {
        code: "create_failed",
        exception: "System.Runtime.InteropServices.COMException",
        data: { detail: "kept" },
    });

    const text = throughToolBoundary(err);
    assert.match(text, /create_failed/);
    assert.match(text, /"detail":"kept"/);
    assert.doesNotMatch(text, /COMException/, "a top-level property survived the boundary");
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

    assert.match(throughToolBoundary(err), /"leftBehind":true/);
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
            assert.ok(seen.includes(full), `${name}: does not name the path -- ${seen}`);
            for (const ok of creatableList().split(", ")) {
                assert.ok(seen.includes(ok), `${name}: does not offer ${ok} -- ${seen}`);
            }
            const ext = path.extname(name);
            if (ext) {
                assert.ok(seen.includes(ext), `${name}: does not name the rejected type -- ${seen}`);
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
        const text = throughToolBoundary(raw);

        assert.match(text, /create_failed/);
        assert.match(text, /"exception":"System\.Runtime\.InteropServices\.COMException"/);
        // Whatever the message says, it must not say why -- nothing here
        // distinguished a locked folder from a bad path from a Word that fell
        // over, and this repo has shipped a correct code with a message that
        // asserted the wrong reason more than once.
        assert.doesNotMatch(text, /lock|locked|permission|read-only|close Word|open in Word/i);
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
        assert.match(cleaned, /"leftBehind":false/);
        assert.match(cleaned, /No document was left behind/);
        assert.doesNotMatch(cleaned, /partial/i);

        const partial = throughToolBoundary(
            await failed(new DocumentAuthor({ reader: stubReader(doc), host: hostSaying(true) }), doc),
        );
        assert.match(partial, /"leftBehind":true/);
        assert.match(partial, /partial document was left/i);
        // And it names the path, because "a partial document exists somewhere"
        // is not something a caller can act on.
        assert.ok(partial.includes(doc));
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
        assert.match(throughToolBoundary(err), /"cause":"ENOENT"/);
        assert.match(throughToolBoundary(err), /"created":false/);
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

        const text = throughToolBoundary(await failed(author, asDirectory));

        assert.equal(hostSaw, asDirectory, "the preflight refused before the host could classify");
        assert.match(text, /invalid_path/);
        assert.match(text, /is a directory/i);
        assert.doesNotMatch(text, /edit_document/i, "a directory was described as an existing document");
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
                }),
            },
        });

        const text = throughToolBoundary(await failed(author, asDirectory));

        assert.match(text, /create_unverified/);
        assert.match(text, /"cause":"EISDIR"/);
        assert.match(text, /"created":null/, "a read that only failed was reported as proof of absence");
        assert.match(text, /Do not create it again/);
        assert.doesNotMatch(text, /there is no file at that path/);
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
        const text = throughToolBoundary(raw);

        assert.match(text, /document_unreadable/);
        assert.match(text, /"created":true/, "the created flag is dropped at the tool boundary");
        assert.match(text, /"cause":"word_timeout"/);
        assert.match(text, /Do not repeat the call/);
        assert.equal(await readFile(doc, "utf8"), "AUTHORED", "the created document was deleted");
    });
});

test("a save that hangs surfaces as a typed word_timeout, on a budget the caller spends down", async () => {
    // The failure #96 is about: `SaveAs2` is an unbounded COM call, the class
    // this repo singles out as where a call "hangs indefinitely rather than
    // failing" (Documents.Open, Quit). If the save wedges, the user must get a
    // typed error they can act on -- never a hang and never a create the agent
    // is tempted to retry into an overwrite.
    //
    // This is the author's half of the guarantee, and it is deliberately narrow
    // about what it proves. `DocumentAuthor` cannot bound anything itself; the
    // timer that actually stops a hang lives in `WordHost.#send`
    // (word-host-timeout.test.mjs covers it). What the author owns is *spending
    // one wall clock down across the operation* -- it hands the host
    // `remaining(...)`, a value bounded by `CREATE_BUDGET_MS` and shrinking as
    // the create proceeds -- and *not reshaping* the host's `word_timeout` on
    // the way back out.
    //
    // What deleting `timeoutMs: remaining("the authoring itself")` at
    // document-author.mjs actually costs, stated straight because the mutant is
    // the evidence: the host's `create` defaults `timeoutMs` to
    // `STARTUP_TIMEOUT_MS` (180 s), so the save does NOT run unbounded -- it
    // still surfaces as `word_timeout`. What is lost is the *shared* clock: the
    // bound stops being spent across authoring and read-back and reverts to a
    // flat per-call default, so a create can outlive the budget its own error
    // text quotes. This test sees that as the handed value going from a finite
    // number `<= CREATE_BUDGET_MS` to `undefined`.
    await withTemp(async (dir) => {
        const doc = path.join(dir, "report.docx");

        let handedTimeout = "unset";
        const author = new DocumentAuthor({
            reader: stubReader(doc),
            host: {
                create: async ({ timeoutMs } = {}) => {
                    handedTimeout = timeoutMs;
                    // What `WordHost.#send` produces when the `create` command
                    // does not answer within its budget: a typed rejection, the
                    // child already killed. No file is written -- a hung save
                    // wrote nothing.
                    throw Object.assign(new Error("Word did not respond to 'create' within 300s."), {
                        code: "word_timeout",
                    });
                },
            },
        });

        const err = await failed(author, doc);
        assert.ok(err, "a hung save resolved instead of failing");

        assert.ok(
            Number.isFinite(handedTimeout) && handedTimeout > 0 && handedTimeout <= 300_000,
            `the save was not handed a spent-down budget (got ${handedTimeout}); the shared wall clock was lost`,
        );

        assert.equal(err.code, "word_timeout", "a hung save reached the caller as something other than word_timeout");
        assert.match(throughToolBoundary(err), /word_timeout/);
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
