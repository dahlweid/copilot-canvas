// The read orchestration, with a stubbed host. Office-free.
//
// Word is not needed to test the part that matters most: that a document which
// changes underneath a read is reported rather than returned as a structure map
// addressing a document that no longer exists.
//
// Run: node --test ".github/extensions/office-canvas/test/unit/*.test.mjs"

import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import { DocumentReader, ReadError } from "../../src/word/document-reader.mjs";
import { flatOpc, paragraph } from "./word-fixtures.mjs";

const MARKUP = flatOpc([paragraph("Kapitel", { styleId: "berschrift1" }), paragraph("Body text.")]);

/**
 * Stands in for the Word host. `onCall` runs before the markup is written, so a
 * test can mutate the document mid-read exactly where Word would be busy.
 */
function stubHost({ markup = MARKUP, onCall = null, meta = {}, failures = 0 } = {}) {
    const calls = [];
    let remainingFailures = failures;
    return {
        calls,
        async structure(args) {
            calls.push(args);
            if (remainingFailures-- > 0) {
                const err = new Error("Word host exited");
                err.code = "word_unavailable";
                throw err;
            }
            await onCall?.(args);
            await writeFile(args.out, markup, "utf8");
            return { out: args.out, bytes: Buffer.byteLength(markup), writable: true, name: path.basename(args.path), ...meta };
        },
    };
}

const withWorkspace = async (fn) => {
    const dir = await mkdtemp(path.join(tmpdir(), "office-reader-"));
    try {
        const docPath = path.join(dir, "demo.docx");
        await writeFile(docPath, "original bytes");
        return await fn({ dir, docPath, workRoot: path.join(dir, "cache") });
    } finally {
        await rm(dir, { recursive: true, force: true });
    }
};

test("a read returns the structure map and the revision token together", async () => {
    await withWorkspace(async ({ docPath, workRoot }) => {
        const host = stubHost();
        const result = await new DocumentReader({ host, workRoot }).read(docPath);

        assert.match(result.revisionToken, /^sha256:[0-9a-f]{16}$/);
        assert.equal(result.paragraphCount, 2);
        assert.equal(result.paragraphs[0].headingLevel, 1);
        assert.equal(result.paragraphs[0].styleId, "berschrift1");
        assert.match(result.paragraphs[1].address, /^p:[0-9a-f]{12}$/);
        assert.equal(result.name, "demo.docx");
        assert.equal(result.writable, true);
        assert.equal(result.sizeBytes, "original bytes".length);
    });
});

test("the host is asked for the markup exactly once, not once per paragraph", async () => {
    // Walking paragraphs over COM measured 3724ms against 289ms for one
    // WordOpenXML call. A per-paragraph walk is a defect, so pin the call count.
    await withWorkspace(async ({ docPath, workRoot }) => {
        const host = stubHost();
        await new DocumentReader({ host, workRoot }).read(docPath);
        assert.equal(host.calls.length, 1);
    });
});

test("the scratch directory is cleaned up after a read", async () => {
    await withWorkspace(async ({ docPath, workRoot }) => {
        const host = stubHost();
        await new DocumentReader({ host, workRoot }).read(docPath);
        assert.equal(existsSync(host.calls[0].workDir), false);
    });
});

test("a document that changes during the read is retried, and the retry is returned", async () => {
    await withWorkspace(async ({ docPath, workRoot }) => {
        let changed = false;
        const host = stubHost({
            onCall: async () => {
                if (changed) return;
                changed = true;
                await writeFile(docPath, "rewritten while we were reading");
            },
        });
        const result = await new DocumentReader({ host, workRoot }).read(docPath);
        assert.equal(host.calls.length, 2);
        assert.equal(result.paragraphCount, 2);
    });
});

test("a document that keeps changing is refused rather than mis-addressed", async () => {
    await withWorkspace(async ({ docPath, workRoot }) => {
        let n = 0;
        const host = stubHost({ onCall: async () => writeFile(docPath, `revision ${n++}`) });
        await assert.rejects(
            new DocumentReader({ host, workRoot }).read(docPath),
            (err) => err instanceof ReadError && err.code === "document_changed_during_read",
        );
        assert.equal(host.calls.length, 2, "two attempts, then refuse");
    });
});

test("a dead host is retried once and the read still succeeds", async () => {
    // A structure read registers no reopen args, so the host's own replay
    // recovery cannot cover it.
    await withWorkspace(async ({ docPath, workRoot }) => {
        const host = stubHost({ failures: 1 });
        const result = await new DocumentReader({ host, workRoot }).read(docPath);
        assert.equal(host.calls.length, 2);
        assert.equal(result.paragraphCount, 2);
    });
});

test("a host failure that will not clear is surfaced, not swallowed", async () => {
    await withWorkspace(async ({ docPath, workRoot }) => {
        const host = stubHost({ failures: 99 });
        await assert.rejects(new DocumentReader({ host, workRoot }).read(docPath), (err) => err.code === "word_unavailable");
    });
});

test("a locked original is reported as not writable rather than failing the read", async () => {
    // The read itself works on a copy, so it is unaffected -- but an edit citing
    // this token would collide, and the agent should learn that up front.
    await withWorkspace(async ({ docPath, workRoot }) => {
        const host = stubHost({ meta: { writable: false } });
        assert.equal((await new DocumentReader({ host, workRoot }).read(docPath)).writable, false);
    });
});

test("paging is passed through to the structure map", async () => {
    await withWorkspace(async ({ docPath, workRoot }) => {
        const host = stubHost();
        const result = await new DocumentReader({ host, workRoot }).read(docPath, { limit: 1, offset: 1 });
        assert.equal(result.paragraphCount, 2);
        assert.equal(result.returned, 1);
        assert.equal(result.paragraphs[0].text, "Body text.");
    });
});

test("a missing file is a typed error and never reaches Word", async () => {
    await withWorkspace(async ({ dir, workRoot }) => {
        const host = stubHost();
        await assert.rejects(
            new DocumentReader({ host, workRoot }).read(path.join(dir, "absent.docx")),
            (err) => err instanceof ReadError && err.code === "file_not_found",
        );
        assert.equal(host.calls.length, 0, "starting Word to discover a typo would be expensive");
    });
});

test("a directory is rejected the same way a missing file is", async () => {
    await withWorkspace(async ({ dir, workRoot }) => {
        const host = stubHost();
        await assert.rejects(
            new DocumentReader({ host, workRoot }).read(dir),
            (err) => err instanceof ReadError && err.code === "file_not_found",
        );
    });
});

// --- scratch isolation -----------------------------------------------------
//
// The scratch id used to be derived from the document path alone, so two
// overlapping reads of one document shared a working directory, a source copy,
// an output file and the host-side docId. One call's cleanup could delete the
// other's markup mid-read, and one call's close could deregister the other's
// document.

test("two reads of the same document never share a scratch directory", async () => {
    await withWorkspace(async ({ docPath, workRoot }) => {
        const host = stubHost();
        const reader = new DocumentReader({ host, workRoot });
        await reader.read(docPath);
        await reader.read(docPath);

        assert.equal(host.calls.length, 2);
        assert.notEqual(host.calls[0].docId, host.calls[1].docId, "docId must be unique per call");
        assert.notEqual(host.calls[0].workDir, host.calls[1].workDir);
        assert.notEqual(host.calls[0].out, host.calls[1].out);
    });
});

test("concurrent reads of one document do not delete each other's markup", async () => {
    await withWorkspace(async ({ docPath, workRoot }) => {
        // Holds whichever read reaches the host first until the other has been
        // all the way through its own cleanup -- the interleaving that used to
        // lose the held read's structure.xml before it could be read back.
        // Which one arrives first is not ours to decide, so the test must not
        // assume it: awaiting a specific read before releasing deadlocks.
        let release;
        const held = new Promise((resolve) => {
            release = resolve;
        });
        let call = 0;
        const host = stubHost({
            onCall: async () => {
                if (call++ === 0) await held;
            },
        });
        const reader = new DocumentReader({ host, workRoot });

        const first = reader.read(docPath);
        const second = reader.read(docPath);
        await new Promise((resolve) => setTimeout(resolve, 50));
        release();

        for (const result of await Promise.all([first, second])) {
            assert.equal(result.paragraphCount, 2);
        }
    });
});

test("the scratch directory is removed after a successful read", async () => {
    await withWorkspace(async ({ docPath, workRoot }) => {
        const host = stubHost();
        await new DocumentReader({ host, workRoot }).read(docPath);
        assert.ok(!existsSync(host.calls[0].workDir), "scratch left behind");
    });
});

// --- bounded time ----------------------------------------------------------

test("the first attempt gets the full startup budget and a retry gets far less", async () => {
    await withWorkspace(async ({ docPath, workRoot }) => {
        const host = stubHost({ failures: 1 });
        await new DocumentReader({ host, workRoot }).read(docPath);

        assert.equal(host.calls.length, 2);
        assert.equal(host.calls[0].timeoutMs, 180_000);
        assert.ok(
            host.calls[1].timeoutMs <= 45_000,
            `a retry must not cost another full budget, got ${host.calls[1].timeoutMs}ms`,
        );
    });
});

test("a document that keeps changing is refused without unbounded retrying", async () => {
    await withWorkspace(async ({ docPath, workRoot }) => {
        let n = 0;
        const host = stubHost({ onCall: async () => writeFile(docPath, `changed ${n++}`) });
        await assert.rejects(
            () => new DocumentReader({ host, workRoot }).read(docPath),
            (err) => err instanceof ReadError && err.code === "document_changed_during_read",
        );
        assert.equal(host.calls.length, 2, "two attempts, then refuse");
    });
});

// --- filesystem errors must not escape untyped -------------------------------
//
// The revision token is the first thing a read touches on the original, before
// Word is involved at all, so it is where an exclusive lock surfaces. It used
// to surface as a raw `EBUSY` from the read stream: not a `ReadError`, not even
// a `word_*` code, so a caller could not tell "locked" from anything else. A
// consumer building edits on top has to branch on exactly this.

const errnoThrower = (code) => () => {
    const err = new Error(`simulated ${code}`);
    err.code = code;
    return Promise.reject(err);
};

test("an exclusively locked original is a typed file_locked error, not a raw errno", async () => {
    await withWorkspace(async ({ docPath, workRoot }) => {
        const reader = new DocumentReader({
            host: stubHost(),
            workRoot,
            tokenOf: errnoThrower("EBUSY"),
        });
        const err = await reader.read(docPath).then(
            () => null,
            (e) => e,
        );
        assert.ok(err instanceof ReadError, `expected a ReadError, got ${err?.name}: ${err?.message}`);
        assert.equal(err.code, "file_locked");
        assert.notEqual(err.code, "EBUSY", "the raw errno must not be the code the caller sees");
    });
});

test("a permission error is reported as permission_denied, not as a lock", async () => {
    // Measured on Windows: an ACL denying read gives EPERM, while an exclusive
    // FileShare::None lock gives EBUSY. They need different remediation -- a
    // lock may clear on its own, a permission will not -- so collapsing them
    // into one code repeats the single-`word_error` defect in miniature.
    for (const errno of ["EACCES", "EPERM"]) {
        await withWorkspace(async ({ docPath, workRoot }) => {
            const reader = new DocumentReader({
                host: stubHost(),
                workRoot,
                tokenOf: errnoThrower(errno),
            });
            const err = await reader.read(docPath).then(
                () => null,
                (e) => e,
            );
            assert.equal(err.code, "permission_denied", `${errno} should map to permission_denied`);
            assert.doesNotMatch(
                err.message,
                /another process (is )?holding/i,
                "a permission failure must not tell the caller to close another program",
            );
        });
    }
});

test("only a genuine sharing violation is described as another process holding the file", async () => {
    await withWorkspace(async ({ docPath, workRoot }) => {
        const reader = new DocumentReader({
            host: stubHost(),
            workRoot,
            tokenOf: errnoThrower("EBUSY"),
        });
        const err = await reader.read(docPath).then(
            () => null,
            (e) => e,
        );
        assert.equal(err.code, "file_locked");
        assert.match(err.message, /another process/i);
        // The distinction that makes the copy-based read possible at all.
        assert.match(err.message, /open in Word can still be read/i);
    });
});

test("a file that vanishes between the stat and the token read is file_not_found", async () => {
    await withWorkspace(async ({ docPath, workRoot }) => {
        const reader = new DocumentReader({
            host: stubHost(),
            workRoot,
            tokenOf: errnoThrower("ENOENT"),
        });
        const err = await reader.read(docPath).then(
            () => null,
            (e) => e,
        );
        assert.equal(err.code, "file_not_found");
    });
});

test("an unrecognised filesystem error still arrives typed", async () => {
    await withWorkspace(async ({ docPath, workRoot }) => {
        const reader = new DocumentReader({
            host: stubHost(),
            workRoot,
            tokenOf: errnoThrower("EIO"),
        });
        const err = await reader.read(docPath).then(
            () => null,
            (e) => e,
        );
        assert.ok(err instanceof ReadError);
        assert.equal(err.code, "document_unreadable");
    });
});

test("a host that reports success but writes no markup is typed, not a bare ENOENT", async () => {
    await withWorkspace(async ({ docPath, workRoot }) => {
        // Succeeds without ever writing the output file.
        const silentHost = {
            calls: 0,
            async structure() {
                this.calls += 1;
                return { name: "demo.docx", writable: true };
            },
        };
        const reader = new DocumentReader({ host: silentHost, workRoot });
        const err = await reader.read(docPath).then(
            () => null,
            (e) => e,
        );
        assert.ok(err instanceof ReadError, `expected a ReadError, got ${err?.name}`);
        assert.equal(err.code, "document_unreadable");
        assert.match(err.message, /wrote no structure/i);
    });
});
