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
