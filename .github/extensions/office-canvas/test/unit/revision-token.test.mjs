// The revision token. Office-free.
// Run: node --test ".github/extensions/office-canvas/test/unit/*.test.mjs"

import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { fileRevisionToken, revisionTokenOf, TOKEN_PREFIX, tokensMatch } from "../../src/revision-token.mjs";

const withTempDir = async (fn) => {
    const dir = await mkdtemp(path.join(tmpdir(), "office-token-"));
    try {
        return await fn(dir);
    } finally {
        await rm(dir, { recursive: true, force: true });
    }
};

test("a token is prefixed and fixed-length", () => {
    const token = revisionTokenOf(Buffer.from("some document bytes"));
    assert.ok(token.startsWith(TOKEN_PREFIX));
    assert.match(token, /^sha256:[0-9a-f]{16}$/);
});

test("the same bytes always give the same token", () => {
    assert.equal(revisionTokenOf(Buffer.from("abc")), revisionTokenOf(Buffer.from("abc")));
});

test("one changed byte changes the token", () => {
    // The point of hashing the content rather than trusting mtime and size: a
    // regenerated document can reproduce both while changing every word.
    const before = Buffer.from("The quick brown fox jumps over the lazy dog.");
    const after = Buffer.from("The quick brown fox jumps over the lazy dot.");
    assert.equal(before.length, after.length);
    assert.notEqual(revisionTokenOf(before), revisionTokenOf(after));
});

test("a string and its bytes agree", () => {
    assert.equal(revisionTokenOf("hello"), revisionTokenOf(Buffer.from("hello", "utf8")));
});

test("an empty document still has a token", () => {
    assert.match(revisionTokenOf(Buffer.alloc(0)), /^sha256:[0-9a-f]{16}$/);
});

test("a file token matches the token of the same bytes read whole", async () => {
    await withTempDir(async (dir) => {
        // Bigger than one stream chunk, so a multi-chunk hash is actually exercised.
        const bytes = Buffer.alloc(200_000, "docx-ish payload ");
        const file = path.join(dir, "doc.docx");
        await writeFile(file, bytes);
        assert.equal(await fileRevisionToken(file), revisionTokenOf(bytes));
    });
});

test("rewriting a file with different content moves its token", async () => {
    await withTempDir(async (dir) => {
        const file = path.join(dir, "doc.docx");
        await writeFile(file, "first");
        const before = await fileRevisionToken(file);
        await writeFile(file, "second");
        assert.notEqual(await fileRevisionToken(file), before);
    });
});

test("rewriting a file with identical content leaves its token alone", async () => {
    // Mirrors the measured Word behaviour: a save with nothing dirty does not
    // move the token, so inspecting a document never forces a re-read.
    await withTempDir(async (dir) => {
        const file = path.join(dir, "doc.docx");
        await writeFile(file, "same");
        const before = await fileRevisionToken(file);
        await writeFile(file, "same");
        assert.equal(await fileRevisionToken(file), before);
    });
});

test("a missing file rejects rather than returning a token for nothing", async () => {
    await assert.rejects(fileRevisionToken(path.join(tmpdir(), "office-token-does-not-exist.docx")));
});

test("tokensMatch is exact", () => {
    assert.equal(tokensMatch("sha256:0123456789abcdef", "sha256:0123456789abcdef"), true);
    assert.equal(tokensMatch("sha256:0123456789abcdef", "sha256:0123456789abcdee"), false);
});

test("tokensMatch never lets a missing token pass", () => {
    // An edit refuses when the token has moved, so "no token" must read as a
    // mismatch. Two nullish tokens comparing equal would defeat the check.
    for (const [a, b] of [
        [null, null],
        [undefined, undefined],
        ["", ""],
        [null, "sha256:0123456789abcdef"],
        ["sha256:0123456789abcdef", null],
        [undefined, undefined],
    ]) {
        assert.equal(tokensMatch(a, b), false, `expected no match for ${JSON.stringify([a, b])}`);
    }
});
