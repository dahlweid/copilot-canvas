// Document path normalization — the identity function for a document.
//
// Office-free. Importing this pulls in render-cache.mjs, which imports the Word
// host module; that import must stay side-effect free, or nothing here can run
// on a hosted runner. These tests fail loudly if that ever changes.
// Run: node --test ".github/extensions/office-canvas/test/unit/*.test.mjs"

import test from "node:test";
import assert from "node:assert/strict";
import path from "node:path";
import { normalizeDocPath, DocumentError } from "../../src/render-cache.mjs";

const isWindows = process.platform === "win32";

test("a relative path is made absolute", () => {
    const result = normalizeDocPath("doc.docx");
    assert.equal(path.isAbsolute(result), true);
    assert.equal(result, path.resolve("doc.docx"));
});

test("surrounding quotes are stripped", () => {
    // Users paste quoted paths out of Explorer and shells constantly.
    assert.equal(normalizeDocPath('"doc.docx"'), normalizeDocPath("doc.docx"));
});

test("surrounding whitespace is stripped", () => {
    assert.equal(normalizeDocPath("  doc.docx  "), normalizeDocPath("doc.docx"));
});

test("an empty or non-string path is a typed error", () => {
    for (const bad of ["", "   ", null, undefined, 42, {}, []]) {
        assert.throws(
            () => normalizeDocPath(bad),
            (err) => err instanceof DocumentError && err.code === "invalid_path",
            `expected invalid_path for ${JSON.stringify(bad)}`,
        );
    }
});

test("normalization is idempotent", () => {
    const once = normalizeDocPath("sub/../doc.docx");
    assert.equal(normalizeDocPath(once), once);
});

test("traversal segments are resolved away", () => {
    assert.equal(normalizeDocPath("a/b/../doc.docx"), path.resolve("a/doc.docx"));
});

test("separators are normalized for the platform", () => {
    const result = normalizeDocPath("a/b/doc.docx");
    if (isWindows) {
        assert.equal(result.includes("/"), false, "Windows paths must not keep forward slashes");
        assert.equal(result.endsWith("a\\b\\doc.docx"), true);
    } else {
        assert.equal(result.endsWith("a/b/doc.docx"), true);
    }
});

test("an absolute path is preserved", () => {
    const input = isWindows ? "C:\\docs\\report.docx" : "/docs/report.docx";
    assert.equal(normalizeDocPath(input), input);
});
