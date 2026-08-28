// Every fact worth telling the agent survives the tool boundary.
//
// These tests exist because the same defect was found twice: an error carrying
// rich details that `asToolError` strips, so the editor's own tests pass while
// the agent receives a bare code. The first fix patched one call site. This
// file asserts the property at the boundary, over the whole class, which is the
// only place the difference is observable.

import test from "node:test";
import assert from "node:assert/strict";

import { asToolError } from "../../src/tool-error.mjs";
import { EditError } from "../../src/word/document-editor.mjs";

test("a stale-token failure reaches the agent with both tokens", () => {
    // The error most likely to fire in normal use, and the one whose details
    // decide whether the agent should re-read or give up.
    const wrapped = asToolError(
        new EditError("stale_revision_token", "The document changed.", {
            expectedRevisionToken: "aaa",
            actualRevisionToken: "bbb",
        }),
    );

    assert.equal(wrapped.code, "stale_revision_token");
    assert.equal(wrapped.data?.expectedRevisionToken, "aaa");
    assert.equal(wrapped.data?.actualRevisionToken, "bbb");
});

test("every detail an EditError is given survives to the agent", () => {
    // Over the class rather than one instance: the per-site fix is what left
    // the other ten throw sites broken.
    const details = {
        paragraphCount: 12,
        expectedCount: 11,
        detail: "COM said no",
        snapshot: "snap-1.docx",
        rolledBack: false,
    };

    const wrapped = asToolError(new EditError("edit_failed", "Word failed.", details));

    for (const [key, value] of Object.entries(details)) {
        assert.deepEqual(wrapped.data?.[key], value, `${key} was lost at the tool boundary`);
    }
});

test("an EditError with no details does not arrive carrying an empty object", () => {
    const wrapped = asToolError(new EditError("no_snapshot", "Nothing to revert."));
    assert.equal(wrapped.data, undefined);
});

test("the code is folded into the message, because that is what the model reads", () => {
    const wrapped = asToolError(new EditError("file_locked", "It is held."));
    assert.equal(wrapped.message, "file_locked: It is held.");
});

test("a non-EditError still degrades to a usable shape", () => {
    assert.equal(asToolError(undefined).code, "word_error");
    assert.equal(asToolError(new Error("boom")).message, "word_error: boom");
});
