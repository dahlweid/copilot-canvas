// Every fact worth telling the agent survives the tool boundary.
//
// This file was written twice, and the first version is the more useful half of
// its own history. It asserted that `asToolError` forwarded `code`, `message`
// and `data`, and it passed — while the agent, on the other side of that
// boundary, received the string `Tool execution failed` and nothing else. The
// tests were green, the comments described a real mechanism, and the property
// they pinned reached no caller, because the boundary they called "the boundary
// the agent sees" was two hops short of it.
//
// So the assertions here are deliberately shaped around the *measured* channel
// (spikes/tool-errors/FINDINGS.md): the only thing that crosses is
// `textResultForLlm`. Anything asserted about `code` or `data` as separate
// fields would be exactly the earlier mistake in a new spelling — which is why
// every assertion below reads the text.
//
// ## Mutations used to confirm these can go red
//
// Run over this file alone, from the extension folder.
//
//   A. src/tool-error.mjs -> `resultType: "success"`. 8 pass, 2 fail. Pins the
//      shape the SDK checks and the flag the runtime forwards.
//   B. src/tool-error.mjs -> drop the `${toolName} failed.` banner line. 8 pass,
//      2 fail. Pins the one thing that tells the agent this is not a result.
//   C. src/tool-error.mjs -> `renderDetails` returns null always, i.e. the old
//      behaviour of leaving facts on `data`. 7 pass, 3 fail.
//   D. src/tool-error.mjs -> paste `json.slice(0, MAX_DETAIL_CHARS)` instead of
//      stating the size. 9 pass, 1 fail.

import test from "node:test";
import assert from "node:assert/strict";

import { toolFailure } from "../../src/tool-error.mjs";
import { EditError } from "../../src/word/document-editor.mjs";

/**
 * The SDK's own predicate, copied from `copilot-sdk/extension.js`.
 *
 * A copy, and it has to be: the SDK is auto-resolved at runtime and is not on
 * disk for a test. What it buys is the difference between "our object looks
 * right to us" and "our object is one the SDK routes down the result path" — an
 * object failing this check is silently `JSON.stringify`d instead, which is a
 * different wire shape reaching the agent. Pinned to the source it was copied
 * from in spikes/tool-errors/FINDINGS.md.
 */
function isToolResultObject(value) {
    if (typeof value !== "object" || value === null) return false;
    if (!("textResultForLlm" in value) || typeof value.textResultForLlm !== "string") return false;
    if (!("resultType" in value) || typeof value.resultType !== "string") return false;
    return ["success", "failure", "rejected", "denied", "timeout"].includes(value.resultType);
}

test("a failure is returned in the shape the SDK forwards as a result", () => {
    // The whole fix rests on this: a throw is discarded, a result is not. An
    // object the SDK does not recognise takes the JSON.stringify path instead.
    const failure = toolFailure("read_document", new EditError("file_locked", "It is held."));

    assert.ok(isToolResultObject(failure), "the SDK would not route this down the result path");
    assert.equal(failure.resultType, "failure");
});

test("the text says the tool failed, because nothing else the agent receives does", () => {
    // Measured: `resultType: "failure"` arrives looking exactly like a success
    // and the `error` field does not arrive at all. If the text does not say
    // it, the agent is not told it.
    const failure = toolFailure("read_document", new EditError("file_locked", "It is held."));

    assert.match(failure.textResultForLlm, /^read_document failed\./);
});

test("the code and the message both reach the agent, in the text", () => {
    const failure = toolFailure("edit_document", new EditError("stale_revision_token", "The document changed."));

    assert.match(failure.textResultForLlm, /stale_revision_token/);
    assert.match(failure.textResultForLlm, /The document changed\./);
});

test("a load-bearing message crosses verbatim", () => {
    // `file_locked`'s wording is not decoration: it distinguishes a strict
    // holder from Word's own, and the intuitive reading of the code is wrong.
    // Asserted as an exact substring so a reworded or truncated crossing fails.
    const sentence =
        "Another process is holding report.docx open more strictly than Word does. " +
        "A document merely open in Word can still be read.";
    const failure = toolFailure("read_document", new EditError("file_locked", sentence));

    assert.ok(
        failure.textResultForLlm.includes(sentence),
        `the sentence did not cross intact:\n${failure.textResultForLlm}`,
    );
});

test("every detail an EditError is given reaches the agent", () => {
    // Over the class rather than one instance. The previous version of this
    // test asserted the same property against `wrapped.data`, which no caller
    // could read; the details have to be *in the text* or they are nowhere.
    const details = {
        expectedRevisionToken: "aaa",
        actualRevisionToken: "bbb",
        paragraphCount: 12,
        detail: "COM said no",
        snapshot: "snap-1.docx",
        rolledBack: false,
    };

    const text = toolFailure("edit_document", new EditError("edit_failed", "Word failed.", details)).textResultForLlm;

    for (const [key, value] of Object.entries(details)) {
        assert.match(text, new RegExp(key), `${key} was lost at the tool boundary`);
        assert.match(text, new RegExp(String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&")), `${key}'s value was lost`);
    }
});

test("a failure with no details does not carry an empty details line", () => {
    const failure = toolFailure("revert_document", new EditError("no_snapshot", "Nothing to revert."));

    assert.doesNotMatch(failure.textResultForLlm, /Details:/);
});

test("details too large to include are stated, not pasted as broken JSON", () => {
    // Truncating JSON produces something that is not JSON, so an agent parsing
    // the details gets a second failure whose cause is us.
    const failure = toolFailure("edit_document", new EditError("edit_failed", "Word failed.", { blob: "x".repeat(9000) }));

    assert.match(failure.textResultForLlm, /too large to include/);
    assert.ok(failure.textResultForLlm.length < 1000, "the oversized bag was pasted after all");
});

test("details that cannot be serialized are reported rather than thrown", () => {
    // A throw here would land in the SDK's catch — the very channel that
    // discards everything — so an unserializable `data` would turn a legible
    // failure back into `Tool execution failed`.
    const circular = { name: "loop" };
    circular.self = circular;

    const failure = toolFailure("edit_document", new EditError("edit_failed", "Word failed.", circular));

    assert.match(failure.textResultForLlm, /could not be serialized/);
    assert.equal(failure.resultType, "failure");
});

test("a plain Error still arrives with a code and its message", () => {
    const failure = toolFailure("read_document", new Error("boom"));

    assert.match(failure.textResultForLlm, /read_document failed\./);
    assert.match(failure.textResultForLlm, /word_error: boom/);
});

test("a thrown non-error degrades to a usable failure rather than to 'undefined'", () => {
    // `throw undefined` is legal and has happened; so is a rejected string.
    assert.match(toolFailure("read_document", undefined).textResultForLlm, /word_error: Unknown error/);
    assert.match(toolFailure("read_document", "just a string").textResultForLlm, /word_error: just a string/);
});
