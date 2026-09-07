// What `edit_document` hands back to the agent.
//
// The editor's result gained `previousText` for the change overlay (#166): the
// paragraph as it stood before the edit, which `changeRecordFrom` diffs against
// the post-edit text to narrow the highlight. That field exists for the canvas,
// not for the agent -- the agent supplied the replacement and holds the read the
// address came from, so echoing the old paragraph back adds a second copy of it
// to every edit result and one more field a reader has to decide to ignore.
//
// The handler therefore consumes it and drops it. This asserts the drop, which
// is otherwise a line no Office-free test would touch: the real editor needs
// Word, so the stub cache reports the field in its place.

import test, { after, before } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { loadExtension } from "./extension-stubs.mjs";
import { STUB_PREVIOUS_TEXT } from "./stub-render-cache.mjs";

let home;
let sdk;

before(async () => {
    home = await mkdtemp(path.join(tmpdir(), "office-canvas-edit-shape-"));
    process.env.COPILOT_HOME = home;
    sdk = await loadExtension();
});

after(async () => {
    if (home) await rm(home, { recursive: true, force: true });
});

test("the edit result does not carry the paragraph's pre-edit text back to the agent", async () => {
    const tool = sdk.joined.tools.find((t) => t.name === "edit_document");
    assert.ok(tool, "edit_document must be registered");

    const doc = path.join(home, "report.docx");
    const result = await tool.handler({
        path: doc,
        op: "replace_text",
        address: "p:0123456789ab",
        revisionToken: "abc",
        text: "Hello.",
    });

    assert.notEqual(result?.resultType, "failure", `expected success, got ${JSON.stringify(result)?.slice(0, 160)}`);
    assert.ok(!("previousText" in result), "the pre-edit text travelled to the agent");
    assert.ok(
        !JSON.stringify(result).includes(STUB_PREVIOUS_TEXT),
        "the pre-edit text reached the agent under some other key",
    );
    // The rest of the result is untouched -- the drop must not take the document
    // with it, which a mistyped rest-spread would.
    assert.equal(result.document.path, doc);
});
