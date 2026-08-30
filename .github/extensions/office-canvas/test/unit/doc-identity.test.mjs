// How the bar identifies a document once the absolute-path row is gone (#71).
//
// The old row set `textContent` and `title` to the same string, so its tooltip
// revealed what was already on screen. The path now lives in two places that
// each say something the other does not: a tooltip on the name, for a pointer,
// and the accessible name of a real button, for everyone else. This file
// executes that decision -- `app.js`, which applies it, cannot be imported under
// Node at all, so nothing there can be asserted except as text.

import { test } from "node:test";
import assert from "node:assert/strict";

import { copyOutcome, describeDocument, folderOf } from "../../src/ui/doc-identity.mjs";

const WINDOWS_DOC = { name: "report.docx", path: "C:\\Users\\ada\\Documents\\report.docx" };

test("the name's tooltip is the full path, which is not what the name says", async () => {
    // The whole complaint in #71: a tooltip that repeats its own element earns
    // nothing. This one carries the location, which is absent from the name.
    const { nameTitle } = describeDocument(WINDOWS_DOC);

    assert.equal(nameTitle, WINDOWS_DOC.path);
    assert.notEqual(nameTitle, WINDOWS_DOC.name);
});

test("the path is also the copy button's accessible name, so it survives without a pointer", async () => {
    // `title` is not keyboard-reachable and screen readers treat it
    // inconsistently, so a tooltip alone would have been a regression for
    // exactly the users the old row served. A button's accessible name is
    // announced on focus and reached by tab.
    const { copyLabel, copyTitle, canCopy } = describeDocument(WINDOWS_DOC);

    assert.ok(copyLabel.includes(WINDOWS_DOC.path), "the copy control does not announce the path");
    assert.match(copyLabel, /^Copy/, "the copy control's name does not say what pressing it does");
    assert.equal(copyTitle, WINDOWS_DOC.path);
    assert.equal(canCopy, true);
});

test("a document whose path is just its name gets no tooltip at all", async () => {
    // Rather than reintroducing the emptiness that cost the old row its place.
    const { nameTitle, folder } = describeDocument({ name: "report.docx", path: "report.docx" });

    assert.equal(nameTitle, "");
    assert.equal(folder, null);
});

test("a document with no usable path leaves the copy control with nothing to copy", async () => {
    for (const doc of [{ name: "report.docx" }, { name: "report.docx", path: "" }, { name: "report.docx", path: "  " }]) {
        const described = describeDocument(doc);
        assert.equal(described.canCopy, false, `${JSON.stringify(doc)} was treated as copyable`);
        assert.equal(described.nameTitle, "", `${JSON.stringify(doc)} produced a tooltip`);
        // Still named, because a disabled control is still read out.
        assert.equal(described.copyLabel, "Copy full path");
    }
});

test("the folder is read off the path's own separator, not the host's", async () => {
    // This module runs in a browser: `node:path` is unavailable and
    // `process.platform` says nothing about where the document lives. A POSIX
    // path handled by Windows rules would lose its folder entirely.
    assert.equal(folderOf("C:\\Users\\ada\\Documents\\report.docx"), "C:\\Users\\ada\\Documents");
    assert.equal(folderOf("/home/ada/docs/report.docx"), "/home/ada/docs");
    assert.equal(folderOf("report.docx"), null, "a bare filename has no folder");
    assert.equal(folderOf("\\report.docx"), null, "a root-relative name would report an empty folder");
    assert.equal(folderOf(null), null);
});

test("a failed copy says so without pretending the path is lost", async () => {
    // `writeText` rejects when the document is not focused, and
    // `navigator.clipboard` is absent outside a secure context. Neither is worth
    // an error banner that offers no way forward: the path is still on screen in
    // the tooltip and in the button's own name, so the message points there.
    const ok = copyOutcome(null);
    const bad = copyOutcome(new Error("Document is not focused."));

    assert.equal(ok.error, false);
    assert.match(ok.text, /copied/i);

    assert.equal(bad.error, true);
    assert.match(bad.text, /not copy/i);
    assert.match(bad.text, /tooltip/i, "a failed copy leaves the user with no way to see the path");
});
