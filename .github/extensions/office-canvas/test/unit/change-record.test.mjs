// What the overlay is told about the operation that just ran.
//
// Office-free: the record is built from an `edit_document` result, which is a
// plain object. The results here are shaped from the ones `document-editor.mjs`
// returns, so a change to that shape shows up as a failure rather than as an
// overlay that quietly stops appearing.

import { test } from "node:test";
import assert from "node:assert/strict";

import { changeRecordFrom } from "../../src/change-record.mjs";
import { OPERATION_NAMES } from "../../src/word/edit-intent.mjs";

const CLOCK = () => "2024-01-01T00:00:00.000Z";

/** An `edit_document` result, with the fields the record reads. */
const result = (over = {}) => ({
    applied: { op: "replace_text", description: "Replaced the paragraph text", ...over.applied },
    page: 3,
    paragraph: { address: "p:0123456789ab", text: "The replacement text.", style: "Standard", ...over.paragraph },
    document: { revisionToken: "sha256:abcdef0123456789", ...over.document },
    ...over,
});

test("a record is built for every operation the editor can apply", () => {
    // Derived from the editor's own list rather than a copy of it. A new
    // operation arriving with no thought given to what the overlay should draw
    // fails here instead of silently producing a box on the wrong thing.
    for (const op of OPERATION_NAMES) {
        const record = changeRecordFrom(result({ applied: { op, description: op } }), { now: CLOCK });
        assert.ok(record, `${op} produced no record`);
        assert.equal(record.op, op);
    }
});

test("a replacement carries the post-edit text and the page Word reported", () => {
    const record = changeRecordFrom(result(), { now: CLOCK });
    assert.deepEqual(record, {
        op: "replace_text",
        description: "Replaced the paragraph text",
        page: 3,
        text: "The replacement text.",
        locatable: true,
        at: "2024-01-01T00:00:00.000Z",
    });
});

test("the record carries no address, because an address cannot survive the edit", () => {
    // ADR 0006: an address is a coordinate valid inside one read-then-edit
    // cycle. The result the record is built from *does* contain one, so this
    // asserts a deliberate omission rather than an absence that happened.
    const source = result();
    assert.ok(source.paragraph.address, "the source result must carry an address for this to mean anything");
    const record = changeRecordFrom(source, { now: CLOCK });
    assert.ok(!("address" in record), "the record must not carry an address");
    assert.ok(
        !JSON.stringify(record).includes(source.paragraph.address),
        "the address must not reach the viewer under any key",
    );
});

test("the record carries no revision token, which is not the invalidator", () => {
    // Invalidation is on the render key, which moves on any write. A token
    // beside it would be a second rule that could disagree -- a byte-identical
    // regeneration leaves the token alone while still producing a new render.
    const source = result();
    assert.ok(source.document.revisionToken, "the source result must carry a token for this to mean anything");
    const record = changeRecordFrom(source, { now: CLOCK });
    assert.ok(!JSON.stringify(record).includes(source.document.revisionToken));
});

test("a deletion is recorded but is not locatable", () => {
    // Nothing is left on the page to find. Anchoring to whatever now follows the
    // gap would assert a position the code never determined.
    const record = changeRecordFrom(
        result({ applied: { op: "delete_paragraph", description: "Deleted the paragraph" } }),
        { now: CLOCK },
    );
    assert.equal(record.op, "delete_paragraph");
    assert.equal(record.locatable, false);
    assert.equal(record.text, null, "an unlocatable record must carry no text to search for");
    assert.equal(record.page, 3, "the page is still known, so the marker can still be placed");
});

test("an empty paragraph is not locatable either", () => {
    // A heading demoted to body text can leave an empty paragraph, and the empty
    // string matches at offset zero of every page.
    for (const text of ["", "   ", "\n\t", null, undefined]) {
        const record = changeRecordFrom(result({ paragraph: { text } }), { now: CLOCK });
        assert.equal(record.locatable, false, `${JSON.stringify(text)} should not be locatable`);
        assert.equal(record.text, null);
    }
});

test("the recorded text is normalized, so it can match the page", () => {
    const record = changeRecordFrom(result({ paragraph: { text: "  The   replacement\ttext.  " } }), { now: CLOCK });
    assert.equal(record.text, "The replacement text.");
});

test("a page Word could not report leaves the page null rather than guessing 1", () => {
    // Page 1 is a real page. Defaulting to it would put a marker on a page the
    // edit may not have touched, and nothing downstream could tell the
    // difference between a known page 1 and an unknown one.
    for (const page of [null, undefined, 0, -2, Number.NaN, "3", Infinity]) {
        const record = changeRecordFrom(result({ page }), { now: CLOCK });
        assert.equal(record.page, null, `page ${JSON.stringify(page)} should not be trusted`);
    }
});

test("a fractional page is floored rather than passed through", () => {
    assert.equal(changeRecordFrom(result({ page: 3.7 }), { now: CLOCK }).page, 3);
});

test("a result that applied nothing produces no record at all", () => {
    // Returning a record the viewer would have to know to ignore is how an
    // overlay ends up drawn for an operation that never happened.
    for (const empty of [null, undefined, {}, { applied: null }, { applied: {} }, { applied: { op: "" } }]) {
        assert.equal(changeRecordFrom(empty, { now: CLOCK }), null, `${JSON.stringify(empty)} should produce no record`);
    }
});

test("the timestamp comes from the injected clock", () => {
    const record = changeRecordFrom(result(), { now: () => "2020-02-02T02:02:02.000Z" });
    assert.equal(record.at, "2020-02-02T02:02:02.000Z");
});

test("the default clock produces a parseable ISO timestamp", () => {
    const record = changeRecordFrom(result());
    assert.ok(Number.isFinite(Date.parse(record.at)));
});
