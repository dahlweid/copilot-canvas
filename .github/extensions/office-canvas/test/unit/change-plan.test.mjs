// Which pages the overlay marks.
//
// Every case here was reachable in the running viewer; the one that mattered was
// found there, not here. An edit located correctly on page 3 also badged page 2,
// because the page before the reported one was marked on adjacency before
// anything had been searched. These tests exist so the fix cannot be undone
// without a red run.
//
// Office-free.

import { test } from "node:test";
import assert from "node:assert/strict";

import { planChangeMarks } from "../../src/ui/change-plan.mjs";

// A text item as pdf.js reports it. Only `str` matters to the locator; the
// transform is what the viewer later turns into quads.
const item = (str) => ({ str, transform: [10, 0, 0, 10, 50, 700], width: str.length * 5, height: 10 });

/** A page that has painted, carrying the given lines. */
const painted = (number, ...lines) => ({ number, items: lines.map(item) });

/** A page that has not painted, so it can answer nothing. */
const blank = (number) => ({ number, items: null });

const record = (over = {}) => ({
    op: "replace_text",
    page: 3,
    text: "the replacement paragraph",
    locatable: true,
    at: "2024-01-01T00:00:00.000Z",
    ...over,
});

test("only the page the text was found on is marked", () => {
    const marks = planChangeMarks(
        record(),
        [painted(1, "nothing here"), painted(2, "nor here"), painted(3, "the replacement paragraph")],
    );
    assert.deepEqual(
        marks.map((m) => m.number),
        [3],
    );
    assert.equal(marks[0].found.status, "located");
});

test("the page before the reported one is marked when the text is there", () => {
    // The straddle insurance, doing its job. The reported page has no match and
    // must not be marked just for having been reported.
    const marks = planChangeMarks(
        record(),
        [painted(1, "nothing"), painted(2, "the replacement paragraph"), painted(3, "something else")],
    );
    assert.deepEqual(
        marks.map((m) => m.number),
        [2],
    );
});

test("a paragraph split across the break is marked on both halves", () => {
    const marks = planChangeMarks(
        record({ text: "the replacement paragraph" }),
        [painted(1, "x"), painted(2, "the replacement paragraph"), painted(3, "the replacement paragraph")],
    );
    assert.deepEqual(
        marks.map((m) => m.number),
        [2, 3],
    );
});

test("a page is never marked on adjacency alone", () => {
    // The defect this module was extracted for. Page 2 is searchable, holds
    // nothing, and is adjacent to the reported page -- which used to be enough.
    const marks = planChangeMarks(record(), [painted(2, "unrelated text"), painted(3, "the replacement paragraph")]);
    assert.ok(
        !marks.some((m) => m.number === 2),
        `page 2 was marked with no match behind it: ${JSON.stringify(marks)}`,
    );
});

test("the reported page carries a page-level marker when the text is nowhere", () => {
    const marks = planChangeMarks(record(), [painted(2, "unrelated"), painted(3, "also unrelated")]);
    assert.deepEqual(marks, [{ number: 3, found: null }]);
});

test("nothing is marked while a candidate page has not painted", () => {
    // "Not found there" is unknown, not false, until the page can answer. Falling
    // back now would put a marker on the reported page and leave it there after
    // the real page painted with a match on it.
    assert.deepEqual(planChangeMarks(record(), [painted(2, "unrelated"), blank(3)]), []);
    assert.deepEqual(planChangeMarks(record(), [blank(2), painted(3, "unrelated")]), []);
});

test("an unpainted page does not suppress a match found elsewhere", () => {
    const marks = planChangeMarks(record(), [blank(2), painted(3, "the replacement paragraph")]);
    assert.deepEqual(
        marks.map((m) => m.number),
        [3],
    );
});

test("a deletion marks the reported page and only that page", () => {
    const marks = planChangeMarks(
        record({ op: "delete_paragraph", locatable: false, text: null }),
        [painted(2, "still here"), painted(3, "also still here")],
    );
    assert.deepEqual(marks, [{ number: 3, found: null }]);
});

test("a deletion on an unpainted page is still marked", () => {
    // There is nothing to search for, so painting cannot change the answer and
    // waiting would mean never marking it at all.
    assert.deepEqual(planChangeMarks(record({ locatable: false, text: null }), [blank(2), blank(3)]), [
        { number: 3, found: null },
    ]);
});

test("a reported page outside the document marks nothing", () => {
    assert.deepEqual(planChangeMarks(record({ page: 9 }), [painted(1, "a"), painted(2, "b")]), []);
    assert.deepEqual(planChangeMarks(record({ page: 9, locatable: false, text: null }), [painted(1, "a")]), []);
});

test("no record and no pages mark nothing", () => {
    assert.deepEqual(planChangeMarks(null, [painted(1, "a")]), []);
    assert.deepEqual(planChangeMarks(record(), []), []);
    assert.deepEqual(planChangeMarks(record(), null), []);
});
