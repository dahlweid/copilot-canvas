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

import { candidatePages, planChangeMarks } from "../../src/ui/change-plan.mjs";
import { MIN_PARTIAL_CHARS } from "../../src/ui/locate-text.mjs";

// A text item as pdf.js reports it. Only `str` matters to the locator; the
// transform is what the viewer later turns into quads.
const item = (str) => ({ str, transform: [10, 0, 0, 10, 50, 700], width: str.length * 5, height: 10 });

/** A page that has painted, carrying the given lines. */
const painted = (number, ...lines) => ({ number, items: lines.map(item) });

/** A page that has not painted, so it can answer nothing. */
const blank = (number) => ({ number, items: null });

// A paragraph long enough that each half still clears MIN_PARTIAL_CHARS, split at
// the page break. Composed so the two halves concatenate to the whole -- the
// assertion below would be worthless if the fixture merely put related text on
// two pages.
const STRADDLE_HEAD = "the replacement paragraph that ";
const STRADDLE_TAIL = "runs across the page break";
const STRADDLE_FULL = STRADDLE_HEAD + STRADDLE_TAIL;

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

test("a paragraph genuinely split across the break is marked on both halves", () => {
    // The real straddle: a prefix ends page 2, the suffix begins page 3, and
    // neither page carries the whole search text. `locateText`'s partial matching
    // is what covers this, so the statuses are asserted and not just the page
    // numbers -- marking [2, 3] is reachable with partial matching deleted
    // entirely, by any fixture that puts the complete text on both pages.
    const marks = planChangeMarks(record({ text: STRADDLE_FULL }), [
        painted(1, "x"),
        painted(2, STRADDLE_HEAD),
        painted(3, STRADDLE_TAIL),
    ]);
    assert.ok(
        STRADDLE_HEAD.trim().length >= MIN_PARTIAL_CHARS && STRADDLE_TAIL.length >= MIN_PARTIAL_CHARS,
        "precondition: each half must clear the partial-match floor, or this measures the fallback instead",
    );
    assert.deepEqual(
        marks.map((m) => [m.number, m.found?.status]),
        [
            [2, "partial"],
            [3, "partial"],
        ],
    );
});

test("a split whose halves are each too short to match falls back to a page marker", () => {
    // Partial matching has a floor, so a short paragraph broken across the page
    // break is genuinely unlocatable. The requirement is that it degrades to the
    // page-level marker rather than marking nothing: something did change on the
    // reported page and the user is entitled to know which page.
    const half = "a".repeat(MIN_PARTIAL_CHARS - 1);
    const marks = planChangeMarks(record({ text: `${half} ${half}` }), [painted(2, `${half} `), painted(3, half)]);
    assert.deepEqual(marks, [{ number: 3, found: null }]);
});

test("the same text on two pages is marked on both, which over-marks by design", () => {
    // Not a straddle -- this is one paragraph's text appearing whole on two pages,
    // a running header or a repeated heading. Both are `located`, so both are
    // boxed, and at most one of them is the edit.
    //
    // Recorded as a known over-mark rather than fixed. The alternative is to pick
    // one, and nothing here distinguishes them: `page` narrows the field to two
    // candidates and says nothing about which. Picking would assert a position the
    // code never determined, which is the rule this module exists to keep.
    const marks = planChangeMarks(record(), [
        painted(1, "x"),
        painted(2, "the replacement paragraph"),
        painted(3, "the replacement paragraph"),
    ]);
    assert.deepEqual(
        marks.map((m) => [m.number, m.found?.status]),
        [
            [2, "located"],
            [3, "located"],
        ],
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

// --- the narrowed box (#166) ---------------------------------------------------

test("a record's span narrows the mark to the line that changed", () => {
    // End to end through the plan: the record still joins on the paragraph, and
    // the box lands on the part of it that changed.
    const marks = planChangeMarks(
        record({ text: "the replacement paragraph that changed here", span: "changed here" }),
        [painted(2, "nothing here"), painted(3, "the replacement paragraph that ", "changed here")],
    );
    assert.deepEqual(
        marks.map((m) => m.number),
        [3],
    );
    assert.equal(marks[0].found.narrowed, true);
    assert.deepEqual(marks[0].found.range, { startItem: 1, endItem: 1 });
});

test("a span that cannot be narrowed leaves the paragraph's own box", () => {
    // Every failure to narrow -- ambiguous inside the paragraph, absent from the
    // page, or no span at all -- must land on the pre-#166 behaviour rather than
    // on no box.
    for (const span of [null, "not on this page", "the"]) {
        const marks = planChangeMarks(
            record({ text: "the replacement paragraph the end", span }),
            [painted(3, "the replacement paragraph ", "the end")],
        );
        assert.equal(marks.length, 1, `${JSON.stringify(span)} should still mark the page`);
        assert.equal(marks[0].found.narrowed, false);
        assert.deepEqual(marks[0].found.range, { startItem: 0, endItem: 1 });
    }
});

test("a span never pulls the mark onto a page the paragraph is not on", () => {
    // The straddle page is searched for the *paragraph*. A span that also occurs
    // there, in unrelated text, must not put a box on it -- that is the box
    // asserting a position nothing determined.
    const marks = planChangeMarks(
        record({ text: "the replacement paragraph that changed here", span: "changed here" }),
        [painted(2, "something else changed here entirely"), painted(3, "the replacement paragraph that changed here")],
    );
    assert.deepEqual(
        marks.map((m) => m.number),
        [3],
    );
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

// --- which pages the viewer must re-check when one paints ----------------------
//
// Review round 1 found the defect these pin. `planChangeMarks` searches the
// reported page *and* the one before it, but the viewer decided for itself when
// a freshly painted page was worth re-planning, and asked only
// `record.page === page.number`. A change that actually sat on the straddle page
// was therefore planned but never triggered: painting the page it was on did
// not count as painting a page of interest, so the overlay stayed missing until
// something unrelated forced a re-apply. One rule, two copies, and the copies
// disagreed.

test("the page before the reported one is a candidate", () => {
    // This is the whole defect: 4 must be in the set, or painting page 4 never
    // re-applies the change that is on it.
    assert.deepEqual(candidatePages(record({ page: 5 })), [4, 5]);
});

test("a deletion has no straddle page to check", () => {
    // Nothing to find, so the page before earns no look.
    assert.deepEqual(candidatePages(record({ page: 5, locatable: false, text: null })), [5]);
});

test("page 1 never produces a page 0", () => {
    assert.deepEqual(candidatePages(record({ page: 1 })), [1]);
});

test("a record with no usable page has no candidates", () => {
    for (const page of [null, undefined, 0, -1, 2.5, NaN, "3"]) {
        assert.deepEqual(candidatePages(record({ page })), [], `page ${JSON.stringify(page)}`);
    }
    assert.deepEqual(candidatePages(null), []);
});

test("every page the plan can mark is a page the viewer would re-check", () => {
    // The property that actually matters, stated so neither side can drift: if
    // `planChangeMarks` would put a mark on a page, `candidatePages` must name
    // it -- otherwise that mark is unreachable when the page paints late, which
    // for anything below the fold is the normal case rather than an edge one.
    const cases = [
        record({ page: 5 }),
        record({ page: 5, locatable: false, text: null }),
        record({ page: 1 }),
        record({ page: 2 }),
    ];
    const pages = [painted(1, "alpha"), painted(2, "hello there"), painted(3, "gamma"),
                   painted(4, "hello there"), painted(5, "epsilon")];
    for (const rec of cases) {
        const marked = planChangeMarks(rec, pages).map((mark) => mark.number);
        const candidates = candidatePages(rec);
        for (const number of marked) {
            assert.ok(
                candidates.includes(number),
                `plan marks page ${number} for ${JSON.stringify(rec)} but candidatePages says ${JSON.stringify(candidates)}`,
            );
        }
    }
});

test("the viewer derives the trigger instead of restating it", async () => {
    // A source check, kept for what behaviour cannot reach: the one line that
    // had the bug is a private call inside `#renderPage`, and asserting the line
    // is gone stops the fix being one careless edit from returning. Its original
    // justification -- that `pdf-view.mjs` could not be loaded under Node -- no
    // longer holds; since #76 `pdf-view.test.mjs` executes the re-plan trigger
    // ("a page painted after the change arrived is still marked"). This check now
    // guards the *shape* of the fix, and that one does the work.
    const { readFile } = await import("node:fs/promises");
    const { fileURLToPath } = await import("node:url");
    const source = await readFile(
        fileURLToPath(new URL("../../src/ui/pdf-view.mjs", import.meta.url)),
        "utf8",
    );
    assert.match(source, /import \{[^}]*\bcandidatePages\b[^}]*\} from "\.\/change-plan\.mjs"/);
    assert.match(source, /candidatePages\(this\.#change\)\.includes\(page\.number\)/);
    assert.doesNotMatch(
        source,
        /#change\??\.page\s*===\s*page\.number/,
        "the viewer is deciding for itself which pages matter again",
    );
});
