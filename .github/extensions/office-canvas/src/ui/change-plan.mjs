// Which pages the overlay marks, and why.
//
// Split out of `pdf-view.mjs` because the decision is the part worth testing and
// the drawing is not. It was also, when written, the only way to reach it at
// all: `pdf-view.mjs` imports pdf.js from an absolute `/vendor/` URL and could
// not be loaded under Node. #76 removed that -- `test/unit/ui-harness.mjs`
// resolves the specifier to a stand-in, and `pdf-view.test.mjs` now drives the
// view. The split stands on its first reason alone.
//
// ## The rule
//
// A page is marked when the record's text was found on it. Nothing else earns a
// mark. The reported page gets a page-level marker only when the text was found
// on no candidate *and* every candidate was searchable -- a page that has not
// painted yet has no text to search, so "not found there" is unknown rather
// than false, and concluding from it would put a marker on the wrong page for as
// long as it takes the reader to scroll.
//
// The page before the reported one is searched as well, because a paragraph that
// straddles a page break has its opening lines there. Why the reported page can
// disagree with the rendered one is *not* established: `wdActiveEndPageNumber`
// names the end of the range, which suggests the later page, but the only
// straddle-free observation to hand had them agree. The second page is cheap
// insurance, not a conclusion -- and it is insurance that must not turn into a
// claim, which is what the "found" condition above is for.
//
// Measured: before that condition existed, an edit located correctly on page 3
// also put a "changed" badge on page 2, on the strength of adjacency alone.

import { locateText } from "./locate-text.mjs";

/**
 * The pages that could possibly answer for this record.
 *
 * Exported because the viewer needs the same answer for a different question:
 * when a page finishes painting, is it worth re-running the plan? Deriving both
 * from one function is deliberate. When the view asked its own version of this
 * -- `record.page === page.number` -- the straddle page below was searched by
 * the planner but never re-triggered by a paint, so a change located on the
 * preceding page stayed unmarked until some unrelated event repainted. Two
 * copies of a rule that must agree is the drift this repo has been bitten by
 * three times; there is now one copy.
 */
export function candidatePages(record) {
    if (!record || !Number.isInteger(record.page) || record.page < 1) return [];
    // A deletion has no text to find, so the page before it is not insurance --
    // it is just a page we would have no reason to mark.
    if (!record.locatable || !record.text) return [record.page];
    return [record.page - 1, record.page].filter((number) => number >= 1);
}

/**
 * Decides what the overlay should mark.
 *
 * `pages` is `[{ number, items }]` in page order, `items` being the extracted
 * text content or `null`/`undefined` when the page has not painted. Returns
 * `[{ number, found }]`, where `found` is the locate result for a page the text
 * was found on, or `null` for a page-level marker with no position behind it.
 */
export function planChangeMarks(record, pages) {
    if (!record || !Array.isArray(pages) || pages.length === 0) return [];

    const byNumber = new Map(pages.map((page) => [page.number, page]));
    const reported = byNumber.get(record.page) ?? null;

    // Nothing to find -- a deletion leaves no text behind. Only the reported page
    // can be marked, and only as a page-level marker.
    if (!record.locatable || !record.text) {
        return reported ? [{ number: reported.number, found: null }] : [];
    }

    const candidates = candidatePages(record)
        .map((number) => byNumber.get(number))
        .filter(Boolean);


    const marks = [];
    let searchable = true;
    for (const page of candidates) {
        if (!page.items) {
            searchable = false;
            continue;
        }
        const found = locateText(page.items, record.text);
        if (found.status === "located" || found.status === "partial") {
            marks.push({ number: page.number, found });
        }
    }

    if (marks.length > 0) return marks;
    if (searchable && reported) return [{ number: reported.number, found: null }];
    return [];
}
