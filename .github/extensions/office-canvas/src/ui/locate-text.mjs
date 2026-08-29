// Finds a paragraph's text among the text items pdf.js extracts from a page.
//
// This is how the change overlay decides *where* to draw. It deliberately does
// no geometry: it answers "which text items cover this string", and the viewer
// -- which is the only thing holding a viewport -- turns item indices into
// rectangles. Keeping the seam here is what lets the hard part run on a hosted
// runner with plain objects instead of a rendered page.
//
// ## Why the join is on text and not on an address
//
// An address is a hash of heading path, text and occurrence. It has no page and
// no geometry, and it is a coordinate valid only inside one read-then-edit cycle
// (ADR 0006). Resolving one at paint time would mean holding it across an edit,
// which is the one thing the address model forbids. So the overlay is handed the
// *text* the operation left behind, read back from the document after the edit,
// and the PDF it searches was exported from that same post-edit file. Both sides
// come from one document state, and the join key is structural.
//
// ## The two things that break naive matching
//
//   * **A line break is not a space.** Word's export emits one text item per
//     line, and consecutive items carry no separator, so concatenating `str`
//     directly turns "the quick" / "brown fox" into "the quickbrown fox".
//     pdf.js marks the break with `hasEOL`, which is treated as whitespace here.
//   * **Word splits a line into runs for invisible reasons** -- a spell-check
//     boundary, a stray formatting mark -- so a match routinely spans several
//     items and can start and end mid-item.

/**
 * Collapses whitespace exactly as `structure-map.mjs`'s `normalizeText` does.
 *
 * The two must agree: one side normalizes the text stored in the change record,
 * the other normalizes the page. A difference between them shows up as an
 * overlay that silently never appears.
 */
export function normalizeText(value) {
    return String(value ?? "")
        .replace(/\s+/gu, " ")
        .trim();
}

/**
 * Flattens pdf.js text items into one normalized string, keeping a map from
 * each character back to the item that produced it.
 *
 * `owners[i]` is the index into `items` of the item that contributed
 * `text[i]`. A space synthesised from `hasEOL` is attributed to the item that
 * ended the line, so a match reaching that space still highlights that item.
 */
export function buildPageText(items) {
    let text = "";
    const owners = [];

    const push = (char, owner) => {
        // Collapse runs of whitespace, and never open with one.
        if (char === " " && (text.length === 0 || text.endsWith(" "))) return;
        text += char;
        owners.push(owner);
    };

    for (const [index, item] of (items ?? []).entries()) {
        for (const char of String(item?.str ?? "")) {
            push(/\s/u.test(char) ? " " : char, index);
        }
        // pdf.js signals a line break on the item that ends the line, including
        // on zero-width marker items whose `str` is empty.
        if (item?.hasEOL) push(" ", index);
    }

    // Trim the trailing space a final `hasEOL` leaves behind. Leading space is
    // impossible by construction above, so only this end needs it.
    while (text.endsWith(" ")) {
        text = text.slice(0, -1);
        owners.pop();
    }
    return { text, owners };
}

/** Every start offset at which `needle` occurs in `haystack`. */
export function findOccurrences(haystack, needle) {
    const found = [];
    if (!needle) return found;
    let from = 0;
    for (;;) {
        const at = haystack.indexOf(needle, from);
        if (at === -1) return found;
        found.push(at);
        from = at + 1; // overlapping occurrences still count as ambiguity
    }
}

/**
 * The shortest match this will accept when a paragraph runs across a page break.
 *
 * A partial match is only ever anchored to an edge of the page, so it cannot
 * drift into the middle of unrelated text -- but a very short fragment ("The")
 * would still land on the wrong line. This is a floor on confidence, not a
 * performance bound.
 */
export const MIN_PARTIAL_CHARS = 24;

/**
 * Locates `target` among a page's text items.
 *
 * Returns one of:
 *
 * | `status` | meaning | overlay |
 * | --- | --- | --- |
 * | `located` | exactly one match | quads over `range` |
 * | `partial` | the paragraph crosses a page edge and this page holds one end | quads over `range` |
 * | `ambiguous` | the same text occurs more than once on this page | page marker only |
 * | `not_found` | not on this page | page marker only |
 * | `empty` | nothing to look for | page marker only |
 *
 * `ambiguous` is a deliberate refusal rather than a first-match. The structure
 * map's `occurrence` counts within a heading, not within a page, so there is no
 * sound way to pick between two identical strings here -- and drawing a box on
 * the wrong one asserts a position that was never determined.
 */
export function locateText(items, target, { minPartialChars = MIN_PARTIAL_CHARS } = {}) {
    const needle = normalizeText(target);
    if (!needle) return { status: "empty", range: null };

    const { text, owners } = buildPageText(items);
    if (!text) return { status: "not_found", range: null };

    const rangeFor = (start, length) => ({
        startItem: owners[start],
        endItem: owners[start + length - 1],
    });

    const hits = findOccurrences(text, needle);
    if (hits.length === 1) return { status: "located", matched: needle, range: rangeFor(hits[0], needle.length) };
    if (hits.length > 1) return { status: "ambiguous", occurrences: hits.length, range: null };

    // Not found whole. The paragraph may straddle a page break, in which case
    // this page holds either its head (a prefix of the target, at the very end
    // of the page) or its tail (a suffix, at the very start).
    const limit = Math.min(needle.length - 1, text.length);
    for (let length = limit; length >= minPartialChars; length--) {
        // Tail of the target sitting at the start of the page.
        if (text.startsWith(needle.slice(needle.length - length))) {
            return { status: "partial", where: "start", matched: needle.slice(needle.length - length), range: rangeFor(0, length) };
        }
        // Head of the target sitting at the end of the page.
        if (text.endsWith(needle.slice(0, length))) {
            return {
                status: "partial",
                where: "end",
                matched: needle.slice(0, length),
                range: rangeFor(text.length - length, length),
            };
        }
    }
    return { status: "not_found", range: null };
}
