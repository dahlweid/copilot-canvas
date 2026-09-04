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
        // One owner per UTF-16 code unit, not per character. The loop below
        // iterates code points, so an astral character -- an emoji, CJK ext-B, a
        // mathematical alphanumeric -- is one `char` here but *two* units in
        // `text`. `owners` is indexed by offsets that come from `indexOf`, which
        // counts units, so pushing once per character would shift every later
        // entry left by one per astral character and silently box the wrong item.
        for (let i = 0; i < char.length; i += 1) owners.push(owner);
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
 *
 * `span`, when given, is the part of `target` that actually changed (#166). It
 * is a *refinement of the match*, never a second search: it is looked for only
 * within the characters `target` matched, so it cannot pull the box onto
 * unrelated text elsewhere on the page, and anything less than exactly one
 * occurrence inside that window leaves the full match standing. `narrowed` says
 * which happened, so a caller never has to infer it by comparing ranges.
 */
export function locateText(items, target, { minPartialChars = MIN_PARTIAL_CHARS, span = null } = {}) {
    const needle = normalizeText(target);
    if (!needle) return { status: "empty", range: null };

    const { text, owners } = buildPageText(items);
    if (!text) return { status: "not_found", range: null };

    const rangeFor = (start, length) => ({
        startItem: owners[start],
        endItem: owners[start + length - 1],
    });

    // The narrowed range, or the given one unchanged. `matched` keeps naming the
    // text the *match* was made on, because that is what the status is about;
    // narrowing only moves the box.
    const narrow = (found, start, length) => {
        const wanted = normalizeText(span);
        if (!wanted || wanted === found.matched) return found;
        const window = text.slice(start, start + length);
        const hits = findOccurrences(window, wanted);
        if (hits.length !== 1) return found;
        return { ...found, narrowed: true, range: rangeFor(start + hits[0], wanted.length) };
    };

    const hits = findOccurrences(text, needle);
    if (hits.length === 1) {
        return narrow(
            { status: "located", matched: needle, narrowed: false, range: rangeFor(hits[0], needle.length) },
            hits[0],
            needle.length,
        );
    }
    if (hits.length > 1) return { status: "ambiguous", occurrences: hits.length, range: null };

    // Not found whole. The paragraph may straddle a page break, in which case
    // this page holds either its head (a prefix of the target, at the very end
    // of the page) or its tail (a suffix, at the very start).
    const limit = Math.min(needle.length - 1, text.length);
    for (let length = limit; length >= minPartialChars; length--) {
        // Tail of the target sitting at the start of the page.
        if (text.startsWith(needle.slice(needle.length - length))) {
            return narrow(
                {
                    status: "partial",
                    where: "start",
                    matched: needle.slice(needle.length - length),
                    narrowed: false,
                    range: rangeFor(0, length),
                },
                0,
                length,
            );
        }
        // Head of the target sitting at the end of the page.
        if (text.endsWith(needle.slice(0, length))) {
            return narrow(
                {
                    status: "partial",
                    where: "end",
                    matched: needle.slice(0, length),
                    narrowed: false,
                    range: rangeFor(text.length - length, length),
                },
                text.length - length,
                length,
            );
        }
    }
    return { status: "not_found", range: null };
}
