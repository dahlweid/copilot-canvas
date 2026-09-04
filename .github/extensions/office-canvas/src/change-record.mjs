// The change record: what the overlay is told about the operation that just ran.
//
// ## It is a statement about a moment that has passed
//
// The obvious design -- remember the address the agent edited, resolve it when
// the page paints -- is the one thing ADR 0006 forbids. An address is a
// coordinate, not a handle: deleting one of several identically-worded
// paragraphs renumbers its successors, and renaming a heading moves every
// address beneath it. An address held across an edit stays *valid* while
// pointing at different content, which is worse than one that breaks.
//
// So no address is carried. `edit_document` re-reads the document after saving
// it, and that read mints fresh addresses and returns the paragraph it touched.
// The record takes the **text** from that post-edit read and the page Word
// reported. The PDF the overlay searches is exported from the same post-edit
// file, so both sides describe one document state and the join key -- page and
// text -- is structural. This is the same rule that forbids joining a COM-side
// name to a file-side one: never join on something that can move underneath you.
//
// ## Not everything can be shown
//
// `delete_paragraph` leaves nothing behind to find. There is no honest box to
// draw, and anchoring to the following paragraph would assert a position the
// code never determined -- the paragraph that now follows the gap need not be
// the one that followed the deleted text. Those cases carry `locatable: false`.
//
// A deletion also carries **no page**, so it gets no marker either, only the
// text notice. That is not the viewer declining to draw one: the host returns
// `0` for a deletion's touched index (`word-host.ps1`, the `delete_paragraph`
// arm), the page is only read `if ($touched -gt 0 ...)`, and `document-editor`
// maps the resulting `0` to `null`. So the page is never determined, and
// `candidatePages` refuses to guess one. The page-level marker is reachable for
// the *other* unlocatable case -- an op whose paragraph has no findable text but
// whose page Word did report -- which is why both branches exist.
//
// Capturing the page before the range is deleted would make the marker
// available, but that is a host-side change and is tracked separately rather
// than asserted here.
//
// ## The record does not carry a revision token
//
// It would be the obvious thing to stamp it with, and it is the wrong
// discriminator. A revision token is a hash of the file's *content*, so a
// regeneration that reproduces the same bytes leaves it unchanged; the render
// key is `hash(identity|mtimeMs|size)` and moves on any write at all. The
// overlay is a claim about a rendered image, and it must die whenever that image
// is replaced -- including by a rewrite that happens to be byte-identical, which
// still produces a new render. `ViewerInstance` stamps the record with the
// render key for exactly that reason, and carrying a token beside it would be a
// second rule that could disagree with the one actually enforced.

// ## The record does not carry the editor's description
//
// `edit_document` returns one -- "replace the text of p:8f957157e47d" -- and it
// was on screen for a while. It is written for the agent, it names an address,
// and an address in front of a reader is both meaningless and precisely what
// this design exists to avoid. The viewer words the change itself from `op`
// (`src/ui/change-wording.mjs`), so the field is not carried at all rather than
// carried and trusted not to be shown.

// ## The span narrows the box; it never becomes the join key
//
// The record also carries `span` -- the words that changed, when they can be
// derived -- because a whole-paragraph highlight tells a reader nothing when a
// paragraph is ten lines and three words moved (#166). What it must not do is
// *replace* `text` as the thing the overlay searches for, and the reason is
// measured in the locator's own contract rather than guessed:
//
//   * `locateText` refuses a needle that occurs twice on a page (`ambiguous`),
//     and `planChangeMarks` then draws a page-level marker with no box at all.
//     A whole paragraph is nearly unique; three words are not. So a span used
//     as the key would sometimes produce *less* than the paragraph it replaced.
//   * `candidatePages` searches the page before the reported one, for the
//     paragraph that straddles a break. A short span can match there, uniquely,
//     in text that has nothing to do with the edit -- a box asserting a
//     position nothing determined.
//
// So the paragraph is located exactly as before, and the span is searched only
// inside the character window that match occupies (`locateText`'s `span`
// option). A wrong-place box is then impossible, ambiguity is bounded to "these
// words occur twice inside this one paragraph", and every way of failing to
// narrow lands on the whole-paragraph box that was there before.

import { normalizeText } from "./ui/locate-text.mjs";

/** Operations that leave no text on the page to anchor to. */
const UNLOCATABLE_OPS = new Set(["delete_paragraph"]);

/**
 * The words that changed, or `null` when there is no honest narrower answer.
 *
 * Both arguments are already-normalized paragraph text -- the same
 * `normalizeText` the page join runs on, so a span derived here is a substring
 * of the string the overlay searches for. Nothing here touches a Word range, so
 * the `\r` + `chr(7)` / `End - Start` trap that character arithmetic on a COM
 * range hits is out of reach by construction.
 *
 * `null` is returned wherever a narrower box would assert something this
 * function did not establish. Each case is a branch rather than a fallthrough,
 * because "there is nothing to narrow to" and "the narrowing came out empty"
 * are different facts that happen to want the same behaviour:
 *
 *   * either side empty -- nothing to diff against;
 *   * the two are equal -- the edit changed nothing this normalization can see;
 *   * the replacement span is empty -- a pure deletion inside the paragraph
 *     leaves no text on the page to draw a box around;
 *   * the span covers the whole paragraph -- no prefix and no suffix survived,
 *     so narrowing would carry a second copy of the text already in the record.
 *
 * The span is widened outward to whitespace boundaries. A mid-word fragment
 * ("ello Markus") is both meaningless to a reader and *more* likely to occur
 * twice than the whole word is, and occurring twice is what makes it
 * unlocatable -- so widening helps on both counts and costs a line of code.
 */
export function changedSpan(previous, next) {
    if (!previous || !next || previous === next) return null;

    const max = Math.min(previous.length, next.length);
    let prefix = 0;
    while (prefix < max && previous[prefix] === next[prefix]) prefix += 1;
    // The suffix is bounded so that the span it leaves in `next` cannot be
    // negative, and so the scan cannot walk off the front of `previous`. It is
    // deliberately *not* bounded to keep prefix and suffix from overlapping in
    // `previous`, which is what `Set-ParagraphText` (`word-host.ps1`) does --
    // and the difference is the issue's own case rather than an exotic one.
    // Inserting "Hallo Markus." after "... modular. " makes ". " both the tail
    // of the prefix and part of the suffix, so a non-overlap bound cuts the
    // suffix short and drags the next word into the highlight.
    //
    // The two are allowed to differ because they answer different questions.
    // The host must produce a *decomposition* it can assign a Word range from,
    // so its prefix and suffix must not claim the same characters twice. This
    // only has to name a span of the post-edit text, and
    // prefix + span + suffix reconstructs `next` exactly either way. Where they
    // disagree the consequence is a box one word wider, never wrong text.
    const maxSuffix = Math.min(previous.length, next.length - prefix);
    let suffix = 0;
    while (suffix < maxSuffix && previous[previous.length - 1 - suffix] === next[next.length - 1 - suffix]) {
        suffix += 1;
    }

    let start = prefix;
    let end = next.length - suffix;
    if (start >= end) return null; // pure deletion: nothing was put in its place

    while (start > 0 && next[start - 1] !== " ") start -= 1;
    while (end < next.length && next[end] !== " ") end += 1;

    const span = next.slice(start, end);
    return span.length > 0 && span !== next ? span : null;
}

/**
 * Builds a change record from an `edit_document` result.
 *
 * Returns `null` when there is nothing worth showing at all, rather than a
 * record the viewer would have to know to ignore.
 */
export function changeRecordFrom(result, { now = () => new Date().toISOString() } = {}) {
    if (!result?.applied?.op) return null;

    const op = result.applied.op;
    const page = Number.isFinite(result.page) && result.page > 0 ? Math.floor(result.page) : null;
    const text = normalizeText(result.paragraph?.text ?? "");

    // Every reason a record cannot be located is decided here, once. The viewer
    // is handed a verdict, not the inputs to re-derive one, so the two cannot
    // come to different conclusions about the same edit.
    const locatable = !UNLOCATABLE_OPS.has(op) && text.length > 0;
    const span = locatable ? changedSpan(normalizeText(result.previousText ?? ""), text) : null;

    return {
        op,
        page,
        text: locatable ? text : null,
        span,
        locatable,
        at: now(),
    };
}
