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
// the one that followed the deleted text. Those cases carry `locatable: false`
// and the viewer shows a page-level marker instead.
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

import { normalizeText } from "./ui/locate-text.mjs";

/** Operations that leave no text on the page to anchor to. */
const UNLOCATABLE_OPS = new Set(["delete_paragraph"]);

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

    return {
        op,
        page,
        text: locatable ? text : null,
        locatable,
        at: now(),
    };
}
