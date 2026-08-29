// Reader-facing wording for a change record.
//
// Separate from `pdf-view.mjs` for two reasons. The banner and the on-page badge
// must say the same thing about one edit, and `pdf-view.mjs` imports pdf.js from
// an absolute `/vendor/` URL, so it cannot be imported under Node and anything
// living inside it cannot be unit-tested.
//
// The editor's own `applied.description` is deliberately not used here. It is
// written for the agent -- "replace the text of p:8f957157e47d" -- and naming an
// address in front of a reader is both meaningless and the one thing the overlay
// is built to avoid. It was on screen once, which is how this module came to
// exist.

/** The phrase shown for each operation. Anything absent falls back generically. */
const PHRASES = {
    replace_text: "Text replaced",
    insert_paragraph_after: "Paragraph added",
    insert_paragraph_before: "Paragraph added",
    delete_paragraph: "Paragraph deleted",
    set_heading_level: "Heading level changed",
};

/** The wording used when an operation has no phrase of its own. */
export const GENERIC_PHRASE = "Changed";

/**
 * What to call the change, in words a reader can act on.
 *
 * A table rather than a switch so a test can ask which operations it covers; a
 * switch would let a newly added operation reach the generic phrase with nothing
 * able to notice.
 */
export function describeChange(record) {
    return PHRASES[record?.op] ?? GENERIC_PHRASE;
}

/** The operations this module has wording for. Exported for the drift check. */
export const DESCRIBED_OPS = Object.keys(PHRASES);
