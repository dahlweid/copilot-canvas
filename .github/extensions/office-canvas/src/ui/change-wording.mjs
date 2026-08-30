// Reader-facing wording for a change record.
//
// Separate from `pdf-view.mjs` because the banner and the on-page badge must say
// the same thing about one edit, and one phrase table is how that is guaranteed
// rather than hoped for. A second reason held when this was written and no
// longer does: `pdf-view.mjs` imports pdf.js from an absolute `/vendor/` URL and
// could not be imported under Node. #76's harness resolves that specifier, so
// the first reason is now the only one.
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

/**
 * The whole banner line, and whether there is anywhere to jump to.
 *
 * `page` is null whenever the edit result carried no usable page number. The
 * first version of this interpolated it regardless and put **"page null"** in
 * front of the reader, beside a "Show me" button that scrolled nowhere. Both
 * halves came from assuming a page is always known.
 *
 * The wording says "page not reported" and not "Word could not report the page".
 * All the record preserves is that no usable number arrived; whether Word
 * declined to give one or the editor never asked is not something this code
 * distinguished, and naming either would be asserting a cause we did not
 * measure.
 *
 * `jumpable` is false in that case, so the caller can hide the control rather
 * than offer an action that cannot happen.
 */
export function describeChangeBanner(record) {
    const phrase = describeChange(record);
    const page = Number.isInteger(record?.page) && record.page >= 1 ? record.page : null;

    if (page === null) {
        // Nothing is marked either -- with no page, the plan has no candidate --
        // so this must not promise a marker the reader could go looking for.
        return { text: `${phrase} — page not reported`, jumpable: false };
    }

    const text = record?.locatable
        ? `${phrase} — page ${page}`
        : `${phrase} — page ${page}, marked but not highlighted`;
    return { text, jumpable: true };
}
