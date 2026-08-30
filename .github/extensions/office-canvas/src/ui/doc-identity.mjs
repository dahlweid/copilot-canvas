// How the header bar identifies the open document, once the absolute-path row
// is gone (#71).
//
// This lives in its own module because the decision is worth executing on its
// own. It was written when `app.js` had no reachable test at all -- it could not
// be imported under Node, dying at module resolution of `pdf-view.mjs`'s
// absolute `/vendor/pdf.min.mjs` specifier before any DOM access. #76 removed
// that: `test/unit/ui-harness.mjs` resolves the specifier and supplies a DOM, so
// `app.test.mjs` now drives `app.js` directly. The split still earns its place
// -- the rules here are asserted against no DOM at all, and `app.js` is asserted
// against a stand-in for one -- but "cannot be imported" is no longer the
// reason, and leaving that claim here would be a stale measurement.

/** The separator this path uses, or null when it has none. */
function separatorOf(path) {
    // Windows paths are what the server produces, but a document opened from a
    // POSIX-shaped fixture path must not silently lose its folder. Ask the
    // string rather than the platform: this module runs in a browser, where
    // `node:path` is not available and `process.platform` says nothing about
    // where the document lives.
    const back = path.lastIndexOf("\\");
    const forward = path.lastIndexOf("/");
    if (back < 0 && forward < 0) return null;
    return back > forward ? "\\" : "/";
}

/**
 * The folder containing `path`, or null when the path names no folder.
 *
 * A bare filename has no folder, and neither does a root-relative name like
 * `\report.docx`: the answer there would be the empty string, which would read
 * on screen as a document that lives nowhere.
 */
export function folderOf(path) {
    if (typeof path !== "string") return null;
    const sep = separatorOf(path);
    if (!sep) return null;
    const folder = path.slice(0, path.lastIndexOf(sep));
    return folder.length > 0 ? folder : null;
}

/**
 * What the bar shows and says about a document.
 *
 * Returns the tooltip for the name, the accessible name and tooltip for the
 * copy-path button, and whether that button has anything to copy.
 *
 * The tooltip is deliberately *not* the only place the path survives. A `title`
 * is not keyboard-reachable and screen readers treat it inconsistently, so the
 * path is also the accessible name of a real focusable control -- which is
 * reachable by tab, announced on focus, and does something useful when pressed.
 */
export function describeDocument(doc) {
    const name = typeof doc?.name === "string" ? doc.name : "";
    const path = typeof doc?.path === "string" && doc.path.trim() ? doc.path : null;

    // Nothing to say beyond the name when the path *is* the name: repeating it
    // in a tooltip is exactly the emptiness that cost the old row its place,
    // where `textContent` and `title` were set to the same string.
    const pathAddsSomething = Boolean(path) && path !== name;

    return {
        /** Tooltip on the document name: the full path, or nothing. */
        nameTitle: pathAddsSomething ? path : "",
        /** The folder, for a caller that wants to show location inline. */
        folder: pathAddsSomething ? folderOf(path) : null,
        /** Whether the copy control has a path to copy. */
        canCopy: Boolean(path),
        /**
         * The copy button's accessible name. It carries the full path, which is
         * what makes the path reachable without a pointer -- a keyboard user
         * tabs to this control and hears it.
         */
        copyLabel: path ? `Copy full path: ${path}` : "Copy full path",
        copyTitle: path ? path : "Copy full path",
    };
}

/**
 * What to say after a copy attempt.
 *
 * Split out because the failure is real and not hypothetical: `writeText`
 * rejects when the document is not focused, and `navigator.clipboard` is absent
 * outside a secure context. Neither is worth an error banner -- the path is
 * still on screen in the tooltip and in this control's own name -- so a failure
 * says what did not happen and offers no remedy the user cannot see.
 */
export function copyOutcome(error) {
    if (!error) return { text: "Full path copied.", error: false };
    return { text: "Could not copy the path. It is in this button's tooltip.", error: true };
}
