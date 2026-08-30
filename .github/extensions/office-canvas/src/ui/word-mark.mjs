// Showing the real Word mark, when this machine has one.
//
// The mark is fetched at runtime from `/api/word-icon`, which extracts it from
// the user's own installed Word (#68). Nothing about it is committed here.
//
// The ordering is the whole design: what ships in the markup is nothing at all,
// and the mark appears only after it has actually loaded. A machine with no Word
// therefore changes nothing, shows nothing broken and says nothing -- which is
// the correct amount of noise for a missing decoration.
//
// ## Why there is no `fallback` any more (#87)
//
// This took a `fallback` element to hide once the mark arrived, so the Open in
// Word button could swap a drawn glyph for the brand mark. That parameter is
// gone with its only caller, rather than kept for a future one, for two reasons.
//
// It made a button's appearance depend on the machine: brand mark where Word
// could be found, arrow where it could not. Measured, 136.34px wide against
// 115.34px.
//
// And the swap never worked. `hidden` is an IDL attribute of `HTMLElement`, and
// the fallback was an `<svg>` -- an `SVGElement`, which does not inherit it. So
// `fallback.hidden = true` set an ordinary JavaScript property, reflected to no
// content attribute, matched no `[hidden]` rule and hid nothing. Measured in
// headless Chromium by spikes/viewer-header/probes/probe-mark-alignment.mjs:
// after the assignment the `<svg>` still computes `display: inline` and still
// has a client rect, while an `<img>` given the same assignment computes
// `display: none` and has none.
//
// The tests that covered the fallback passed throughout, because they drove a
// stub object on which `hidden` is a plain data property. Keeping the parameter
// would keep that gap: a branch the product does not use, which the only test
// able to reach it cannot fail on.

/**
 * Reveal an `<img>` once it has actually loaded.
 *
 * There is no failure path to see, and that is deliberate: the document name
 * carried no decoration before #68, so a machine with no Word is a machine where
 * this does nothing.
 */
export function showWordMark({ img, src = "/api/word-icon" }) {
    // Set before the request, not after: an `<img>` with a `src` that 404s
    // renders the browser's broken-image affordance, and a panel showing a
    // broken image is worse than one showing no decoration at all.
    img.hidden = true;

    img.addEventListener(
        "load",
        () => {
            img.hidden = false;
        },
        { once: true },
    );

    img.addEventListener(
        "error",
        () => {
            // Explicit rather than "leave it as it was". A retry, a second call,
            // or a load that succeeds and is later replaced must all land in the
            // same state, and that state is the one the markup shipped.
            img.hidden = true;
        },
        { once: true },
    );

    img.src = src;
}
