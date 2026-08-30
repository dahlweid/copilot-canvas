// Swapping the drawn glyph for the real Word mark, when this machine has one.
//
// The mark is fetched at runtime from `/api/word-icon`, which extracts it from
// the user's own installed Word (#68). Nothing about it is committed here.
//
// The ordering is the whole design: what ships in the markup is the fallback,
// and the mark replaces it only after it has actually loaded. A machine with no
// Word therefore changes nothing, shows nothing broken and says nothing -- which
// is the correct amount of noise for a missing decoration.

/**
 * Wire an `<img>` to take over from a fallback element once it loads.
 *
 * `fallback` is optional: at the document name there was no glyph before, so
 * there is nothing to hide and nothing to restore -- the image simply stays
 * hidden. Returns the element that ends up visible in the failure case, which is
 * what a caller asserts on.
 */
export function showWordMark({ img, fallback = null, src = "/api/word-icon" }) {
    // Set before the request, not after: an `<img>` with a `src` that 404s
    // renders the browser's broken-image affordance, and a panel showing a
    // broken image is worse than one showing the glyph it already had.
    img.hidden = true;
    if (fallback) fallback.hidden = false;

    img.addEventListener(
        "load",
        () => {
            img.hidden = false;
            if (fallback) fallback.hidden = true;
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
            if (fallback) fallback.hidden = false;
        },
        { once: true },
    );

    img.src = src;
    return fallback ?? img;
}
