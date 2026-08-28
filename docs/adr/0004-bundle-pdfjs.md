# Ship pdf.js instead of using the host's native PDF viewer

v1 renders the exported PDF by pointing an iframe at it and letting the
browser's built-in PDF plugin display it. We are replacing that with a bundled
pdf.js. Shipping a PDF renderer when the host already has one looks like
gratuitous work, so the reason needs recording.

Once the agent is the editor, the canvas has to show *what the agent changed* —
a highlight on an edited paragraph, a marker on an inserted table. Nothing can
be drawn over the content of a native PDF plugin from the embedding page. The
plugin renders in its own process and exposes no DOM, no coordinates, and no
scroll position. This is structurally impossible rather than merely awkward, and
the symptom is already visible in v1: changing page requires reassigning the
iframe `src`, because there is no API to ask the plugin to scroll.

pdf.js renders to a canvas element we own, which makes overlays, page-anchored
annotations, scroll synchronisation, and text-layer selection ordinary work.

## Consequences

Word embeds fonts into its exported PDFs, so pdf.js's standard font pack is
probably unnecessary — worth confirming, because those files are binary and
gist-based sharing refuses binary content.

Accessibility needs a deliberate check rather than an assumption. An earlier
draft claimed parity with the native viewer on tagged PDFs; that claim was
withdrawn because nothing had been tested.
