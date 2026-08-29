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

## Accessibility

An earlier draft claimed parity with the native viewer on tagged PDFs. That
claim was withdrawn because nothing had been tested, and it is **not**
reinstated here: parity was never measured, by anyone.

It also cannot be measured this way. The native plugin renders in its own
process and exposes no DOM to the embedding page — the very fact this decision
rests on — so there is nothing to enumerate on that side and no like-for-like
comparison to run. Establishing parity would mean driving a real screen reader
against both viewers, which is outside what a probe in this repo can do. So what
follows is an absolute statement of what pdf.js gives us, not a comparison.

Measured by `spikes/pdfjs/probes/probe-accessibility.mjs` against a five-page
document exported through our own pipeline (pdf.js 6.2.108):

| question | answer |
| --- | --- |
| Is Word's export tagged? | Yes — `/MarkInfo << /Marked true >>`, `/StructTreeRoot` and `/Lang` all present |
| Does pdf.js expose the structure? | Yes — `getStructTree()` returns `Root › Document › H1/H2/P`, 19 nodes |
| Do headings survive as headings? | Yes — `H1` and `H2` roles, matching the document's outline |
| Does the structure reference the text? | Yes — 17 marked-content refs in the tree |
| Is the text layer in reading order? | Yes — DOM order matched extraction order exactly, 31 spans |

Word is already exporting with `DocStructureTags = $true` (see the export call
in `word-host.ps1`), which is why the tagging is there to find.

### The gap, named

`spansWithMarkedContentIds: 0`. The structure tree carries 17 refs into the page
content, but the spans our `TextLayer` mounts carry none, so nothing connects
them: a screen reader walking our DOM sees 31 correctly-ordered `<span>`s with
`role="presentation"` and no headings. The structure is *available* and we are
not *using* it. pdf.js's own viewer wires this up with a separate structure-tree
layer; we mount only the text layer.

That is a real, bounded piece of work rather than an unknown, and it is not in
this layer's scope. What must not happen is for the measurements above to be
read as parity — the roles exist in the file, not in the accessibility tree the
user's software actually sees.

Not measured by anyone: actual assistive-technology output, from either viewer.
