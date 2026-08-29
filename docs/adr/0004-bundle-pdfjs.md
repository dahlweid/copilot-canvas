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

It also cannot be measured this way, and the reason matters more than the
verdict. The tempting phrasing — "the plugin exposes no DOM, so there is nothing
to compare" — is an outcome-level reason and it is subtly wrong: a screen reader
does not read the DOM either. It reads an accessibility tree. pdf.js publishes
into that tree *via* the DOM the browser builds it from; the native plugin
publishes into it directly, from its own process, through its own provider. Both
sides therefore have something a screen reader can consume. What only one side
has is a channel a probe running **inside the page** can enumerate. The
asymmetry is in our instrument, not in the two viewers, and stating it the other
way would credit pdf.js with an advantage this repo has not shown it to have.

So, to the standard this repo now holds negatives to — what a successful
comparison would have looked like, precisely enough that its absence is a
finding and not a shrug:

- **Instrument.** A real screen reader (NVDA's speech-viewer log is the cheapest
  capture) driven over a real desktop, not a probe in the page.
- **Arms.** The same tagged PDF, exported through our own pipeline, opened once
  in the canvas and once in the native plugin.
- **Measurement.** The announced sequence for a full read-through of one page:
  the order items are spoken in, and whether each heading is announced *as* a
  heading at its level.
- **Verdict.** Parity iff the two sequences agree on reading order and on
  heading level. A difference in either direction is the finding; the direction
  decides whether the gap below is worth fixing or whether we have understated
  pdf.js.
- **Control.** The instrument has to be shown able to produce both answers —
  a document with deliberately broken tags must come back as a *mismatch*, or a
  match means nothing. This is the arm that would be easiest to skip and would
  invalidate the whole result.

None of that has been run, by anyone. Everything below is an absolute statement
of what pdf.js gives us, and no part of it is a comparison.

Measured by `spikes/pdfjs/probes/probe-accessibility.mjs` against a five-page
document exported through our own pipeline (pdf.js 6.2.108):

| question | answer |
| --- | --- |
| Is Word's export tagged? | Yes — `/MarkInfo << /Marked true >>`, `/StructTreeRoot` and `/Lang` all present |
| Does pdf.js expose the structure? | Yes — `getStructTree()` returns `Root › Document › H1/H2/P`, 19 nodes |
| Do headings survive as headings? | Yes — `H1` and `H2` roles, matching the document's outline |
| Does the structure reference the text? | Yes — 17 marked-content refs in the tree |
| Is the text layer in reading order? | Yes — DOM order matched extraction order exactly, 31 spans |

The tagging is there to find because the export asks for it: `DocStructureTags`
on `ExportAsFixedFormat`. That flag is not written by name anywhere in the host.
The call passes **fourteen positional arguments** and the parameter names live in
a comment above it, so the flag is identified by its index and by nothing else —
an argument inserted or removed ahead of it re-points it silently, leaves the
comment still reading correctly, and produces no symptom other than an untagged
PDF.

So this ADR does not restate the value. `test/unit/export-tagging.test.mjs`
derives it — names from the comment, index from the names, value from the call —
and separately asserts the two lists are the same length, which is the only
assertion that catches the shift. Three mutants cover it, including the shift
case, which the value check alone passes.

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

Not measured, and **not measurable from inside the page**: actual
assistive-technology output, from either viewer. The instrument for it is the
screen reader named above, driven over a real desktop; nothing running in the
canvas can stand in for it, because what a probe in the page can enumerate is the
DOM and what a screen reader consumes is the accessibility tree. That is why this
gap is stated here rather than closed — the measurement is absent, and the means
of taking it is absent too.
