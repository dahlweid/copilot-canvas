# One extension hosting several canvases, not one extension per application

Word, Excel, and PowerPoint ship as a single extension that declares a canvas
per application (`canvases?: Canvas[]` accepts an array), rather than three
independent extensions.

The motivation is a runtime constraint, not tidiness. Extensions may not carry
a `package.json` or `node_modules`, and `install_extension` copies exactly one
folder which must run as committed. There is therefore no way for three
extension folders to share a library: three extensions would mean three
vendored copies of pdf.js, three copies of the render cache, server, and
watcher, three installs for the user, and three sets of drift. Gist sharing
caps at roughly 5 MB and refuses binary files, so the duplication is not merely
wasteful — it approaches a hard limit.

Separate canvas ids within that one extension are kept because the agent picks
a canvas from its description, and "Word document" is a clearer signal than a
generic canvas with a mode parameter. Their input schemas also legitimately
differ: a page number for Word, a slide for PowerPoint, a sheet and range for
Excel.

## Consequences

The shared pdf.js rendering surface is a good fit for Word and PowerPoint,
where pagination is inherent to the format. It is a poor fit for Excel, where
pages are invented by print setup and a PDF cannot show a formula at all.
**Excel is therefore out of scope for now**: we ship Word and PowerPoint, and
revisit Excel once we know what its display surface should be, since it is
almost certainly a grid rather than a page.

That deferral costs nothing precisely because the canvases are already
separate. Excel can arrive later with a different view, a different input
schema, and a different verification signal, without disturbing the two that
render as pages. Only the packaging is shared, and that is the part the runtime
forces us to share.

## Amendment: PowerPoint is not automated the way Word is

Added after `spikes/powerpoint` measured PowerPoint for the first time; until
then every platform number in this repo came from Word. See
[`spikes/powerpoint/FINDINGS.md`](../../spikes/powerpoint/FINDINGS.md).

**The rendering premise is confirmed, and more strongly than for Word.**
`ExportAsFixedFormat` on a `.pptx` produces exactly one PDF page per slide, with
`/MediaBox` matching `PageSetup.SlideWidth/Height` to 0.000 pt, at 94–110 ms per
slide against Word's 168 ms per page. There is no rendering reason to split the
extension, and PowerPoint should not follow Excel out of scope.

**The shared-host premise does not hold as stated.** "The hidden-COM host
pattern" is listed among the substance Word and PowerPoint share, but the two
applications do not expose the same process model:

- Word is multi-instance. Each `New-Object -ComObject Word.Application` starts
  its own `WINWORD.EXE`.
- PowerPoint registers **one automation server per session**. A second
  `New-Object` attaches to the PowerPoint already running — including the user's
  — and launching `POWERPNT.EXE` on the current desktop hands off and exits.
  Verified with Word as the control, so this is a property of the application
  and not of the method.

A private PowerPoint is still obtainable, by the separate-desktop launch and
`OBJID_NATIVEOM` bind that `spikes/isolation` already uses for Word, binding on
`mdiClass` rather than Word's `_WwG`. But it is a different mechanism with a
different cost: **~12 s to a bound instance, against ~1 s for a COM attach.**

Two consequences follow for the host, both correctness rather than hygiene:

- Never call `Quit()` on a COM-attached PowerPoint and never kill its pid; both
  destroy the user's session. `Quit()` does not terminate `POWERPNT.EXE` anyway
  (15/15 cycles), and a killed PowerPoint makes the next launch show a modal
  safe-mode prompt.
- The ~12 s bind cost rules out opening an isolated instance lazily on first
  interaction. Start it ahead of time, or serve the first render from the
  OOXML/PDF path, which needs no PowerPoint at all.

The transient-lock model in [ADR 0005](0005-transient-lock-write-model.md)
transfers **unchanged**: a genuinely separate PowerPoint opening a held deck
hangs silently and needs an external kill, exactly as Word does.

