# Render through the Office print pipeline, not a live window

We investigated streaming a live, invisible Word window into the canvas by
capturing it with `PrintWindow` — including running it on a `CreateDesktop`
non-input desktop so it never appears in the taskbar. It works: COM crosses the
desktop boundary, and capture succeeds. We rejected it anyway, because a
single-page `ExportAsFixedFormat` measured **168 ms** and was flat regardless of
page position, which removes the latency argument that motivated streaming at
all.

The deciding factor is fidelity, and it is structural rather than a list of
bugs. Capture shows Word's *editing* view: dark page in dark mode, spelling
squiggles, revision bars, the comment margin, licensing nags. Fixed-format
export uses the *print* pipeline, which has no view state to leak. v1's
immunity to all of that is a property of the pipeline, not something we
maintain.

Two further findings sealed it. `PrintWindow` on a hidden desktop returns pure
black except with the undocumented `PW_RENDERFULLCONTENT` flag, so the whole
approach rests on DWM compositing a desktop that is never displayed — with no
fallback when DWM parks its surfaces at lock, RDP transition, sleep, or GPU
reset. And the input channel was unreliable: 2 of 5 `TypeText` calls produced no
confirming frame.

## Consequences

The hidden-desktop apparatus is dropped entirely, and with it the Job Object
framing and the EDR security review it would have required. Note that v1
already runs Word with `Visible = false`, so invisibility was never the thing
the desktop bought us — it only ever existed to give `PrintWindow` a renderable
window.

Expect this to be re-proposed, because "just show the real Word" is the obvious
idea. The counter-argument is the 168 ms measurement plus the editing-view
category error, not implementation difficulty.
