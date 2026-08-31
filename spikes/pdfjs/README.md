# pdf.js probes

Evidence for issue #9: *Render pages with pdf.js instead of the app's PDF
viewer.*

**Read [`FINDINGS.md`](FINDINGS.md) for the answers and the numbers.** This file
just says how to run the probes.

These were run *before* #9 was written, so that layer inherits measurements
rather than reasoning. They inverted both halves of the prediction the
restructure plan had made about pdf.js (see [`FINDINGS.md`](FINDINGS.md) §1), and
corrected three claims in the issue itself.

## Running them

```powershell
cd spikes\pdfjs
node .\probes\probe-pdf-fonts.mjs
node .\probes\probe-worker-split.mjs
```

Each probe is independent and re-runnable. They need `pdfjs-dist` resolvable and
a Word-exported PDF; both are obtained per run rather than committed, since C2
forbids a `package.json` in the extension and no binary belongs in this repo.

| Probe | Question it answers |
| --- | --- |
| `probe-pdf-fonts.mjs` | Does a Word PDF embed everything, or does it need `standard_fonts/` and `cmaps/`? |
| `probe-font-embedding.mjs` | Does any installed face forbid embedding, and would that break the answer above? |
| `probe-worker-split.mjs` | Can the oversized worker be split, served and actually rendered from? |
| `probe-range-requests.mjs` | What Range requests does pdf.js really issue against our server? |

`probe-worker-split.mjs` loads the reassembled worker **in the app's own
webview** through the browser canvas, not a headless browser — that is the
surface the canvas actually uses, and no Chrome is installed on this machine
anyway. It asserts a rendered page is non-blank by counting non-white pixels
rather than trusting that `render()` resolved.

## Why the split lives in vendoring, not packaging

`pdf.worker.min.mjs` is 1,262,398 bytes against a 1,000,000-byte per-file
install cap, so it cannot ship whole through any path `install_extension`
supports. The split must happen at vendoring time **with the parts committed**:
ADR 0003's constraint is that the extension folder runs exactly as committed, and
a repo-folder install never runs `tools/package-extension.mjs`. Splitting during
packaging would produce an installable artifact from a source tree that is not
itself installable — the divergence between packaging and validation that is
worse than no packaging at all.

Split by cap, never by a hard-coded part count: the file needs three parts, and a
size listing that rounds to "1.20 MB" suggests two.

## Process safety

These probes start a local HTTP server and a browser canvas; they touch no Office
process and hold no document open. `probe-worker-split.mjs` serves on an
ephemeral port and shuts down in a `finally`, so a failed run does not leave a
listener behind. If a browser canvas is left pointing at a dead port, that is
cosmetic — the port is not reused.
