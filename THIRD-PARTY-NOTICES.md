# Third-party notices

This repository redistributes third-party code **verbatim, in the tree**. It has to:
an extension is installed by copying one folder, which must run exactly as committed —
no build step, no `package.json`, no `npm install`. There is no dependency resolution
step at install time to fetch anything, so anything the canvas needs is committed.

That makes this project a redistributor, and the terms below apply to anyone who copies
this repository or an extension folder out of it.

The project's own licence is [MIT](LICENSE). It does not replace anything here: the
files listed below stay under their own terms.

---

## pdf.js

| | |
| --- | --- |
| **Component** | Mozilla pdf.js (`pdfjs-dist`) |
| **Version** | 6.2.108 |
| **Copyright** | Copyright Mozilla Foundation and pdf.js contributors |
| **Licence** | Apache License, Version 2.0 |
| **Licence text** | [`.github/extensions/office-canvas/src/ui/vendor/LICENSE`](.github/extensions/office-canvas/src/ui/vendor/LICENSE) |
| **Upstream** | <https://github.com/mozilla/pdf.js> |
| **Vendored from** | <https://cdn.jsdelivr.net/npm/pdfjs-dist@6.2.108> |

pdf.js renders the exported PDF in the canvas. The vendored files, all under
`.github/extensions/office-canvas/src/ui/vendor/`, are recorded with their sizes and
digests in `pdfjs.manifest.json`, and are produced by `tools/vendor-pdfjs.mjs`:

| File | What it is |
| --- | --- |
| `pdf.min.mjs` | The pdf.js bundle, verbatim. Carries its own `@licstart` notice. |
| `pdf.worker.min.mjs.part0` … `.part2` | The pdf.js worker, verbatim, split into three files. `part0` carries the `@licstart` notice; the other two are continuation bytes of the same file. |
| `pdf-text-layer.css` | **Modified.** The `.textLayer` rule cut out of upstream `web/pdf_viewer.css`; the other style layers this viewer never mounts are dropped. Carries upstream's copyright and licence notice, and states the change. |
| `LICENSE` | Apache License 2.0, verbatim, as shipped by `pdfjs-dist@6.2.108`. |
| `pdfjs.manifest.json` | Ours, not Mozilla's — a record of what was vendored. |

The worker is split because extension install enforces a 1,000,000-byte per-file limit
and the worker is 1,262,398 bytes. The split is a packaging measure only: the bytes are
reassembled unmodified before pdf.js is given them, and the manifest's digest is checked
on every reassembly.

### On Apache-2.0 §4(d)

Section 4(d) applies only where the redistributed work includes a `NOTICE` file.
`mozilla/pdf.js` ships none — an explicit Contents-API request for `NOTICE` at the
repository root returns HTTP 404 — so no `NOTICE` propagation is owed, and adding
machinery for it would be complying with an obligation that does not exist.

---

## Microsoft Word

Not redistributed, and not an open-source dependency. The document tools drive an
installation of Microsoft Word that is already on the user's machine, through COM. No
Microsoft code is included in this repository or in any extension folder.

Your use of Word is governed by **your own licence agreement with Microsoft**. Nothing
in this repository's licence grants, restricts or otherwise affects it, and this project
cannot supply Word to anyone who lacks it.
