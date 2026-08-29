# pdf.js viability probes — findings

Run before writing any of issue #9, to avoid handing that layer a design that
had only been reasoned about. Every number here was produced by a probe in
[`probes/`](probes/), on Windows, against `pdfjs-dist@6.2.108` and Word-exported
PDFs.

## 1. The plan's prediction was wrong in both halves

`docs/repo-restructure.md` §3.2 predicted the pdf.js **worker would fit** the C4
envelope and the **binary font files would be the blocker**. Measurement inverted
both.

**Fonts are a non-issue.** A Word-exported PDF embeds everything it uses:

| Metric | Value |
| --- | --- |
| `/FontFile*` entries | 8 |
| Unique `BaseFont`s | 4, all subset-tagged |
| Standard-14 referenced | none |
| Type0/CID fonts | 2, all Identity encoding |
| Non-Identity CMaps | 0 |

So pdf.js's `standard_fonts/` (780 KB) and `cmaps/` (1.17 MB, 169 binary files)
never need to ship — which also removes the binary-file problem entirely, since
C4 refuses binaries.

Scanning all 446 font faces in `C:\Windows\Fonts` found **none that
unambiguously forbids embedding**. Exactly one, `ENGR.TTF`, sets
`fsType=0x000e` — bits 1, 2 and 3 together, which the OpenType spec defines as
mutually exclusive. That is malformed, not a restriction. It remains the ideal
test case precisely because it shows how Word resolves a contradictory `fsType`.

**The worker is the real blocker**, at 1,262,398 bytes against a 1,000,000-byte
per-file cap.

## 2. The cap is universal and decimal

Measured by the packaging work (#10) by pushing oversized files through both
install paths, not inferred from error strings:

| | limit | applies to |
| --- | --- | --- |
| per file | 1,000,000 bytes | gist **and** repo-folder install |
| total | 5,000,000 bytes | gist **and** repo-folder install |

Decimal, not binary. `tools/validate-extensions.mjs` had these as `1024*1024`
and was therefore too permissive by 48,576 bytes per file.

## 3. Splitting the worker works, end to end

`probe-worker-split.mjs` splits, serves the parts concatenated behind one route,
and loads the result **in the app's own webview** — the surface the canvas
actually uses, not a headless browser.

| Assertion | Result |
| --- | --- |
| splits into parts under the cap | 3 parts: 600,000 + 600,000 + 62,398 |
| every part independently valid UTF-8 | pass |
| concatenation byte-exact (SHA-256) | pass |
| served route reassembles to original | pass |
| pdf.js loads a Word PDF from it | 3 pages |
| page renders non-blank | 595×841, 64,945 non-white pixels |
| text extraction | matches source |

Three corrections fell out of running it:

1. The worker needs **three** parts, not two — 1,262,398 bytes, not the ~1.20 MB
   a size listing suggested. Split by cap; never hard-code a part count.
2. The split must happen at **vendoring time with the parts committed**. C3
   requires the folder to run exactly as committed, and repo-folder install never
   runs the packager — a packaging-time split would yield an installable artifact
   from a source tree that is not installable.
3. The UTF-8 rule is *not* "boundary byte < 0x80". It is that the boundary must
   not fall inside a multi-byte sequence, i.e. must not be a continuation byte
   (`0x80`–`0xBF`). Newline boundaries are preferable for diffability but **did
   not fire** on minified JS — the continuation-byte fallback did all the work,
   so both paths are needed.

`checkPackagingEnvelope` decodes with `new TextDecoder("utf-8", {fatal:true})`,
so a mis-split fails validation on `ubuntu-latest` before it can reach an install.

## 4. pdf.js never issues suffix ranges

Issue #9 originally said L4 would exercise `parseByteRange` "via pdf.js suffix
ranges". Wrong twice:

- **Default config issues no Range requests at all.** One unranged GET; the whole
  body comes back; `parseByteRange` is never called.
- **Ranges appear only with `disableStream: true`** — 14 requests for the 265 KB
  fixture, every one parsed correctly by the real `parseByteRange`.
- **Zero requests used the suffix form `bytes=-N`.** pdf.js always sends a bounded
  `start-end`, computing the tail itself from `Content-Length`.

Since our PDFs are local and small, ranged mode costs 14 round trips to save
nothing. Leave streaming on; keep serving `content-length` and `accept-ranges`
so pdf.js can make the single-request choice. `parseByteRange`'s suffix branch
stays unit-test-only — do not describe it as pdf.js-driven.

## Budget

| Item | Bytes |
| --- | --- |
| worker (3 parts) | 1,262,398 |
| `pdf.min.mjs` | 454,669 |
| existing extension artifact | 107,232 |
| `standard_fonts/` + `cmaps/` | 0 — not needed |
| **total** | **1,824,299 of 5,000,000 (36%)** |

Assert against `dist/office-canvas-<version>.package.json` rather than by eye:
it carries `.totalBytes`, `.budget.maxFileBytes`, `.budget.maxTotalBytes`,
`.budget.largestFile.{path,bytes}`, `.budget.headroomBytes`, and `.files[]` with
per-file `path`/`bytes`/`sha256`.

## Accessibility — measured, and one half that cannot be

`probe-accessibility.mjs`, five-page document exported through our own render
pipeline, pdf.js 6.2.108.

**The comparison ADR 0004 asked for is structurally impossible.** The native
plugin exposes no DOM to the embedding page — the fact the whole decision rests
on — so there is no accessibility tree to enumerate on that side. Parity cannot
be established this way at all, only by driving a real screen reader against
both. What follows is therefore absolute, not comparative.

| measurement | result |
| --- | --- |
| `/MarkInfo`, `/Marked true`, `/StructTreeRoot`, `/Lang` | all present |
| `getStructTree()` | 19 nodes, `Root > Document > H1/H2/P` |
| distinct roles | `Root`, `Document`, `H1`, `H2`, `P` |
| marked-content refs in the tree | 17 |
| text-layer spans | 31 |
| DOM order vs extraction order | identical |
| **spans carrying a marked-content id** | **0** |

Word already exports with `DocStructureTags = $true`, which is why the tagging
is there.

**The gap, named:** the structure exists in the file and pdf.js hands it over,
but a bare `TextLayer` does not wire it into the DOM. A screen reader walking our
markup sees 31 correctly-ordered spans with `role="presentation"` and no
headings. pdf.js's own viewer mounts a separate structure-tree layer to close
this; we mount only the text layer. Bounded work, not an unknown — but it means
the roles above must not be read as accessibility parity.

**Two methodological corrections made while probing, both of the "measured
nothing" kind:**

- A first version counted structure roles with a regex over the raw PDF bytes and
  reported `{"/L": 1}` for a document full of headings. Word writes the structure
  tree into a compressed object stream, so those tokens are not present as bytes;
  the single hit was a coincidence. Roles must be read through pdf.js, which
  decompresses them.
- A first version walked the struct tree recording `hasRef: Boolean(node.id)` on
  roled nodes and found none, which reads like "the tree has no refs". The refs
  live on `{type: "content", id}` leaves, which carry no `role` and so were never
  visited. Counted properly there are 17.

Not measured by anyone: actual assistive-technology output, from either viewer.
