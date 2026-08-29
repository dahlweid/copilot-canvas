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
- A third, made after the fact and worth as much as the other two: this file used
  to justify the impossibility of a comparison with "the plugin exposes no DOM".
  That is an outcome-level reason and it is wrong at the mechanism — a screen
  reader consumes an accessibility tree, not a DOM. pdf.js reaches that tree
  through the DOM; the plugin reaches it directly from its own process. Both are
  legible to AT. Only one is legible to a probe running in the page, so the
  asymmetry belongs to our instrument, and writing it as a property of the
  viewers would have credited pdf.js with an advantage nothing here measured.

**What a successful comparison would have looked like** — recorded so that "not
measured" is a finding with a shape, not a shrug. Instrument: a real screen
reader's speech log, driven over a desktop. Arms: one tagged PDF from our own
pipeline, opened in the canvas and in the native plugin. Measurement: the
announced sequence for one page — reading order, and whether each heading is
announced as a heading at its level. Verdict: parity iff both agree on order and
level; a difference either way is the result, and its direction says whether the
gap above matters or whether pdf.js is being undersold here. Control: a document
with deliberately broken tags must come back a *mismatch*, or a match proves
only that the instrument is insensitive.

Not measured by anyone: actual assistive-technology output, from either viewer.

## Two mutants that ran and one class of mutant that cannot

A mutation gate reports three outcomes, not two. `KILLED` and `SURVIVED` both
mean the mutant was applied; a mutant whose anchor is absent was **never
applied**, and counting it either way is a lie. `mutate-pdfjs.ps1` therefore has
a third category, `MISSING`, and fails on it. It has now caught two separate
mechanisms in this branch alone, neither of which produced a red line before it
existed.

**Mechanism 1 — the tool never reads the file you mutated.** `.gitattributes`
marking the vendored parts `-text` is a guard like any other, so the mutant moved
it aside. But git resolves attributes from the **index** when the working-tree
file is absent, so once `.gitattributes` was committed the mutant stopped
disabling anything and both of its tests went `KILLED` to `SURVIVED` silently.
The mutant must `git rm --cached` as well. **Generally: mutating a tracked
configuration file is inert whenever the consuming tool reads the committed copy
rather than the working one** — and that is the normal case for git's own
configuration, not an exotic one.

**Mechanism 2 — the anchor is mangled before it is ever compared.**
`mutate-pdfjs.ps1` is run with `powershell -File`, which is Windows PowerShell
5.1, and 5.1 decodes a BOM-less file as ANSI. Every anchor in it was ASCII until
an em dash appeared in one, at which point three mutants stopped matching and
were never applied. The failure is worse than it looks: the mutants with ASCII
anchors keep passing, so the run still ends in a wall of `KILLED` and only the
count moves. Fixed by giving the script a UTF-8 BOM, and guarded by a self-check
that asserts an em dash is one character before any mutant runs — placed first,
because a partial run is the hardest kind to notice. The self-check was itself
mutation-checked by stripping the BOM from a copy: it exits 1.

Both belong to one family: **a mutation gate can report a kill it never made.**
Three sessions have now found different members of it independently — a moved
anchor, a config read from the index, and a mis-decoded anchor — which suggests
the category is worth naming in `CONTEXT.md` rather than the instances.

## Two font probes, declined

Issue #9 lists an `ENGR.TTF` export probe and a not-installed-font probe. Neither
was run, and this is a decision rather than an omission.

`standard_fonts/` and `cmaps/` are not in the package — Word embeds every font as
a subset, so pdf.js never reaches for a standard font file (finding 1). No result
either probe could produce changes a byte of what ships: a surprising result
would tell us something about Word's embedding, not about the viewer, and there
is no code path it would send us to. A probe that cannot change a decision is not
evidence, it is cost — and on a machine with a dozen concurrent Word instances it
would produce numbers needing a caveat longer than the finding.

Re-open both if the font directories are ever shipped, which is the condition
under which the answer would start to matter.

## The unit-test step could have gone green having run nothing

`node --test <pattern>` expands the pattern **itself**, and a pattern it matches
nothing with runs nothing and **exits 0**. Measured from this repo root:

| invocation | files | result |
| --- | --- | --- |
| `node --test "**/test/unit/*.test.mjs"` | — | `tests 0`, **exit 0** |
| `node --test ".github/extensions/office-canvas/test/unit/*.test.mjs"` | 28 | `tests 379`, exit 0 |

The first fails because `**` does not descend into a dot-directory and the
extension lives under `.github/`. CI used the second, correct form, so nothing
was broken — but nothing *checked* that it stayed correct either, and the failure
is silent: a moved path would leave the step green having tested not one line.

Fixed by removing the second globber rather than guarding it. The shell expands
the file set, the set is checked non-empty, and Node is handed discrete paths:

```bash
files=(.github/extensions/office-canvas/test/unit/*.test.mjs)
[ "${#files[@]}" -eq 0 ] && exit 1
node --test "${files[@]}"
```

One parser instead of two, which is the same reasoning the repo already applies
to paths crossing `cmd.exe` and PowerShell.

**A count check was written first and then removed, because it could not fail.**
The intuition — "assert Node reports at least one test" — does not survive
measurement: a file containing no `test()` calls is still reported as `tests 1`,
so once the file list is known non-empty, "collected zero tests" is not a
reachable state. It would have been a guard with no producible failure, sitting
in CI looking like protection.

Verified with both arms against the real tree: 28 files and 379 tests, exit 0;
and from a directory without the path, `exit 1` with the error. The first arm is
the control — without it the check passes trivially by never being exercised.

**The probe that found this measured nothing on its first two attempts,** and
both were path-translation errors rather than anything interesting: Windows paths
handed to Git Bash matched no files, so every arm agreed and "failed" for the
same wrong reason; then bash-native `/tmp` paths were unresolvable by the Windows
`node`, so the positive control failed too. Every arm agreeing is the signature.
