# pdf.js viability probes — findings

Run before writing any of issue #9, to avoid handing that layer a design that
had only been reasoned about. Every number here was produced by a probe in
[`probes/`](probes/), on Windows, against `pdfjs-dist@6.2.108` and Word-exported
PDFs.

## 1. The plan's prediction was wrong in both halves

The restructure plan this work was scoped against predicted the pdf.js **worker
would fit** the C4 envelope and the **binary font files would be the blocker**.
Measurement inverted both. (The plan's own account of this inversion is at
`git show 93c3536:docs/repo-restructure.md` §3.2, which reads: *"this section
originally predicted where C4 would bite. Both halves of the prediction turned
out to be wrong"* and names `standard_fonts/` as the predicted blocker.)

**Fonts are a non-issue.** The Word-exported PDF measured — `demo.docx`, one
document — embedded everything it used:

| Metric | Value |
| --- | --- |
| `/FontFile*` entries | 8 |
| Unique `BaseFont`s | 4, all subset-tagged |
| Standard-14 referenced | none |
| Type0/CID fonts | 2, all Identity encoding |
| Non-Identity CMaps | 0 |

`probe-pdf-fonts.mjs` states its own scope — "for THIS document" — and that scope
is kept here deliberately: one document cannot establish what Word does with
every font, and nothing below needs it to. So pdf.js's `standard_fonts/` (780 KB)
and `cmaps/` (1.17 MB, 169 binary files) are not shipped — which also removes the
binary-file problem entirely, since C4 refuses binaries. Were some other document
to reference a standard-14 font, the consequence would be a fallback face in the
rendered page, not a failure to render; that is the risk the narrow claim leaves
open, and it is accepted rather than unmeasured-and-unmentioned.

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

**On the strength of that last bullet.** It was established before
`probe-range-requests.mjs` could tell a failed pdf.js from a genuine absence:
the in-page script reported its outcome only into the DOM, the report tick keyed
on the requests pdf.js had *made*, and SIGINT was the only exit. Measured under
issue #109, serving 500 to the second range request produced a report whose
verdict line — `suffix-form ranges (bytes=-N): NONE` — was byte-identical to a
healthy run's, and exited 0. The probe now posts the browser's outcome back to
the server, prints that negative only on `ok:true`, and exits non-zero
otherwise. Nothing here says the finding is wrong; it says the run that
established it could not have told you, and a re-run now can.

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

The tagging is there because the export asks for it. The flag is passed
positionally — 12th of fourteen arguments, named only in a comment above the
call — so its value is not restated here; `test/unit/export-tagging.test.mjs`
derives it from the source and guards the alignment it depends on.

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
placed before any mutant, because a partial run is the hardest kind to notice.

The guard tests a **literal** em dash read from the script's own text —
`'X—X'.Length -ne 3` — which is 3 decoded as UTF-8 and 5 as ANSI, the em dash
arriving as its three UTF-8 bytes. Mutation-checked by stripping the BOM from a
copy: it exits 1, and passes with the BOM present.

**The guard originally had a second condition, and that one could never fire.**
It read `"$([char]0x2014)".Length -ne 1`, which builds the character from a
*number* — so it is one character however the file was decoded, measured at 1 in
both arms. It came first and read as the primary check, so a later simplification
would plausibly have kept the inert half and the guard would have died with the
file still green. Removed in `edb2c1c`. It was found by asking whether the
missing BOM was the *exclusive* cause of the three `MISSING` mutants rather than
merely a sufficient one: holding anchors and target fixed, the em-dash anchor
matches only with the BOM, while an ASCII anchor in the same two arms matches
either way — which is also the measurement behind the "only the count moves"
claim above.

Mechanisms 1 and 2 belong to one family: **a mutation gate can report a kill it
never made.** The two above are this branch's contribution to it; other sessions
have found others — a moved anchor and a runner-level variant — and the family is
enumerated in `CONTEXT.md`, which is where the count belongs. It is deliberately
not restated here, because it grows whenever another session finds a member and a
copy of it would be wrong without anything having changed in this file. The
generalisation that covers every member so far: **the mutant must be applied to
the artefact the tool actually reads.**

The dead condition above is a *different* failure and should not be counted among
them: nothing reported a false kill, and the guard worked throughout. It is an
instrument carrying a part that cannot move — the hazard is what a later reader
would keep, not what it did.

## Two font probes, declined

Issue #9 lists an `ENGR.TTF` export probe and a not-installed-font probe. Neither
was run, and this is a decision rather than an omission.

`standard_fonts/` and `cmaps/` are not in the package. In the one document
measured, Word embedded **every font it used** as a subset — 8 `/FontFile*`, zero
standard-14 references — so pdf.js never reached for a standard font file
(finding 1). That is a claim about `demo.docx`, not about Word: one document
cannot establish what Word does with every font, and `probe-pdf-fonts.mjs` says
"for THIS document" in its own conclusion.

The decision does not rest on the general claim, which is why it is not worth
widening the probe. No result either font probe could produce changes a byte of
what ships: the directories are already absent, a surprising result would tell us
something about Word's embedding rather than about the viewer, and there is no
code path it would send us to. A probe that cannot change a decision is not
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

## Reviewing this layer against itself found four defects, and two units of measure

Run before the last review round rather than after, on the argument that the
diff's *content* is stable across the pending rebase even though its ancestry is
not. Four findings, all in this layer's own files, all reproduced before being
believed.

### `owners` counted code points while `indexOf` counted code units

`buildPageText` builds a string and a parallel `owners` array mapping character
offset back to the pdf.js text item that produced it. The fill loop is
`for (const char of str)`, which yields **code points**; `text += char` appends
**one or two UTF-16 units**. One `owners.push` per iteration therefore desynchronises
by one entry per astral character — an emoji, CJK ext-B, a mathematical
alphanumeric. `locateText` then finds its match with `indexOf`, which counts
units, and indexes `owners` with that offset.

Measured, three emoji ahead of the target on one page:

| fixture | `text.length` | `owners.length` | located range |
| --- | --- | --- | --- |
| `😀😀😀 Titel` + `AAAA` + … | 33 | **30** | `{startItem: 1, endItem: 2}` |
| `XXX Titel` + `AAAA` + … | 30 | 30 | `{startItem: 1, endItem: 1}` |

The status is `located` in both. This is not the overlay degrading to a marker —
it is the overlay drawing a box over an item the edit never touched, which is the
one outcome `change-record.mjs` is written to prevent.

**An ASCII fixture cannot observe it**, because there the two units of measure
coincide. The existing test asserted exactly the right invariant
(`owners.length === text.length`) and could not fail: every fixture it had was
ASCII. That is this repo's recurring shape — a correct assertion with an input
set that cannot produce the failure — and it is why the new fixture is astral
rather than merely longer.

### A refresh that joined a no-op stamped the record against the pre-edit render

`ViewerInstance.refresh` returns early when `cache.refresh` reports the file has
not moved, **without re-rendering and without advancing `this.doc`**. The join
path awaited the in-flight refresh and then stamped unconditionally, never
consulting `joined.changed`.

The comment directly above `refresh` names this exact hazard — *"stamping it with
the current key here would instead publish the new edit's text over the pre-edit
render"* — and the join path three lines below committed it. Reachable when a
watcher echo fires before Word finishes saving: that refresh is a no-op, the
edit's own forced refresh joins it, and `force` was silently discarded so no
render containing the edit ever happened.

**No test could reach it.** `FakeCache.refresh` returned `changed: true`
unconditionally, so the early return at the heart of the bug was unreachable from
the unit suite — including from the test named *"a record arriving during an
in-flight refresh is not swallowed"*, which covers that very code path.

Both halves are now separately mutated, because either alone leaves the defect:
gating the stamp without fixing `force` loses the render, fixing `force` without
gating the stamp still publishes against a stale key.

### `vendor_missing` was reported for every read failure

The vendored-file route's `catch` never inspected `err.code`, so `EACCES`,
`EISDIR` and `EMFILE` all returned *"is missing. Run `node tools/vendor-pdfjs.mjs`"*.
That is a code naming a cause the code never distinguished, and the remedy is
actively misleading: for a permissions failure the prescribed script appears to
succeed while the file stays unservable. Split on `err.code === "ENOENT"`, which
is what the platform distinguishes; injected with a directory in the file's place.

### A deletion can never carry a page, so its documented marker was unreachable

`change-record.mjs` claimed unlocatable changes *"carry `locatable: false` and the
viewer shows a page-level marker instead"*. For `delete_paragraph` — the op that
paragraph is about — the host returns `0` as the touched index, the page is only
read `if ($touched -gt 0 ...)`, and `document-editor` maps `0` to `null`. So a
deletion has no page, `candidatePages` refuses to guess one, and only the text
notice appears. The prose promised a behaviour the pipeline cannot produce.

The comment is corrected rather than the code: capturing the page before the
range is deleted is a host-side change and is not asserted here. The
page-level-marker branch is *not* dead — it is reached by an unlocatable op whose
page Word did report — which is why both branches remain.

**The verification of this section's own mutation run was wrong first time.**
`Select-String -SimpleMatch` with a `[regex]::Escape`-d pattern searches for the
escaping backslashes, so all four restored anchors reported LOST. The files were
fine; the instrument was looking for text that never existed. Same class as every
other entry here, one level up: the check that confirms a fix has to be able to
tell the two answers apart.
