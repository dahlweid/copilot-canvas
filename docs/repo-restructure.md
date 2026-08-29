# Repository restructure: `copilot-canvas`, and building on the server

**Status:** §2 and §3 executed; §4 workflows not yet written
**Covers:** (1) renaming the repo for a multi-canvas future, (2) the directory and workflow
structure needed to validate, version and package extensions in CI.

---

## 1. The constraints this plan has to live inside

These are not preferences. They come from the extension runtime and from the code as it
stands, and three of them rule out the obvious approach.

| | Constraint | Source | Consequence |
| --- | --- | --- | --- |
| **C1** | The entry file must be `.github/extensions/<name>/extension.mjs`. Discovery scans **only immediate subdirectories** of `.github/extensions/`. | runtime | One folder per extension, no nesting, no grouping folder |
| **C2** | `@github/copilot-sdk` is auto-resolved. **Do not add a `package.json` or `node_modules` for it.** | runtime, stated explicitly | No conventional npm build |
| **C3** | `install_extension` can install **a folder in a GitHub repo** (`/tree/<ref>/path/to/ext`). | tool contract | Each extension folder must be **runnable exactly as committed** |
| **C4** | `install_extension` enforces **1,000,000 bytes per file and 5,000,000 bytes total** (decimal, not MiB) on **both** install paths — gist *and* repo folder. Gist sharing additionally flattens to one level and refuses non-UTF-8 files. | measured — see §3.2 | Binary assets cannot ship via gist at all, and no install path escapes the size caps |
| **C5** | **Every existing test drives Word through `powershell.exe`.** `cache-smoke`, `host-smoke`, `server-smoke` and `make-fixture.ps1` all need Word installed. | measured — grep of `test/` | GitHub-hosted runners cannot run **any** current test |

**C2 + C3 together mean this repository cannot have a conventional build.** There is no
bundling step whose output differs from its source, because the runtime loads the committed
source directly. "Build and package on the server" therefore has to mean something other than
`npm run build`, and §4 defines what.

**C5 is the expensive one** and is stated first deliberately: a CI plan written without it
produces three workflow files that cannot run.

---

## 2. Renaming to `copilot-canvas`

Low risk, and worth doing before a second canvas exists rather than after.

### 2.1 Steps

1. **GitHub:** Settings → rename `copilot-canvas-word` → `copilot-canvas`. GitHub keeps
   redirects for both git and web URLs indefinitely, *unless* someone later creates a repo at
   the old name. Redirects are a safety net, not an interface — update references rather than
   relying on them.
2. **Remotes**, in every clone and worktree:
   `git remote set-url origin https://github.com/dahlweid/copilot-canvas.git`
3. **README:** the `# copilot-canvas-word` heading and the framing, which currently describes
   a single-purpose repo. It becomes an index of canvases.
4. **Nothing else in-tree** references the repo name. Verified by grep.

### 2.2 What must *not* change

| Identifier | Value | Why it is frozen |
| --- | --- | --- |
| `canvasId` | `word-doc` | Already generic; agents and any saved state key off it |

The repo name and the extension name are independent. Renaming the repo touches neither.

**Correction, on execution.** This section previously froze the extension folder at
`word-canvas` on the grounds that `install_extension` URLs point at it and user installs live
at `$COPILOT_HOME/extensions/word-canvas/`. ADR 0003 makes that untenable — one extension
hosting a Word *and* a PowerPoint canvas cannot be called `word-canvas` — so the folder was
renamed to `office-canvas`, and with it the derived `extensionId`.

The freeze was also based on a misreading. The only thing at
`~/.copilot/extensions/word-canvas/` on this machine was an `artifacts/` directory — the
render cache and recents file written at runtime by `store.mjs`, with no `extension.mjs`
beside it. It was never an install. So the migration cost was one directory move, and the
cache is regenerable anyway; only `recents.json` carried anything a user would miss.

Anyone who *had* installed the extension keeps a working `word-canvas` copy that no longer
receives updates, and would need to install `office-canvas` and delete the old folder — both
register `canvasId` `word-doc`, so leaving both in place is a collision. That cost is real
but is at its lowest now, before distribution, which is the argument for doing the rename
first rather than later.

### 2.3 The one non-obvious cost

Local directory names do **not** follow a GitHub rename. The Copilot app has this project
registered against a filesystem path (`C:\Git\...\copilot-canvas-word\`), and renaming that
directory breaks the registration. Either leave local paths alone — they are just names — or
re-add the project through the app afterwards. Do not rename the directory and expect the app
to notice.

**On execution:** local paths were deliberately left alone, so the working copy still lives
under `copilot-canvas-word\`. The remotes in both the main checkout and this worktree were
repointed at the new URL and a fetch verified against it; GitHub's redirect was not relied on.

---

## 3. Directory structure

```
copilot-canvas/
├─ .github/
│  ├─ extensions/                    # C1: immediate subdirs only
│  │  └─ office-canvas/              # ONE extension, a canvas per app (ADR 0003)
│  │     ├─ extension.mjs            # C1: entry, exact name
│  │     ├─ copilot-extension.json   # {name, version:1, productVersion}; the install gate
│  │     ├─ src/
│  │     │  ├─ word/                 # Word COM host — app-specific
│  │     │  ├─ powerpoint/           # (next) PowerPoint host
│  │     │  ├─ ui/                   # viewer front-end; pdf.js vendored here
│  │     │  └─ *.mjs                 # app-agnostic: cache, server, watcher, store
│  │     └─ test/
│  │        ├─ unit/                 # Office-free  → hosted runners  (EMPTY — §4.4)
│  │        └─ integration/          # needs Office → self-hosted only
│  └─ workflows/
│     ├─ validate.yml                # every push/PR, ubuntu-latest
│     ├─ integration-windows.yml     # self-hosted + Office; nightly and manual
│     └─ release.yml                 # on tag; packages and attaches artifacts
├─ spikes/                           # outside the extension — see below
├─ tools/
│  ├─ validate-extensions.mjs
│  └─ package-extension.mjs
├─ docs/
│  └─ adr/
├─ CONTEXT.md
└─ README.md
```

**There is no `src/shared/`.** The tree above previously showed one, but a `shared/`
directory naming what is shared between one implementation and a hypothetical second is a
guess at a seam. Application-specific code goes in `src/<app>/`; everything else sits at
`src/` and is app-agnostic by default. The one genuine seam — `render-cache.mjs` importing
`word/word-host.mjs` — is left as a visible direct import, to be broken when PowerPoint
provides a second implementation to generalise against.

**`spikes/` moved to the repository root**, out of the extension folder. C3 copies an
extension folder wholesale, and the spike carried 300 KB of JPEG screenshots that no install
needs — which C4 would also refuse outright, since gists reject binary files. Extracting it
took the extension from 619 KB to 141 KB with zero binary files, so the C4 envelope now has
headroom for pdf.js (§3.2).

Excel would arrive later as `src/excel/` with its own canvas, once its display
surface is decided — it will not be a page renderer.

### 3.1 The shared-code problem, which is the real decision

Word, Excel and PowerPoint canvases will share most of their substance: the loopback server,
the render cache, the hidden-COM host pattern, the file watcher, and the pdf.js viewer from
`PLAN.md` §15.5. But C1 and C3 require each extension folder to be self-contained — a
top-level `shared/` is neither copied by `install_extension` nor resolvable at runtime from
inside `.github/extensions/<name>/`. A relative import such as `../../../shared/x.mjs` works
in a git checkout and breaks the moment the folder is installed on its own.

Three options:

| | Approach | Verdict |
| --- | --- | --- |
| **a** | Duplicate shared code into each extension folder | Simple, self-contained, **drifts within weeks**. Rejected |
| **b** | `shared/` is the source of truth; `tools/sync-shared.mjs` copies it into each `src/vendor/shared/`, and CI fails if a copy is stale | Keeps folders installable standalone with one source of truth, at the price of committing generated copies and policing staleness |
| **c** | **One extension registering several canvases.** `joinSession` accepts `canvases: Canvas[]` | **Chosen.** Sidesteps sharing entirely — see ADR 0003 |

**Option c is the decision** (ADR 0003), reversing this document's earlier
recommendation of **b**. Two things changed. Excel is deferred because a PDF
cannot show a formula, so the near-term set is Word and PowerPoint — two
canvases that share not "most of their substance" but effectively all of it,
since both paginate natively and both render through the same export pipeline.
The sync-shared machinery would exist to serve a distinction that carries no
weight between them. And §3.2's pdf.js constraint bites hardest here: option b
vendors pdf.js once per extension folder against a 5,000,000-byte install cap,
while option c vendors it once.

The two objections raised against **c** above are worth answering rather than
dropping:

- *"Users cannot install Word alone."* True, and acceptable — one install of
  "Office document tools" is a better story than three, which is also what the
  user-scope global install is for. It does impose a design requirement:
  **no COM object may be created at startup**, so that a machine with Word but
  no PowerPoint fails only when the PowerPoint canvas is actually used.
- *"One process failure takes down all canvases."* Largely illusory. The risky
  component is the Office process, which is already separate and already
  per-application. Shared code is byte-identical under option b as well, so a
  defect in it affects both canvases either way.

The reversibility argument for **b** does not survive scrutiny either:
splitting one extension into two is copying a folder and deleting the canvas
registrations you do not want, which is no harder than the collapse it was
meant to keep cheap.

Under option c there is no `shared/` directory and no `sync-shared.mjs`,
because there is nothing to sync — one folder holds the code and imports it
normally.

The "build" this repository has is therefore **copy-and-verify, not bundle.** That is the
honest shape of it under C2.

### 3.2 A packaging constraint that lands on the pdf.js decision

`PLAN.md` §15.5 commits to vendoring pdf.js, and this section originally *predicted* where
C4 would bite. Both halves of the prediction turned out to be wrong, so what follows is the
measurement that replaced it.

**The caps are decimal and they apply to every install path.** `install_extension` is
implemented in the app binary, not in `copilot-sdk/`. Pushing deliberately oversized inputs
through all four paths gives:

| Path | Over per-file | Over total |
| --- | --- | --- |
| `share_extension` (gist) | `too large (1536000 bytes > 1000000 byte limit)` | `too large to share via gist (>5000000 bytes)` |
| `install_extension` (repo folder) | same `> 1000000 byte limit` | `Folder contents exceed the 5000000 byte total limit.` |

So the limits are **1,000,000 bytes per file and 5,000,000 bytes total — decimal, not
`1024*1024` and `5*1024*1024`.** The earlier claim here that *"repo-folder install and release
zips have no such limit and become the primary channels"* is false: repo-folder install has
**identical** limits, and a release zip escapes them only because nothing installs a zip —
`install_extension` accepts a gist or a repo folder and nothing else, so a zip is not a
distribution channel in any useful sense. Gist sharing is not the constrained path; it is
merely the path that *additionally* flattens to one level and refuses non-UTF-8 files.

The decimal/binary distinction is not pedantry: `tools/validate-extensions.mjs` used the
binary constants, which left a 48,576-byte window in which a file passed validation and was
then rejected at install. That is corrected, and `test/unit/validator.test.mjs` now pins the
boundary from both sides so the values cannot be "tidied" back to `1024*1024`.

**Fonts are a non-issue; the worker is the constraint.** The predicted blocker was pdf.js's
binary `standard_fonts/`. Probing one Word-exported PDF found 8 `/FontFile*` entries, 4
subset-tagged `BaseFont`s, zero standard-14 references and Identity-encoded CID fonts — that
document embeds every font it uses, so `standard_fonts/` (780 KB binary) and `cmaps/` (1.17 MB,
169 binary files) are not shipped, and C4's binary refusal never fires. One document does not
establish the general case and the probe does not claim it; see `spikes/pdfjs/FINDINGS.md` §1
for the scope kept and the risk accepted. What does bite is
`pdf.worker.min.mjs` at **1,262,398 bytes**, which exceeds the *per-file* cap on both paths.
It is UTF-8 JavaScript, so it is not refused for being binary — it is simply too large for any
single file. §15.5's design therefore splits the worker across three committed parts under
`src/vendor/` reassembled behind one server route, rather than switching channels; splitting
at packaging time would violate C3 by producing an installable artefact from a source tree
that is not itself installable.

Against the real 5,000,000-byte total, the packaged extension is **107,232 bytes today**
(`tools/package-extension.mjs` reports it), and pdf.js projects to about 1.8 MB — roughly 36%
of budget.

### 3.3 The manifest field that made the extension uninstallable

Packaging surfaced a defect that no amount of local testing could have: the extension **ran
perfectly in the development checkout and could not be installed by anyone.**

`copilot-extension.json` carried `"version": "1.0.0"`. The in-place loader never looks at it,
so nothing complained. `install_extension` does, and refuses the whole extension:

```
Failed to parse `copilot-extension.json`: invalid type: string "1.0.0", expected u32
```

`version` is the **manifest format version**, not the product version — the app's own authoring
guide gives the shape as `{ "name": "<name>", "version": 1 }`, and `mobile-canvas`, the one
extension installed on this machine, uses `1`. Unknown keys *are* tolerated (measured: a
manifest carrying `productVersion` installs cleanly), so the product version gets its own key
and `release.yml` tags from that.

Two further behaviours were measured while proving this, both worth knowing:

- **The installer does not copy `copilot-extension.json` into the installed folder.** It is
  purely an install-time gate. The installed tree is the other ten files, byte-identical to
  the packaged artefact.
- **Gist keys round-trip correctly.** `src\ui\app.js` in the flat bundle is restored as
  `src/ui/app.js` on install.

This is the strongest argument for the round-trip criterion in issue #10. A packaging tool
that only checked file layout would have produced a green build for an artefact nobody could
install, and `validate-extensions.mjs` — the very file whose job is to prevent that — was
asserting the version *had* to be semver. Both are now fixed and pinned by fault-injection
tests.

---

## 4. Workflows

### 4.1 `validate.yml` — every push and PR, `ubuntu-latest`, no Word

The point is fast signal at zero infrastructure cost. It cannot test rendering, and must not
be described as if it does.

- `node --check` every `.mjs`.
- **Reject `console.log` in the extension process.** This is the documented failure mode that
  corrupts the JSON-RPC channel and takes the extension down. `src/` and `extension.mjs` are
  **clean today** (verified), so this is a regression guard rather than a fix — and it must
  **exclude `src/ui/`**, which runs in the canvas iframe where `console` is legitimate. A
  naive repo-wide rule would be wrong.
- Assert each `.github/extensions/*/` has `extension.mjs` and `copilot-extension.json`, and
  that the manifest `name` matches its folder name.
- Assert **no `package.json`** exists in any extension folder (C2).
- Enforce the C4 envelope: UTF-8, per-file ≤ 1,000,000 bytes, total ≤ 5,000,000 bytes (decimal
  — see §3.2) — so neither gist share nor repo-folder install can fail in a way only discovered
  at install time.
- Verify no file under an extension folder imports from outside it (C3), which
  is what makes the folder installable as committed.
- Run `test/unit/` — currently **empty**, see §4.4.

### 4.2 `integration-windows.yml` — self-hosted, Office installed

Runs `test/integration/` (the three existing suites). Nightly plus manual dispatch, not on
every PR: it needs a machine we own.

**This requires a self-hosted Windows runner with a licensed Office install, and the licensing
is the blocker, not the runner.** Microsoft 365 Apps on shared or cloud VM infrastructure
generally requires Shared Computer Activation and an appropriate subscription. Confirm the
licence position **before** provisioning anything.

Until that exists, the honest state is: **rendering is verified by a human running the suites
locally before a release.** That should be written into the release checklist rather than
papered over with a green badge.

### 4.3 `release.yml` — on tag

- Tag format `v1.2.0`. With a single extension there is nothing to version
  independently; revisit only if the repo ever ships a second one.
- Assert the tag version equals `copilot-extension.json`'s **`productVersion`**. Note that the
  manifest's `version` key is *not* the product version — see §3.3.
- Produce the installable folder, a zip of it, and a flattened gist bundle.

One detail worth recording, because it changed the shape of the output: the
flattened gist bundle **cannot be written as files on Windows**. Gist keys encode
a nested path with a backslash (verified against a real shared gist:
`src\ui\nested.js`), and on Windows the backslash *is* the path separator, so
writing that key creates a subdirectory instead of a flat entry. The bundle is
therefore emitted as a **JSON API request body** —
`dist/office-canvas-<version>.gist.json`, `POST`-able to `/gists` unmodified —
rather than as a directory. It is checked against C4 either way.

### 4.4 The prerequisite nobody can skip

**There are currently no Word-free tests, so `validate.yml` would start with nothing to run.**
Splitting the suites is a real task, not a `mkdir`:

- Genuinely Word-free and worth extracting now: `normalizeDocPath`, cache key derivation and
  eviction, path resolution, HTTP routing and range handling in `server.mjs`, `watcher.mjs`
  debounce logic.
- Irreducibly Word-dependent: everything in `host-smoke.mjs`, PDF export, and fixture
  generation.

Do this split **before** writing `validate.yml`, otherwise the workflow is decorative.

---

## 5. Order of work

| | Step | Depends on | State |
| --- | --- | --- | --- |
| 1 | Rename the repo; update remotes and README | — | **done** |
| 2 | Collapse to one `office-canvas` extension; move `spikes/` out; add `copilot-extension.json` | — | **done** |
| 3 | Split tests into `unit/` and `integration/` | — | **done** — 42 Office-free tests |
| 4 | `tools/validate-extensions.mjs` + `validate.yml` | 2, 3 | **done** — green on `ubuntu-latest` |
| 5 | `tools/package-extension.mjs` + `release.yml` | 2 | **done** — see §5.3 |
| 6 | Probe: does pdf.js need its binary font files? (§3.2) | — | **done** — no; the worker is the constraint |
| 7 | Self-hosted Windows+Office runner, licence permitting | licence check | open |

Step 3 is done. It was the one that mattered: the directories existed but `unit/` had
nothing in it, so step 4 would have produced a workflow that runs zero tests and reports
green. §5.2 records what running them actually found.

Step 7 may never be worth it. Step 6 is answered in §3.2, and the answer inverted the
prediction: the fonts are unnecessary and the worker is what does not fit.

The former step 7 — "decide option b vs c for shared code" — is gone: ADR 0003 decided it,
and §3.1 records the reasoning.

### 5.1 Verification of steps 1–2

The restructure was behaviour-preserving, checked rather than assumed: all three integration
suites pass unchanged against the moved tree — 18/18 host, 16/16 cache, 29/29 server, with no
`WINWORD.EXE` left behind. `node --check` passes on every `.mjs`, the manifest name matches
its folder, no `package.json` exists in the extension (C2), and the folder now contains zero
binary files (C4).

### 5.2 Verification of steps 3–4

The worry above — a workflow that runs zero tests and reports green — is answered by
running the suite somewhere Office does not exist. **42/42 pass on Linux x86_64 (WSL,
Node 22.11.0) with no Word and no PowerShell on `PATH`**, and the same 42 pass on Windows.
The workflow does exactly what that check did, so the badge is load-bearing.

Two things were found by writing the tests rather than by reading the code:

- **`bytes=-500` was served as bytes 0–500.** A suffix range means the *last* 500 bytes
  (RFC 9110 §14.1.2). Extracting `parseByteRange` for testability exposed it, and the fix is
  right on the spec — but the reason given here for why it mattered was wrong. This
  paragraph used to say pdf.js "fetches the trailer that way". It does not:
  `spikes/pdfjs/probes/probe-range-requests.mjs` recorded **zero** suffix ranges under every
  configuration tried, including `disableStream` with an 8 KB chunk size. pdf.js computes the
  tail's offset itself from `Content-Length` and asks for an explicit `bytes=start-end`.
  Under the default configuration it issues no Range request at all. So the bug was real and
  the correction is worth keeping; what it was not is a bug pdf.js would ever have hit. The
  suffix branch has no consumer in this repo, and must not be "verified" by assuming the
  viewer exercises it.
- **The C3 self-containment check missed bare `import "…"`.** It matched only the
  `from`-bearing and dynamic forms, so a side-effect import escaping the folder passed. The
  fault-injection test caught this on its first run — which is the argument for having
  written it. All five import shapes are now covered.

`validator.test.mjs` breaks one invariant at a time in a staged copy of the repo and asserts
the validator rejects it *and* accepts the untouched copy, so the failures cannot be an
artefact of everything being broken.

### 5.3 Verification of step 5

The same Office-free check applies to packaging: **62/62 unit tests pass on Linux x86_64 (WSL,
Node 22.11.0)**, 17 of them covering the packager, and the same 62 pass on Windows. More
usefully, the packaged artefact is **byte-identical across the two platforms** —
`sha256:859370ecfc85feb6b35c4b47346b2faf10ff45717d00c316bf568654f27ff49f` from both — which is
what makes the reproducibility check in `release.yml` meaningful rather than a tautology.
Determinism comes from a sorted file list, verbatim byte copies (C3 forbids normalising line
endings, so nothing is rewritten), every output mtime pinned to `SOURCE_DATE_EPOCH`, and a
manifest that contains no timestamp.

Two decisions are worth recording:

- **The packager runs `tools/validate-extensions.mjs` rather than reimplementing it**, and
  runs it **twice** — once on the source, once on the artefact. The second run is not
  redundant: the first refuses to package anything CI would reject, including material an
  exclusion rule would otherwise hide; the second catches an exclusion that removed something
  the extension actually needs. Both directions have tests.
- **Exclusion is a small named ruleset, not a glob list**, and each rule carries a `why` that
  is printed at package time. Only `test/` is expected to match in practice; the rest
  (`spikes/`, `artifacts/`, VCS and editor metadata, OS junk, secrets, build output) exist so
  that a future stray file is excluded *deliberately* and visibly. `artifacts/` matters more
  than it looks: on a development machine it holds exported PDFs of the user's own documents,
  which must never be shipped.

The one criterion in issue #10 that cannot be met on CI is the end-to-end round trip —
install the packaged artefact and confirm the canvas renders a document — because that needs
Word (C5). It was run by hand and it earned its place: it is what found §3.3. The procedure
was: package, `POST` the gist body to `/gists`, `install_extension` from that gist, move the
development copy out of `.github/extensions/` so the `word-doc` canvas id is unambiguous,
open a five-page Word fixture, and confirm `get_info` and `get_outline` return real pagination
(5 pages, 1,399 words, 10 headings across pages 1–5) rather than a status string. The gist and
the installed copy were then deleted and the development copy restored. This belongs in the
§4.2 release checklist, not in `release.yml`.
