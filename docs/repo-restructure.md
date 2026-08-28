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
| **C4** | Gist sharing is **flat, UTF-8 only**, ~1 MB per file, ~5 MB total; binary files are refused. | runtime | Binary assets cannot ship via gist at all |
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
│  │     ├─ copilot-extension.json   # {name, version}; C4 requires it for gist install
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
vendors pdf.js once per extension folder against a gist cap of roughly 5 MB
that refuses binary files, while option c vendors it once.

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

`PLAN.md` §15.5 commits to vendoring pdf.js. C4 caps a gist at ~5 MB total and **refuses
binary files**. pdf.js's worker is large but is UTF-8 JavaScript and should fit; its
**standard font files are binary and would be rejected outright**.

So: gist sharing may stop being a viable distribution channel for this extension once pdf.js
lands. Repo-folder install (C3) and release zips have no such limit and become the primary
channels. Whether the font files are needed at all is testable — Word's PDF export embeds
fonts, so the standard-font fallback may never fire. **That is a probe, not an assumption**,
and it belongs next to the §15.5 accessibility probe.

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
- Enforce the C4 envelope: UTF-8, per-file ≤ 1 MB, total ≤ 5 MB — so gist share cannot fail
  in a way only discovered at share time.
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
- Assert the tag version equals `copilot-extension.json`'s `version`.
- Produce a folder zip, and a flattened gist-format bundle (`/` → `\`) when it
  fits inside C4.
- Attach both to a GitHub Release.

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
| 3 | Split tests into `unit/` and `integration/` | — | **dirs only** — `unit/` is empty |
| 4 | `tools/validate-extensions.mjs` + `validate.yml` | 2, 3 | blocked on 3 |
| 5 | `tools/package-extension.mjs` + `release.yml` | 2 | ready |
| 6 | Probe: does pdf.js need its binary font files? (§3.2) | — | open |
| 7 | Self-hosted Windows+Office runner, licence permitting | licence check | open |

Step 3 is the one that matters: the directories exist but `unit/` has nothing in it, so
step 4 would produce a workflow that runs zero tests and reports green. **Write the first
Office-free tests before writing `validate.yml`**, or the badge lies.

Step 7 may never be worth it. Step 6 is a decision this document deliberately leaves open
rather than guessing at.

The former step 7 — "decide option b vs c for shared code" — is gone: ADR 0003 decided it,
and §3.1 records the reasoning.

### 5.1 Verification of steps 1–2

The restructure was behaviour-preserving, checked rather than assumed: all three integration
suites pass unchanged against the moved tree — 18/18 host, 16/16 cache, 29/29 server, with no
`WINWORD.EXE` left behind. `node --check` passes on every `.mjs`, the manifest name matches
its folder, no `package.json` exists in the extension (C2), and the folder now contains zero
binary files (C4).
