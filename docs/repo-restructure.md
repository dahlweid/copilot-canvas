# Repository restructure: `copilot-canvas`, and building on the server

**Status:** plan, not executed
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
| Extension folder | `word-canvas` | `install_extension` repo-folder URLs point at it, and existing user installs live at `$COPILOT_HOME/extensions/word-canvas/` |
| `extensionId` | `project:word-canvas` (auto-derived from the folder) | Changing the folder changes the id |
| `canvasId` | `word-doc` | Already generic; agents and any saved state key off it |

The repo name and the extension name are independent. Renaming the repo touches neither.

### 2.3 The one non-obvious cost

Local directory names do **not** follow a GitHub rename. The Copilot app has this project
registered against a filesystem path (`C:\Git\...\copilot-canvas-word\`), and renaming that
directory breaks the registration. Either leave local paths alone — they are just names — or
re-add the project through the app afterwards. Do not rename the directory and expect the app
to notice.

---

## 3. Directory structure

```
copilot-canvas/
├─ .github/
│  ├─ extensions/                    # C1: immediate subdirs only
│  │  ├─ word-canvas/
│  │  │  ├─ extension.mjs            # C1: entry, exact name
│  │  │  ├─ copilot-extension.json   # NEW — {name, version}; C4 requires it for gist install
│  │  │  ├─ src/
│  │  │  │  ├─ vendor/shared/        # GENERATED — see §3.1
│  │  │  │  └─ ui/
│  │  │  └─ test/
│  │  │     ├─ unit/                 # Word-free  → hosted runners
│  │  │     └─ integration/          # needs Word → self-hosted only
│  │  ├─ excel-canvas/               # later, out of scope per PLAN.md §14.3
│  │  └─ powerpoint-canvas/          # later
│  └─ workflows/
│     ├─ validate.yml                # every push/PR, ubuntu-latest
│     ├─ integration-windows.yml     # self-hosted + Office; nightly and manual
│     └─ release.yml                 # on tag; packages and attaches artifacts
├─ shared/                           # source of truth for cross-canvas code
├─ tools/
│  ├─ validate-extensions.mjs
│  ├─ sync-shared.mjs                # copies shared/ → each src/vendor/shared/
│  └─ package-extension.mjs
├─ docs/
└─ README.md
```

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
| **b** | `shared/` is the source of truth; `tools/sync-shared.mjs` copies it into each `src/vendor/shared/`, and CI fails if a copy is stale | **Recommended.** Keeps folders installable standalone with one source of truth. The generated copies are committed, which is the price of C3 |
| **c** | **One extension registering three canvases.** The SDK allows several `createCanvas` calls in one `joinSession` | Sidesteps sharing entirely. Real cost: users cannot install Word alone, and one process failure takes down all three canvases |

Option **b** is the recommendation, mainly because it is reversible: collapsing three vendored
folders into option **c** later is easy, whereas splitting a combined extension apart is not.
The decision does not have to be made now — Excel and PowerPoint are out of scope — but the
structure must not foreclose it, and this one does not.

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
- Verify `src/vendor/shared/` matches `shared/` (§3.1b).
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

- Tag format `word-canvas/v1.2.0`, so extensions version independently.
- Assert the tag version equals `copilot-extension.json`'s `version`.
- Produce, per extension: a folder zip, and a flattened gist-format bundle (`/` → `\`) when it
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

| | Step | Depends on |
| --- | --- | --- |
| 1 | Rename the repo; update remotes and README | — |
| 2 | Add `copilot-extension.json` to `word-canvas` | — |
| 3 | Split tests into `unit/` and `integration/` | — |
| 4 | `tools/validate-extensions.mjs` + `validate.yml` | 2, 3 |
| 5 | `tools/package-extension.mjs` + `release.yml` | 2 |
| 6 | Probe: does pdf.js need its binary font files? (§3.2) | — |
| 7 | Decide option **b** vs **c** for shared code (§3.1) | only when a second canvas is real |
| 8 | Self-hosted Windows+Office runner, licence permitting | licence check |

Steps 1–5 are a few hours and unblock everything else. Step 8 may never be worth it; steps 6
and 7 are decisions this document deliberately leaves open rather than guessing at.
