# The Word icon — findings

Written for issue #68. One probe, three questions.

## The questions

The issue asks for the real Word icon in three places: the canvas **tab**, the
document **name** in the bar, and the **Open in Word** button. The user added a
hard constraint — *the icon must not be a resource in the repo; load it at
runtime* — which the repo going public turns from a preference into a licensing
matter. That raises three questions the change cannot be made without:

1. **Is there an open-licensed Word mark?** If simple-icons carried one,
   vendoring it would be legitimate and the rest of this apparatus unnecessary.
2. **Does the user's own installed Word yield one, without launching Word?**
3. **Can the canvas tab carry an icon at all?** The SDK is said to accept
   `canvas.declaration.icon = "assets/icon.png"` — a *file path relative to the
   entrypoint*, which fights the constraint head-on. This was the one part of
   #68 that had never been measured.

## The probe

`probes/probe-icon-sources.mjs`, three arms. Each reports `skipped` rather than
`ok` when its prerequisite is missing, because an arm that could not run and
reports success is worse than no arm.

**No Word is launched.** Arm 2 counts `WINWORD.EXE` before and after, so "does
not launch Word" is measured rather than read out of the API documentation.

## What it measured

Run on Node v24.18.0, Windows, Copilot CLI 1.0.80.

```
ok      simple-icons has no Microsoft mark
        microsoft=404 microsoftword=404 microsoft-word=404 microsoftoffice=404 libreoffice=200
ok      the installed Word yields a PNG
        32x32, 2349 bytes, 3132 base64 chars, 1483ms; WINWORD processes 14 -> 14
ok      the canvas tab can carry an icon
        1.0.80: 17 source files; "icon" appears in [generated\rpc.d.ts:18,
        generated\session-events.d.ts:13]; declaration built from
        [id, displayName, description, inputSchema, actions], spread: false

3 arm(s), 0 failed, 0 skipped
```

1. **There is no open-licensed Word mark.** Four spellings of the Microsoft
   family all 404 while `libreoffice` returns 200 — so the instrument works and
   the family is genuinely absent, which is what a trademark removal looks like.
   Using LibreOffice Writer's mark for a Word document would name the wrong
   product, so that door is closed too.
2. **The installed Word yields one, and does not start.** 32×32, 2349 bytes,
   ~1.5 s. The `WINWORD.EXE` count is unchanged across the extraction:
   `ExtractAssociatedIcon` reads the executable's resources rather than
   automating the application. Nothing is redistributed — the bytes never enter
   the repository and come from the machine that displays them.
3. **The canvas tab cannot carry an icon from a Node extension.** This
   contradicts the premise in the issue, so it is worth being precise about what
   was measured rather than what was concluded:
   - The string `icon` occurs in the SDK **only in two generated `.d.ts` files**
     — `generated/rpc.d.ts` and `generated/session-events.d.ts`, which describe
     the *wire protocol*. It occurs in **none of the SDK's runtime JavaScript**.
   - `Canvas` builds `this.declaration` from an explicit five-field object
     literal — `id, displayName, description, inputSchema, actions` — **with no
     spread**. An `icon` passed to `createCanvas` is therefore dropped before the
     declaration exists, whatever its value: relative path, absolute path or data
     URI alike.
   - The wire *does* have `DiscoveredCanvas.icon` ("Host-local PNG path for the
     canvas icon, when supplied"), which is presumably why the capability looked
     available. Nothing in the Node SDK populates it.

   So the tab icon is not reachable by any value this extension could pass. The
   only route would be committing a file for the host to find — which is the one
   thing the constraint forbids, and which `tools/validate-extensions.mjs` counts
   besides.

## What shipped

Placements 2 and 3 — the document name and the Open in Word button — both from
`GET /api/word-icon`, extracted once per process and held in memory. Placement 1
does not ship, on the evidence above.

## What this is not

- **Not evidence about other Word layouts.** One machine, one installation
  (`Office16` under `Program Files`). `word-icon.ps1` consults the shell's own
  `App Paths` registration first, so a non-default directory should resolve, but
  that fallback is unexercised here: the first candidate matched.
- **Not evidence about future SDK versions.** The measurement is of the SDK
  bundled with CLI 1.0.80. `DiscoveredCanvas.icon` existing on the wire suggests
  the host end is already built, so a later SDK may well expose it. Re-run arm 3
  before concluding the tab is still closed.
- **Not a measurement of the browser end.** That the `<img>` reveals itself on
  load and stays hidden on error is measured in `test/unit/word-mark.test.mjs`
  against stub elements — the ordering is a property of the module. Nothing in
  this repo executes `src/ui/app.js`, which calls it; see
  `spikes/viewer-connection/FINDINGS.md` for why.
- **Not evidence about icon *quality*.** 32×32 is what
  `ExtractAssociatedIcon` returns; the markup renders it at 16px to match the
  drawn glyphs beside it. Whether a higher-resolution variant is extractable
  from the same executable was not investigated.
