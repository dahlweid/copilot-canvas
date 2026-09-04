# Mutation check for the webview unit tests.
#
# Not part of the suite: a test that cannot fail is invisible to a green run, so
# each defect below is reintroduced, the tests are run, and the run must go red.
# Anything reported as SURVIVED is a test that proves nothing.
#
# ## Why this file exists at all
#
# `app.js` and `pdf-view.mjs` had no reachable assertion until #76, and the two
# test files that now cover them passed on their first run. A suite that goes
# green the moment it is written is exactly the suite whose liveness has not been
# measured -- the oldest rule here is that a probe on which every arm agrees has
# measured nothing. So every assertion family in `app.test.mjs` and
# `pdf-view.test.mjs` has a mutant below, and the harness itself is mutated too:
# a stub DOM that quietly swallows a wrong value would make the whole suite
# decorative.
#
# ## This file must keep its UTF-8 BOM
#
# It is run with `powershell -File`, which is Windows PowerShell 5.1, and 5.1
# decodes a BOM-less script as ANSI. Two anchors here contain a literal em dash
# (the tab title and the on-page badge), and under a mis-decode they would stop
# matching and be silently never applied -- a mutation gate reporting a kill it
# never made. `mutate-pdfjs.ps1` documents the family this comes from; the same
# self-check is repeated below rather than shared, because a check that has to be
# imported to work is one a copied script loses.
#
# Run from the extension root:
#   powershell -File test/unit/mutate-webview.ps1

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

# An em dash is one character decoded as UTF-8 and three decoded as ANSI. This
# has to come before any mutant runs: a mangled anchor still produces a green
# `KILLED` line for every mutant whose anchor happens to be ASCII, and a partial
# run is the thing hardest to notice.
if ('X—X'.Length -ne 3) {
    Write-Host "This script was decoded as ANSI, not UTF-8 -- its BOM is missing." -ForegroundColor Red
    Write-Host "Every non-ASCII anchor would silently fail to match. Restore the BOM." -ForegroundColor Red
    exit 1
}

$suite = @(
    'test/unit/app.test.mjs',
    'test/unit/pdf-view.test.mjs'
)

# Anchors are single lines on purpose. A multi-line anchor carries the line
# ending of *this* file, so it matches a CRLF checkout and silently misses an LF
# one -- the MISSING category would catch it, but only on the machine that has
# the other ending. Every mutant here is one line in, one line out.
$mutants = @(
    # --- app.js: the change banner -------------------------------------------

    # The state arrives over SSE, so every push carries a fresh object. Comparing
    # records by identity un-dismisses the change the reader just dismissed.
    @{ name = 'a dismissed change is remembered by object identity'
       file = 'src/ui/app.js'
       from = '        ? `${record.at}\u0000${record.op}\u0000${record.page}\u0000${record.text ?? ""}\u0000${record.span ?? ""}`'
       to   = '        ? record' }

    @{ name = '"Show me" is offered for a change with nowhere to go'
       file = 'src/ui/app.js'
       from = '    el.jumpToChange.hidden = !banner.jumpable;'
       to   = '    el.jumpToChange.hidden = banner.jumpable;' }

    @{ name = 'the banner is drawn but the page is not marked'
       file = 'src/ui/app.js'
       from = '    view.setChange(record);'
       to   = '    void record;' }

    # --- app.js: identity and the bar ----------------------------------------

    @{ name = 'the tab is named for the product, not the document'
       file = 'src/ui/app.js'
       from = '        document.title = `${doc.name} — ${PRODUCT}`;'
       to   = '        document.title = PRODUCT;' }

    @{ name = 'the empty state is never shown'
       file = 'src/ui/app.js'
       from = '    el.emptyState.classList.toggle("visible", showEmpty);'
       to   = '    el.emptyState.classList.toggle("visible", false);' }

    @{ name = 'the bar is shown with no document behind it'
       file = 'src/ui/app.js'
       from = '    el.bar.hidden = !state?.doc;'
       to   = '    el.bar.hidden = false;' }

    # --- app.js: copy --------------------------------------------------------

    # What is copied must be what the control's accessible name just promised.
    @{ name = 'copy takes the document name instead of the path it announced'
       file = 'src/ui/app.js'
       from = '    const path = el.copyPath.title;'
       to   = '    const path = el.docName.textContent;' }

    @{ name = 'a failed copy is reported as a copy'
       file = 'src/ui/app.js'
       from = '    const outcome = copyOutcome(failure);'
       to   = '    const outcome = copyOutcome(null);' }

    # --- app.js: the render cache -------------------------------------------

    @{ name = 'every state push tears down and re-parses the document'
       file = 'src/ui/app.js'
       from = '    if (!force && loadedPdfUrl === state.pdfUrl) {'
       to   = '    if (false) {' }

    @{ name = 'Reload asks for a re-render without forcing one'
       file = 'src/ui/app.js'
       from = '        await api("/api/refresh", { method: "POST", body: JSON.stringify({ force: true }) });'
       to   = '        await api("/api/refresh", { method: "POST", body: JSON.stringify({ force: false }) });' }

    # --- app.js: errors ------------------------------------------------------

    # An error body carries the host's own typed message; the status code alone
    # names no cause the reader can act on.
    @{ name = 'a request failure is reported by its status code'
       file = 'src/ui/app.js'
       from = '        const err = new Error(body?.error?.message ?? `Request failed (${res.status})`);'
       to   = '        const err = new Error(`Request failed (${res.status})`);' }

    @{ name = 'a host-side error state is not shown at all'
       file = 'src/ui/app.js'
       from = '    if (state?.status === "error" && state.error) {'
       to   = '    if (false) {' }

    # The defect this suite found. `render()` refuses to draw a document whose
    # refresh failed; `applyState` used to load it anyway on the first state a
    # panel saw, and the failure replaced the host's typed message.
    @{ name = 'applyState loads the render regardless of the error state'
       file = 'src/ui/app.js'
       from = '        // Deliberately no `showPdf` here. `render()` above has already loaded'
       to   = '        showPdf(currentPage);' }

    @{ name = 'the viewer is shown over a document that failed to open'
       file = 'src/ui/app.js'
       from = '    el.viewer.classList.toggle("visible", !showEmpty && hasDoc);'
       to   = '    el.viewer.classList.toggle("visible", !showEmpty);' }

    @{ name = 'a failed document is drawn anyway'
       file = 'src/ui/app.js'
       from = '    const hasDoc = Boolean(state?.doc && state.status !== "error");'
       to   = '    const hasDoc = Boolean(state?.doc);' }

    # --- app.js: outline and search ------------------------------------------

    @{ name = 'the outline reloads on every state push'
       file = 'src/ui/app.js'
       from = '    if (state?.doc && state.doc.key !== previousKey) {'
       to   = '    if (state?.doc) {' }

    @{ name = 'a new document keeps the previous document outline'
       file = 'src/ui/app.js'
       from = '    if (state?.doc && state.doc.key !== previousKey) {'
       to   = '    if (false) {' }

    @{ name = 're-opening the sidebar re-fetches the outline'
       file = 'src/ui/app.js'
       from = '    if (outlineLoadedFor === state.doc.key) return;'
       to   = '    if (false) return;' }

    @{ name = 'a document with no headings shows an outline heading'
       file = 'src/ui/app.js'
       from = '                textContent: headings.length ? "Outline" : "No headings in this document",'
       to   = '                textContent: "Outline",' }

    @{ name = 'Enter waits for the debounce like any other key'
       file = 'src/ui/app.js'
       from = '    if (event.key === "Enter") {'
       to   = '    if (false) {' }

    @{ name = 'a jump does not tell the host where the reader went'
       file = 'src/ui/app.js'
       from = '    api("/api/page", { method: "POST", body: JSON.stringify({ page: currentPage }) }).catch(() => {});'
       to   = '    void currentPage;' }

    # --- pdf-view.mjs: the lifetime of a document ----------------------------

    # pdf.js holds a worker-side handle per document. Leaking one per refresh --
    # and a refresh follows every edit -- accumulates for as long as the canvas
    # is open.
    @{ name = 'the previous document is dropped rather than destroyed'
       file = 'src/ui/pdf-view.mjs'
       from = '        if (doc) await doc.destroy().catch(() => {});'
       to   = '        if (doc) void doc;' }

    @{ name = 'the previous observer keeps watching removed pages'
       file = 'src/ui/pdf-view.mjs'
       from = '        this.#observer?.disconnect();'
       to   = '        void this.#observer;' }

    @{ name = 'the previous document pages are left in the container'
       file = 'src/ui/pdf-view.mjs'
       from = '        this.#container.replaceChildren();'
       to   = '        void this.#container;' }

    @{ name = 'page proxies are never cleaned up'
       file = 'src/ui/pdf-view.mjs'
       from = '            page.proxy?.cleanup();'
       to   = '            void page.proxy;' }

    # Two refreshes in quick succession: the first document arrives late and,
    # without the guard, overwrites the newer one and is never destroyed.
    @{ name = 'a load overtaken by a newer one installs itself anyway'
       file = 'src/ui/pdf-view.mjs'
       from = '        if (generation !== this.#generation) {'
       to   = '        if (false) {' }

    # --- pdf-view.mjs: layout and laziness -----------------------------------

    @{ name = 'pages are drawn at 1:1 regardless of the container width'
       file = 'src/ui/pdf-view.mjs'
       from = '        return clampScale(available / (mode === "height" ? unscaled.height : unscaled.width));'
       to   = '        return 1;' }

    @{ name = 'goToPage does not clamp to the document'
       file = 'src/ui/pdf-view.mjs'
       from = '        const page = this.#pages[Math.max(1, Math.min(number, this.#pages.length)) - 1];'
       to   = '        const page = this.#pages[number - 1];' }

    @{ name = 'every page is painted at load rather than when it is reached'
       file = 'src/ui/pdf-view.mjs'
       from = '        this.#observe();'
       to   = '        this.#observe(); for (const p of this.#pages) this.#renderPage(p).catch(() => {});' }

    @{ name = 'a page already painted is painted again on every intersection'
       file = 'src/ui/pdf-view.mjs'
       from = '        if (page.rendered || !page.proxy) return;'
       to   = '        if (!page.proxy) return;' }

    # A text layer over the page root rather than its own div sits on nothing,
    # and selection lands away from the glyphs.
    @{ name = 'the text layer is built over the page root'
       file = 'src/ui/pdf-view.mjs'
       from = '            container: page.textLayerDiv,'
       to   = '            container: page.root,' }

    @{ name = 'the visible page is never reported'
       file = 'src/ui/pdf-view.mjs'
       from = '        this.#onPageChange(this.#visiblePage());'
       to   = '        void this.#visiblePage();' }

    @{ name = 'the visible-page tolerance is dropped'
       file = 'src/ui/pdf-view.mjs'
       from = '            if (page.root.getBoundingClientRect().top - viewerTop <= 8) best = page.number;'
       to   = '            if (page.root.getBoundingClientRect().top - viewerTop <= 0) best = page.number;' }

    # --- pdf-view.mjs: the change overlay ------------------------------------

    # The lazy case is the normal one for a change below the fold: the record is
    # set while the page is an empty box with no text to search.
    @{ name = 'a page painted after the record arrived is never marked'
       file = 'src/ui/pdf-view.mjs'
       from = '        if (this.#change && candidatePages(this.#change).includes(page.number)) {'
       to   = '        if (false) {' }

    @{ name = 'the overlay box is not flipped out of PDF user space'
       file = 'src/ui/pdf-view.mjs'
       from = '        rects.push({ left: x, top: y - height, width, height });'
       to   = '        rects.push({ left: x, top: y, width, height });' }

    @{ name = 'the item scale is ignored when sizing the box'
       file = 'src/ui/pdf-view.mjs'
       from = '        const width = (item.width ?? 0) * this.#scale;'
       to   = '        const width = 1;' }

    # A page the plan named but whose quads came out empty is still the page the
    # text is on. Dropping it reinstates the miss the plan just avoided.
    @{ name = 'a page with no quads is dropped instead of marked'
       file = 'src/ui/pdf-view.mjs'
       from = '            this.#markPage(page, this.#change, mark.found ? this.#quadsFor(page, mark.found) : []);'
       to   = '            if (mark.found) this.#markPage(page, this.#change, this.#quadsFor(page, mark.found));' }

    @{ name = 'the badge does not distinguish a located change from a page-level one'
       file = 'src/ui/pdf-view.mjs'
       from = '        badge.textContent = quads.length ? describeChange(record) : `${describeChange(record)} — on this page`;'
       to   = '        badge.textContent = describeChange(record);' }

    @{ name = 'a dismissed change leaves its boxes on the page'
       file = 'src/ui/pdf-view.mjs'
       from = '            page.overlay.replaceChildren();'
       to   = '            void page.overlay;' }

    @{ name = 'a dismissed change leaves the page flagged as changed'
       file = 'src/ui/pdf-view.mjs'
       from = '            page.root.classList.remove("has-change");'
       to   = '            void page.root;' }

    # --- the harness itself ---------------------------------------------------
    #
    # A stub that swallows a wrong value would make every assertion above
    # decorative, so the two stand-ins that could plausibly do so are mutated
    # too. These are not product defects; they are the question "is this harness
    # capable of observing anything", asked in the only way that answers it.

    @{ name = 'HARNESS: the stub classList reports every class as present'
       file = 'test/unit/ui-dom.mjs'
       from = '        return this.#names.has(name);'
       to   = '        return true;' }

    @{ name = 'HARNESS: the stub matrix multiply returns the identity'
       file = 'test/unit/stub-pdfjs.mjs'
       from = '            m[0] * t[4] + m[2] * t[5] + m[4],'
       to   = '            t[4],' }
)

$survived = @()
$missing = @()

# A mutant is only informative against a suite that is otherwise green: a red
# baseline kills every mutant and the run reads as a clean sweep.
Push-Location $root
try {
    & node --test @suite *> $null
    if ($LASTEXITCODE -ne 0) {
        Write-Host 'Baseline is red. Fix the suite before mutating; a red baseline kills every mutant.' -ForegroundColor Red
        exit 1
    }

    foreach ($m in $mutants) {
        $file = Join-Path $root $m.file
        $original = [IO.File]::ReadAllText($file)
        if (-not $original.Contains($m.from)) {
            # A stale anchor is its own failure, not a survivor. Reported as one
            # it reads as "the tests do not defend this", when in fact nothing
            # was mutated at all.
            Write-Host ("MISSING  {0} -- anchor not found in {1}" -f $m.name, $m.file) -ForegroundColor Yellow
            $missing += $m.name
            continue
        }

        [IO.File]::WriteAllText($file, $original.Replace($m.from, $m.to))
        try {
            & node --test @suite *> $null
            $red = $LASTEXITCODE -ne 0
        } finally {
            [IO.File]::WriteAllText($file, $original)
        }

        if ($red) {
            Write-Host ("KILLED   {0}" -f $m.name) -ForegroundColor Green
        } else {
            Write-Host ("SURVIVED {0}" -f $m.name) -ForegroundColor Red
            $survived += $m.name
        }
    }
} finally {
    Pop-Location
}

Write-Host ''
Write-Host ("{0} of {1} mutants killed" -f ($mutants.Count - $survived.Count - $missing.Count), $mutants.Count)
if ($missing.Count -gt 0) {
    Write-Host 'Never applied (stale anchor -- these measured nothing):' -ForegroundColor Yellow
    $missing | ForEach-Object { Write-Host ("  - {0}" -f $_) -ForegroundColor Yellow }
}
if ($survived.Count -gt 0) {
    Write-Host 'Unkilled:' -ForegroundColor Red
    $survived | ForEach-Object { Write-Host ("  - {0}" -f $_) -ForegroundColor Red }
}
if ($survived.Count -gt 0 -or $missing.Count -gt 0) { exit 1 }
