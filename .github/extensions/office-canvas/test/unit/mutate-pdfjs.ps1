# Mutation check for the pdf.js viewer and change-overlay unit tests.
#
# Not part of the suite: a test that cannot fail is invisible to a green run, so
# each defect below is reintroduced, the tests are run, and the run must go red.
# Anything reported as SURVIVED is a test that proves nothing.
#
# Where a property is defended in more than one place, the mutant removes *every*
# guard of it. A single-guard mutation here would "confirm" a test that was only
# half live -- which has already happened once in this repo.
#
# ## This file must keep its UTF-8 BOM
#
# It is run with `powershell -File`, which is Windows PowerShell 5.1, and 5.1
# decodes a BOM-less script as ANSI. Every anchor here was pure ASCII until an
# em dash appeared in one, at which point three mutants stopped matching their
# anchors and were silently never applied -- a mutation gate reporting a kill it
# never made, which is the third distinct mechanism for that this repo has found.
# The BOM makes 5.1 read the file as UTF-8 and the anchors match again.
#
# The `MISSING` category below is what caught it, and is why it exists: a mutant
# whose anchor is absent has measured nothing, and counting it as a survivor or
# as a kill would both be wrong.
#
# Run from the extension root:
#   powershell -File test/unit/mutate-pdfjs.ps1

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$repo = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $root))

# Does this script's own text survive being read? An em dash is one character
# when the file is decoded as UTF-8 and three when it is decoded as ANSI, so this
# fails loudly on the mis-decode that silently unhooked three mutants once. It
# has to come before any mutant runs: a mangled anchor still produces a green
# `KILLED` line for every mutant whose anchor happens to be ASCII, and a partial
# run is the thing hardest to notice.
if ("$([char]0x2014)".Length -ne 1 -or 'X—X'.Length -ne 3) {
    Write-Host "This script was decoded as ANSI, not UTF-8 -- its BOM is missing." -ForegroundColor Red
    Write-Host "Every non-ASCII anchor would silently fail to match. Restore the BOM." -ForegroundColor Red
    exit 1
}

$suite = @(
    'test/unit/vendor-pdfjs.test.mjs',
    'test/unit/worker-route.test.mjs',
    'test/unit/locate-text.test.mjs',
    'test/unit/change-record.test.mjs',
    'test/unit/viewer-state.test.mjs',
    'test/unit/vendor-checkout.test.mjs',
    'test/unit/ui-contract.test.mjs',
    'test/unit/change-plan.test.mjs',
    'test/unit/change-wording.test.mjs'
)

# `tools/` lives above the extension folder, so its mutants are anchored against
# the repository root rather than $root.
$mutants = @(
    # --- the split ------------------------------------------------------------

    # Both guards of "a boundary never lands mid-sequence" go at once: the
    # walk-back that avoids it and the throw that would notice it could not.
    @{ name = 'split boundary ignores continuation bytes'
       file = 'tools/vendor-pdfjs.mjs'; at = 'repo'
       from = '                while (end > start && (buffer[end] & 0xc0) === 0x80) end--;
                if (end === start) throw new Error("no safe UTF-8 boundary within one part");'
       to   = '                /* boundary taken as-is */' }

    @{ name = 'newline boundary preference dropped'
       file = 'tools/vendor-pdfjs.mjs'; at = 'repo'
       from = '            if (newline > start && end - newline < NEWLINE_SEARCH_WINDOW) {'
       to   = '            if (false) {' }

    @{ name = 'the text-layer contract is not checked against the bundle'
       file = 'tools/vendor-pdfjs.mjs'; at = 'repo'
       from = 'const TEXT_LAYER_CONTRACT = ["--font-height", "--scale-x", "--rotate", "--min-font-size", "--total-scale-factor"];'
       to   = 'const TEXT_LAYER_CONTRACT = [];' }

    # --- reassembly -----------------------------------------------------------

    @{ name = 'parts concatenated in filename order, not manifest order'
       file = '.github/extensions/office-canvas/src/vendor-assets.mjs'; at = 'repo'
       from = '    for (const part of manifest.worker.parts) {'
       to   = '    for (const part of [...manifest.worker.parts].sort((a, b) => String(a.name).localeCompare(String(b.name)))) {' }

    # The digest is the only guard that can see a permutation -- two parts are
    # both exactly 600,000 bytes, so every ordering totals the same. Removing it
    # must leave the swap test with nothing to notice.
    @{ name = 'reassembled worker not verified against its digest'
       file = '.github/extensions/office-canvas/src/vendor-assets.mjs'; at = 'repo'
       from = '    if (digest !== manifest.worker.sha256) {'
       to   = '    if (false) {' }

    @{ name = 'a missing part is skipped instead of failing the request'
       file = '.github/extensions/office-canvas/src/vendor-assets.mjs'; at = 'repo'
       from = '        } catch (err) {
            throw new VendorAssetError(
                "vendor_incomplete",'
       to   = '        } catch (err) {
            if (err) continue;
            throw new VendorAssetError(
                "vendor_incomplete",' }

    @{ name = 'an empty part list is accepted'
       file = '.github/extensions/office-canvas/src/vendor-assets.mjs'; at = 'repo'
       from = '    if (!Array.isArray(parts) || parts.length === 0) {'
       to   = '    if (false) {' }

    @{ name = 'a part name from the manifest is not basenamed'
       file = '.github/extensions/office-canvas/src/vendor-assets.mjs'; at = 'repo'
       from = '        const file = path.join(dir, path.basename(String(part.name)));'
       to   = '        const file = path.join(dir, String(part.name));' }

    # --- the route ------------------------------------------------------------

    @{ name = 'every request answered 304, whatever it asked for'
       file = '.github/extensions/office-canvas/src/server.mjs'; at = 'repo'
       from = '            if (req.headers["if-none-match"] === worker.etag) {'
       to   = '            if (true) {' }

    @{ name = 'the worker cached immutably at a versionless URL'
       file = '.github/extensions/office-canvas/src/server.mjs'; at = 'repo'
       from = '                "Cache-Control": "private, must-revalidate",
                ETag: worker.etag,'
       to   = '                "Cache-Control": "public, max-age=31536000, immutable",
                ETag: worker.etag,' }

    @{ name = 'any file in the vendor directory is served'
       file = '.github/extensions/office-canvas/src/server.mjs'; at = 'repo'
       from = '        if (!VENDOR_FILES.has(name)) {'
       to   = '        if (false) {' }

    @{ name = 'a vendored file the UI imports is renamed out of the allowlist'
       file = '.github/extensions/office-canvas/src/server.mjs'; at = 'repo'
       from = 'const VENDOR_FILES = new Set(["pdf.min.mjs", "pdf-text-layer.css"]);'
       to   = 'const VENDOR_FILES = new Set(["pdf.min.mjs", "pdf-text-layer-v2.css"]);' }

    # --- the locator ----------------------------------------------------------

    @{ name = 'a line break contributes nothing, so lines run together'
       file = '.github/extensions/office-canvas/src/ui/locate-text.mjs'; at = 'repo'
       from = '        if (item?.hasEOL) push(" ", index);'
       to   = '        /* line break dropped */' }

    @{ name = 'the owner map is not trimmed with the text'
       file = '.github/extensions/office-canvas/src/ui/locate-text.mjs'; at = 'repo'
       from = '        text = text.slice(0, -1);
        owners.pop();'
       to   = '        text = text.slice(0, -1);' }

    # Both guards of "never guess between duplicates": the exactly-one test and
    # the more-than-one refusal. Relaxing only the first would leave the second
    # still returning `ambiguous`, and the test would pass on a broken locator.
    @{ name = 'the first of several matches is picked'
       file = '.github/extensions/office-canvas/src/ui/locate-text.mjs'; at = 'repo'
       from = '    if (hits.length === 1) return { status: "located", matched: needle, range: rangeFor(hits[0], needle.length) };
    if (hits.length > 1) return { status: "ambiguous", occurrences: hits.length, range: null };'
       to   = '    if (hits.length >= 1) return { status: "located", matched: needle, range: rangeFor(hits[0], needle.length) };' }

    @{ name = 'the partial-match confidence floor is removed'
       file = '.github/extensions/office-canvas/src/ui/locate-text.mjs'; at = 'repo'
       from = '    for (let length = limit; length >= minPartialChars; length--) {'
       to   = '    for (let length = limit; length >= 1; length--) {' }

    @{ name = 'a partial match is allowed to float in the middle of the page'
       file = '.github/extensions/office-canvas/src/ui/locate-text.mjs'; at = 'repo'
       from = '        if (text.startsWith(needle.slice(needle.length - length))) {'
       to   = '        if (text.includes(needle.slice(needle.length - length))) {' }

    # The other half of the same property. Mutating one branch left the other
    # still refusing the fixture, and the mutant survived against a test that
    # looked like it covered both.
    @{ name = 'a partial head match is allowed to float in the middle of the page'
       file = '.github/extensions/office-canvas/src/ui/locate-text.mjs'; at = 'repo'
       from = '        if (text.endsWith(needle.slice(0, length))) {'
       to   = '        if (text.includes(needle.slice(0, length))) {' }

    @{ name = 'an empty target matches at the start of every page'
       file = '.github/extensions/office-canvas/src/ui/locate-text.mjs'; at = 'repo'
       from = '    if (!needle) return { status: "empty", range: null };'
       to   = '    if (false) return { status: "empty", range: null };' }

    @{ name = 'the locator stops trimming, so it disagrees with the structure map'
       file = '.github/extensions/office-canvas/src/ui/locate-text.mjs'; at = 'repo'
       from = '    return String(value ?? "")
        .replace(/\s+/gu, " ")
        .trim();'
       to   = '    return String(value ?? "").replace(/\s+/gu, " ");' }

    # --- the record -----------------------------------------------------------

    @{ name = 'a deletion is treated as locatable'
       file = '.github/extensions/office-canvas/src/change-record.mjs'; at = 'repo'
       from = 'const UNLOCATABLE_OPS = new Set(["delete_paragraph"]);'
       to   = 'const UNLOCATABLE_OPS = new Set();' }

    @{ name = 'an empty paragraph is treated as locatable'
       file = '.github/extensions/office-canvas/src/change-record.mjs'; at = 'repo'
       from = '    const locatable = !UNLOCATABLE_OPS.has(op) && text.length > 0;'
       to   = '    const locatable = !UNLOCATABLE_OPS.has(op);' }

    @{ name = 'an unknown page defaults to page 1'
       file = '.github/extensions/office-canvas/src/change-record.mjs'; at = 'repo'
       from = '    const page = Number.isFinite(result.page) && result.page > 0 ? Math.floor(result.page) : null;'
       to   = '    const page = Number.isFinite(result.page) && result.page > 0 ? Math.floor(result.page) : 1;' }

    @{ name = 'a result that applied nothing still produces a record'
       file = '.github/extensions/office-canvas/src/change-record.mjs'; at = 'repo'
       from = '    if (!result?.applied?.op) return null;'
       to   = '    if (false) return null;' }

    @{ name = 'the address is carried to the viewer'
       file = '.github/extensions/office-canvas/src/change-record.mjs'; at = 'repo'
       from = '    return {
        op,
        page,'
       to   = '    return {
        address: result.paragraph?.address ?? null,
        op,
        page,' }

    # The editor''s own description is agent-facing and spells out an address --
    # "replace the text of p:0df86a1a0e1e". It reached the change bar once.
    @{ name = "the editor's description is carried to the viewer"
       file = '.github/extensions/office-canvas/src/change-record.mjs'; at = 'repo'
       from = '    return {
        op,
        page,'
       to   = '    return {
        op,
        description: result.applied?.description ?? null,
        page,' }

    # --- the overlay lifetime -------------------------------------------------

    @{ name = 'the overlay outlives the render it describes'
       file = '.github/extensions/office-canvas/src/server.mjs'; at = 'repo'
       from = '        return this.change.docKey === this.doc.key ? this.change : null;'
       to   = '        return this.change;' }

    @{ name = 'opening a document leaves the previous overlay up'
       file = '.github/extensions/office-canvas/src/server.mjs'; at = 'repo'
       from = '            this.change = null;
            this.status = "ready";'
       to   = '            this.status = "ready";' }

    @{ name = 'an omitted change is treated as a request to clear'
       file = '.github/extensions/office-canvas/src/server.mjs'; at = 'repo'
       from = '        if (change !== undefined) this.change = change;'
       to   = '        this.change = change ?? null;' }

    @{ name = 'an explicit null is stamped into a record instead of staying cleared'
       file = '.github/extensions/office-canvas/src/server.mjs'; at = 'repo'
       from = '        if (this.change) this.change = { ...this.change, docKey };'
       to   = '        this.change = { ...this.change, docKey };' }

    @{ name = 'a record arriving during an in-flight refresh is swallowed'
       file = '.github/extensions/office-canvas/src/server.mjs'; at = 'repo'
       from = '            if (change !== undefined) {
                this.#stampChange(this.doc?.key ?? null);'
       to   = '            if (false) {
                this.#stampChange(this.doc?.key ?? null);' }

    @{ name = 'the record is never tied to the render it describes'
       file = '.github/extensions/office-canvas/src/server.mjs'; at = 'repo'
       from = '            if (change !== undefined) this.#stampChange(result.key);'
       to   = '            /* not stamped */' }

    # --- the change bar -------------------------------------------------------
    #
    # Both directions of the same contract, because the two files can drift
    # either way: markup renamed under a live lookup, and a lookup dropped while
    # the markup stays. One mutant would leave half of it unmeasured.

    @{ name = 'the markup renames an id the script still looks up'
       file = 'src/ui/index.html'
       from = 'id="jumpToChange"'
       to   = 'id="jumpToChangeX"' }

    @{ name = 'the script stops looking up an id the markup defines'
       file = 'src/ui/app.js'
       from = '    jumpToChange: $("jumpToChange"),'
       to   = '    jumpToChange: null,' }

    # Both spellings of the same regression, because they fail differently: the
    # rule deleted outright, and the rule kept but demoted to a priority that
    # loses to every component that sets its own display.
    @{ name = 'the [hidden] rule is dropped entirely'
       file = 'src/ui/app.css'
       from = '[hidden] {
    display: none !important;
}'
       to   = '/* no [hidden] rule */' }

    @{ name = 'the [hidden] rule loses its priority'
       file = 'src/ui/app.css'
       from = '    display: none !important;'
       to   = '    display: none;' }

    # --- which pages get marked -----------------------------------------------

    # The defect found in the running viewer: every candidate marked before
    # anything was searched, so the page before the change carried a badge with
    # no match behind it.
    @{ name = 'a page is marked on adjacency rather than on a match'
       file = 'src/ui/change-plan.mjs'
       from = '        if (found.status === "located" || found.status === "partial") {
            marks.push({ number: page.number, found });
        }'
       to   = '        marks.push({ number: page.number, found: found.status === "located" || found.status === "partial" ? found : null });' }

    # "Not found on a page that has not painted" is unknown, not false. Without
    # the guard the reported page is marked immediately and keeps its marker even
    # after the page that really holds the text paints.
    @{ name = 'an unpainted page counts as searched and not matching'
       file = 'src/ui/change-plan.mjs'
       from = '    if (searchable && reported) return [{ number: reported.number, found: null }];'
       to   = '    if (reported) return [{ number: reported.number, found: null }];' }

    @{ name = 'a deletion marks its neighbour too'
       file = 'src/ui/change-plan.mjs'
       from = '        return reported ? [{ number: reported.number, found: null }] : [];'
       to   = '        return [record.page - 1, record.page].map((n) => byNumber.get(n)).filter(Boolean).map((p) => ({ number: p.number, found: null }));' }

    @{ name = 'the page before the reported one is never searched'
       file = 'src/ui/change-plan.mjs'
       from = '    return [record.page - 1, record.page].filter((number) => number >= 1);'
       to   = '    return [record.page];' }

    # --- the banner, and the trigger that feeds it -----------------------------

    # Round 1 of review found both of these in the running viewer's logic.

    # The defect exactly as it shipped: the view answering "is this page worth
    # re-planning?" for itself, and getting a different answer than the planner.
    # Only the source check can see this -- `pdf-view.mjs` cannot be imported
    # under Node -- which is precisely why the source check exists.
    @{ name = 'the viewer decides for itself which painted pages matter'
       file = 'src/ui/pdf-view.mjs'
       from = '        if (this.#change && candidatePages(this.#change).includes(page.number)) {
            this.#applyChange();
        }'
       to   = '        if (this.#change?.page === page.number) this.#applyChange();' }

    # Both halves of the null-page handling go together. Removing only the text
    # branch would leave `jumpable` correct and vice versa, and either on its own
    # would "confirm" tests that were only half live.
    @{ name = 'a missing page is interpolated into the banner and still offered'
       file = 'src/ui/change-wording.mjs'
       from = '    if (page === null) {
        // Nothing is marked either -- with no page, the plan has no candidate --
        // so this must not promise a marker the reader could go looking for.
        return { text: `${phrase} — page not reported`, jumpable: false };
    }

    const text = record?.locatable'
       to   = '    const text = record?.locatable' }

    # The narrower half, so a test that only ever checks the wording is caught.
    @{ name = 'a change with no page is still offered as a destination'
       file = 'src/ui/change-wording.mjs'
       from = '        return { text: `${phrase} — page not reported`, jumpable: false };'
       to   = '        return { text: `${phrase} — page not reported`, jumpable: true };' }

    # And the reverse: the wording alone, with the offer left correct.
    @{ name = 'the unreported-page wording names a cause it never measured'
       file = 'src/ui/change-wording.mjs'
       from = '        return { text: `${phrase} — page not reported`, jumpable: false };'
       to   = '        return { text: `${phrase} — page unknown, Word could not report it`, jumpable: false };' }
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
        $base = if ($m.at -eq 'repo') { $repo } else { $root }
        $file = Join-Path $base $m.file
        $original = [IO.File]::ReadAllText($file)
        if (-not $original.Contains($m.from)) {
            # A stale anchor is its own failure, not a survivor. Reported as one it
            # reads as "the tests do not defend this", when in fact nothing was
            # mutated at all -- measured here: removing `description` from the
            # record left this mutant unappliable and it reported as unkilled.
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

# --- the checkout guard, which is not a source edit ---------------------------
#
# `.gitattributes` marking the vendored files `-text` is a guard like any other,
# but two things make the obvious mutant inert.
#
# Deleting the working-tree file does nothing on its own: git reads attributes
# from the index when the working copy is absent, so a committed `.gitattributes`
# keeps applying to a checkout that cannot see it. Measured -- the first version
# of this mutant moved the file aside, reported KILLED while it was still
# untracked, and silently went green the moment it was committed. The mutant must
# drop the index entry too.
#
# And removing the guard changes no bytes until git next writes those files, so
# both forms are run:
#
#   A  guard removed, working tree untouched -- the state a developer is in the
#      moment the line is deleted. Only the attribute test can see this.
#   B  guard removed and the files re-checked out -- the state a fresh clone on
#      Windows lands in, and the one that actually ships broken.
#
# Reporting only A would understate the guard; reporting only B would hide that
# the attribute test is the earlier detector. Measured here, A leaves the
# byte-comparison test green and B takes both red.
$attrFile = Join-Path $repo '.gitattributes'
$vendorRel = '.github/extensions/office-canvas/src/ui/vendor'
$checkoutMutants = @(
    @{ name = 'vendored files not marked -text (attribute only)'; recheckout = $false }
    @{ name = 'vendored files not marked -text (fresh checkout)'; recheckout = $true }
)

Push-Location $root
try {
    foreach ($m in $checkoutMutants) {
        Move-Item $attrFile "$attrFile.mutant" -Force
        & git -C $repo rm --cached --quiet .gitattributes
        try {
            if ($m.recheckout) {
                Remove-Item (Join-Path $repo "$vendorRel/*") -Force
                & git -C $repo checkout -- $vendorRel
            }
            & node --test @suite *> $null
            $red = $LASTEXITCODE -ne 0
        } finally {
            Move-Item "$attrFile.mutant" $attrFile -Force
            & git -C $repo add .gitattributes
            if ($m.recheckout) {
                Remove-Item (Join-Path $repo "$vendorRel/*") -Force
                & git -C $repo checkout -- $vendorRel
            }
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

$mutantCount = $mutants.Count + $checkoutMutants.Count

Write-Host ''
Write-Host ("{0} of {1} mutants killed" -f ($mutantCount - $survived.Count - $missing.Count), $mutantCount)
if ($missing.Count -gt 0) {
    Write-Host 'Never applied (stale anchor -- these measured nothing):' -ForegroundColor Yellow
    $missing | ForEach-Object { Write-Host ("  - {0}" -f $_) -ForegroundColor Yellow }
}
if ($survived.Count -gt 0) {
    Write-Host 'Unkilled:' -ForegroundColor Red
    $survived | ForEach-Object { Write-Host ("  - {0}" -f $_) -ForegroundColor Red }
}
if ($survived.Count -gt 0 -or $missing.Count -gt 0) { exit 1 }
