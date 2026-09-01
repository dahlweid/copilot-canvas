# Probe: does ExportAsFixedFormat on a .pptx give one page per slide, and what
# does it cost?
#
# This is the gate question for ADR 0003. That ADR commits Word and PowerPoint
# to one extension and one pdf.js canvas on the reasoning that "both export to
# PDF and pagination is inherent to the format". Every number behind that claim
# was measured on Word. Nothing here has ever been measured.
#
# Three things have to be true for the claim to survive:
#
#   E1. Page count == slide count. Not "roughly"; exactly.
#   E2. Page geometry == slide geometry, so a page number addresses a slide and
#       the rendered page is the slide, not a handout or a notes page.
#   E3. The cost is in the same class as Word's 168 ms single page / 664 ms
#       whole document, otherwise the shared render cache is sized wrong.
#
# It also measures the single-slide re-export, which is the operation the
# optimistic-overlay edit loop depends on.

param([string]$Fixture)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '_common.ps1')

if (-not $Fixture) { $Fixture = Join-Path (Split-Path $PSScriptRoot -Parent) '.fixtures\deck.pptx' }
if (-not (Test-Path $Fixture)) { throw "fixture missing: $Fixture (run make-fixture.ps1 first)" }

$out = Join-Path $env:TEMP ("pptexport-" + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Force -Path $out | Out-Null
$src = Join-Path $out 'work.pptx'
Copy-Item $Fixture $src

# ppFixedFormatTypePDF = 2, ppFixedFormatIntentPrint = 2
# ppPrintOutputSlides  = 1, ppPrintAll = 1, ppPrintSlideRange = 4
# msoFalse = 0, ppPrintHandoutVerticalFirst = 2
$PDF = 2; $INTENT_PRINT = 2; $OUT_SLIDES = 1; $ALL = 1; $RANGE = 4; $NO = 0; $HANDOUT = 2

$ctx = $null
$pres = $null
try {
    Rep "POWERPNT pids before" ($(if (Get-PptPids) { (Get-PptPids) -join ',' } else { '(none)' }))
    # --- C1: cold start of a fresh instance ----------------------------------
    # The prior Word figure is omitted: it was a single first-export engine-load
    # measurement from a different spike, not a comparable repeated range.
    # Measure the equivalent: process creation through first PDF on disk.
    $coldSw = [Diagnostics.Stopwatch]::StartNew()
    $ctx = New-PowerPointInstance
    $app = $ctx.App
    $appMs = $coldSw.Elapsed.TotalMilliseconds
    Rep "new POWERPNT pids seen" ($(if ($ctx.NewPids) { $ctx.NewPids -join ',' } else { '(none appeared - attached)' }))
    Rep "COM instance creation" ("{0:F0} ms" -f $appMs)

    # The instance dying mid-probe is itself a finding, so report liveness at
    # every phase boundary rather than discovering it as an RPC error later.
    function Alive($phase) {
        $ok = $true
        foreach ($p in $ctx.NewPids) { if (-not (Get-Process -Id $p -ErrorAction SilentlyContinue)) { $ok = $false } }
        if (-not $ok) { Rep "  !! POWERPNT DIED after" $phase }
    }

    # Open(path, ReadOnly, Untitled, WithWindow) -- WithWindow msoFalse.
    $openSw = [Diagnostics.Stopwatch]::StartNew()
    $pres = $app.Presentations.Open($src, $NO, $NO, $NO)
    $openSw.Stop()
    Rep "Presentations.Open (cold)" ("{0:F0} ms" -f $openSw.Elapsed.TotalMilliseconds)

    $slides = $pres.Slides.Count
    $slideW = $pres.PageSetup.SlideWidth
    $slideH = $pres.PageSetup.SlideHeight
    Rep "slides in deck" $slides
    Rep "slide size (points)" ("{0} x {1}" -f $slideW, $slideH)

    function Export-Range($path, $from, $to) {
        # The PrintRange argument may not be $null or Missing -- late binding
        # throws NullReferenceException / ArgumentException for both. A real
        # Ranges object has to be supplied even for a whole-deck export.
        $pres.PrintOptions.Ranges.ClearAll()
        $r = $pres.PrintOptions.Ranges.Add($from, $to)
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $pres.ExportAsFixedFormat($path, $PDF, $INTENT_PRINT, $NO, $HANDOUT, $OUT_SLIDES, $NO,
            $r, $RANGE, "", $false, $false, $false, $true, $false)
        $sw.Stop()
        $sw.Elapsed.TotalMilliseconds
    }

    function Export-All($path) { Export-Range $path 1 $slides }
    function Export-Slide($path, $from, $to) { Export-Range $path $from $to }

    $firstMs = Export-All (Join-Path $out 'cold.pdf')
    $coldSw.Stop()
    Rep "first export (cold, engine load)" ("{0:F0} ms" -f $firstMs)
    Rep "COLD TOTAL (new instance -> pdf)" ("{0:F0} ms" -f $coldSw.Elapsed.TotalMilliseconds)

    # --- E1 + E2: is a page a slide? -----------------------------------------
    "== E1/E2: page count and page geometry =="
    $pages = Get-PdfPageCount (Join-Path $out 'cold.pdf')
    Rep "  pdf pages / deck slides" ("$pages / $slides")
    Rep "  ONE PAGE PER SLIDE" ($(if ($pages -eq $slides) { 'YES' } else { 'NO' }))

    $boxes = Get-PdfPageBoxes (Join-Path $out 'cold.pdf')
    if ($boxes) {
        $distinct = @($boxes | ForEach-Object { "{0:F2}x{1:F2}" -f $_.Width, $_.Height } | Sort-Object -Unique)
        Rep "  distinct /MediaBox sizes" ($distinct -join ' | ')
        $dw = [math]::Abs($boxes[0].Width - $slideW)
        $dh = [math]::Abs($boxes[0].Height - $slideH)
        Rep "  page vs slide delta (points)" ("{0:F3} x {1:F3}" -f $dw, $dh)
        Rep "  PAGE-ACCURATE" ($(if ($dw -lt 1 -and $dh -lt 1) { 'YES' } else { 'NO' }))
    }
    else { Rep "  /MediaBox" "not readable - page geometry UNVERIFIED" }

    # --- E3: warm cost --------------------------------------------------------
    "== E3: warm export cost =="
    $full = @(); for ($i = 0; $i -lt 3; $i++) { $full += Export-All (Join-Path $out "full$i.pdf") }
    Rep "  whole deck export (mean)" ("{0:F0} ms" -f (($full | Measure-Object -Average).Average))
    Rep "  whole deck export (min)" ("{0:F0} ms" -f (($full | Measure-Object -Minimum).Minimum))
    Rep "  whole deck pdf size" ("{0:N0} bytes" -f (Get-Item (Join-Path $out 'full0.pdf')).Length)
    Rep "  per slide, amortised over deck" ("{0:F0} ms" -f ((($full | Measure-Object -Average).Average) / $slides))

    $one = @(); for ($i = 0; $i -lt 5; $i++) { $one += Export-Slide (Join-Path $out "s1_$i.pdf") 1 1 }
    Rep "  single slide 1 (mean)" ("{0:F0} ms" -f (($one | Measure-Object -Average).Average))
    $p1 = Get-PdfPageCount (Join-Path $out 's1_0.pdf')
    Rep "  single slide 1 pdf pages" ("$p1 (expect 1)")

    if ($slides -ge 7) {
        $mid = @(); for ($i = 0; $i -lt 5; $i++) { $mid += Export-Slide (Join-Path $out "s7_$i.pdf") 7 7 }
        Rep "  single slide 7 (mean)" ("{0:F0} ms" -f (($mid | Measure-Object -Average).Average))
        Rep "  flat in slide position?" ($(if ([math]::Abs((($mid | Measure-Object -Average).Average) - (($one | Measure-Object -Average).Average)) -lt 100) { 'YES' } else { 'NO' }))
    }

    # --- E4: the edit loop ----------------------------------------------------
    "== E4: edit + single-slide re-export =="
    $loop = @()
    for ($i = 0; $i -lt 5; $i++) {
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $null = $pres.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.InsertAfter("x")
        $pres.PrintOptions.Ranges.ClearAll()
        $r = $pres.PrintOptions.Ranges.Add(1, 1)
        $pres.ExportAsFixedFormat((Join-Path $out "edit$i.pdf"), $PDF, $INTENT_PRINT, $NO, $HANDOUT,
            $OUT_SLIDES, $NO, $r, $RANGE, "", $false, $false, $false, $true, $false)
        $sw.Stop(); $loop += $sw.Elapsed.TotalMilliseconds
    }
    Rep "  edit + 1-slide re-export (mean)" ("{0:F0} ms" -f (($loop | Measure-Object -Average).Average))
    Rep "  edit + 1-slide re-export (min)" ("{0:F0} ms" -f (($loop | Measure-Object -Minimum).Minimum))
    Alive "edit loop"

    # --- E5: notes pages, the thing that is NOT a slide -----------------------
    # ppPrintOutputNotesPages = 5 is the nearest neighbour to a slide export.
    # It lives in probe-notes-control.ps1, which runs the A/B that isolates it as
    # a cause of instance death; running it here would poison the numbers above.

    $pres.Saved = -1
    $pres.Close()
    $pres = $null
    Alive "Presentation.Close"
}
catch { Rep "ERROR" $_.Exception.Message }
finally {
    # Ours, opened from our own temp root. Closing it here stops a failure above
    # leaving our deck open in an instance we may only have attached to.
    try { if ($pres) { $pres.Saved = -1; $pres.Close() } } catch { }
    $pres = $null
    Close-PowerPointInstance $ctx
    Remove-Item $out -Recurse -Force -ErrorAction SilentlyContinue
    Rep "POWERPNT pids after cleanup" ($(if (Get-PptPids) { (Get-PptPids) -join ',' } else { '(none)' }))
}
