# Probe (A/B control): is ppPrintOutputNotesPages what kills the instance?
#
# probe-export.ps1 originally exported notes pages as a control at the end of its
# run and then failed on the very next COM call with "RPC server is unavailable"
# -- the POWERPNT process was gone. That could have been the notes export, or it
# could have been the five preceding edits, or repeated export in general.
#
# So: two runs that differ in exactly one call.
#
#   A (control) : open -> slides export -> slides export -> Close -> Quit
#   B (test)    : open -> slides export -> NOTES  export -> Close -> Quit
#
# If A closes cleanly and B does not, the notes export is the cause. Anything
# else and the crash is somewhere we have not looked yet.
#
# It also records the notes page geometry, which is the second half of the
# page-accuracy answer: a notes page is NOT the slide, so the canvas must pin
# OutputType to ppPrintOutputSlides rather than inherit the user's last print
# setting.

param([string]$Fixture, [int]$Repeats = 2)

$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot '_common.ps1')

if (-not $Fixture) { $Fixture = Join-Path (Split-Path $PSScriptRoot -Parent) '.fixtures\deck.pptx' }
if (-not (Test-Path $Fixture)) { throw "fixture missing: $Fixture (run make-fixture.ps1 first)" }

$PDF = 2; $INTENT_PRINT = 2; $OUT_SLIDES = 1; $OUT_NOTES = 5; $RANGE = 4; $NO = 0; $HANDOUT = 2

function Invoke-Arm([string]$Label, [int]$SecondOutputType) {
    $out = Join-Path $env:TEMP ("pptnotes-" + [guid]::NewGuid().ToString('N').Substring(0, 8))
    New-Item -ItemType Directory -Force -Path $out | Out-Null
    $src = Join-Path $out 'work.pptx'
    Copy-Item $Fixture $src

    $ctx = $null
    $pres = $null
    try {
        $ctx = New-PowerPointInstance
        $app = $ctx.App
        $pres = $app.Presentations.Open($src, $NO, $NO, $NO)
        $slides = $pres.Slides.Count

        function Export($path, $outputType) {
            $pres.PrintOptions.Ranges.ClearAll()
            $r = $pres.PrintOptions.Ranges.Add(1, $slides)
            $sw = [Diagnostics.Stopwatch]::StartNew()
            $pres.ExportAsFixedFormat($path, $PDF, $INTENT_PRINT, $NO, $HANDOUT, $outputType, $NO,
                $r, $RANGE, "", $false, $false, $false, $true, $false)
            $sw.Stop()
            $sw.Elapsed.TotalMilliseconds
        }

        $ms1 = Export (Join-Path $out 'one.pdf') $OUT_SLIDES
        Rep "  [$Label] slides export" ("{0:F0} ms, {1} pages" -f $ms1, (Get-PdfPageCount (Join-Path $out 'one.pdf')))

        $ms2 = Export (Join-Path $out 'two.pdf') $SecondOutputType
        $boxes = Get-PdfPageBoxes (Join-Path $out 'two.pdf')
        $geom = if ($boxes) { "{0:F0} x {1:F0} pt" -f $boxes[0].Width, $boxes[0].Height } else { 'unreadable' }
        Rep "  [$Label] second export (type $SecondOutputType)" ("{0:F0} ms, {1} pages, {2}" -f $ms2, (Get-PdfPageCount (Join-Path $out 'two.pdf')), $geom)

        Start-Sleep -Milliseconds 500
        $aliveAfterExport = $true
        foreach ($p in $ctx.NewPids) { if (-not (Get-Process -Id $p -ErrorAction SilentlyContinue)) { $aliveAfterExport = $false } }
        Rep "  [$Label] process alive after export" $aliveAfterExport

        try {
            $pres.Saved = -1
            $pres.Close()
            $pres = $null
            Rep "  [$Label] Presentation.Close" "OK"
        }
        catch {
            Rep "  [$Label] Presentation.Close" ("FAILED -> " + $_.Exception.Message.Split([char]10)[0])
        }
    }
    finally {
        # Ours, opened from our own temp root -- close it before releasing the
        # application, which may be one we merely attached to.
        try { if ($pres) { $pres.Saved = -1; $pres.Close() } } catch { }
        $pres = $null
        Close-PowerPointInstance $ctx
        Remove-Item $out -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Rep "POWERPNT pids before" ($(if (Get-PptPids) { (Get-PptPids) -join ',' } else { '(none)' }))

for ($i = 1; $i -le $Repeats; $i++) {
    "== run $i =="
    "  A (control): slides then slides"
    Invoke-Arm "A$i" $OUT_SLIDES
    "  B (test): slides then NOTES"
    Invoke-Arm "B$i" $OUT_NOTES
}

Rep "POWERPNT pids after" ($(if (Get-PptPids) { (Get-PptPids) -join ',' } else { '(none)' }))
