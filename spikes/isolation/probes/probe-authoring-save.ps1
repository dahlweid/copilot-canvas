# probe-authoring-save.ps1
#
# What this settles
# -----------------
# Whether a document Word *authored in memory* -- headings, list styles, tables --
# can be written to disk with `Document.SaveAs2`, and what style identity comes
# back when the saved file is reopened.
#
# This matters because nothing in the shipping host has ever authored a document.
# `word-host.ps1` only ever calls `Save()` on a document it opened, so `SaveAs2`
# was entirely unmeasured ground before `create_document`.
#
# A retraction this probe exists to correct
# -----------------------------------------
# An earlier probe (`probe-save-hang.ps1`, deleted) reported that `SaveAs2` hangs
# indefinitely once a style is applied, a table is added, or `AutoFormat` is run,
# and that only plain text survives. That was **wrong**, and it was wrong in a way
# worth recording rather than quietly deleting: *its own control arm hung too*.
# Two paragraphs of plain text, no styling at all, hung at `SaveAs2` in 6 runs out
# of 6 -- so every arm agreed, and a probe whose arms all agree has separated
# nothing. The "styles hang the save" reading came from comparing hung arms
# against each other.
#
# The discriminator was a second worker that does the same COM work and saves in
# ~120 ms, run alternately with the first in the same directory, to the same file
# name, against the same live Word population:
#
#     good run1  winword_before=13  OK
#     bad  run2  winword_before=14  HANG
#     good run3  winword_before=14  OK
#     bad  run4  winword_before=15  HANG
#     good run5  winword_before=16  OK
#
# Perfectly alternating, so the cause is the worker script and not machine load,
# Word instance count, the file name, the directory, the paragraph text, or
# stdout redirection -- each of which was mutated separately and none of which
# changed the outcome. The remaining mechanism is **not identified**, and this
# comment does not guess at one: an unexplained probe artefact is what was
# measured, and asserting a cause here would repeat the mistake being corrected.
# The product conclusion does not depend on it, because the arms below establish
# the positive result directly.
#
# Arms
# ----
# Each arm builds a document, saves it, reopens it and reports what survived.
#
#   none    two paragraphs of plain text          (control)
#   style   + a Heading 1 via the numeric constant
#   list    + a bullet list style
#   table   + a 2x2 table
#   both    + a heading and a table together
#
# The control here can fail: if `SaveAs2` were genuinely fragile the `none` arm
# would hang like every other arm, and the run would be void again rather than
# silently reinterpreted.
#
# Styles are set by **numeric constant only**. Word's UI here is German;
# `Range.Style = 'Heading 1'` and the OOXML id `'berschrift1'` both throw
# `InvalidCastException` (measured in probe-authoring.ps1, arm R).
#
# Values reach the worker as discrete argv elements, never interpolated into a
# command string, so an apostrophe or an '&' in a path cannot reach a parser.
# Cleanup kills only PIDs absent before the probe started -- several WINWORD.EXE
# belonging to other sessions are alive on this machine and killing one destroys
# someone's unsaved work.

param(
    [string[]] $Only,
    [int] $DeadlineSeconds = 120
)

$ErrorActionPreference = 'Stop'

function Get-WordPids {
    @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }) | Sort-Object
}

$pidsBefore = Get-WordPids
Write-Host "WINWORD before: $($pidsBefore.Count) alive"

$root = Join-Path $env:TEMP ("au-" + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $root | Out-Null
$worker = Join-Path $root 'worker.ps1'

@'
param(
    [Parameter(Mandatory = $true)][string] $Feature,
    [Parameter(Mandatory = $true)][string] $DocPath,
    [Parameter(Mandatory = $true)][string] $Trace
)

# $DocPath, not $Doc: a [string] param constraint follows the variable for the
# whole scope and PowerShell names are case-insensitive, so a later
# `$doc = $app.Documents.Add()` is silently coerced to a string rather than
# rejected, and the failure surfaces lines later on a correct line.

$ErrorActionPreference = 'Stop'
$sw = [Diagnostics.Stopwatch]::StartNew()
function Step([string] $what) {
    Add-Content -LiteralPath $Trace -Value ("{0,7}ms  {1}" -f $sw.ElapsedMilliseconds, $what)
}

$WD_DO_NOT_SAVE  = 0
$WD_FORMAT_DOCX  = 16
$WD_CHARACTER    = 1
$WD_HEADING_1    = -2
$WD_LIST_BULLET  = -49

function Set-ParagraphText($para, [string] $text) {
    # The trim is derived from the position span, never from string length:
    # inside a table a cell paragraph's Range.Text ends with "`r" plus chr(7),
    # two characters in the string, but End - Start counts the end-of-cell mark
    # as one position.
    $r = $para.Range
    $visible = ([string]$r.Text).TrimEnd("`r", [char]7, "`n")
    $drop = ($r.End - $r.Start) - $visible.Length
    if ($drop -gt 0) { $r.MoveEnd($WD_CHARACTER, -$drop) | Out-Null }
    $r.Text = $text
}

$app = $null
try {
    $app = New-Object -ComObject Word.Application
    $app.Visible = $false
    $app.DisplayAlerts = 0
    Step 'Word ready'

    $document = $app.Documents.Add()
    Set-ParagraphText $document.Paragraphs.Item(1) 'First paragraph.'
    $document.Paragraphs.Add() | Out-Null
    Set-ParagraphText $document.Paragraphs.Item(2) 'Second paragraph.'
    Step 'wrote 2 paragraphs'

    switch ($Feature) {
        'style' {
            $document.Paragraphs.Item(1).Range.Style = $WD_HEADING_1
            Step "applied $WD_HEADING_1"
        }
        'list' {
            $document.Paragraphs.Item(2).Range.Style = $WD_LIST_BULLET
            Step "applied $WD_LIST_BULLET"
        }
        'table' {
            $document.Paragraphs.Add() | Out-Null
            $t = $document.Tables.Add($document.Paragraphs.Item(3).Range, 2, 2)
            Step "Tables.Add $($t.Rows.Count)x$($t.Columns.Count)"
        }
        'both' {
            $document.Paragraphs.Item(1).Range.Style = $WD_HEADING_1
            $document.Paragraphs.Add() | Out-Null
            $t = $document.Tables.Add($document.Paragraphs.Item(3).Range, 2, 2)
            Step 'applied heading and table'
        }
        default { Step 'plain text only (control)' }
    }

    Step 'calling SaveAs2'
    $t0 = $sw.ElapsedMilliseconds
    $document.SaveAs2($DocPath, $WD_FORMAT_DOCX)
    Step "SaveAs2 returned in $($sw.ElapsedMilliseconds - $t0)ms"

    $document.Close($WD_DO_NOT_SAVE)
    Step "Close returned; on disk = $((Get-Item -LiteralPath $DocPath).Length) bytes"

    # Reopen and report what identity the styles actually carry on disk. Style
    # ids are read verbatim and never constructed or compared -- Word mints them
    # from the *localized* style name with non-ASCII dropped.
    $reopened = $app.Documents.Open($DocPath, $false, $true)
    Step "reopened, $($reopened.Paragraphs.Count) paragraphs"
    for ($i = 1; $i -le $reopened.Paragraphs.Count; $i++) {
        $p = $reopened.Paragraphs.Item($i)
        $s = $p.Range.Style
        $text = ([string]$p.Range.Text).TrimEnd("`r", [char]7, "`n")
        Step ("  para {0}: outline={1} nameLocal='{2}' text='{3}'" -f $i, $p.OutlineLevel, $s.NameLocal, $text)
    }
    $reopened.Close($WD_DO_NOT_SAVE)
    Step 'reopened doc closed'
} catch {
    Step "ERROR $($_.Exception.GetType().FullName): $($_.Exception.Message)"
} finally {
    if ($null -ne $app) {
        try { $app.Quit($WD_DO_NOT_SAVE) } catch { }
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch { }
    }
    Step 'quit'
}
'@ | Set-Content -LiteralPath $worker -Encoding UTF8

function Invoke-Arm([string] $feature) {
    if ($Only -and $Only -notcontains $feature) { return }
    $trace = Join-Path $root "$feature.trace"
    $doc = Join-Path $root "$feature.docx"
    New-Item -ItemType File -Path $trace | Out-Null

    Write-Host ""
    Write-Host "=== $feature ==="

    $p = Start-Process -FilePath 'powershell.exe' -PassThru -WindowStyle Hidden -ArgumentList @(
        '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass'
        '-File', $worker
        '-Feature', $feature
        '-DocPath', $doc
        '-Trace', $trace
    )

    # Poll to a deadline. Word's teardown is load-dependent, so a flat sleep
    # would decide the outcome on how quiet the machine is rather than on
    # behaviour.
    $deadline = (Get-Date).AddSeconds($DeadlineSeconds)
    while (-not $p.HasExited -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 250 }
    if (-not $p.HasExited) {
        Write-Host "  NO EXIT within ${DeadlineSeconds}s"
        try { Stop-Process -Id $p.Id -Force } catch { }
    }

    Get-Content -LiteralPath $trace | ForEach-Object { Write-Host "  $_" }
}

Invoke-Arm 'none'
Invoke-Arm 'style'
Invoke-Arm 'list'
Invoke-Arm 'table'
Invoke-Arm 'both'

$leaked = @(Get-WordPids | Where-Object { $pidsBefore -notcontains $_ })
Write-Host ""
if ($leaked.Count -gt 0) {
    Write-Host "cleaning up WINWORD started by this probe: $($leaked -join ', ')"
    foreach ($leakedPid in $leaked) {
        try { Stop-Process -Id $leakedPid -Force -ErrorAction SilentlyContinue } catch { }
    }
} else {
    Write-Host "no WINWORD left behind"
}

Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
