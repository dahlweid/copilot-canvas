# probe-save-bisect.ps1
#
# Why this probe exists
# ---------------------
# An earlier probe of mine -- deleted, because what follows voided it -- reported
# that `Document.SaveAs2` hangs after a style, a
# table or an AutoFormat. Then its own *control* arm -- two paragraphs of plain
# text, no styling at all -- hung as well. A control that fails means the arms
# measured the harness, not the document, and every conclusion drawn from that
# run is void.
#
# It is named nowhere here on purpose: a probe that no longer ships is not
# evidence a reader of this branch can follow, which is the whole point of the
# retraction and is enforced by tools/check-citations.mjs.
#
# probe-saveas-apartment.ps1 had already saved a document in a real
# `powershell.exe` in 149 ms. Two probes, contradictory results, so the cause is
# one of the differences between them rather than anything about the document.
#
# The two harnesses differ in exactly two respects:
#
#   1. the deleted probe passed -RedirectStandardOutput / -RedirectStandardError
#      to Start-Process; probe-saveas-apartment.ps1 does not.
#   2. the deleted probe waited 45 s; probe-saveas-apartment.ps1 waits 90 s.
#
# and in one respect of the document itself:
#
#   3. the save-hang worker adds a second paragraph via Paragraphs.Add().
#
# (2) is not a footnote. "Hung" and "took longer than the deadline" are
# different findings with different fixes, and this repo has twice shipped an
# assertion whose result depended on how quiet the machine was. So this probe
# waits 180 s and *reports the elapsed time of the SaveAs2 call itself*, which
# turns the question from a verdict into a number.
#
# Arms: a 2x2 over redirection and document shape, so neither is confounded
# with the other.
#
#   A  one paragraph,  no redirection    (reproduces the known-good result)
#   B  one paragraph,  redirection
#   C  two paragraphs, no redirection
#   D  two paragraphs, redirection       (reproduces probe-save-hang's control)
#
# Values reach the worker as discrete argv elements, never interpolated into a
# command string. Cleanup kills only PIDs absent before the probe started --
# several WINWORD.EXE belonging to other sessions are alive on this machine and
# killing one destroys someone's unsaved work.

param(
    [string[]] $Only,
    [int] $DeadlineSeconds = 180
)

$ErrorActionPreference = 'Stop'

function Get-WordPids {
    @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }) | Sort-Object
}

$pidsBefore = Get-WordPids
Write-Host "WINWORD before: $($pidsBefore.Count) alive -- $($pidsBefore -join ', ')"

$root = Join-Path $env:TEMP ("bi-" + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $root | Out-Null
$worker = Join-Path $root 'worker.ps1'

@'
param(
    [Parameter(Mandatory = $true)][string] $Arm,
    [Parameter(Mandatory = $true)][ValidateSet('1', '2')][string] $Paragraphs,
    [Parameter(Mandatory = $true)][string] $DocPath,
    [Parameter(Mandatory = $true)][string] $Trace
)

# $DocPath, not $Doc: a [string] param constraint follows the variable for the
# whole scope and PowerShell names are case-insensitive, so `$doc = ...Add()`
# would be silently coerced to a string.

$ErrorActionPreference = 'Stop'
$sw = [Diagnostics.Stopwatch]::StartNew()
function Step([string] $what) {
    Add-Content -LiteralPath $Trace -Value ("{0,7}ms  {1}" -f $sw.ElapsedMilliseconds, $what)
}

$WD_DO_NOT_SAVE = 0
$WD_FORMAT_DOCX = 16
$WD_CHARACTER   = 1

function Set-ParagraphText($para, [string] $text) {
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
    Step 'Documents.Add() returned'

    Set-ParagraphText $document.Paragraphs.Item(1) 'First paragraph.'
    if ($Paragraphs -eq '2') {
        $document.Paragraphs.Add() | Out-Null
        Set-ParagraphText $document.Paragraphs.Item(2) 'Second paragraph.'
    }
    Step "wrote $Paragraphs paragraph(s); document now has $($document.Paragraphs.Count)"

    # Time the call itself. A deadline gives a verdict; a number gives a finding.
    Step 'calling SaveAs2'
    $t0 = $sw.ElapsedMilliseconds
    $document.SaveAs2($DocPath, $WD_FORMAT_DOCX)
    $saveMs = $sw.ElapsedMilliseconds - $t0
    Step "SaveAs2 returned in ${saveMs}ms; file exists = $(Test-Path -LiteralPath $DocPath)"

    $t1 = $sw.ElapsedMilliseconds
    $document.Close($WD_DO_NOT_SAVE)
    Step "Close returned in $($sw.ElapsedMilliseconds - $t1)ms"
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

function Invoke-Arm([string] $arm, [string] $paragraphs, [bool] $redirect, [string] $title) {
    if ($Only -and $Only -notcontains $arm) { return }
    $trace = Join-Path $root "$arm.trace"
    $docPath = Join-Path $root "$arm.docx"
    New-Item -ItemType File -Path $trace | Out-Null

    Write-Host ""
    Write-Host "=== Arm ${arm}: $title ==="

    $argv = @(
        '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass'
        '-File', $worker
        '-Arm', $arm
        '-Paragraphs', $paragraphs
        '-DocPath', $docPath
        '-Trace', $trace
    )

    # The single variable under test in the left column of the 2x2. Everything
    # else about the two launches is identical.
    if ($redirect) {
        $out = Join-Path $root "$arm.out"
        $err = Join-Path $root "$arm.err"
        $p = Start-Process -FilePath 'powershell.exe' -PassThru -WindowStyle Hidden `
            -RedirectStandardOutput $out -RedirectStandardError $err -ArgumentList $argv
    } else {
        $p = Start-Process -FilePath 'powershell.exe' -PassThru -WindowStyle Hidden -ArgumentList $argv
    }

    $wall = [Diagnostics.Stopwatch]::StartNew()
    $deadline = (Get-Date).AddSeconds($DeadlineSeconds)
    while (-not $p.HasExited -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 250 }

    if (-not $p.HasExited) {
        Write-Host "  NO EXIT within ${DeadlineSeconds}s"
        try { Stop-Process -Id $p.Id -Force } catch { }
    } else {
        Write-Host "  worker exited $($p.ExitCode) after $($wall.ElapsedMilliseconds)ms wall"
    }

    Get-Content -LiteralPath $trace | ForEach-Object { Write-Host "  $_" }
    Write-Host "  file on disk: $(Test-Path -LiteralPath $docPath)"
}

Invoke-Arm 'A' '1' $false 'one paragraph,  no redirection  (known-good shape)'
Invoke-Arm 'B' '1' $true  'one paragraph,  redirection'
Invoke-Arm 'C' '2' $false 'two paragraphs, no redirection'
Invoke-Arm 'D' '2' $true  'two paragraphs, redirection     (save-hang control)'

$leaked = @(Get-WordPids | Where-Object { $pidsBefore -notcontains $_ })
Write-Host ""
if ($leaked.Count -gt 0) {
    Write-Host "cleaning up WINWORD started by this probe: $($leaked -join ', ')"
    foreach ($p in $leaked) { try { Stop-Process -Id $p -Force -ErrorAction SilentlyContinue } catch { } }
} else {
    Write-Host "no WINWORD left behind"
}

Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
