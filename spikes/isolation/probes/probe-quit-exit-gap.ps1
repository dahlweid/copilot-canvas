# Measures the gap between Application.Quit() returning and WINWORD actually
# exiting, under whichever interpreter runs this. Exists because the repo
# carries two mutually inconsistent measured figures for that gap:
#   word-host.ps1:1148      "~120 ms"
#   document-editor.mjs:404 "seconds"
# and a full-suite A/B showed a fixed 300 ms wait was NOT enough (the fallback
# kill still fired), which falsifies ~120 ms on this machine.
#
# Attribution: census-diff at creation, and it ABORTS rather than guessing if
# the diff is not exactly one pid (concurrent creation misattributes -- #25).
# Kills only the pid it minted, and only if Quit failed to take it.

$ErrorActionPreference = 'Stop'
"interpreter : $($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition))"

function Get-WordPids { @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }) }

$baseline = Get-WordPids
"pre-existing WINWORD: $($baseline.Count) [$($baseline -join ',')]  (never touched)"
""

$reps = 3
foreach ($rep in 1..$reps) {
    $before = Get-WordPids
    $app = New-Object -ComObject Word.Application
    $app.Visible = $false
    $app.DisplayAlerts = 0

    # Give the process a moment to appear in the census before diffing.
    Start-Sleep -Milliseconds 400
    $after = Get-WordPids
    $new = @($after | Where-Object { $before -notcontains $_ })

    if ($new.Count -ne 1) {
        "rep $rep : UNATTRIBUTABLE -- census diff was $($new.Count) pids [$($new -join ',')]."
        "         Refusing to measure or kill. Aborting run."
        try { $app.Quit() } catch { }
        break
    }
    $pid_ = [int]$new[0]

    # The measurement: stopwatch spans Quit() returning -> process gone.
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $threw = $null
    try { $app.Quit() } catch { $threw = $_.Exception.Message }
    $returnedAt = $sw.ElapsedMilliseconds

    $exitedAt = $null
    while ($sw.ElapsedMilliseconds -lt 15000) {
        if (-not (Get-Process -Id $pid_ -ErrorAction SilentlyContinue)) { $exitedAt = $sw.ElapsedMilliseconds; break }
        Start-Sleep -Milliseconds 20
    }

    $gap = if ($null -ne $exitedAt) { "$($exitedAt - $returnedAt) ms" } else { "NEVER (>15000 ms)" }
    $thr = if ($threw) { "THREW: $threw" } else { "no throw" }
    "rep $rep : pid $pid_  Quit() returned at $returnedAt ms, $thr"
    "         process exited at $(if ($null -ne $exitedAt) { "$exitedAt ms" } else { 'never' })  ->  gap after return: $gap"
    "         would a fixed 300 ms wait have seen it exit? $(if ($null -ne $exitedAt -and $exitedAt -le 300) { 'YES' } else { 'NO' })"

    if ($null -eq $exitedAt) {
        "         killing the one pid I minted ($pid_)"
        Stop-Process -Id $pid_ -Force -ErrorAction SilentlyContinue
    }
    try { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch { }
    $app = $null
    Start-Sleep -Milliseconds 600
}

""
$final = Get-WordPids
"final WINWORD: $($final.Count) [$($final -join ',')]"
$leaked = @($final | Where-Object { $baseline -notcontains $_ })
"leaked by this probe: $($leaked.Count) [$($leaked -join ',')]"

# --- second measurement: what the old count-based wait actually budgeted -------
# Stop-Word used `for ($i=0; $i -lt 30; $i++) { Start-Sleep -Milliseconds 100; Get-Process ... }`.
# Nominally 3000 ms. The claim under test is that its real budget is materially
# larger, because Get-Process is not free -- i.e. that the wait worked on the
# incidental cost of an unrelated API rather than on a chosen number.
$sw = [Diagnostics.Stopwatch]::StartNew()
for ($i = 0; $i -lt 30; $i++) { Start-Sleep -Milliseconds 100; $null = Get-Process -Id $PID -ErrorAction SilentlyContinue }
$withGetProcess = $sw.ElapsedMilliseconds

$sw = [Diagnostics.Stopwatch]::StartNew()
for ($i = 0; $i -lt 30; $i++) { Start-Sleep -Milliseconds 100 }
$sleepOnly = $sw.ElapsedMilliseconds

""
"30x100ms loop, with Get-Process : $withGetProcess ms"
"30x100ms loop, sleep only       : $sleepOnly ms   (control)"
"Get-Process cost per call       : $([math]::Round(($withGetProcess - $sleepOnly) / 30, 1)) ms"
"nominal budget                  : 3000 ms"
