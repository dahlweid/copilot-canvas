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
# Kills nothing. The pid it once called "minted" was differenced (#136).

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot '_common.ps1')

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
        # Deliberately no ReleaseComObject here, and this was raised in review.
        # Two measured reasons, and a third thing that is NOT one:
        #   1. `Quit()` (no-arg) binds under PS 5.1 and Word does exit -- #34's
        #      control arm measured exactly this: the bare form is reaped after
        #      the owning process exits, and it was the *swallowed* `Quit(0)`,
        #      which throws and never binds, that leaked.
        #   2. The RCW does not outlive this process; the script breaks and ends
        #      within milliseconds, and `Quit()` returns 3-28 ms before Word
        #      goes, so the release would not be what ends the instance anyway.
        # The decline stands on 1 and 2 alone.
        #
        # It previously cited a third reason -- PR #16 round 5 measuring
        # ReleaseComObject *causing* the leak it was meant to prevent -- and
        # summarised it as "ReleaseComObject is not neutral". That is a real
        # measurement carried to a scope it never covered, which is this repo's
        # named recurring failure. Read the source at
        # word-host.ps1 `Cmd-Edit`: it measured `ReleaseComObject($window)` on a
        # *Protected View window* RCW, across *multiple* operations of a
        # long-lived host, and explicitly declines to pin the mechanism beyond
        # naming Protected View's uninstrumented second WINWORD as the likely
        # route. None of that is in play for an `Application` RCW in a script
        # that exits milliseconds later.
        #
        # The check that settles it is internal: `Stop-Word` in that same file
        # releases the Application RCW itself. If the #16 result generalised,
        # it would convict the host's own teardown -- so it does not generalise,
        # and it was never evidence about this line.
        # What this path genuinely cannot do is verify the outcome: it aborts
        # precisely because attribution failed, so there is no pid it is
        # entitled to poll or to kill. Reporting and stopping is the honest end.
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
        # This used to kill $pid_, describing it as "the one pid I minted". It
        # was not minted, it was DIFFERENCED, and the abort above narrows that
        # without repairing it: the exact-count guard catches the 2-for-1 case
        # probe-init-attribution.ps1 measured, but not the case where our own
        # pid misses the 400 ms window at the diff above while a stranger's
        # WINWORD appears inside it -- the diff is then 1, and it is the wrong
        # process. A census control measured 2 strangers' WINWORDs appearing in
        # a 40 s window with nothing launched, so that window is not empty in
        # practice. Report; do not kill (#136).
        "         pid $pid_ never exited -- NOT killed, see the note in the source"
        Write-CensusSurvivors @($pid_)
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
