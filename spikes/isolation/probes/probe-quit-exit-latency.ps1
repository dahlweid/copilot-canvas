# probe-quit-exit-latency.ps1 -- how long does a Word we started actually take
# to go, once Quit() has returned?
#
# Why this exists. main's Stop-Word now polls for the owned process instead of
# sleeping a flat 300 ms, and the comment beside it says the corrected Quit()
# "exits on its own and this branch never fires". That poll is 30 iterations of
# 100 ms, so it gives up after 3 s and kills.
#
# Every probe in this repo that waits for the same event uses a 30 s deadline,
# with the comment "the process outlives it by seconds, longer under load".
# Those two numbers cannot both be right. If teardown routinely crosses 3 s then
# the poll expires, the kill fires, and the state the fix removed is silently
# restored -- with a comment asserting it does not happen. That is this repo's
# most-repeated defect class, so it is worth a measurement rather than an
# argument.
#
# THIS PROBE FIXES NOTHING and does not touch word-host.ps1. Stop-Word is owned
# elsewhere. It reports a number.
#
# Two conditions, because the interesting one is not the idle one:
#   A. one Word at a time, quit and timed  -- one teardown in flight.
#   B. three Words alive, quit together    -- closer to what the suite does, and
#      the reason "typical-case, not bounds" is a standing rule here.
#
# NEITHER CONDITION IS A QUIET MACHINE, and an earlier revision of this file
# called condition A "the quiet-machine case", which it had not earned: the
# census during those runs was 16 WINWORDs, most of them predating the session
# and none attributable to it. That label would have let the numbers below be
# quoted as clean-machine figures when they are biased toward a loaded idle.
# Condition A is one teardown *of ours* in flight, which is a different claim.
# So the census is now recorded at both ends of the run and printed beside the
# figures, because a number whose conditions travel with it cannot be quoted
# out of them.
#
# Only processes this probe started are ever touched, identified by PID
# differencing against a census taken first.
#
# OBSERVED RUNS, recorded here because the single number is not the finding.
# Both runs below are unmutated, on the same machine, with a census of 16 at
# both ends of each:
#
#   run 2   min 2563  max 3593  mean 3150   over budget:  8 of 11
#   run 4   min  779  max 3009  mean 2638   over budget:  1 of 11
#
# The exceedance rate swings eight-fold between two runs whose census is
# identical, so the census is NOT the discriminator and no proportion from a
# single run may be quoted as the rate. An earlier report of this probe said
# the kill branch "fires on most teardowns" on the strength of run 2 alone;
# that is a one-run proportion transported into a general claim and it does
# not survive run 4. What survives both runs, and is the actual finding, is
# narrower and enough: teardown crosses 3 s here, so the branch commented as
# never firing does fire. THE VARIANCE IS THE RESULT. A budget whose
# exceedance rate moves that far between two runs on one machine is not a
# budget anyone measured.
#
# WHAT THIS PROBE THEN RECOMMENDED, AND WHY THAT WAS WRONG.
# An earlier revision of this header concluded "...and that argues for a
# deadline far outside the distribution rather than for a slightly larger one
# near it." The measurement is unaffected; the recommendation is withdrawn.
#
# PR #36 measured the half this probe cannot see. dispose() sends `quit` with a
# 20 s *client* timeout, and expiry kills the host outright: a 10 s bound
# returns cleanly at 10279 ms, a 30 s bound times out at 20007 ms with the host
# killed mid-wait. So the repo's 30 s is not generous, it is UNREACHABLE -- the
# client kills first. Combined with the spread below and a Word observed
# surviving a 30 s poll, NO REACHABLE BOUND LIES OUTSIDE THE DISTRIBUTION.
# There is no number that turns this into a guarantee, which is what the
# withdrawn sentence assumed there was.
#
# The corrected shape is not a bigger number but a different kind of number: a
# budget whose EXPIRY IS SAFE, falling through to an identity-proving kill,
# rather than one asserted to be large enough that expiry cannot happen.
#
# The instructive part is why this probe could not have caught its own bad
# advice. It measures how long WINWORD takes to leave -- correctly, twice, with
# its conditions recorded. The recommendation ranged over the client-side
# timeout, a variable this instrument never observed and had no way to observe.
# Measuring soundly does not license recommending past the edge of what was
# measured, and nothing about a correct measurement flags when its prose has
# stepped over that edge. Two of the three corrections against this file are
# now of that shape rather than of the numbers.
#
# Run: powershell.exe -File spikes\isolation\probes\probe-quit-exit-latency.ps1

$ErrorActionPreference = 'Stop'

$POLL_BUDGET_MS = 3000   # what Stop-Word allows before it kills
$HARD_DEADLINE_MS = 60000

function Get-WordPids {
    @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
}

# Waits for one pid to leave, to a generous deadline. Returns ms, or -1.
function Wait-Exit([int]$wordPid, [int]$budgetMs) {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    while ($sw.ElapsedMilliseconds -lt $budgetMs) {
        if ($null -eq (Get-Process -Id $wordPid -ErrorAction SilentlyContinue)) {
            return [int]$sw.ElapsedMilliseconds
        }
        Start-Sleep -Milliseconds 25
    }
    return -1
}

$samples = New-Object System.Collections.ArrayList
$strays = New-Object System.Collections.ArrayList
$censusStart = (Get-WordPids).Count

function Start-OwnWord {
    $before = Get-WordPids
    $app = New-Object -ComObject Word.Application
    $app.Visible = $false
    $deadline = [Diagnostics.Stopwatch]::StartNew()
    $newPid = $null
    while ($deadline.ElapsedMilliseconds -lt 30000) {
        $diff = @(Get-WordPids | Where-Object { $before -notcontains $_ })
        if ($diff.Count -eq 1) { $newPid = $diff[0]; break }
        Start-Sleep -Milliseconds 50
    }
    if ($null -eq $newPid) {
        # Exactly one new pid, or none: three foreign WINWORDs are routinely
        # alive on this machine and another session starting one inside this
        # window makes the difference ambiguous. Refusing to guess is the only
        # safe branch -- mis-attributing here would end someone else's Word.
        #
        # But refusing is not enough on its own. `New-Object` already succeeded,
        # so an instance exists that this probe started and can no longer name,
        # and throwing here would leak it -- the finally block below only knows
        # pids. Quit through the COM object instead: we hold the reference, so
        # that is attributable by construction and needs no pid at all.
        try { $app.Quit() } catch { }
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch { }
        throw 'ambiguous WINWORD census -- another session started one inside the attribution window; quit our own instance and gave up rather than guess. Re-run.'
    }
    [void]$strays.Add($newPid)
    return [pscustomobject]@{ App = $app; Pid = $newPid }
}

function Record($condition, $wordPid, $quitMs, $exitMs) {
    [void]$samples.Add([pscustomobject]@{
        Condition = $condition
        Pid       = $wordPid
        QuitMs    = $quitMs
        ExitMs    = $exitMs
    })
    $verdict = if ($exitMs -lt 0) { 'NEVER (within deadline)' }
               elseif ($exitMs -gt $POLL_BUDGET_MS) { "OVER $POLL_BUDGET_MS ms -- Stop-Word would kill" }
               else { 'within budget' }
    Write-Host ("  {0}  pid {1,-6}  Quit() returned {2,5} ms   process gone {3,6} ms   {4}" -f `
        $condition, $wordPid, $quitMs, $exitMs, $verdict)
}

try {
    Write-Host "`n=== A. one at a time (one teardown of ours in flight) ==="
    for ($i = 1; $i -le 5; $i++) {
        $w = Start-OwnWord
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $w.App.Quit()
        $quitMs = [int]$sw.ElapsedMilliseconds
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($w.App) | Out-Null } catch { }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        Record 'A' $w.Pid $quitMs (Wait-Exit $w.Pid $HARD_DEADLINE_MS)
    }

    Write-Host "`n=== B. three alive, quit together (closer to the suite) ==="
    for ($round = 1; $round -le 2; $round++) {
        $ws = @(Start-OwnWord; Start-OwnWord; Start-OwnWord)
        $quits = @{}
        foreach ($w in $ws) {
            $sw = [Diagnostics.Stopwatch]::StartNew()
            $w.App.Quit()
            $quits[$w.Pid] = [int]$sw.ElapsedMilliseconds
            try { [Runtime.InteropServices.Marshal]::ReleaseComObject($w.App) | Out-Null } catch { }
        }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()

        # All three are polled against ONE origin taken after the last Quit.
        #
        # The obvious loop -- Wait-Exit each pid in turn -- does not measure
        # this. The first call blocks for seconds, by which time the other two
        # have already gone, so they score ~1 ms and the run reports a load
        # condition that is faster than the idle one. That is not a result, it
        # is the harness timing its own waiting. Measured before this was
        # fixed: 1 ms, 1 ms and 280 ms, none of which is a teardown latency.
        $origin = [Diagnostics.Stopwatch]::StartNew()
        $pending = @{}
        foreach ($w in $ws) { $pending[$w.Pid] = $true }
        $exitAt = @{}
        while ($pending.Count -gt 0 -and $origin.ElapsedMilliseconds -lt $HARD_DEADLINE_MS) {
            foreach ($key in @($pending.Keys)) {
                if ($null -eq (Get-Process -Id $key -ErrorAction SilentlyContinue)) {
                    $exitAt[$key] = [int]$origin.ElapsedMilliseconds
                    $pending.Remove($key)
                }
            }
            if ($pending.Count -gt 0) { Start-Sleep -Milliseconds 25 }
        }
        foreach ($w in $ws) {
            $exitMs = if ($exitAt.ContainsKey($w.Pid)) { $exitAt[$w.Pid] } else { -1 }
            Record 'B' $w.Pid $quits[$w.Pid] $exitMs
        }
    }
} finally {
    # Only pids this probe attributed to itself.
    foreach ($strayPid in $strays) {
        $p = Get-Process -Id $strayPid -ErrorAction SilentlyContinue
        if ($null -ne $p -and $p.ProcessName -eq 'WINWORD') {
            Write-Host "  cleanup: killing own pid $strayPid"
            try { $p.Kill() } catch { }
        }
    }
}

$exits = @($samples | Where-Object { $_.ExitMs -ge 0 } | Select-Object -ExpandProperty ExitMs)
$over = @($samples | Where-Object { $_.ExitMs -lt 0 -or $_.ExitMs -gt $POLL_BUDGET_MS })

Write-Host "`n=== summary ==="
# Informational, deliberately NOT asserted equal. Another session may start or
# stop a Word mid-run, so census-end != census-start has causes this probe
# cannot distinguish. The leak check that CAN fail is the finally block above,
# which only ever looks at pids this probe attributed to itself.
$censusEnd = (Get-WordPids).Count
Write-Host ("WINWORD census: {0} at start, {1} at end -- these are NOT clean-machine" -f $censusStart, $censusEnd)
Write-Host  "figures and must not be quoted as such. Foreign instances are routine"
Write-Host  "here and are not attributable, so the figures below are biased toward a"
Write-Host  "loaded idle. That biases them in the direction that WEAKENS the finding"
Write-Host  "if you want a small budget, and strengthens it if you want a large one."
if ($exits.Count -gt 0) {
    $stats = $exits | Measure-Object -Minimum -Maximum -Average
    Write-Host ("exit after Quit(): min {0} ms, max {1} ms, mean {2} ms, n={3}" -f `
        [int]$stats.Minimum, [int]$stats.Maximum, [int]$stats.Average, $exits.Count)
}
Write-Host ("samples exceeding Stop-Word's {0} ms poll: {1} of {2}" -f `
    $POLL_BUDGET_MS, $over.Count, $samples.Count)

if ($over.Count -gt 0) {
    Write-Host ("`nVERDICT: the {0} ms poll is NOT a bound on this machine." -f $POLL_BUDGET_MS)
    Write-Host "Those samples would have been killed, not quit. The comment beside"
    Write-Host "the poll says the kill branch never fires; it fired here. That is"
    Write-Host "all this shows -- it does NOT show what the comment's author"
    Write-Host "observed, and nothing here measured the machine they measured on."
    Write-Host "Do not quote this run's proportion as the rate: see the run log in"
    Write-Host "the header, where it swings 8x at an identical census."
    Write-Host "Reported, not fixed: Stop-Word is owned elsewhere."
    exit 1
}

Write-Host ("`nVERDICT: every sample fit inside the {0} ms poll on this run." -f $POLL_BUDGET_MS)
Write-Host "That is a typical-case result and not a bound -- this probe cannot"
Write-Host "show a poll is sufficient, only that it was here, now."
exit 0
