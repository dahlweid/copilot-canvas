# Can we attribute a WINWORD.EXE to *our* COM instance soundly?
#
# Today both `Initialize-Word` (shipping) and `word-pids.mjs` (tests) attribute
# by PID-set differencing: snapshot before, create, and treat whatever is new as
# ours. That is unsound whenever anything else starts a Word inside the window,
# and the window is not small -- a cold `New-Object Word.Application` is ~4.5 s
# here, plus up to 1.5 s of polling afterwards. The consequence is not cosmetic:
# `Stop-Word` calls `$p.Kill()` on the pid it picked, and it picks `$new[0]`.
#
# Arms:
#   A  differencing under a *deliberately* concurrent start -- does the set ever
#      contain a pid that is not ours, and can it sort first?
#   B  caption tagging -- set `$App.Caption` to a GUID on a hidden instance and
#      look for it in the process list. If the hidden OpusApp window carries the
#      title, attribution becomes a fact we minted rather than an inference.
#
# Arm B is the one that matters: it either yields a sound mechanism or rules it
# out. Arm A only sizes a defect we already know the shape of.

$ErrorActionPreference = 'Stop'

function Get-WordPids { @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id) }

$created = @()

function New-TaggedWord([string]$tag) {
    $app = New-Object -ComObject Word.Application
    $app.Visible = $false
    try { $app.Caption = $tag } catch { Write-Host "    Caption set failed: $($_.Exception.Message)" }
    return $app
}

Write-Host "=== baseline ==="
$baseline = Get-WordPids
Write-Host "WINWORD alive before: $($baseline.Count)"

# Everything that creates a Word runs inside this `try`, so that cleanup and the
# census below run even when an arm throws. `$ErrorActionPreference = 'Stop'` is
# right for a probe -- an arm that fails must not be reported as a measurement --
# but combined with straight-line cleanup it means any error anywhere skips the
# teardown entirely. That is not hypothetical: the first run of this file's own
# leak fix aborted on `New-Item -LiteralPath`, which Windows PowerShell 5.1 does
# not accept, and leaked the four instances the fix was written to prevent.
$foreign = $null
$ready = Join-Path ([IO.Path]::GetTempPath()) ("word-probe-ready-" + [guid]::NewGuid().ToString('N'))
$release = Join-Path ([IO.Path]::GetTempPath()) ("word-probe-release-" + [guid]::NewGuid().ToString('N'))
try {

# --- Arm B: caption tagging --------------------------------------------------
Write-Host ""
Write-Host "=== Arm B: is a minted caption readable from the process list? ==="

$tagA = "word-canvas-probe-" + [guid]::NewGuid().ToString('N')
$t0 = Get-Date
$appA = New-TaggedWord $tagA
$created += , $appA
$startMs = [int]((Get-Date) - $t0).TotalMilliseconds
Write-Host "instance A created in ${startMs}ms, caption='$tagA'"

# Windows may not publish the title immediately; poll briefly like the shipping
# code polls for the pid, so the comparison is like-for-like.
$match = @()
for ($i = 0; $i -lt 20; $i++) {
    $match = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowTitle -like "*$tagA*" })
    if ($match.Count -gt 0) { break }
    Start-Sleep -Milliseconds 100
}

if ($match.Count -eq 1) {
    Write-Host "  RESULT: caption FOUND on pid $($match[0].Id) -- sound attribution is available"
} elseif ($match.Count -gt 1) {
    Write-Host "  RESULT: caption matched $($match.Count) processes -- ambiguous, NOT sound"
} else {
    Write-Host "  RESULT: caption NOT visible in the process list"
    Write-Host "  (titles seen: " + (@(Get-Process -Name WINWORD -ErrorAction SilentlyContinue |
            ForEach-Object { "'" + $_.MainWindowTitle + "'" }) -join ', ') + ")"
}

# Does a *second* hidden instance carry its own distinct caption? If both land
# on the same title the mechanism cannot separate two of our own instances.
$tagB = "word-canvas-probe-" + [guid]::NewGuid().ToString('N')
$appB = New-TaggedWord $tagB
$created += , $appB
Start-Sleep -Milliseconds 500
$matchB = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue |
    Where-Object { $_.MainWindowTitle -like "*$tagB*" })
$matchA2 = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue |
    Where-Object { $_.MainWindowTitle -like "*$tagA*" })
Write-Host "  after a second instance: A matches $($matchA2.Count), B matches $($matchB.Count)"
if ($matchA2.Count -eq 1 -and $matchB.Count -eq 1 -and $matchA2[0].Id -ne $matchB[0].Id) {
    Write-Host "  RESULT: two instances separate cleanly by caption"
}

# --- Arm A: how wide is the differencing window, and what gets caught in it? --
Write-Host ""
Write-Host "=== Arm A: differencing under a concurrent start ==="

$before = Get-WordPids
# A foreign Word, started by a *different process*, standing in for another
# session. Word is multi-instance, so this genuinely is a separate WINWORD.
#
# The helper owns its Word's whole lifetime and quits it itself. An earlier
# version slept a fixed 25 s and was force-killed from here, which orphans the
# COM server: killing the client does not quit Word, and a hidden WINWORD with
# no documents and no live client stays up indefinitely. That leaked four
# instances into a machine this repo measures Word start-up on, where the cost
# is not hypothetical -- `Documents.Add` was measured going ~700 ms to 9-11 s
# as the population grew. A probe about attributing Word processes must not
# manufacture unattributable ones.
$foreign = Start-Process powershell.exe -PassThru -WindowStyle Hidden -ArgumentList @(
    '-NoProfile', '-NonInteractive', '-Command',
    "`$w = New-Object -ComObject Word.Application; `$w.Visible = `$false; " +
    "Set-Content -LiteralPath '$ready' -Value `$PID; " +
    "for (`$i = 0; `$i -lt 600; `$i++) { if (Test-Path -LiteralPath '$release') { break }; Start-Sleep -Milliseconds 100 }; " +
    "try { `$w.Quit() } catch { }; " +
    "try { [Runtime.InteropServices.Marshal]::ReleaseComObject(`$w) | Out-Null } catch { }; " +
    "[System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()"
)

# Wait for the helper to signal that its Word exists, rather than sleeping a
# fixed interval and hoping. Cold `New-Object Word.Application` measured 4029 ms
# here, so the 1200 ms this used to sleep was usually a false ready: arm A would
# then difference against a foreign Word that had not started yet and report
# "the foreign Word did not land inside the window" -- a null result produced by
# the harness rather than by the mechanism under test.
$foreignReady = $false
for ($i = 0; $i -lt 300; $i++) {
    if (Test-Path -LiteralPath $ready) { $foreignReady = $true; break }
    if ($foreign.HasExited) { break }
    Start-Sleep -Milliseconds 100
}
if ($foreignReady) {
    Write-Host "foreign Word ready after $([int]($i * 100))ms"
} else {
    Write-Host "  WARNING: the foreign Word never signalled ready -- arm A measures nothing below"
}

$tagC = "word-canvas-probe-" + [guid]::NewGuid().ToString('N')
$appC = New-TaggedWord $tagC
$created += , $appC
Start-Sleep -Milliseconds 800

$new = @(Get-WordPids | Where-Object { $before -notcontains $_ })
$mineByCaption = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue |
    Where-Object { $_.MainWindowTitle -like "*$tagC*" })

Write-Host "differencing says new pids: $($new -join ', ')  (count $($new.Count))"
if ($mineByCaption.Count -eq 1) {
    $mine = $mineByCaption[0].Id
    Write-Host "caption says ours is:      $mine"
    $sorted = @($new | Sort-Object)
    Write-Host "shipping code would pick:  $($new[0])   (`$new[0], unsorted)"
    if ($new.Count -gt 1) {
        Write-Host "  RESULT: differencing caught $($new.Count) pids for ONE instance we created."
        if ($new[0] -ne $mine) {
            Write-Host "  RESULT: *** \$new[0] = $($new[0]) is NOT ours ($mine). Stop-Word would kill a foreign Word. ***"
        } else {
            Write-Host "  RESULT: \$new[0] happened to be ours this run -- but $($new.Count - 1) foreign pid(s) were candidates."
        }
    } else {
        Write-Host "  RESULT: only one new pid this run; the foreign Word did not land inside the window."
    }
} else {
    Write-Host "  (caption unavailable, so arm A cannot name the true owner)"
}

# --- cleanup -----------------------------------------------------------------
} finally {
Write-Host ""
Write-Host "=== cleanup ==="
# Report rather than swallow. An empty `catch` here is the "evidence code lies
# rather than fails" trap in its purest form: if `Quit` refuses, the probe still
# prints a tidy cleanup section and the leak shows up later as somebody else's
# problem. It did exactly that -- three runs of this file leaked every instance
# it created, and the reason was only ever one `catch { }` away:
#
#   $a.Quit(0)  ->  Argument "1" must be System.Management.Automation.PSReference
#
# `Application.Quit` takes its three parameters as `VARIANT*`, so the literal `0`
# has to bind by reference. The no-argument form takes the same default
# (`wdSaveChanges` is only consulted for dirty documents, and these instances
# never open one) and sidesteps the binding entirely, which is why the four
# probes here that already used `$word.Quit()` never leaked.
#
# Honest limit on that explanation: the failure does *not* reproduce in a minimal
# script -- `Quit(0)` on a function-returned, array-stored instance binds fine
# there. So the trigger is something this file does that the reduction does not,
# and it is unidentified. What is measured is that the `(0)` form threw here on
# every instance of every run, and that the no-argument form reaps them.
$i = 0
foreach ($a in $created) {
    $i++
    try { $a.Quit() } catch { Write-Host "  instance ${i}: Quit threw -- $($_.Exception.Message.Split([char]10)[0])" }
    try {
        $rc = [Runtime.InteropServices.Marshal]::ReleaseComObject($a)
        if ($rc -ne 0) { Write-Host "  instance ${i}: $rc reference(s) still held after release" }
    } catch { Write-Host "  instance ${i}: release threw -- $($_.Exception.Message.Split([char]10)[0])" }
}
$created = @()
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

# Release the helper and let it quit its own Word. Force-killing is the fallback
# and not the plan, because it is the thing that leaked; if we ever take it, say
# so, since the census below will then be reporting our own damage.
#
# `-Path`, not `-LiteralPath`: Windows PowerShell 5.1's `New-Item` has no
# `-LiteralPath`, and under `Stop` the resulting parameter-binding error killed
# the whole teardown. These are GUID names in the temp directory, so there is no
# wildcard to protect against.
if ($foreign) {
    try { New-Item -ItemType File -Path $release -Force | Out-Null } catch { }
    $foreign | Wait-Process -Timeout 45 -ErrorAction SilentlyContinue
    if (-not $foreign.HasExited) {
        Write-Host "  WARNING: helper did not exit in 45s -- force-killing, which orphans its Word"
        try { Stop-Process -Id $foreign.Id -Force -ErrorAction SilentlyContinue } catch { }
    }
}
Remove-Item -Path $ready, $release -Force -ErrorAction SilentlyContinue

# --- census: did this probe leave anything behind? ---------------------------
# Poll rather than sleep flat. `Quit()` returns in ~120 ms while the process
# outlives it by seconds, and the tail is load-dependent: a Quit-to-exit
# measured at 2.7-6.1 s idle survived a 30 s poll with another session driving
# Word concurrently.
Write-Host ""
Write-Host "=== census ==="
$survivors = @()
for ($i = 0; $i -lt 120; $i++) {
    $survivors = @(Get-WordPids | Where-Object { $baseline -notcontains $_ })
    if ($survivors.Count -eq 0) { break }
    Start-Sleep -Milliseconds 500
}
if ($survivors.Count -eq 0) {
    Write-Host "  RESULT: clean -- WINWORD population back to its $($baseline.Count) at start"
} else {
    # Differencing again, and unsound here for exactly the reason arm A measures:
    # a Word another session started during this run lands in this set too. It is
    # the right instrument anyway, because over-reporting a leak is a false alarm
    # while under-reporting one is the defect. Do not kill these.
    Write-Host "  RESULT: *** $($survivors.Count) WINWORD still alive after 60s: $($survivors -join ', ') ***"
    Write-Host "  (some may be another session's -- this set is differenced, which arm A"
    Write-Host "   just showed is unsound. Reported, not killed.)"
}
Write-Host "WINWORD alive after: $((Get-WordPids).Count) (baseline was $($baseline.Count))"
}
