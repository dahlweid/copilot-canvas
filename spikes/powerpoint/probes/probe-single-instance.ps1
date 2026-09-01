# Probe: is PowerPoint single-instance? Can we get a PRIVATE hidden instance at
# all, or do we always end up driving the user's PowerPoint?
#
# This is the most consequential question in the whole spike and it was not on
# the original list. It surfaced from an anomaly: probe-lock.ps1's "second
# PowerPoint" opened a file that the first instance was holding, in 36 ms, with
# no sharing violation and no new POWERPNT process -- because there was no
# second PowerPoint. New-Object appeared to hand back the instance already
# running.
#
# If that is right, the consequences are severe and they are not about speed:
#
#   - There is no such thing as "our" hidden PowerPoint. Every COM call lands in
#     whatever PowerPoint the user already has open, with their decks in it.
#   - Quit() would close the user's PowerPoint and their unsaved work.
#   - Killing "our" PID would kill theirs.
#   - Application-level settings we set (DisplayAlerts, and anything else) apply
#     to their session too.
#
# Word is multi-instance: every CreateObject starts a new WINWORD, which is why
# the isolation spike could bind to its own instance via OBJID_NATIVEOM and
# never touch the user's. So the arms are:
#
#   S1  Two New-Object calls in THIS process. Same PID? Same Presentations?
#   S2  A New-Object from a SEPARATE process (job). Does it start a new POWERPNT
#       or attach to the running one?
#   S3  CONTROL: the identical S1 test against Word. If Word forks a second
#       WINWORD and PowerPoint does not, the difference is the application, not
#       the method or this machine.
#
# SAFETY: this probe never calls Quit() on an instance it did not create, and
# never kills a PID that existed before it started. If a sibling session or the
# user has Office open, it degrades to reporting that rather than interfering.

param([string]$Fixture)

$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot '_common.ps1')

if (-not $Fixture) { $Fixture = Join-Path (Split-Path $PSScriptRoot -Parent) '.fixtures\deck.pptx' }
if (-not (Test-Path $Fixture)) { throw "fixture missing: $Fixture (run make-fixture.ps1 first)" }

$root = Join-Path $env:TEMP ("pptsingle-" + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Force -Path $root | Out-Null

$pptBefore = Get-PptPids
$wordBefore = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
Rep "POWERPNT pids before" ($(if ($pptBefore) { $pptBefore -join ',' } else { '(none)' }))
Rep "WINWORD pids before" ($(if ($wordBefore) { $wordBefore -join ',' } else { '(none)' }))
if ($pptBefore.Count -or $wordBefore.Count) {
    Say "  NOTE: Office already running. This probe will not quit or kill anything it did not start."
}

$app1 = $null; $app2 = $null; $pres = $null
try {
    Say "== S1: two New-Object calls in ONE process =="
    $app1 = New-Object -ComObject PowerPoint.Application
    Start-Sleep -Milliseconds 600
    $after1 = Get-PptPids
    $owned1 = @($after1 | Where-Object { $pptBefore -notcontains $_ })
    Rep "  after 1st New-Object, new pids" ($(if ($owned1) { $owned1 -join ',' } else { '(none - ATTACHED)' }))

    $app2 = New-Object -ComObject PowerPoint.Application
    Start-Sleep -Milliseconds 600
    $after2 = Get-PptPids
    $owned2 = @($after2 | Where-Object { $after1 -notcontains $_ })
    Rep "  after 2nd New-Object, new pids" ($(if ($owned2) { $owned2 -join ',' } else { '(none - ATTACHED)' }))
    Rep "  total POWERPNT processes now" $after2.Count
    Rep "  SINGLE-INSTANCE (in-process)" ($(if ($owned2.Count -eq 0) { 'YES - the 2nd call returned the 1st instance' } else { 'NO - a separate process started' }))

    # Behavioural confirmation: does instance 2 see instance 1's presentation?
    # Identical PIDs could in principle still be two isolated COM apartments.
    $src = Join-Path $root 'shared.pptx'; Copy-Item $Fixture $src
    $before2 = 0
    try { $before2 = $app2.Presentations.Count } catch { }
    $pres = $app1.Presentations.Open($src, 0, 0, 0)
    Start-Sleep -Milliseconds 400
    $seen = -1
    try { $seen = $app2.Presentations.Count } catch { }
    Rep "  app2.Presentations.Count before/after app1 opened" ("$before2 -> $seen")
    Rep "  SAME OBJECT MODEL" ($(if ($seen -gt $before2) { 'YES - app2 can see the deck app1 opened' } else { 'NO' }))
    if ($seen -gt 0) {
        try { Rep "  app2 can name app1's deck" $app2.Presentations.Item($seen).Name } catch { }
    }
    try { $pres.Saved = -1; $pres.Close(); $pres = $null } catch { }

    Say "== S2: New-Object from a SEPARATE process =="
    $pidFile = Join-Path $root 'jobpids.txt'
    $job = Start-Job -ArgumentList $pidFile -ScriptBlock {
        param($pidFile)
        $before = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | ForEach-Object Id)
        $a = New-Object -ComObject PowerPoint.Application
        Start-Sleep -Milliseconds 800
        $after = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | ForEach-Object Id)
        $new = @($after | Where-Object { $before -notcontains $_ })
        Set-Content -Path $pidFile -Value ($new -join ',')
        $n = -1; try { $n = $a.Presentations.Count } catch { }
        # Do NOT Quit: if this attached to an instance we did not create,
        # quitting would close it out from under its owner.
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($a) | Out-Null } catch { }
        "new pids: $($(if ($new) { $new -join ',' } else { '(none - ATTACHED)' })); Presentations.Count = $n"
    }
    $done = Wait-Job $job -Timeout 90
    if ($done) { Say ("  " + (Receive-Job $job)) } else { Say "  job timed out"; Stop-Job $job -ErrorAction SilentlyContinue }
    Remove-Job $job -Force -ErrorAction SilentlyContinue
    if (Test-Path $pidFile) {
        # The job above (:98-100) correctly declines to Quit these, on the
        # grounds that it may have attached to an instance it did not create.
        # This block used to force-kill the very same pids, which discarded that
        # reasoning at the point it mattered most. Report only (issue #139).
        $still = @((Get-Content $pidFile) -split ',' | Where-Object { $_ } |
            Where-Object { Get-Process -Id ([int]$_) -ErrorAction SilentlyContinue })
        if ($still.Count -gt 0) {
            Say "  POSSIBLY left running: $($still -join ',') -- confirm and close by hand"
        }
    }
    Rep "  SINGLE-INSTANCE (cross-process)" ($(if ((Test-Path $pidFile) -and -not (Get-Content $pidFile)) { 'YES - a separate process attached to the running instance' } else { 'NO' }))
}
catch { Rep "ERROR (PowerPoint arms)" $_.Exception.Message.Split([char]10)[0] }
finally {
    # Our own deck, opened at the top of S1 from our temp root. Close it here so
    # a failure above does not leave it open in an instance we attached to.
    try { if ($pres) { $pres.Saved = -1; $pres.Close() } } catch { }
    $pres = $null
    foreach ($a in @($app2, $app1)) {
        if ($a) { try { [Runtime.InteropServices.Marshal]::ReleaseComObject($a) | Out-Null } catch { } }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Start-Sleep -Seconds 2
    # Report, never kill. This probe's own S1 result is that New-Object ATTACHES,
    # so a pid appearing between $pptBefore and now is not evidence this probe
    # created it -- the old message here said "(created by this probe)", which
    # was a cause the code could not know, and it was attached to a force-kill.
    $appeared = @(Get-PptPids | Where-Object { $pptBefore -notcontains $_ })
    if ($appeared.Count -gt 0) {
        Say "  POWERPNT appeared during this probe and is still up: $($appeared -join ',') -- confirm and close by hand"
    }
}

# --- S3: the Word control -----------------------------------------------------
Say "== S3: CONTROL - the identical test against Word =="
$w1 = $null; $w2 = $null
$wOwned1 = @(); $wOwned2 = @()
try {
    $w1 = New-Object -ComObject Word.Application
    Start-Sleep -Milliseconds 600
    $wAfter1 = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
    $wOwned1 = @($wAfter1 | Where-Object { $wordBefore -notcontains $_ })
    Rep "  after 1st New-Object, new pids" ($(if ($wOwned1) { $wOwned1 -join ',' } else { '(none - ATTACHED)' }))

    $w2 = New-Object -ComObject Word.Application
    Start-Sleep -Milliseconds 600
    $wAfter2 = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
    $wOwned2 = @($wAfter2 | Where-Object { $wAfter1 -notcontains $_ })
    Rep "  after 2nd New-Object, new pids" ($(if ($wOwned2) { $wOwned2 -join ',' } else { '(none - ATTACHED)' }))
    Rep "  Word MULTI-INSTANCE" ($(if ($wOwned2.Count -gt 0) { 'YES - each call started its own WINWORD' } else { 'NO' }))
}
catch { Rep "  ERROR (Word control)" $_.Exception.Message.Split([char]10)[0] }
finally {
    # Quit the instance this iteration holds, unconditionally. Both earlier
    # versions of this gate consulted a pid census before allowing Quit(); the
    # per-instance narrowing above fixed *which* population was consulted, but
    # consulting one at all is the remaining defect, and it is mine (#136).
    #
    # Why no gate is the right gate here. Quit() is invoked on $w, an RCW that
    # refers to one specific instance. It is not a pid, it cannot be recycled,
    # and it cannot be another session's Word by accident -- so unlike a
    # Stop-Process on a differenced pid, the destructive reach of this call is
    # already bounded by the object we hold. The only question the census could
    # answer is "is this instance ours", and probe-newobject-attach.ps1 answered
    # that directly and better: New-Object Word.Application never attached to a
    # running Word in either arm. Word is multi-instance, so an instance we
    # created is ours.
    #
    # Note this is the OPPOSITE of the PowerPoint situation in S1/S2 above, and
    # importing that conclusion here would be the mistake that split #136 from
    # #139: PowerPoint is single-instance, a New-Object may attach to the user's
    # running instance, and quitting it destroys their work. Word never attaches.
    #
    # What the gate actually cost. $wOwned1/$wOwned2 are census differences, and
    # a difference can come up EMPTY for an instance we did own -- Quit() returns
    # 2.7-6.1 s before its process exits (ADR 0005), so $w2's pid can still be in
    # $wAfter1 when $wOwned1 is computed, and the 600 ms sleep is not a guarantee
    # either. Every such miss took the else branch and released the RCW without
    # quitting, leaking a WINWORD this probe created. The gate could only ever
    # turn a Word we own into a leak; it could never prevent a wrong Quit,
    # because Quit() was never aimed at the census in the first place.
    #
    # $wOwned1/$wOwned2 stay as REPORTED evidence below. They are no longer
    # permission.
    foreach ($inst in @(
            [pscustomobject]@{ App = $w2; Owned = @($wOwned2); Label = '2nd' },
            [pscustomobject]@{ App = $w1; Owned = @($wOwned1); Label = '1st' })) {
        $w = $inst.App
        if ($w) {
            if ($inst.Owned.Count -eq 0) {
                Rep "  $($inst.Label) New-Object: census attributed no pid" 'quitting anyway - the RCW is the evidence, not the census'
            }
            try {
                # Quit(), never Quit(<arg>): under Windows PowerShell 5.1 -- the
                # runtime every .ps1 here runs under -- the argument form throws
                # and the Word survives, and process exit does not reap it either
                # (spikes/isolation/probes/probe-quit0-leak.ps1).
                $w.Quit()
            }
            catch { Rep "  Quit() FAILED (Word may leak)" $_.Exception.Message.Split([char]10)[0] }
            try { [Runtime.InteropServices.Marshal]::ReleaseComObject($w) | Out-Null } catch { }
        }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    # $ourPids is pid-differenced, and the comment above already concedes that is
    # "not sound attribution" -- a stranger's Word appearing in the census window
    # lands in this set at a measured rate. It used to be force-killed anyway, and
    # Stop-Process -Force destroys unsaved work with no prompt. Removed (#139).
    # The Quit() gate above no longer consults it either (#136); it survives here
    # only as something to REPORT on, which is all a difference can support.
    #
    # The wait stays, because it is the instrument: Quit() reaps slowly, so
    # without it a survivor count would be a stopwatch artefact rather than a
    # leak. What follows the deadline is now a report, not a kill.
    $ourPids = @($wOwned1 + $wOwned2 | Where-Object { $_ } | Select-Object -Unique)
    $deadline = (Get-Date).AddSeconds(30)
    while ((Get-Date) -lt $deadline -and @($ourPids | Where-Object { (Get-Process -Id $_ -ErrorAction SilentlyContinue).ProcessName -eq 'WINWORD' }).Count -gt 0) {
        Start-Sleep -Milliseconds 250
    }
    foreach ($p in $ourPids) {
        if ((Get-Process -Id $p -ErrorAction SilentlyContinue).ProcessName -eq 'WINWORD') {
            Rep "  WINWORD pid $p still up 30 s after Quit()" 'LEAKED - not killed; confirm and close by hand'
        }
        else { Say "  WINWORD pid $p exited on Quit()" }
    }
    # "unattributed", not "not ours" -- the code knows only that this pid is
    # absent from two census reads, and this file's own reasoning above admits a
    # census can miss a pid we made. Asserting "not this probe's" is the same
    # error as #136 read backwards: a difference used as an attribution.
    foreach ($p in (@(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id) | Where-Object { $wordBefore -notcontains $_ -and $ourPids -notcontains $_ })) {
        Rep "  new WINWORD pid $p appeared during this run and is unattributed" 'not killed; may be ours or another session''s -- confirm by hand'
    }
    Remove-Item $root -Recurse -Force -ErrorAction SilentlyContinue
}

Rep "POWERPNT pids after" ($(if (Get-PptPids) { (Get-PptPids) -join ',' } else { '(none)' }))
Rep "WINWORD pids after" ($(if (Get-Process WINWORD -ErrorAction SilentlyContinue) { (@(Get-Process WINWORD | ForEach-Object Id)) -join ',' } else { '(none)' }))
