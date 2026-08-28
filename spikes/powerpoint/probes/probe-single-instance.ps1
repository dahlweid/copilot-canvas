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

$app1 = $null; $app2 = $null
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
    try { $pres.Saved = -1; $pres.Close() } catch { }

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
        foreach ($p in ((Get-Content $pidFile) -split ',' | Where-Object { $_ })) {
            if (Get-Process -Id ([int]$p) -ErrorAction SilentlyContinue) {
                Stop-Process -Id ([int]$p) -Force -ErrorAction SilentlyContinue
                Say "  swept job-created pid $p"
            }
        }
    }
    Rep "  SINGLE-INSTANCE (cross-process)" ($(if ((Test-Path $pidFile) -and -not (Get-Content $pidFile)) { 'YES - a separate process attached to the running instance' } else { 'NO' }))
}
catch { Rep "ERROR (PowerPoint arms)" $_.Exception.Message.Split([char]10)[0] }
finally {
    foreach ($a in @($app2, $app1)) {
        if ($a) { try { [Runtime.InteropServices.Marshal]::ReleaseComObject($a) | Out-Null } catch { } }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Start-Sleep -Seconds 2
    foreach ($p in (Get-PptPids | Where-Object { $pptBefore -notcontains $_ })) {
        if (Get-Process -Id $p -ErrorAction SilentlyContinue) {
            Stop-Process -Id $p -Force -ErrorAction SilentlyContinue
            Say "  swept POWERPNT pid $p (created by this probe)"
        }
    }
}

# --- S3: the Word control -----------------------------------------------------
Say "== S3: CONTROL - the identical test against Word =="
$w1 = $null; $w2 = $null
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
    # Only quit Word instances this probe created; a sibling session may be
    # driving Word right now.
    foreach ($w in @($w2, $w1)) {
        if ($w) {
            try {
                $newNow = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id | Where-Object { $wordBefore -notcontains $_ })
                if ($newNow.Count -gt 0) { $w.Quit(0) }
            }
            catch { }
            try { [Runtime.InteropServices.Marshal]::ReleaseComObject($w) | Out-Null } catch { }
        }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Start-Sleep -Seconds 2
    foreach ($p in (@(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id) | Where-Object { $wordBefore -notcontains $_ })) {
        if (Get-Process -Id $p -ErrorAction SilentlyContinue) {
            Stop-Process -Id $p -Force -ErrorAction SilentlyContinue
            Say "  swept WINWORD pid $p (created by this probe)"
        }
    }
    Remove-Item $root -Recurse -Force -ErrorAction SilentlyContinue
}

Rep "POWERPNT pids after" ($(if (Get-PptPids) { (Get-PptPids) -join ',' } else { '(none)' }))
Rep "WINWORD pids after" ($(if (Get-Process WINWORD -ErrorAction SilentlyContinue) { (@(Get-Process WINWORD | ForEach-Object Id)) -join ',' } else { '(none)' }))
