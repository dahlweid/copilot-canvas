# Probe: does PowerPoint BLOCK on a held file the way Word does?
#
# CORRECTION -- READ BEFORE TRUSTING T3/T3c.
# T3 and T3c are INVALID and are kept only as a record of how the mistake was
# made. They assume New-Object -ComObject PowerPoint.Application starts a second
# PowerPoint. It does not: probe-single-instance.ps1 later showed it ATTACHES to
# the running one, so T3 measured a single instance re-opening its own file and
# duly reported "does not block" -- the opposite of the truth. The question is
# answered properly in probe-cross-instance-lock.ps1, which launches a genuinely
# separate process and finds that PowerPoint DOES hang, exactly like Word.
#
# T1, T2, T4 and T5 are unaffected: they never needed a second instance.
#
# Word does not arbitrate access, it blocks: while a hidden Word held a document,
# every external write pattern failed with a sharing violation, and a second Word
# opening the same path hung indefinitely -- with DisplayAlerts already off, so
# no dialog existed to dismiss. Both processes needed an external kill. That is
# what forced the transient-lock model.
#
#   T1  What lock artefacts appear next to the file?
#   T2  Can an external writer replace the file while we hold it?
#       (a) direct overwrite  (b) write-temp-then-rename  (c) exclusive handle
#   T3  INVALID, see above. Can a SECOND PowerPoint open the held file?
#   T3c INVALID, see above. A/B control against a different file.
#   T4  Transient-lock round trip: open, edit, save, close. Word: 228 ms.
#   T5  Lock detection by write-handle test. Word: 4 ms held / 9 ms free.
#
#   -TimeoutSec  how long a blocked open is allowed to run before we call it a
#                hang (default 45; Word never returned at all)

param([string]$Fixture, [int]$TimeoutSec = 45)

$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot '_common.ps1')

if (-not $Fixture) { $Fixture = Join-Path (Split-Path $PSScriptRoot -Parent) '.fixtures\deck.pptx' }
if (-not (Test-Path $Fixture)) { throw "fixture missing: $Fixture (run make-fixture.ps1 first)" }

$root = Join-Path $env:TEMP ("pptlock-" + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $root | Out-Null
$deck = Join-Path $root 'original.pptx'
$other = Join-Path $root 'control.pptx'      # never opened by us: the A/B control target
$regen = Join-Path $root 'regen-source.pptx'
Copy-Item $Fixture $deck
Copy-Item $Fixture $other
Copy-Item $Fixture $regen

function Show-Dir($label) {
    Say ("  [$label] " + ((Get-ChildItem $root -Force | Select-Object -ExpandProperty Name) -join ', '))
}

function Test-Locked([string]$Path) {
    try { $fs = [IO.File]::Open($Path, 'Open', 'Write', 'None'); $fs.Close(); $false } catch { $true }
}

# Is the presentation handle we are holding still usable? Word's handle survived
# everything done to the file underneath it; this checks whether PowerPoint's
# does, because a silently-dead handle would be a worse failure than a refusal.
function Test-PresAlive($pres, $label) {
    try { $null = $pres.Slides.Count; Rep "  handle still valid after $label" 'YES' }
    catch { Rep "  handle still valid after $label" ('NO -> ' + $_.Exception.Message.Split([char]10)[0]) }
}

# Open a path in a SEPARATE PowerPoint, in a job, under a timeout. The job
# records the PIDs it created so the parent can sweep exactly those and nothing
# a sibling session owns.
function Invoke-SecondInstanceOpen([string]$Path, [string]$Label) {
    $pidFile = Join-Path $root ("pids-" + $Label + ".txt")
    $job = Start-Job -ArgumentList $Path, $pidFile -ScriptBlock {
        param($p, $pidFile)
        $before = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | ForEach-Object Id)
        $app = New-Object -ComObject PowerPoint.Application
        $app.DisplayAlerts = 1     # ppAlertsNone -- nothing may surface a dialog
        Start-Sleep -Milliseconds 300
        $owned = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | ForEach-Object Id) |
            Where-Object { $before -notcontains $_ }
        Set-Content -Path $pidFile -Value ($owned -join ',')
        $sw = [Diagnostics.Stopwatch]::StartNew()
        try {
            $pres = $app.Presentations.Open($p, 0, 0, 0)
            $sw.Stop()
            $r = "OPENED in $($sw.ElapsedMilliseconds) ms; ReadOnly=$($pres.ReadOnly)"
            try { $pres.Close() } catch { }
            $r
        }
        catch {
            $sw.Stop()
            "REFUSED after $($sw.ElapsedMilliseconds) ms -> " + $_.Exception.Message.Split([char]10)[0]
        }
        finally { try { $app.Quit() } catch { } }
    }

    $done = Wait-Job $job -Timeout $TimeoutSec
    if ($done) {
        Say ("  [$Label] " + (Receive-Job $job))
    }
    else {
        Say "  [$Label] HUNG -- no return after $TimeoutSec s (this is Word's behaviour)"
        Stop-Job $job -ErrorAction SilentlyContinue
    }
    Remove-Job $job -Force -ErrorAction SilentlyContinue

    # Sweep only the PIDs that job created.
    if (Test-Path $pidFile) {
        foreach ($p in ((Get-Content $pidFile) -split ',' | Where-Object { $_ })) {
            if (Get-Process -Id ([int]$p) -ErrorAction SilentlyContinue) {
                Stop-Process -Id ([int]$p) -Force -ErrorAction SilentlyContinue
                Say "  [$Label] swept second-instance pid $p"
            }
        }
    }
    return $done -ne $null
}

$ctx = $null
try {
    Rep "POWERPNT pids before" ($(if (Get-PptPids) { (Get-PptPids) -join ',' } else { '(none)' }))
    Say "== T0: baseline =="
    Show-Dir 'before open'

    $ctx = New-OwnedPowerPoint
    Rep "owned pid" ($(if ($ctx.Owned) { $ctx.Owned -join ',' } else { '(attached - will NOT kill)' }))

    $sw = [Diagnostics.Stopwatch]::StartNew()
    $pres = $ctx.App.Presentations.Open($deck, 0, 0, 0)   # ReadOnly = msoFalse
    $sw.Stop()
    Rep "opened read-write in" ("{0} ms; ReadOnly={1}" -f $sw.ElapsedMilliseconds, $pres.ReadOnly)

    Say "== T1: lock artefacts while held =="
    Show-Dir 'while open'
    Rep "  write-handle test says locked" (Test-Locked $deck)

    Say "== T2: can an external writer replace the file? =="
    # -ErrorAction Stop on every one of these: Copy-Item and Move-Item raise
    # NON-terminating errors, so without it the catch never fires and a failed
    # overwrite is reported as SUCCEEDED.
    try { Copy-Item $regen $deck -Force -ErrorAction Stop; Say "  (a) Copy-Item overwrite: SUCCEEDED" }
    catch { Say ("  (a) Copy-Item overwrite: FAILED -> " + $_.Exception.Message.Split([char]10)[0]) }
    try {
        $tmp = Join-Path $root 'new.tmp'
        Copy-Item $regen $tmp -Force -ErrorAction Stop
        Move-Item $tmp $deck -Force -ErrorAction Stop
        Say "  (b) write-temp-then-rename: SUCCEEDED"
    }
    catch { Say ("  (b) write-temp-then-rename: FAILED -> " + $_.Exception.Message.Split([char]10)[0]) }
    finally { Remove-Item (Join-Path $root 'new.tmp') -Force -ErrorAction SilentlyContinue }
    try { $fs = [IO.File]::Open($deck, 'Open', 'Write', 'None'); $fs.Close(); Say "  (c) exclusive write handle: SUCCEEDED" }
    catch { Say ("  (c) exclusive write handle: FAILED -> " + $_.Exception.Message.Split([char]10)[0]) }
    Test-PresAlive $pres 'T2'

    Say "== T3: a SECOND PowerPoint opens the HELD file =="
    $t3 = Invoke-SecondInstanceOpen $deck 'T3-held'
    Test-PresAlive $pres 'T3'

    Say "== T3c: A/B CONTROL -- second PowerPoint opens a DIFFERENT, free file =="
    $t3c = Invoke-SecondInstanceOpen $other 'T3-control'
    Rep "  control returned at all" $t3c
    Rep "  held-file open returned at all" $t3
    Test-PresAlive $pres 'T3c'

    Say "== T4: transient-lock round trip =="
    try { $pres.Saved = -1; $pres.Close() } catch { Say ("  close failed -> " + $_.Exception.Message.Split([char]10)[0]) }
    Rep "  file free after close" (-not (Test-Locked $deck))

    # Discard the first iteration: it carries one-off engine load.
    try {
        $p = $ctx.App.Presentations.Open($deck, 0, 0, 0)
        $null = $p.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.InsertAfter("warm")
        $p.Save(); $p.Close()
    }
    catch { Say ("  warmup failed -> " + $_.Exception.Message.Split([char]10)[0]) }

    $times = @()
    for ($i = 1; $i -le 5; $i++) {
        try {
            $sw = [Diagnostics.Stopwatch]::StartNew()
            $p = $ctx.App.Presentations.Open($deck, 0, 0, 0)
            $null = $p.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.InsertAfter("e$i")
            $p.Save()
            $p.Close()
            $sw.Stop()
            $times += $sw.ElapsedMilliseconds
        }
        catch { Say ("  iteration $i failed -> " + $_.Exception.Message.Split([char]10)[0]); break }
    }
    if ($times.Count) {
        Rep "  per-iteration ms" ($times -join ', ')
        Rep "  open+edit+save+close (mean)" ("{0:F0} ms" -f (($times | Measure-Object -Average).Average))
    }

    Say "== T5: lock detection cost =="
    $sw = [Diagnostics.Stopwatch]::StartNew(); $freeLocked = Test-Locked $deck; $sw.Stop()
    Rep "  file FREE  -> locked=$freeLocked" ("{0} ms (expect False)" -f $sw.ElapsedMilliseconds)
    $held = $null
    try {
        $held = $ctx.App.Presentations.Open($deck, 0, 0, 0)
        $sw = [Diagnostics.Stopwatch]::StartNew(); $heldLocked = Test-Locked $deck; $sw.Stop()
        Rep "  file HELD  -> locked=$heldLocked" ("{0} ms (expect True)" -f $sw.ElapsedMilliseconds)
        $held.Saved = -1; $held.Close()
    }
    catch { Say ("  T5 held check failed -> " + $_.Exception.Message.Split([char]10)[0]) }
}
catch { Rep "ERROR" $_.Exception.Message }
finally {
    Close-OwnedPowerPoint $ctx
    Remove-Item $root -Recurse -Force -ErrorAction SilentlyContinue
    Rep "POWERPNT pids after cleanup" ($(if (Get-PptPids) { (Get-PptPids) -join ',' } else { '(none)' }))
}
