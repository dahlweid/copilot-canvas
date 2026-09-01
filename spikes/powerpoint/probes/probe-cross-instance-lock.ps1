# Probe: does a SECOND PowerPoint instance block on a file the first one holds?
#
# This is a redo. probe-lock.ps1 asked this question and got "no blocking", but
# probe-single-instance.ps1 later showed that its "second instance" was the SAME
# instance re-opening its own file -- which is not the Word scenario at all. The
# Word finding was: a genuinely separate WINWORD, opening a file another WINWORD
# held, hung indefinitely even with DisplayAlerts suppressed, and both processes
# needed an external kill. That is what forced the transient-lock model in
# docs/adr/0005.
#
# So this probe uses _isolated.ps1 to launch a genuinely separate POWERPNT.EXE,
# and repeats the Word test faithfully.
#
#   A  TEST     isolated instance opens the deck the anchor instance is HOLDING
#   B  CONTROL  isolated instance opens a DIFFERENT, free deck
#
# B is what makes A mean anything: if A hangs and B does not, the held file is
# the cause. If both bind quickly, PowerPoint simply does not block. Without B a
# slow A could just be cold start.
#
# Both arms are bounded by a timeout, and the isolated process is killed by pid
# either way, so a hang cannot leave an orphan behind.

param([string]$Fixture, [int]$TimeoutSeconds = 40)

$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot '_common.ps1')
. (Join-Path $PSScriptRoot '_isolated.ps1')

if (-not $Fixture) { $Fixture = Join-Path (Split-Path $PSScriptRoot -Parent) '.fixtures\deck.pptx' }
if (-not (Test-Path $Fixture)) { throw "fixture missing: $Fixture (run make-fixture.ps1 first)" }

$root = Join-Path $env:TEMP ("pptxlock-" + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Force -Path $root | Out-Null
$held = Join-Path $root 'held.pptx'      # the anchor will hold this one open
$free = Join-Path $root 'free.pptx'      # control target, nobody holds it
Copy-Item $Fixture $held
Copy-Item $Fixture $free

function Try-Arm {
    param([string]$Label, [string]$Target, [int]$Timeout)
    Say "== $Label : isolated instance opens $(Split-Path $Target -Leaf) =="
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $iso = $null
    try {
        $iso = Start-IsolatedPowerPoint -File $Target -TimeoutSeconds $Timeout
        $sw.Stop()
        $procAlive = [bool](Get-Process -Id $iso.Pid -ErrorAction SilentlyContinue)
        $bound = [bool]$iso.App
        Rep "  launched pid" $iso.Pid
        Rep "  bound to object model" $bound
        Rep "  time to bind" $(if ($bound) { "$($iso.BoundMs) ms" } else { "TIMED OUT after $($sw.ElapsedMilliseconds) ms" })
        Rep "  process still alive" $procAlive
        if (-not $bound) { Rep "  diagnostic" $iso.Diag }
        if ($iso.Dialogs.Count) {
            Rep "  modal dialogs seen" $iso.Dialogs.Count
            foreach ($d in $iso.Dialogs) { Say ("      " + $d.Substring(0, [Math]::Min(220, $d.Length))) }
        }
        else { Rep "  modal dialogs seen" 0 }
        if ($bound) {
            try { Rep "  deck it opened" $iso.Pres.Name } catch { }
            try { Rep "  opened read-only" $iso.Pres.ReadOnly } catch { Rep "  opened read-only" 'unreadable' }
            try { Rep "  slides visible to it" $iso.Pres.Slides.Count } catch { }
        }
        Rep "  OUTCOME" $(if ($bound) { 'OPENED - did not block' }
            elseif ($procAlive) { 'HUNG - alive but no object model (Word behaviour)' }
            else { 'EXITED - refused the file and quit' })
        return [pscustomobject]@{ Bound = $bound; Ms = $iso.BoundMs; Alive = $procAlive }
    }
    finally { Stop-IsolatedPowerPoint $iso }
}

$anchor = $null
$anchorPres = $null
try {
    $before = @(Get-PptPids)
    Rep "POWERPNT pids before" ($(if ($before) { $before -join ',' } else { '(none)' }))

    $anchor = New-PowerPointInstance
    $anchorPres = $anchor.App.Presentations.Open($held, 0, 0, 0)   # read-write, holds the lock
    Rep "anchor holds" ("{0} (ReadOnly={1})" -f $anchorPres.Name, $anchorPres.ReadOnly)
    Rep "anchor pid" ($(if ($anchor.NewPids) { $anchor.NewPids -join ',' } else { '(attached)' }))

    # CONTROL runs FIRST on purpose. A hung arm can only be cleaned up with a
    # kill, and a killed PowerPoint poisons the next launch with a safe-mode
    # prompt. Running the control first means it is measured from a clean state;
    # if it ran second, a hang in the test arm would break it for a reason that
    # has nothing to do with file locks.
    $b = Try-Arm 'B  CONTROL' $free $TimeoutSeconds
    $a = Try-Arm 'A  TEST   ' $held $TimeoutSeconds

    Say "== verdict =="
    Rep "  A held file" $(if ($a.Bound) { "opened in $($a.Ms) ms" } elseif ($a.Alive) { 'HUNG' } else { 'exited' })
    Rep "  B free file" $(if ($b.Bound) { "opened in $($b.Ms) ms" } elseif ($b.Alive) { 'HUNG' } else { 'exited' })
    if ($a.Bound -and $b.Bound) {
        Rep "  CONCLUSION" 'PowerPoint does NOT block on a held file (Word hangs indefinitely)'
        Rep "  cost of the lock" ("{0} ms vs {1} ms control" -f $a.Ms, $b.Ms)
    }
    elseif (-not $a.Bound -and $b.Bound) {
        Rep "  CONCLUSION" 'PowerPoint DOES block on a held file - the lock is the cause (control opened fine)'
    }
    elseif (-not $a.Bound -and -not $b.Bound) {
        Rep "  CONCLUSION" 'INCONCLUSIVE - even the control failed, so the method is at fault, not the lock'
    }
    else { Rep "  CONCLUSION" 'anomalous - held opened but free did not' }

    try { $anchorPres.Saved = -1; $anchorPres.Close(); $anchorPres = $null } catch { }
}
catch { Rep "ERROR" $_.Exception.Message.Split([char]10)[0] }
finally {
    # Ours, opened from our own temp root -- close it before releasing the
    # application, which may be one we merely attached to.
    try { if ($anchorPres) { $anchorPres.Saved = -1; $anchorPres.Close() } } catch { }
    $anchorPres = $null
    Close-PowerPointInstance $anchor
    Remove-Item $root -Recurse -Force -ErrorAction SilentlyContinue
    Rep "POWERPNT pids after" ($(if (Get-PptPids) { (Get-PptPids) -join ',' } else { '(none)' }))
}
