# Probe: how often does a hidden PowerPoint instance die under us, and is
# reusing a warm instance the cause?
#
# This probe exists because of a result nobody went looking for. probe-export.ps1
# and probe-notes-control.ps1 both failed on the COM call that FOLLOWED a
# successful ExportAsFixedFormat, with 0x800706BA "RPC server is unavailable" --
# the POWERPNT process was simply gone. The A/B in probe-notes-control.ps1 ruled
# out the notes export as the cause: its control arm, which exported slides twice
# and nothing else, died too.
#
# Intermittent means it needs a rate, not an anecdote, and a rate is only useful
# next to a control. Two arms:
#
#   FRESH  N cycles, a new PowerPoint context per cycle: open, export, close.
#   WARM   one POWERPNT, one open presentation, N exports in a row.
#
# WARM is the architecture ADR 0003/0005 assume -- a long-lived hidden instance
# that re-exports on demand. FRESH is the control that says whether the cost sits
# in "export" or in "reuse".
#
# WHAT THIS PROBE NO LONGER ESTABLISHES
#
# The FRESH arm used to read "a brand new POWERPNT per cycle: open, export,
# close, quit", and it reported how often Quit() left the process running --
# the source of the 15/15 figure in FINDINGS.md. It got that by calling Quit()
# on the instance and then force-killing the census difference between cycles.
#
# Both of those are gone (see issue #139 and _common.ps1's header): New-Object
# attaches to a running PowerPoint, so the instance may be the user's, and a
# census difference cannot establish otherwise. Neither Quit() nor a kill is
# defensible on an object obtained that way.
#
# The consequence is honest and worth stating rather than absorbing: with no
# kill between cycles, nothing makes each cycle's instance a new process, so
# the arm no longer establishes a fresh process per cycle and may attach to a
# previous one. It still measures the export path per cycle, which is what the
# FRESH/WARM comparison of export cost needs. It cannot reproduce the 15/15
# Quit()-reaping figure, and that figure is NOT re-derivable from this file as
# it now stands -- it was measured on a clean machine, where the census really
# was empty and the sweep really did make each cycle fresh, and it stands as
# that historical measurement.
#
# Re-establishing a genuinely per-cycle process needs the route _isolated.ps1
# already proves: CreateProcess, so the pid is the kernel's answer rather than
# an inference. That is a redesign of this arm, not a change to it.
#
# It also reports whether the OS logged a fault, which distinguishes "PowerPoint
# crashed" from "PowerPoint decided to exit".
#
#   -Cycles   exports per arm (default 12)
#   -Arm      fresh | warm | both (default both)

param([string]$Fixture, [int]$Cycles = 12, [ValidateSet('fresh', 'warm', 'both')][string]$Arm = 'both')

$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot '_common.ps1')

if (-not $Fixture) { $Fixture = Join-Path (Split-Path $PSScriptRoot -Parent) '.fixtures\deck.pptx' }
if (-not (Test-Path $Fixture)) { throw "fixture missing: $Fixture (run make-fixture.ps1 first)" }

$PDF = 2; $INTENT_PRINT = 2; $OUT_SLIDES = 1; $RANGE = 4; $NO = 0; $HANDOUT = 2
$startedAt = Get-Date

$root = Join-Path $env:TEMP ("pptstab-" + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Force -Path $root | Out-Null

function Export-Deck($pres, $path) {
    $pres.PrintOptions.Ranges.ClearAll()
    $r = $pres.PrintOptions.Ranges.Add(1, $pres.Slides.Count)
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $pres.ExportAsFixedFormat($path, $PDF, $INTENT_PRINT, $NO, $HANDOUT, $OUT_SLIDES, $NO,
        $r, $RANGE, "", $false, $false, $false, $true, $false)
    $sw.Stop()
    $sw.Elapsed.TotalMilliseconds
}

function Sweep($ctx) {
    if ($null -eq $ctx) { return 0 }
    Close-PowerPointInstance $ctx
    # Report, do not kill. $ctx.NewPids is a census difference and cannot
    # establish that this probe created anything -- see _common.ps1's header.
    @(Get-PptPids | Where-Object { $ctx.NewPids -contains $_ }).Count
}

function Invoke-FreshArm {
        Say "== FRESH: a new PowerPoint context per export (control) =="
        Say "   NOTE: no longer guaranteed to be a NEW PROCESS per cycle -- see the header."
    $results = @()
    $survivors = 0
    for ($c = 1; $c -le $Cycles; $c++) {
        $src = Join-Path $root "fresh$c.pptx"
        Copy-Item $Fixture $src
        $stage = 'create'; $ctx = $null; $pres = $null; $ms = 0
        try {
            $ctx = New-PowerPointInstance
            $stage = 'open'
            $pres = $ctx.App.Presentations.Open($src, $NO, $NO, $NO)
            $stage = 'export'
            $ms = Export-Deck $pres (Join-Path $root "fresh$c.pdf")
            $stage = 'post-export-property'
            $null = $pres.Slides.Count
            $stage = 'close'
            $pres.Saved = -1; $pres.Close(); $pres = $null
            $results += [pscustomobject]@{ N = $c; Failed = $null; Ms = $ms }
        }
        catch {
            $alive = $false
            foreach ($p in $ctx.NewPids) { if (Get-Process -Id $p -ErrorAction SilentlyContinue) { $alive = $true } }
            $results += [pscustomobject]@{ N = $c; Failed = $stage; Ms = $ms }
            Rep "  cycle $c" ("DIED at '$stage' (process alive: $alive) -> " + $_.Exception.Message.Split([char]10)[0])
        }
        finally {
            # Ours: opened from our own temp copy. Closing it here stops a
            # failure above leaving our deck open in an attached instance.
            try { if ($pres) { $pres.Saved = -1; $pres.Close() } } catch { }
            $pres = $null
            $survivors += (Sweep $ctx)
        }
    }
    Rep "  POWERPNT still up after release" ("$survivors / $Cycles cycles (not killed -- see header)")
    , $results
}

function Invoke-WarmArm {
    Say "== WARM: one POWERPNT, one presentation, repeated export (the architecture) =="
    $src = Join-Path $root 'warm.pptx'
    Copy-Item $Fixture $src
    $results = @()
    $ctx = $null
    $pres = $null
    try {
        $ctx = New-PowerPointInstance
        $pres = $ctx.App.Presentations.Open($src, $NO, $NO, $NO)
        for ($c = 1; $c -le $Cycles; $c++) {
            $stage = 'export'; $ms = 0
            try {
                $ms = Export-Deck $pres (Join-Path $root "warm$c.pdf")
                $stage = 'post-export-property'
                $null = $pres.Slides.Count
                $results += [pscustomobject]@{ N = $c; Failed = $null; Ms = $ms }
            }
            catch {
                $alive = $false
                foreach ($p in $ctx.NewPids) { if (Get-Process -Id $p -ErrorAction SilentlyContinue) { $alive = $true } }
                $results += [pscustomobject]@{ N = $c; Failed = $stage; Ms = $ms }
                Rep "  export $c" ("DIED at '$stage' (process alive: $alive) -> " + $_.Exception.Message.Split([char]10)[0])
                break   # the instance is gone; remaining iterations would fail identically
            }
        }
        try { $pres.Saved = -1; $pres.Close(); $pres = $null } catch { }
    }
    finally {
        try { if ($pres) { $pres.Saved = -1; $pres.Close() } } catch { }
        $pres = $null
        $null = Sweep $ctx
    }
    , $results
}

function Show-Summary($label, $results) {
    $fail = @($results | Where-Object { $_.Failed })
    $ok = @($results | Where-Object { -not $_.Failed })
    Rep "  [$label] attempted / clean" ("{0} / {1}" -f $results.Count, $ok.Count)
    Rep "  [$label] died" ("{0}  ({1:P0} of attempts)" -f $fail.Count, ($fail.Count / [double][math]::Max(1, $results.Count)))
    if ($fail.Count) { Rep "  [$label] first death at attempt #" $fail[0].N }
    if ($ok.Count) { Rep "  [$label] export ms (clean, mean)" ("{0:F0} ms" -f (($ok | Measure-Object Ms -Average).Average)) }
}

Rep "POWERPNT pids before" ($(if (Get-PptPids) { (Get-PptPids) -join ',' } else { '(none)' }))

$freshR = $null; $warmR = $null
if ($Arm -eq 'fresh' -or $Arm -eq 'both') { $freshR = Invoke-FreshArm }
if ($Arm -eq 'warm' -or $Arm -eq 'both') { $warmR = Invoke-WarmArm }

""
"== summary =="
if ($null -ne $freshR) { Show-Summary 'FRESH' $freshR }
if ($null -ne $warmR) { Show-Summary 'WARM ' $warmR }

""
"== Application event log, POWERPNT, since probe start =="
$evt = Get-WinEvent -FilterHashtable @{ LogName = 'Application'; StartTime = $startedAt } -ErrorAction SilentlyContinue |
Where-Object { $_.Message -match 'POWERPNT' }
if ($evt) {
    $evt | Select-Object -First 6 | ForEach-Object {
        "  [{0}] {1}: {2}" -f $_.TimeCreated.ToString('HH:mm:ss'), $_.ProviderName, ($_.Message.Split([char]10)[0].Trim())
    }
}
else { "  (no POWERPNT entries -- any exit was not logged as a fault)" }

Remove-Item $root -Recurse -Force -ErrorAction SilentlyContinue
Rep "POWERPNT pids after" ($(if (Get-PptPids) { (Get-PptPids) -join ',' } else { '(none)' }))
