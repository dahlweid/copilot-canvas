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
#   FRESH  N cycles, a brand new POWERPNT per cycle: open, export, close, quit.
#   WARM   one POWERPNT, one open presentation, N exports in a row.
#
# WARM is the architecture ADR 0003/0005 assume -- a long-lived hidden instance
# that re-exports on demand. FRESH is the control that says whether the cost sits
# in "export" or in "reuse".
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
    try { if ($ctx.App) { [Runtime.InteropServices.Marshal]::ReleaseComObject($ctx.App) | Out-Null } } catch { }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Start-Sleep -Milliseconds 1200
    $killed = 0
    foreach ($p in $ctx.Owned) {
        if (Get-Process -Id $p -ErrorAction SilentlyContinue) {
            Stop-Process -Id $p -Force -ErrorAction SilentlyContinue
            $killed++
        }
    }
    $killed
}

function Invoke-FreshArm {
        Say "== FRESH: a new POWERPNT per export (control) =="
    $results = @()
    $quitLeaks = 0
    for ($c = 1; $c -le $Cycles; $c++) {
        $src = Join-Path $root "fresh$c.pptx"
        Copy-Item $Fixture $src
        $stage = 'create'; $ctx = $null; $ms = 0
        try {
            $ctx = New-OwnedPowerPoint
            $stage = 'open'
            $pres = $ctx.App.Presentations.Open($src, $NO, $NO, $NO)
            $stage = 'export'
            $ms = Export-Deck $pres (Join-Path $root "fresh$c.pdf")
            $stage = 'post-export-property'
            $null = $pres.Slides.Count
            $stage = 'close'
            $pres.Saved = -1; $pres.Close()
            $stage = 'quit'
            $ctx.App.Quit()
            $results += [pscustomobject]@{ N = $c; Failed = $null; Ms = $ms }
        }
        catch {
            $alive = $false
            foreach ($p in $ctx.Owned) { if (Get-Process -Id $p -ErrorAction SilentlyContinue) { $alive = $true } }
            $results += [pscustomobject]@{ N = $c; Failed = $stage; Ms = $ms }
            Rep "  cycle $c" ("DIED at '$stage' (process alive: $alive) -> " + $_.Exception.Message.Split([char]10)[0])
        }
        finally { $quitLeaks += (Sweep $ctx) }
    }
    Rep "  Quit() left the process running" ("$quitLeaks / $Cycles cycles needed an external kill")
    , $results
}

function Invoke-WarmArm {
    Say "== WARM: one POWERPNT, one presentation, repeated export (the architecture) =="
    $src = Join-Path $root 'warm.pptx'
    Copy-Item $Fixture $src
    $results = @()
    $ctx = $null
    try {
        $ctx = New-OwnedPowerPoint
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
                foreach ($p in $ctx.Owned) { if (Get-Process -Id $p -ErrorAction SilentlyContinue) { $alive = $true } }
                $results += [pscustomobject]@{ N = $c; Failed = $stage; Ms = $ms }
                Rep "  export $c" ("DIED at '$stage' (process alive: $alive) -> " + $_.Exception.Message.Split([char]10)[0])
                break   # the instance is gone; remaining iterations would fail identically
            }
        }
        try { $pres.Saved = -1; $pres.Close(); $ctx.App.Quit() } catch { }
    }
    finally { $null = Sweep $ctx }
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
