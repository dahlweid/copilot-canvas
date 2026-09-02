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
#   FRESH  N cycles, a genuinely new PowerPoint PROCESS per cycle (through the
#          CreateProcess route in _isolated.ps1): open, export, close, Quit(),
#          and measure whether Quit() reaps the process.
#   WARM   one POWERPNT, one open presentation, N exports in a row.
#
# WARM is the architecture ADR 0003/0005 assume -- a long-lived hidden instance
# that re-exports on demand. FRESH is the control that says whether the cost sits
# in "export" or in "reuse".
#
# WHAT THE FRESH ARM MEASURES
#
# The FRESH arm launches each cycle through Start-IsolatedPowerPoint: a
# CreateProcess into a private desktop, so the pid is the kernel's answer and the
# instance is provably not the user's. New-Object attaches to a running
# PowerPoint (probe-single-instance.ps1) and cannot be written to or quit safely,
# which is why #139 removed the old Quit()-and-kill and why this arm was rebuilt
# onto the isolated route rather than patched in place.
#
# Per cycle it opens the deck, exports, closes the deck, calls Quit(), and polls
# whether the process reaps within the same window the teardown itself waits. The
# count of cycles where Quit() left the process running is the measurement behind
# the 15/15 figure in FINDINGS.md, now taken on an instance whose ownership is not
# in doubt. The between-cycle sweep is Stop-IsolatedPowerPoint, which routes any
# survivor through Stop-VerifiedPpt -- a kill gated on the recorded pid AND
# StartTime that declines rather than guessing -- so a survivor is reaped without
# the census-difference inference #139 removed.
#
# The original 15/15 stands as a historical measurement taken on a clean machine
# with an empty census per cycle; FINDINGS.md records it as such. Any number this
# arm produces now is recorded ALONGSIDE it, with its own conditions, never over
# it.
#
# It also reports whether the OS logged a fault, which distinguishes "PowerPoint
# crashed" from "PowerPoint decided to exit".
#
#   -Cycles   exports per arm (default 12)
#   -Arm      fresh | warm | both (default both)

param([string]$Fixture, [int]$Cycles = 12, [ValidateSet('fresh', 'warm', 'both')][string]$Arm = 'both')

$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot '_common.ps1')
. (Join-Path $PSScriptRoot '_isolated.ps1')

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
    Say "== FRESH: a genuinely new PowerPoint process per cycle (control) =="
    Say "   Each cycle: CreateProcess (Start-IsolatedPowerPoint) -> export -> close -> Quit()."
    Say "   Reports whether Quit() reaps the process; any survivor is swept by verified"
    Say "   pid+StartTime (Stop-IsolatedPowerPoint -> Stop-VerifiedPpt), never by census."
    $results = @()
    $quitReaped = 0        # Quit() alone left no process
    $quitSurvived = 0      # process outlived Quit() and needed the verified sweep
    $selfExited = 0        # closing the last deck exited the process before Quit()
    for ($c = 1; $c -le $Cycles; $c++) {
        $src = Join-Path $root "fresh$c.pptx"
        Copy-Item $Fixture $src
        $stage = 'launch'; $iso = $null; $ms = 0
        try {
            $iso = Start-IsolatedPowerPoint -File $src
            if (-not $iso.App) {
                $results += [pscustomobject]@{ N = $c; Failed = 'launch'; Ms = 0 }
                Rep "  cycle $c" "NOT BOUND -- $($iso.Diag)"
                continue
            }
            $stage = 'export'
            $ms = Export-Deck $iso.Pres (Join-Path $root "fresh$c.pdf")
            $stage = 'post-export-property'
            $null = $iso.Pres.Slides.Count
            $stage = 'close'
            try { $iso.Pres.Saved = -1; $iso.Pres.Close() } catch { }
            Start-Sleep -Milliseconds 500
            if (-not (Get-Process -Id $iso.Pid -ErrorAction SilentlyContinue)) {
                # Closing the last deck can itself exit the process (FINDINGS Q4);
                # report it apart from a Quit() reap so the reaping rate is not
                # confounded by an exit that happened before Quit() was called.
                $selfExited++
                $results += [pscustomobject]@{ N = $c; Failed = $null; Ms = $ms }
                Rep "  cycle $c" ("export {0:F0} ms; process self-exited on closing its last deck (before Quit)" -f $ms)
                continue
            }
            $stage = 'quit'
            try { $iso.App.Quit() } catch { }
            $reaped = $false
            for ($i = 0; $i -lt 16; $i++) {
                if (-not (Get-Process -Id $iso.Pid -ErrorAction SilentlyContinue)) { $reaped = $true; break }
                Start-Sleep -Milliseconds 500
            }
            if ($reaped) { $quitReaped++ } else { $quitSurvived++ }
            $results += [pscustomobject]@{ N = $c; Failed = $null; Ms = $ms }
            $verdict = $(if ($reaped) { 'reaped the process' } else { 'left it running -> verified sweep' })
            Rep "  cycle $c" ("export {0:F0} ms; Quit() {1}" -f $ms, $verdict)
        }
        catch {
            $alive = [bool]($iso -and $iso.Pid -and (Get-Process -Id $iso.Pid -ErrorAction SilentlyContinue))
            $results += [pscustomobject]@{ N = $c; Failed = $stage; Ms = $ms }
            Rep "  cycle $c" ("DIED at '$stage' (process alive: $alive) -> " + $_.Exception.Message.Split([char]10)[0])
        }
        finally {
            # The sound between-cycle sweep. Stop-IsolatedPowerPoint quits again
            # (idempotent), routes any survivor through Stop-VerifiedPpt -- which
            # re-checks name and the recorded StartTime and declines rather than
            # trusting the pid -- and closes the private desktop. The pid it acts
            # on came from CreateProcess, not a census difference.
            Stop-IsolatedPowerPoint $iso
        }
    }
    Rep "  Quit() reaped / survived / self-exited" "$quitReaped / $quitSurvived / $selfExited  (of $Cycles cycles)"
    Rep "  Quit() failed to reap" "$quitSurvived / $($quitReaped + $quitSurvived) cycles where Quit() was reached"
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
