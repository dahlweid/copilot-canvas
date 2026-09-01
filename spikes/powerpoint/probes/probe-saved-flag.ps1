# Probe (A/B): does the Word idiom `doc.Saved = $true` port to PowerPoint?
#
# Hypothesis, arrived at from a failure rather than a guess. Three probes here
# died on the call right after a successful export, with 0x800706BA "RPC server
# is unavailable", and the Windows Application log records a matching
# POWERPNT.EXE access violation (0xc0000005) in combase.dll. The call they all
# died on was the same one:
#
#     $pres.Saved = $true
#
# which is a direct port of the Word host's `$doc.Saved = $true`. But Word's
# Saved is a VBA Boolean and PowerPoint's is an MsoTriState -- msoTrue is -1,
# msoFalse is 0. Assigning a VARIANT_BOOL where an enum is expected is exactly
# the shape of bug that produces an access violation inside the marshaller.
#
# Two arms differing in one literal:
#
#   A (correct)   $pres.Saved = -1        (msoTrue)
#   B (Word idiom) $pres.Saved = $true    (VARIANT_BOOL)
#
# Crash attribution is by PID against the Application event log, so a crash from
# a sibling session's Office process cannot be counted as ours.
#
#   -Cycles   iterations per arm (default 8)

param([string]$Fixture, [int]$Cycles = 8)

$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot '_common.ps1')

if (-not $Fixture) { $Fixture = Join-Path (Split-Path $PSScriptRoot -Parent) '.fixtures\deck.pptx' }
if (-not (Test-Path $Fixture)) { throw "fixture missing: $Fixture (run make-fixture.ps1 first)" }

$root = Join-Path $env:TEMP ("pptsaved-" + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Force -Path $root | Out-Null
$startedAt = Get-Date

function Invoke-Arm([string]$Label, $SavedValue) {
    $died = 0
    $ourPids = @()
    for ($c = 1; $c -le $Cycles; $c++) {
        $src = Join-Path $root "$Label$c.pptx"
        Copy-Item $Fixture $src
        $ctx = $null
        $pres = $null
        try {
            $ctx = New-PowerPointInstance
            $ourPids += $ctx.NewPids
            $pres = $ctx.App.Presentations.Open($src, 0, 0, 0)
            $null = $pres.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.InsertAfter("x")
            $pres.Saved = $SavedValue
            $pres.Close()
            $pres = $null
            $null = $ctx.App.Presentations.Count   # a call that needs the process alive
        }
        catch {
            $died++
            Rep "  [$Label] cycle $c" ('FAILED -> ' + $_.Exception.Message.Split([char]10)[0])
        }
        finally {
            # No Quit() and no kill: $ctx.App may be the user's PowerPoint
            # (New-Object attaches), and $ctx.NewPids is a census difference,
            # which is evidence of nothing. See _common.ps1's header.
            #
            # The presentation IS ours -- we opened $src, a copy of the fixture
            # in our own temp root -- so closing it is both safe and necessary:
            # an exception above would otherwise leave our deck open in
            # somebody else's instance. Close only this object; never enumerate
            # Presentations.
            try { if ($pres) { $pres.Saved = -1; $pres.Close() } } catch { }
            $pres = $null
            Close-PowerPointInstance $ctx
        }
    }
    [pscustomobject]@{ Label = $Label; Died = $died; Pids = $ourPids }
}

Rep "POWERPNT pids before" ($(if (Get-PptPids) { (Get-PptPids) -join ',' } else { '(none)' }))

Say "== A (correct): `$pres.Saved = -1  (msoTrue) =="
$a = Invoke-Arm 'A' (-1)
Say "== B (Word idiom): `$pres.Saved = `$true  (VARIANT_BOOL) =="
$b = Invoke-Arm 'B' $true

Say ""
Say "== summary =="
Rep "  A  Saved = -1 (msoTrue)" ("{0} / {1} cycles failed" -f $a.Died, $Cycles)
Rep "  B  Saved = `$true" ("{0} / {1} cycles failed" -f $b.Died, $Cycles)

# Correlate logged crashes with the pids that appeared while each arm ran. The
# event message carries the faulting process id in hex.
#
# This is a correlation, not an attribution: the pid set comes from a census
# difference, which over-reports (probe-init-attribution.ps1) and cannot
# establish that this arm created the process. The labels below say "coincided
# with" rather than "caused by" for that reason.
$evt = @(Get-WinEvent -FilterHashtable @{ LogName = 'Application'; ProviderName = 'Application Error'; StartTime = $startedAt } -ErrorAction SilentlyContinue |
    Where-Object { $_.Message -match 'POWERPNT' })
$crashPids = @()
foreach ($e in $evt) {
    $m = [regex]::Match($e.Message, '0x([0-9A-Fa-f]{1,8})\s*$', 'Multiline')
    foreach ($mm in [regex]::Matches($e.Message, '(?:ID|id)[^0-9A-Fa-fx]*0x([0-9A-Fa-f]+)')) {
        $crashPids += [Convert]::ToInt32($mm.Groups[1].Value, 16)
    }
}
$crashPids = @($crashPids | Sort-Object -Unique)
Rep "  logged POWERPNT crash events" $evt.Count
Rep "  crashes coinciding with arm A pids" (@($a.Pids | Where-Object { $crashPids -contains $_ }).Count)
Rep "  crashes coinciding with arm B pids" (@($b.Pids | Where-Object { $crashPids -contains $_ }).Count)
foreach ($e in ($evt | Select-Object -First 3)) {
    Say ("    [{0}] {1}" -f $e.TimeCreated.ToString('HH:mm:ss'), (($e.Message -split "`n")[0..3] -join ' | ').Trim())
}

Remove-Item $root -Recurse -Force -ErrorAction SilentlyContinue
Rep "POWERPNT pids after" ($(if (Get-PptPids) { (Get-PptPids) -join ',' } else { '(none)' }))
