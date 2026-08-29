# Does `New-Object -ComObject Word.Application` ever hand back a Word we did not
# create? Every one of this repo's Word probes obtains its instance that way and
# then calls Quit() on it, so if the answer is ever "yes" those Quit()s can close
# a Word belonging to the user or to a concurrent session -- prompting on unsaved
# work in a visible instance, or hanging on a modal prompt in a hidden one.
#
# The repo currently holds both answers. PLAN.md 2.3 states the call "goes through
# the Running Object Table and may hand back the *user's* Word", and uses that to
# justify the HWND bind. probe-single-instance.ps1 S3 reports Word as
# multi-instance. Neither is a measurement of the question the Quit() sites ask,
# which is specifically: with a Word already running and held by another process,
# does a fresh New-Object attach to it?
#
# Attribution is ActiveWindow.Hwnd + GetWindowThreadProcessId, never pid
# differencing -- an external producer creates WINWORDs on this machine at a
# measured rate, so a pid that is merely new attributes to nothing.
#
# Safety: a worker quits ONLY an instance whose attributed pid is absent from the
# baseline census it was handed. If it ever attributes to a pre-existing pid, that
# is the finding, and it is reported and left strictly alone -- not quit, not
# killed. The probe must not commit the very act it exists to detect.
$ErrorActionPreference = 'Stop'

Add-Type -Namespace Native -Name U -MemberDefinition @'
[DllImport("user32.dll")] public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);
'@

function Get-WordPids { @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id) }

$worker = @'
param([string]$OutFile, [string]$Baseline, [string]$Mode, [string]$SignalFile)
$ErrorActionPreference = 'Stop'
Add-Type -Namespace Native -Name U -MemberDefinition @"
[DllImport("user32.dll")] public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);
"@
$base = @($Baseline -split ',' | Where-Object { $_ } | ForEach-Object { [int]$_ })
$lines = New-Object System.Collections.ArrayList
$apps = New-Object System.Collections.ArrayList

function New-Attributed {
    # Returns @(app, pid, owned). A document is required: ActiveWindow throws
    # without one.
    $app = New-Object -ComObject Word.Application
    $app.Visible = $false
    $app.DisplayAlerts = 0
    $doc = $app.Documents.Add()
    $wp = 0
    [void][Native.U]::GetWindowThreadProcessId([IntPtr]$app.ActiveWindow.Hwnd, [ref]$wp)
    $proc = Get-Process -Id $wp -ErrorAction SilentlyContinue
    if ($wp -le 4 -or -not $proc -or $proc.ProcessName -ne 'WINWORD') {
        throw "attribution failed: hwnd resolved to pid $wp ($($proc.ProcessName))"
    }
    # Close only the document this call added, never the instance. Close(0) is
    # deliberate: argument-less Close() prompts on a dirty document.
    try { $doc.Close(0) } catch { }
    ,@($app, $wp, (-not ($base -contains $wp)))
}

try {
    [IO.File]::WriteAllText($OutFile, "START")

    $r1 = New-Attributed
    [void]$apps.Add($r1[0])
    [void]$lines.Add("PID $($r1[1]) $(if ($r1[2]) { 'OWNED' } else { 'ATTACHED-PREEXISTING' })")
    [IO.File]::WriteAllLines($OutFile, @('PARTIAL') + $lines)

    if ($Mode -eq 'twice') {
        $r2 = New-Attributed
        [void]$apps.Add($r2[0])
        [void]$lines.Add("PID $($r2[1]) $(if ($r2[2]) { 'OWNED' } else { 'ATTACHED-PREEXISTING' })")
    }

    if ($Mode -eq 'hold') {
        # Keep the instance alive and referenced while the sibling worker runs, so
        # the sibling faces a genuinely running, held Word. Bounded so a lost
        # signal cannot hang the probe.
        $deadline = (Get-Date).AddSeconds(120)
        while (-not (Test-Path -LiteralPath $SignalFile) -and (Get-Date) -lt $deadline) {
            Start-Sleep -Milliseconds 250
        }
    }

    [IO.File]::WriteAllLines($OutFile, @('OK') + $lines)
}
catch {
    [IO.File]::WriteAllText($OutFile, "ERR $($_.Exception.Message -replace '\s+', ' ')")
    exit 1
}
finally {
    # Quit only what this worker created. An instance that attributed to a
    # pre-existing pid is released, never quit: quitting it is the data-loss
    # event this probe exists to detect.
    for ($i = 0; $i -lt $apps.Count; $i++) {
        $line = $lines[$i]
        if ($line -like '*OWNED*') {
            try { $apps[$i].Quit() } catch { }
        }
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($apps[$i]) | Out-Null } catch { }
    }
}
'@

$workerPath = Join-Path $env:TEMP ("attachworker-{0}.ps1" -f ([guid]::NewGuid().ToString('N')))
[IO.File]::WriteAllText($workerPath, $worker)
$signal = Join-Path $env:TEMP ("attachsignal-{0}.txt" -f ([guid]::NewGuid().ToString('N')))

function Start-Worker {
    param([string]$Mode, [string]$OutFile, [string]$Baseline)
    # -File with discrete argv elements: no value crosses a command-line parser.
    Start-Process -FilePath 'powershell.exe' -ArgumentList @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $workerPath,
        '-OutFile', $OutFile, '-Baseline', $Baseline, '-Mode', $Mode, '-SignalFile', $signal
    ) -WindowStyle Hidden -PassThru
}

function Read-Pids {
    param([string]$OutFile)
    # The unary comma is load-bearing. Without it a single matching line is
    # unrolled by the function's output pipeline and returned as a bare String,
    # whose [0] is the character 'P'. Measured: that silently produced an empty
    # pid on the cross-process arm, which then compared equal to nothing and
    # printed "DIFFERENT - own instance" regardless of the pids -- a verdict of
    # "safe" that the data never supported, in the one arm this probe exists for.
    if (-not (Test-Path -LiteralPath $OutFile)) { return ,@() }
    $raw = ([IO.File]::ReadAllText($OutFile)).Trim()
    ,@(($raw -split '\r?\n') | Where-Object { $_ -like 'PID *' })
}

function Get-PidField {
    # 0 means "did not parse". Callers must refuse a verdict on 0 rather than
    # comparing it, for the same reason attribution refuses hwnd-pid 0.
    param([string]$Line)
    $m = [regex]::Match([string]$Line, '^PID (\d+) ')
    if (-not $m.Success) { return 0 }
    [int]$m.Groups[1].Value
}

$baseline = Get-WordPids
Write-Host ("baseline WINWORD count: {0}" -f $baseline.Count)
$baseStr = ($baseline -join ',')

$outHold = Join-Path $env:TEMP ("attachout-hold-{0}.txt" -f ([guid]::NewGuid().ToString('N')))
$outProbe = Join-Path $env:TEMP ("attachout-probe-{0}.txt" -f ([guid]::NewGuid().ToString('N')))
$outTwice = Join-Path $env:TEMP ("attachout-twice-{0}.txt" -f ([guid]::NewGuid().ToString('N')))
$holdProc = $null
$probeProc = $null
$twiceProc = $null
$holdPids = @(); $probePids = @(); $twicePids = @()

try {
    # --- Arm 1: a second *process* creates Word while the first still holds one ---
    $holdProc = Start-Worker -Mode 'hold' -OutFile $outHold -Baseline $baseStr

    $deadline = (Get-Date).AddSeconds(180)
    while ((Get-Date) -lt $deadline -and -not (Read-Pids $outHold)) { Start-Sleep -Milliseconds 250 }
    $holdPids = Read-Pids $outHold
    if (-not $holdPids) { throw "hold worker never published a pid (state: '$(if (Test-Path $outHold) { (Get-Content $outHold -Raw).Trim() })')" }
    Write-Host ("hold worker:  {0}" -f ($holdPids -join ' ; '))

    $probeProc = Start-Worker -Mode 'probe' -OutFile $outProbe -Baseline $baseStr
    $probeProc.WaitForExit(180000) | Out-Null
    $probePids = Read-Pids $outProbe
    Write-Host ("probe worker: {0}" -f ($probePids -join ' ; '))

    [IO.File]::WriteAllText($signal, 'go')
    $holdProc.WaitForExit(180000) | Out-Null

    # --- Arm 2: two New-Object calls inside one process ---
    $twiceProc = Start-Worker -Mode 'twice' -OutFile $outTwice -Baseline $baseStr
    $twiceProc.WaitForExit(180000) | Out-Null
    $twicePids = Read-Pids $outTwice
    Write-Host ("twice worker: {0}" -f ($twicePids -join ' ; '))
}
finally {
    foreach ($p in @($holdProc, $probeProc, $twiceProc)) {
        if ($p -and -not $p.HasExited) { try { $p.Kill() } catch { } }
    }
    Remove-Item -LiteralPath $workerPath, $signal, $outHold, $outProbe, $outTwice -ErrorAction SilentlyContinue
}

Write-Host ''
$attached = @($holdPids + $probePids + $twicePids | Where-Object { $_ -like '*ATTACHED-PREEXISTING*' })
if ($attached.Count -gt 0) {
    Write-Host "FINDING: New-Object attributed to a WINWORD that predates this probe."
    $attached | ForEach-Object { Write-Host "  $_" }
    Write-Host "Those instances were NOT quit and NOT killed."
}

# The question each Quit() site asks is whether a fresh New-Object can return an
# instance the caller did not create. Two independent ways for that to show:
# attributing to a baseline pid, or two calls resolving to one pid.
# Every verdict below is derived from a pid that actually parsed. A comparison
# between values that did not parse is not a measurement, and reports as such.
$inconclusive = $false
$holdPid = if ($holdPids.Count -ge 1) { Get-PidField $holdPids[0] } else { 0 }
$probePid = if ($probePids.Count -ge 1) { Get-PidField $probePids[0] } else { 0 }
Write-Host ''
if ($holdPid -le 0 -or $probePid -le 0) {
    $inconclusive = $true
    Write-Host ("cross-process: INCONCLUSIVE - a pid did not parse (holder line '{0}', fresh line '{1}')" -f
        $(if ($holdPids.Count -ge 1) { $holdPids[0] } else { '<none>' }),
        $(if ($probePids.Count -ge 1) { $probePids[0] } else { '<none>' }))
}
else {
    Write-Host ("cross-process: holder pid {0} vs fresh pid {1} -> {2}" -f $holdPid, $probePid,
        $(if ($holdPid -eq $probePid) { 'SAME - attached' } else { 'DIFFERENT - own instance' }))
}

if ($twicePids.Count -eq 2) {
    $t1 = Get-PidField $twicePids[0]; $t2 = Get-PidField $twicePids[1]
    if ($t1 -le 0 -or $t2 -le 0) {
        $inconclusive = $true
        Write-Host ("same-process:  INCONCLUSIVE - a pid did not parse ('{0}', '{1}')" -f $twicePids[0], $twicePids[1])
    }
    else {
        Write-Host ("same-process:  first pid {0} vs second pid {1} -> {2}" -f $t1, $t2,
            $(if ($t1 -eq $t2) { 'SAME - attached' } else { 'DIFFERENT - own instance' }))
    }
}
else {
    $inconclusive = $true
    Write-Host ("same-process:  two New-Object calls in one process did not both attribute ({0} result(s))" -f $twicePids.Count)
}

Write-Host ''
Write-Host ("PSVersion: {0}" -f $PSVersionTable.PSVersion)
if ($baseline.Count -eq 0) {
    Write-Host "INCONCLUSIVE: no WINWORD was running at baseline, so the cross-process arm had nothing to attach to."
    exit 2
}
if ($inconclusive) {
    Write-Host "INCONCLUSIVE: at least one arm produced no comparable measurement."
    exit 2
}
