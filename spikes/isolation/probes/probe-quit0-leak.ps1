# Does a swallowed `Quit(0)` under Windows PowerShell 5.1 actually leak a WINWORD,
# or does normal process exit release the RCW and let Word go anyway?
#
# The question matters because six probes in this repo ended with
# `try { $word.Quit(0) } catch { }`. `Quit(0)` is measured to throw under 5.1 and
# leave the process alive -- but that measurement was taken *while the parent
# process kept running*. If the RCW finalizer at process exit is enough to release
# Word, those probes are clean and the leak claim is false.
#
# RUN THIS UNDER `powershell.exe`, NOT `pwsh`. The binding failure is specific to
# Windows PowerShell 5.1; under PowerShell 7 every `Quit` form binds and both arms
# reap, so the probe is INCONCLUSIVE (exit 2) there. Running it in a 7.x shell is
# the exact mistake that hid this defect for the whole project.
#
#   powershell.exe -NoProfile -ExecutionPolicy Bypass -File probe-quit0-leak.ps1
#
# Attribution is by the sound route only: open a document, read
# ActiveWindow.Hwnd, resolve the owning pid via GetWindowThreadProcessId. Pid-set
# differencing is unsound on this machine (measured: 2 new pids for 1 instance
# created, under concurrent sessions) and is not used here.
#
# Arms:
#   A  swallowed Quit(0), then exit  -- the shape the probes used to ship
#   B  Quit(),            then exit  -- the control that must reap
#
# Exits 2 if the arms agree, because a probe whose arms agree has measured nothing.

$ErrorActionPreference = 'Stop'

$worker = @'
param([string]$Mode, [string]$OutFile)
$ErrorActionPreference = 'Stop'

Add-Type -Namespace Native -Name U -MemberDefinition @"
[DllImport("user32.dll")]
public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);
"@

try {
    # Publish a marker before COM activation. This cannot carry a pid -- attribution
    # needs ActiveWindow, which needs a document, so the pid does not exist yet --
    # but it lets the parent distinguish "hung before any Word existed" from "hung
    # after creating one it can no longer name". Measured: this window is real. A
    # seed block in probe-desktop.ps1 with the same ordering blocked in Documents.Add
    # for over nine minutes and left a Word (/Automation -Embedding) that no sound
    # route could attribute, and that outlived its client.
    [IO.File]::WriteAllText($OutFile, "START")

    $app = New-Object -ComObject Word.Application
    $app.Visible = $false
    $app.DisplayAlerts = 0

    # ActiveWindow throws with no document open, so one is required for attribution.
    $doc = $app.Documents.Add()
    $wordPid = 0
    [void][Native.U]::GetWindowThreadProcessId([IntPtr]$app.ActiveWindow.Hwnd, [ref]$wordPid)

    # An unattributed pid must abort rather than be polled. Pid 0 is the Idle
    # process, which never exits, so an unchecked 0 would report a false "leaked";
    # a pid that is not WINWORD would report a false "reaped" once it exits. Both
    # directions are wrong, so neither is allowed through.
    $proc = Get-Process -Id $wordPid -ErrorAction SilentlyContinue
    if ($wordPid -le 4 -or -not $proc -or $proc.ProcessName -ne 'WINWORD') {
        throw "attribution failed: hwnd resolved to pid $wordPid ($($proc.ProcessName))"
    }

    # Publish the pid at the first moment it can be known -- immediately after
    # attribution and before the Quit arm under test -- so a worker that hangs in
    # the measured section still leaves the parent able to name and clean up the
    # Word it started. Note the limit of that guarantee: everything above this line
    # can hang too, and a hang there is covered only by the START marker, not by a
    # pid. An earlier version of this comment claimed the pid was published before
    # anything that can hang, which the code above disproves.
    [IO.File]::WriteAllText($OutFile, "PID $wordPid")
    $threw = 'no'
    $why = ''
    # Close(<arg>) is measured to bind fine under 5.1 and is deliberately left in
    # the argument form: argument-less Close() prompts on a dirty document, and a
    # modal prompt in a hidden Word is a hang.
    try { $doc.Close(0) } catch { }
    if ($Mode -eq 'quit0') {
        # Exactly the shape the repo's probes shipped: argument form, swallowed.
        try { $app.Quit(0) } catch { $threw = 'yes'; $why = $_.Exception.Message }
    } else {
        try { $app.Quit() } catch { $threw = 'yes'; $why = $_.Exception.Message }
    }

    # Line-oriented, so the error text (which contains spaces, and is localized)
    # travels back intact rather than being shredded by a field split.
    [IO.File]::WriteAllLines($OutFile, @('OK', $wordPid, $threw, ($why -replace '\s+', ' ')))
    # Deliberately no ReleaseComObject and no kill: the point is what an ordinary
    # process exit does on its own.
}
catch {
    # A worker that dies silently is indistinguishable from one that measured
    # nothing, so the failure travels back on the same channel as a result.
    [IO.File]::WriteAllText($OutFile, "ERR $($_.Exception.Message -replace '\s+', ' ')")
    exit 1
}
'@

$workerPath = Join-Path $env:TEMP ("quit0worker-{0}.ps1" -f ([guid]::NewGuid().ToString('N')))
[IO.File]::WriteAllText($workerPath, $worker)

# Only ever act on a pid this probe minted and printed.
function Stop-OwnedWord {
    param([int]$WordPid, [string]$Why)
    if ($WordPid -le 4) { return }
    $proc = Get-Process -Id $WordPid -ErrorAction SilentlyContinue
    if ($proc -and $proc.ProcessName -eq 'WINWORD') {
        Write-Host ("cleanup: killing pid {0} ({1})" -f $WordPid, $Why)
        # Measured: Get-Process reports nothing once the process has gone, and
        # Kill() on an exited Process throws InvalidOperationException -- so the
        # check above can pass and the call still fail, purely from the race.
        # Swallowing the call is not enough on its own: a kill that failed for a
        # real reason would then be indistinguishable from one that raced a normal
        # exit, and this probe exists because of Words that outlive their owner.
        # So the exception is swallowed and the *outcome* is observed and reported.
        try { $proc.Kill() } catch { }
        $deadline = (Get-Date).AddSeconds(15)
        while ((Get-Date) -lt $deadline -and (Get-Process -Id $WordPid -ErrorAction SilentlyContinue).ProcessName -eq 'WINWORD') {
            Start-Sleep -Milliseconds 250
        }
        if ((Get-Process -Id $WordPid -ErrorAction SilentlyContinue).ProcessName -eq 'WINWORD') {
            Write-Host ("cleanup: pid {0} STILL ALIVE after kill -- leaked" -f $WordPid)
        }
        else {
            Write-Host ("cleanup: pid {0} is gone" -f $WordPid)
        }
    }
}

function Invoke-Arm {
    param([string]$Mode)

    $outFile = Join-Path $env:TEMP ("quit0out-{0}.txt" -f ([guid]::NewGuid().ToString('N')))
    # -File with discrete argv elements: no value crosses a command-line parser.
    $p = Start-Process -FilePath 'powershell.exe' `
        -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $workerPath, '-Mode', $Mode, '-OutFile', $outFile) `
        -WindowStyle Hidden -PassThru
    $p.WaitForExit(120000) | Out-Null

    $raw = if (Test-Path -LiteralPath $outFile) { ([IO.File]::ReadAllText($outFile)).Trim() } else { '' }
    Remove-Item -LiteralPath $outFile -ErrorAction SilentlyContinue
    $lines = $raw -split '\r?\n'
    $head = $lines[0].Split(' ')

    if (-not $p.HasExited) {
        # Same race as the Word kill above: HasExited can be false and the process
        # gone by the time Kill() runs.
        try { $p.Kill() } catch { }
        if ($head[0] -eq 'PID') { Stop-OwnedWord -WordPid ([int]$head[1]) -Why "$Mode worker hung" }
        elseif ($head[0] -eq 'START') {
            # Hung before attribution completed. A Word may exist and cannot be
            # named, so it is reported and deliberately not killed -- pid-set
            # differencing is unsound on a machine with a live external producer,
            # and killing an unattributable WINWORD can destroy another session's
            # work. Reporting an orphan beats guessing which one it is.
            Write-Host "WARNING: $Mode worker hung before attribution. If it had activated Word, that instance cannot be named and was NOT killed."
        }
        throw "$Mode worker did not exit within 120 s (last state: '$raw')"
    }
    if ($head[0] -eq 'ERR') { throw "$Mode worker failed: $raw" }
    if ($head[0] -ne 'OK') { throw "$Mode worker wrote no outcome (last state: '$raw')" }

    $wordPid = [int]$lines[1]
    $threw = $lines[2]
    $why = if ($lines.Count -gt 3) { $lines[3] } else { '' }

    # The worker process is gone. Poll the Word it owned to a deadline -- measured
    # Quit-to-exit is 2.7-6.1 s idle and longer under load, so a short sleep would
    # report a false "reaped". The name check stops a recycled pid from reading as
    # a survivor.
    $deadline = (Get-Date).AddSeconds(30)
    $alive = $true
    while ((Get-Date) -lt $deadline) {
        $proc = Get-Process -Id $wordPid -ErrorAction SilentlyContinue
        if (-not $proc -or $proc.ProcessName -ne 'WINWORD') { $alive = $false; break }
        Start-Sleep -Milliseconds 500
    }

    [pscustomobject]@{ Mode = $Mode; WordPid = $wordPid; Threw = $threw; AliveAfterExit = $alive; Error = $why }
}

$results = @()
try {
    foreach ($m in @('quit0', 'quit')) { $results += Invoke-Arm -Mode $m }
} finally {
    Remove-Item -LiteralPath $workerPath -ErrorAction SilentlyContinue
    # Runs even on a failed arm, so a thrown probe still tidies the arms that ran.
    foreach ($r in $results) {
        if ($r.AliveAfterExit) { Stop-OwnedWord -WordPid $r.WordPid -Why "leaked by arm $($r.Mode)" }
    }
}

$results | Format-Table -AutoSize Mode, WordPid, Threw, AliveAfterExit | Out-String | Write-Host
foreach ($r in $results) {
    if ($r.Error) { Write-Host ("{0}: Quit threw -> {1}" -f $r.Mode, $r.Error) }
}

$a = ($results | Where-Object Mode -eq 'quit0').AliveAfterExit
$b = ($results | Where-Object Mode -eq 'quit').AliveAfterExit
if ($a -eq $b) {
    Write-Host "INCONCLUSIVE: both arms agree (alive=$a). The probe separated nothing."
    Write-Host "PSVersion was $($PSVersionTable.PSVersion). Under 7.x every Quit form binds; use powershell.exe."
    exit 2
}
Write-Host "VERDICT: swallowed Quit(0) leaks past process exit = $a ; Quit() control leaks = $b"
