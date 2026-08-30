# Can `Initialize-Word` learn the pid of the Word it just created, soundly?
#
# probe-word-ownership-hwnd.ps1 already established the mechanism:
# `Application.ActiveWindow.Hwnd` + `GetWindowThreadProcessId` yields a pid that
# is ours by construction, and it needs a document open because `ActiveWindow`
# throws without one. This probe measures the **shipping shape** instead of the
# mechanism -- the exact sequence word-host.ps1's `Initialize-Word` would run --
# because that is the thing the host's comment will cite, and the two differ in
# ways that matter:
#
#   * the host must open a document it did not want, purely to be attributed,
#     and close it again. What that costs, and whether Word is left usable, is
#     not something the mechanism probe asked.
#   * the host is a long-lived process whose `Visible` handling is gated on
#     ownership -- so whether a fresh automation instance starts hidden decides
#     whether a scratch document can flash on the user's screen.
#   * `Add-Type` compiles a P/Invoke at host startup, which is a cost the
#     mechanism probe never paid because it hoists it out of the measurement.
#
# Arms:
#   A  is a fresh `New-Object -ComObject Word.Application` already hidden?
#   B  does `ActiveWindow` throw on a hidden instance with no document -- the
#      exact state `Initialize-Word` is in, and the whole reason a scratch
#      document has to exist at all?
#   C  per-step cost of the whole init sequence, on this machine, right now.
#   D  under a Word created *concurrently by another process*, does hwnd name
#      exactly the instance we created while differencing cannot?
#
# Arms B and D are the ones that can come back the wrong way, and they are the
# reason to run this. If B does not throw, the scratch document is unnecessary
# and this design is more expensive than it needs to be. If D shows hwnd
# agreeing with the foreign instance, the mechanism is unsound and the issue
# needs a different answer entirely. A and C cannot fail informatively; they
# only price the thing.
#
# What this deliberately does NOT measure, because the design does not depend on
# it: reading `ActiveWindow` *late*. The host learns its pid once, during
# `Initialize-Word`, and caches it -- so teardown never has to ask a Word that
# may by then be hung, mid-Protected-View, or gone. That ordering is what keeps
# the awkward states out of the attribution path rather than any property of
# `ActiveWindow`, and a probe of a late read would be measuring a call the
# shipping code does not make. It matters if anyone ever moves the read: a
# Protected View window belongs to a *second* WINWORD the bridge never holds a
# handle to, so a late `ActiveWindow` could name the sandbox process.
#
# Safety: three other sessions drive Word on this machine. Nothing here is
# killed, ever, and an instance is quit ONLY when this probe attributed it by
# hwnd *and* its pid was absent from the baseline census. An instance that
# attributes to a pre-existing pid is released and left strictly alone -- a probe
# about attributing Word processes must not commit the act it exists to detect.

$ErrorActionPreference = 'Stop'

$typeMs = [int](Measure-Command {
    Add-Type -Namespace WordCanvasProbe -Name Win -MemberDefinition @'
[DllImport("user32.dll", SetLastError = true)]
public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint lpdwProcessId);
'@
}).TotalMilliseconds

function Get-WordPids { @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id) }

# Returns the pid behind a window handle, or 0 when it cannot be trusted.
# 0 is "did not resolve", never "probably ours": pids 0 and 4 are the idle and
# system processes, and a handle that has gone stale resolves to nothing.
function Get-PidFromHwnd($hwnd) {
    [uint32]$procId = 0
    [void][WordCanvasProbe.Win]::GetWindowThreadProcessId([IntPtr][int64]$hwnd, [ref]$procId)
    if ($procId -le 4) { return 0 }
    $p = Get-Process -Id ([int]$procId) -ErrorAction SilentlyContinue
    if ($null -eq $p -or $p.ProcessName -ne 'WINWORD') { return 0 }
    return [int]$procId
}

$baseline = Get-WordPids
Write-Host "baseline WINWORD count: $($baseline.Count)"
Write-Host "Add-Type (GetWindowThreadProcessId P/Invoke): ${typeMs}ms"
Write-Host "PSVersion: $($PSVersionTable.PSVersion)"

# Everything that creates a Word is recorded here as @{ app; pid }, so the
# `finally` can quit exactly what this probe made and nothing else.
$created = New-Object System.Collections.ArrayList
$holder = $null
$holderOut = Join-Path ([IO.Path]::GetTempPath()) ("initattr-holder-{0}.txt" -f [guid]::NewGuid().ToString('N'))
$holderRelease = Join-Path ([IO.Path]::GetTempPath()) ("initattr-release-{0}.txt" -f [guid]::NewGuid().ToString('N'))
$workerPath = Join-Path ([IO.Path]::GetTempPath()) ("initattr-worker-{0}.ps1" -f [guid]::NewGuid().ToString('N'))

# The foreign Word, standing in for another session. It attributes its own
# instance by hwnd and quits only that -- never anything it found running.
$worker = @'
param([string]$OutFile, [string]$ReleaseFile)
$ErrorActionPreference = 'Stop'
Add-Type -Namespace WordCanvasWorker -Name Win -MemberDefinition @"
[DllImport("user32.dll", SetLastError = true)]
public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint lpdwProcessId);
"@
$app = $null
try {
    $app = New-Object -ComObject Word.Application
    $app.Visible = $false
    $app.DisplayAlerts = 0
    $doc = $app.Documents.Add()
    [uint32]$wp = 0
    [void][WordCanvasWorker.Win]::GetWindowThreadProcessId([IntPtr][int64]$app.ActiveWindow.Hwnd, [ref]$wp)
    try { $doc.Close(0) } catch { }
    [IO.File]::WriteAllText($OutFile, "PID $wp")
    $deadline = (Get-Date).AddSeconds(180)
    while (-not (Test-Path -LiteralPath $ReleaseFile) -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 250 }
}
catch {
    try { [IO.File]::WriteAllText($OutFile, "ERR $($_.Exception.Message -replace '\s+', ' ')") } catch { }
}
finally {
    # Quit() with no argument: under Windows PowerShell 5.1 -- what this runs
    # under -- Quit(0) throws and leaves the process alive.
    if ($null -ne $app) {
        try { $app.Quit() } catch { }
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch { }
    }
}
'@
[IO.File]::WriteAllText($workerPath, $worker)

try {
    # --- Arm D's foreign Word starts first, so its creation overlaps ours -----
    # Started before our own New-Object rather than after, and deliberately
    # without waiting for it to be ready: the whole point is that both
    # activations are in flight at once, which is the condition under which
    # differencing was measured returning 2 new pids for 1 instance created.
    # -File with discrete argv elements, so no path crosses a command parser.
    $holder = Start-Process -FilePath 'powershell.exe' -PassThru -WindowStyle Hidden -ArgumentList @(
        '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', $workerPath,
        '-OutFile', $holderOut, '-ReleaseFile', $holderRelease
    )

    $before = Get-WordPids

    Write-Host ''
    Write-Host '=== Arm A + B + C: the shipping init sequence ==='

    $t = [Diagnostics.Stopwatch]::StartNew()
    $app = New-Object -ComObject Word.Application
    $newObjectMs = [int]$t.ElapsedMilliseconds
    [void]$created.Add(@{ app = $app; pid = 0 })

    # Read Visible BEFORE touching it. This is arm A, and it only means anything
    # unread: the host wants to know whether a scratch document could appear on
    # someone's screen, and setting the property first would answer its own
    # question.
    $visible = $null
    try { $visible = $app.Visible } catch { $visible = 'threw' }
    Write-Host "  A: Visible on a fresh automation instance = $visible"

    $t.Restart()
    try { $app.DisplayAlerts = 0 } catch { Write-Host "     DisplayAlerts threw -- $($_.Exception.Message.Split([char]10)[0])" }
    try { $app.AutomationSecurity = 3 } catch { Write-Host "     AutomationSecurity threw -- $($_.Exception.Message.Split([char]10)[0])" }
    $optionsMs = [int]$t.ElapsedMilliseconds

    # Arm B. This is the arm that can sink the design: if `ActiveWindow` yields
    # a handle here, with no document open, then the scratch document below is
    # pure cost and the sequence should read the handle straight after
    # New-Object. It is run before the Add, on the instance in exactly the state
    # `Initialize-Word` holds it.
    $noDocHwnd = $null
    try {
        $noDocHwnd = $app.ActiveWindow.Hwnd
        Write-Host "  B: ActiveWindow.Hwnd with NO document = $noDocHwnd"
        $noDocPid = Get-PidFromHwnd $noDocHwnd
        if ($noDocPid -gt 0) {
            Write-Host "     *** FINDING: it resolved to WINWORD pid $noDocPid. The scratch document is unnecessary. ***"
        }
        else {
            Write-Host '     but it did not resolve to a live WINWORD, so it attributes nothing.'
        }
    }
    catch {
        Write-Host "  B: ActiveWindow with no document throws -- $($_.Exception.Message.Split([char]10)[0])"
        Write-Host '     so a document is required, and the scratch Add below is load-bearing.'
    }

    $t.Restart()
    $doc = $app.Documents.Add()
    $addMs = [int]$t.ElapsedMilliseconds

    $t.Restart()
    $hwnd = $app.ActiveWindow.Hwnd
    $hwndMs = [int]$t.ElapsedMilliseconds

    $t.Restart()
    $ourPid = Get-PidFromHwnd $hwnd
    $mapMs = [int]$t.ElapsedMilliseconds
    $created[0].pid = $ourPid

    $t.Restart()
    try { $doc.Close(0) } catch { Write-Host "     scratch Close threw -- $($_.Exception.Message.Split([char]10)[0])" }
    $closeMs = [int]$t.ElapsedMilliseconds

    $docsLeft = -1
    try { $docsLeft = [int]$app.Documents.Count } catch { }

    Write-Host "  C: New-Object            ${newObjectMs}ms"
    Write-Host "     DisplayAlerts+Security ${optionsMs}ms"    Write-Host "     Documents.Add          ${addMs}ms"
    Write-Host "     ActiveWindow.Hwnd      ${hwndMs}ms  -> $hwnd"
    Write-Host "     hwnd -> pid            ${mapMs}ms  -> $ourPid"
    Write-Host "     scratch Close(0)       ${closeMs}ms  -> Documents.Count $docsLeft"
    Write-Host "     attribution total      $($optionsMs + $addMs + $hwndMs + $mapMs + $closeMs)ms on top of New-Object"

    # --- Arm C: what could differencing have concluded? ----------------------
    Write-Host ''
    Write-Host '=== Arm D: hwnd vs differencing, with a concurrent foreign Word ==='

    # Wait for the foreign worker to publish, so its pid is comparable. Its Word
    # was created concurrently with ours regardless of when it reports.
    $foreignPid = 0
    $foreignLine = ''
    $deadline = (Get-Date).AddSeconds(180)
    while ((Get-Date) -lt $deadline) {
        if (Test-Path -LiteralPath $holderOut) {
            $foreignLine = ([IO.File]::ReadAllText($holderOut)).Trim()
            if ($foreignLine) { break }
        }
        if ($holder.HasExited) { break }
        Start-Sleep -Milliseconds 250
    }
    $m = [regex]::Match($foreignLine, '^PID (\d+)$')
    if ($m.Success) { $foreignPid = [int]$m.Groups[1].Value }

    $new = @(Get-WordPids | Where-Object { $before -notcontains $_ })
    Write-Host "  differencing new pids: $($new -join ', ')  (count $($new.Count))"
    Write-Host "  hwnd says ours is:     $ourPid"
    Write-Host "  the foreign Word is:   $(if ($foreignPid -gt 0) { $foreignPid } else { "not reported ('$foreignLine')" })"

    # Every verdict below is refused unless the values it compares actually
    # parsed. A comparison against 0 is not a measurement. Arm C's census is
    # deliberately reported even when a verdict is refused.
    if ($ourPid -le 0) {
        Write-Host '  INCONCLUSIVE: hwnd did not resolve to a live WINWORD, so nothing here is comparable.'
    }
    elseif ($foreignPid -le 0) {
        Write-Host '  INCONCLUSIVE: the foreign worker never published a pid, so the concurrency arm measures nothing.'
    }
    else {
        if ($ourPid -eq $foreignPid) {
            Write-Host '  *** FINDING: our hwnd attributed to the foreign instance. The mechanism is NOT sound. ***'
        }
        else {
            Write-Host '  hwnd named a different instance from the foreign one, as required.'
        }
        if ($new.Count -gt 1) {
            Write-Host "  the ambiguous condition was reached: $($new.Count) new pids for 1 instance created."
            Write-Host "  the old shipping code picks `$new[0] = $($new[0]) -- $(if ($new[0] -eq $ourPid) { 'ours this run, by luck' } else { "NOT ours ($ourPid). Differencing would have quit and killed a foreign Word." })"
        }
        else {
            Write-Host '  the ambiguous condition was NOT reached this run: only one new pid, so'
            Write-Host '  differencing happened to agree. That is a null result about differencing,'
            Write-Host '  not a clearance -- re-run, or run it while another session drives Word.'
        }
        if ($new -notcontains $ourPid) {
            Write-Host '  *** FINDING: our own pid is absent from the differenced set. ***'
        }
    }
}
finally {
    Write-Host ''
    Write-Host '=== cleanup ==='
    foreach ($entry in $created) {
        # Quit only what this probe both created and attributed. An instance
        # whose pid was already alive at baseline is somebody else's by
        # definition, and is released rather than quit.
        $attributed = ($entry.pid -gt 0)
        $preexisting = $attributed -and ($baseline -contains $entry.pid)
        if ($preexisting) {
            Write-Host "  pid $($entry.pid) predates this probe -- released, NOT quit."
        }
        elseif (-not $attributed) {
            # Quit through the handle is still addressed to the object we made,
            # so it cannot reach a process we never bound to. Only a kill needs
            # a proven pid, and nothing here kills.
            Write-Host '  an instance could not be attributed -- quitting through its own handle, killing nothing.'
            try { $entry.app.Quit() } catch { Write-Host "    Quit threw -- $($_.Exception.Message.Split([char]10)[0])" }
        }
        else {
            try { $entry.app.Quit() } catch { Write-Host "    Quit threw -- $($_.Exception.Message.Split([char]10)[0])" }
        }
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($entry.app) | Out-Null }
        catch { Write-Host "    release threw -- $($_.Exception.Message.Split([char]10)[0])" }
    }
    [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()

    if ($holder) {
        try { [IO.File]::WriteAllText($holderRelease, '') }
        catch { Write-Host "  WARNING: could not release the holder -- it will time out, orphaning its Word." }
        $holder | Wait-Process -Timeout 60 -ErrorAction SilentlyContinue
        if (-not $holder.HasExited) {
            Write-Host '  WARNING: the holder did not exit in 60s. NOT killed -- killing a COM client orphans its Word.'
        }
    }
    Remove-Item -LiteralPath $workerPath, $holderOut, $holderRelease -Force -ErrorAction SilentlyContinue

    # Census. Differenced, and therefore unsound in exactly the way arm C
    # measures -- reported anyway, because over-reporting a leak is a false
    # alarm while under-reporting one is the defect. Nothing here is killed.
    Write-Host ''
    Write-Host '=== census ==='
    $survivors = @()
    for ($i = 0; $i -lt 180; $i++) {
        $survivors = @(Get-WordPids | Where-Object { $baseline -notcontains $_ })
        if ($survivors.Count -eq 0) { break }
        Start-Sleep -Milliseconds 500
    }
    if ($survivors.Count -eq 0) {
        Write-Host "  clean -- WINWORD population back to its $($baseline.Count) at start"
    }
    else {
        Write-Host "  $($survivors.Count) WINWORD alive 90s after teardown: $($survivors -join ', ')"
        Write-Host '  (differenced, so some may be another session s. Reported, not killed.)'
    }
    Write-Host "WINWORD alive after: $((Get-WordPids).Count) (baseline was $($baseline.Count))"
}
