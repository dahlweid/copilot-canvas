# Follow-on to probe-word-ownership.ps1, which ruled out caption tagging: a
# hidden WINWORD reports an empty MainWindowTitle, so nothing we mint is visible
# from the process list.
#
# This asks the remaining question: can we get the pid from the COM object
# itself? `Application.Hwnd` does not exist on Word (PowerPoint has it, Word does
# not), but `Application.ActiveWindow.Hwnd` does -- once a document is open.
# GetWindowThreadProcessId on that handle yields a pid that is ours *by
# construction* rather than by inference, because the handle came out of our own
# instance.
#
# The cost question matters as much as the feasibility one: if it needs a
# document, it is free for a host that opens documents anyway and expensive for
# one that does not.

$ErrorActionPreference = 'Stop'

Add-Type -Namespace Win32 -Name Wnd -MemberDefinition @'
[DllImport("user32.dll", SetLastError = true)]
public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint lpdwProcessId);
'@

function Get-WordPids { @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id) }

function Get-PidFromHwnd([IntPtr]$hwnd) {
    [uint32]$procId = 0
    [void][Win32.Wnd]::GetWindowThreadProcessId($hwnd, [ref]$procId)
    return [int]$procId
}

$baseline = Get-WordPids
Write-Host "WINWORD alive before: $($baseline.Count)"

Write-Host ""
Write-Host "=== does ActiveWindow.Hwnd exist without a document? ==="
$before = Get-WordPids
$app = New-Object -ComObject Word.Application
$app.Visible = $false
try {
    $h = $app.ActiveWindow.Hwnd
    Write-Host "  ActiveWindow.Hwnd = $h  (unexpected -- no document is open)"
} catch {
    Write-Host "  throws, as expected: $($_.Exception.Message.Split([char]10)[0])"
}

Write-Host ""
Write-Host "=== with a document open ==="
$t0 = Get-Date
$doc = $app.Documents.Add()
$addMs = [int]((Get-Date) - $t0).TotalMilliseconds
$t1 = Get-Date
$hwnd = [IntPtr][int64]$app.ActiveWindow.Hwnd
$hwndMs = [int]((Get-Date) - $t1).TotalMilliseconds
$t2 = Get-Date
$pidFromHwnd = Get-PidFromHwnd $hwnd
$mapMs = [int]((Get-Date) - $t2).TotalMilliseconds

$new = @(Get-WordPids | Where-Object { $before -notcontains $_ })

Write-Host "  Documents.Add            ${addMs}ms"
Write-Host "  ActiveWindow.Hwnd        ${hwndMs}ms  -> $hwnd"
Write-Host "  GetWindowThreadProcessId ${mapMs}ms  -> pid $pidFromHwnd"
Write-Host "  differencing new pids:   $($new -join ', ')  (count $($new.Count))"

if ($new -contains $pidFromHwnd) {
    Write-Host "  AGREE: the hwnd pid is among the newly-appeared pids"
} else {
    Write-Host "  *** DISAGREE: hwnd says $pidFromHwnd, differencing says $($new -join ', ') ***"
}
if ($new.Count -eq 1 -and $new[0] -eq $pidFromHwnd) {
    Write-Host "  and differencing was unambiguous this run"
}

Write-Host ""
Write-Host "=== is it still the right pid for a document opened later? ==="
# The host opens the caller's document, not a blank one. Confirm the handle
# tracks whichever document is active rather than the one that created it.
$doc2 = $app.Documents.Add()
$hwnd2 = [IntPtr][int64]$app.ActiveWindow.Hwnd
$pid2 = Get-PidFromHwnd $hwnd2
Write-Host "  second document -> hwnd $hwnd2, pid $pid2"
Write-Host "  same pid as the first: $($pid2 -eq $pidFromHwnd)"

Write-Host ""
Write-Host "=== cleanup ==="
try { $doc.Close(0) } catch { }
try { $doc2.Close(0) } catch { }
# `Quit()`, not `Quit(0)`: the argument form binds `VARIANT*` parameters and was
# measured throwing "Argument 1 must be ... PSReference" in probe-word-ownership.ps1,
# inside a `catch { }` that hid it while every instance leaked.
try { $app.Quit() } catch { Write-Host "  Quit threw -- $($_.Exception.Message.Split([char]10)[0])" }
try { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch { }
[System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()

# Poll rather than sleep flat: Quit-to-exit is 2.7-6.1s idle and longer loaded.
for ($i = 0; $i -lt 60; $i++) {
    if (-not (Get-Process -Id $pidFromHwnd -ErrorAction SilentlyContinue)) { break }
    Start-Sleep -Milliseconds 250
}
$gone = -not (Get-Process -Id $pidFromHwnd -ErrorAction SilentlyContinue)
Write-Host "  our instance ($pidFromHwnd) exited: $gone"
Write-Host "WINWORD alive after: $((Get-WordPids).Count) (baseline was $($baseline.Count))"
