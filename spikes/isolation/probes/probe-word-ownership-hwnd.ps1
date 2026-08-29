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
# Reporting catches, not swallowing ones. `Close(0)` was measured **under 5.1**,
# this file's actual runtime, and does *not* throw -- `Close(0)`, `Close()` and
# `Close([ref]0)` all returned OK and left `Documents.Count` at 0. So this is not
# a fix for a live leak, and saying otherwise would be the overclaiming-comment
# defect in the commit that cites it. It is the instrument.
#
# That asymmetry is itself measured, not assumed: `Quit(0)` throws under 5.1 and
# `Close(0)` does not, despite both taking by-value VARIANT arguments. Why they
# differ is unknown; that they differ is why neither may be inferred from the
# other.
#
# The argument is kept. Unlike `Quit`, `Document.Close()` with no argument
# *prompts* when the document is dirty, and a modal prompt in a hidden Word is a
# hang -- so `0` (wdDoNotSaveChanges) is load-bearing here rather than a default
# spelled out, and swapping it for the argument-less form to dodge a binding
# hazard would trade a measured non-problem for a real deadlock.
try { $doc.Close(0) } catch { Write-Host "  doc.Close threw -- $($_.Exception.Message.Split([char]10)[0])" }
try { $doc2.Close(0) } catch { Write-Host "  doc2.Close threw -- $($_.Exception.Message.Split([char]10)[0])" }
# `Quit()`, not `Quit(0)`. Under Windows PowerShell 5.1 -- this file's runtime --
# `Quit(0)` and `Quit($var)` both throw "Argument 1 must be ... PSReference" *and
# leave the process alive*; only the no-argument form binds and reaps it. Under
# PowerShell 7.6.5 all three forms work, which is exactly how this stayed hidden:
# the reduction that "cleared" it was run under 7.x.
try { $app.Quit() } catch { Write-Host "  Quit threw -- $($_.Exception.Message.Split([char]10)[0])" }
try { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch { Write-Host "  ReleaseComObject threw -- $($_.Exception.Message.Split([char]10)[0])" }
[System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()

# Poll rather than sleep flat: Quit-to-exit is 2.7-6.1s idle and longer loaded.
for ($i = 0; $i -lt 60; $i++) {
    if (-not (Get-Process -Id $pidFromHwnd -ErrorAction SilentlyContinue)) { break }
    Start-Sleep -Milliseconds 250
}
$gone = -not (Get-Process -Id $pidFromHwnd -ErrorAction SilentlyContinue)
Write-Host "  our instance ($pidFromHwnd) exited: $gone"
Write-Host "WINWORD alive after: $((Get-WordPids).Count) (baseline was $($baseline.Count))"
