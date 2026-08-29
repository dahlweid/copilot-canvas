# Can we attribute a WINWORD.EXE to *our* COM instance soundly?
#
# Today both `Initialize-Word` (shipping) and `word-pids.mjs` (tests) attribute
# by PID-set differencing: snapshot before, create, and treat whatever is new as
# ours. That is unsound whenever anything else starts a Word inside the window,
# and the window is not small -- a cold `New-Object Word.Application` is ~4.5 s
# here, plus up to 1.5 s of polling afterwards. The consequence is not cosmetic:
# `Stop-Word` calls `$p.Kill()` on the pid it picked, and it picks `$new[0]`.
#
# Arms:
#   A  differencing under a *deliberately* concurrent start -- does the set ever
#      contain a pid that is not ours, and can it sort first?
#   B  caption tagging -- set `$App.Caption` to a GUID on a hidden instance and
#      look for it in the process list. If the hidden OpusApp window carries the
#      title, attribution becomes a fact we minted rather than an inference.
#
# Arm B is the one that matters: it either yields a sound mechanism or rules it
# out. Arm A only sizes a defect we already know the shape of.

$ErrorActionPreference = 'Stop'

function Get-WordPids { @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id) }

$created = @()

function New-TaggedWord([string]$tag) {
    $app = New-Object -ComObject Word.Application
    $app.Visible = $false
    try { $app.Caption = $tag } catch { Write-Host "    Caption set failed: $($_.Exception.Message)" }
    return $app
}

Write-Host "=== baseline ==="
$baseline = Get-WordPids
Write-Host "WINWORD alive before: $($baseline.Count)"

# --- Arm B: caption tagging --------------------------------------------------
Write-Host ""
Write-Host "=== Arm B: is a minted caption readable from the process list? ==="

$tagA = "word-canvas-probe-" + [guid]::NewGuid().ToString('N')
$t0 = Get-Date
$appA = New-TaggedWord $tagA
$created += , $appA
$startMs = [int]((Get-Date) - $t0).TotalMilliseconds
Write-Host "instance A created in ${startMs}ms, caption='$tagA'"

# Windows may not publish the title immediately; poll briefly like the shipping
# code polls for the pid, so the comparison is like-for-like.
$match = @()
for ($i = 0; $i -lt 20; $i++) {
    $match = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowTitle -like "*$tagA*" })
    if ($match.Count -gt 0) { break }
    Start-Sleep -Milliseconds 100
}

if ($match.Count -eq 1) {
    Write-Host "  RESULT: caption FOUND on pid $($match[0].Id) -- sound attribution is available"
} elseif ($match.Count -gt 1) {
    Write-Host "  RESULT: caption matched $($match.Count) processes -- ambiguous, NOT sound"
} else {
    Write-Host "  RESULT: caption NOT visible in the process list"
    Write-Host "  (titles seen: " + (@(Get-Process -Name WINWORD -ErrorAction SilentlyContinue |
            ForEach-Object { "'" + $_.MainWindowTitle + "'" }) -join ', ') + ")"
}

# Does a *second* hidden instance carry its own distinct caption? If both land
# on the same title the mechanism cannot separate two of our own instances.
$tagB = "word-canvas-probe-" + [guid]::NewGuid().ToString('N')
$appB = New-TaggedWord $tagB
$created += , $appB
Start-Sleep -Milliseconds 500
$matchB = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue |
    Where-Object { $_.MainWindowTitle -like "*$tagB*" })
$matchA2 = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue |
    Where-Object { $_.MainWindowTitle -like "*$tagA*" })
Write-Host "  after a second instance: A matches $($matchA2.Count), B matches $($matchB.Count)"
if ($matchA2.Count -eq 1 -and $matchB.Count -eq 1 -and $matchA2[0].Id -ne $matchB[0].Id) {
    Write-Host "  RESULT: two instances separate cleanly by caption"
}

# --- Arm A: how wide is the differencing window, and what gets caught in it? --
Write-Host ""
Write-Host "=== Arm A: differencing under a concurrent start ==="

$before = Get-WordPids
# A foreign Word, started by a *different process*, standing in for another
# session. Word is multi-instance, so this genuinely is a separate WINWORD.
$foreign = Start-Process powershell.exe -PassThru -WindowStyle Hidden -ArgumentList @(
    '-NoProfile', '-NonInteractive', '-Command',
    '$w = New-Object -ComObject Word.Application; $w.Visible = $false; Start-Sleep -Seconds 25'
)
Start-Sleep -Milliseconds 1200

$tagC = "word-canvas-probe-" + [guid]::NewGuid().ToString('N')
$appC = New-TaggedWord $tagC
$created += , $appC
Start-Sleep -Milliseconds 800

$new = @(Get-WordPids | Where-Object { $before -notcontains $_ })
$mineByCaption = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue |
    Where-Object { $_.MainWindowTitle -like "*$tagC*" })

Write-Host "differencing says new pids: $($new -join ', ')  (count $($new.Count))"
if ($mineByCaption.Count -eq 1) {
    $mine = $mineByCaption[0].Id
    Write-Host "caption says ours is:      $mine"
    $sorted = @($new | Sort-Object)
    Write-Host "shipping code would pick:  $($new[0])   (`$new[0], unsorted)"
    if ($new.Count -gt 1) {
        Write-Host "  RESULT: differencing caught $($new.Count) pids for ONE instance we created."
        if ($new[0] -ne $mine) {
            Write-Host "  RESULT: *** \$new[0] = $($new[0]) is NOT ours ($mine). Stop-Word would kill a foreign Word. ***"
        } else {
            Write-Host "  RESULT: \$new[0] happened to be ours this run -- but $($new.Count - 1) foreign pid(s) were candidates."
        }
    } else {
        Write-Host "  RESULT: only one new pid this run; the foreign Word did not land inside the window."
    }
} else {
    Write-Host "  (caption unavailable, so arm A cannot name the true owner)"
}

# --- cleanup -----------------------------------------------------------------
Write-Host ""
Write-Host "=== cleanup ==="
foreach ($a in $created) {
    try { $a.Quit(0) } catch { }
    try { [Runtime.InteropServices.Marshal]::ReleaseComObject($a) | Out-Null } catch { }
}
try { Stop-Process -Id $foreign.Id -Force -ErrorAction SilentlyContinue } catch { }
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
Start-Sleep -Seconds 3
Write-Host "WINWORD alive after: $((Get-WordPids).Count) (baseline was $($baseline.Count))"
