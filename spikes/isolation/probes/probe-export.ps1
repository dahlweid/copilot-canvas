# Probe: how expensive is a SINGLE-PAGE re-export via ExportAsFixedFormat?
# This is the gate between two very different architectures (plan phase 2b):
#   - if a one-page re-export is fast enough to hide behind an optimistic overlay,
#     option C wins (static colour-correct PDF display + COM edit channel)
#   - if it is slow, only pixel streaming can give a live editing feel
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot '_common.ps1')

function Rep($l, $v) { Write-Output ("{0,-38} {1}" -f $l, $v) }

$out = "$env:TEMP\export-bench"
Remove-Item $out -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force -Path $out | Out-Null
$src = "$out\work.docx"
Copy-Item "$env:TEMP\desktop-probe.docx" $src

# Census before, used ONLY as a negative -- to recognise a pid that already
# existed, never to select one. The comment here used to read "only ever kill a
# PID that did not exist before we started", which asserted an attribution this
# code does not have: differencing over-reports, so a pid absent from $before can
# still be another session's Word (#136). Nothing on this path kills anything.
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
Rep "WINWORD pids before" ($(if ($before) { $before -join ',' } else { '(none)' }))

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$after = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$appeared = @($after | Where-Object { $before -notcontains $_ })
Rep "WINWORD that appeared" ($(if ($appeared) { $appeared -join ',' } else { '(none)' }))

try {
    # ReadOnly:=$false because the edit-loop measurement has to actually modify the doc.
    $doc = $word.Documents.Open($src, $false, $false, $false)
    $pages = $doc.Content.Information(4)
    Rep "document pages" $pages

    # wdExportFormatPDF=17, wdExportAllDocument=0, wdExportFromTo=3
    function Export-Range($path, $rangeKind, $from, $to) {
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $doc.ExportAsFixedFormat($path, 17, $false, 0, $rangeKind, $from, $to, 0, $false, $false, 0)
        $sw.Stop()
        return $sw.Elapsed.TotalMilliseconds
    }

    Rep "first export (cold)" ("{0:F0} ms" -f (Export-Range "$out\warm.pdf" 0 1 1))

    $full = @(); for ($i = 0; $i -lt 3; $i++) { $full += Export-Range "$out\full$i.pdf" 0 1 1 }
    Rep "full document export (mean)" ("{0:F0} ms" -f (($full | Measure-Object -Average).Average))
    Rep "   full pdf size" ("{0:N0} bytes" -f (Get-Item "$out\full0.pdf").Length)

    $one = @(); for ($i = 0; $i -lt 5; $i++) { $one += Export-Range "$out\p1_$i.pdf" 3 1 1 }
    Rep "single page 1 export (mean)" ("{0:F0} ms" -f (($one | Measure-Object -Average).Average))
    Rep "   single-page pdf size" ("{0:N0} bytes" -f (Get-Item "$out\p1_0.pdf").Length)

    if ($pages -ge 7) {
        $mid = @(); for ($i = 0; $i -lt 5; $i++) { $mid += Export-Range "$out\p7_$i.pdf" 3 7 7 }
        Rep "single page 7 export (mean)" ("{0:F0} ms" -f (($mid | Measure-Object -Average).Average))
    }

    # The realistic edit loop: mutate the document, then re-export only the edited page.
    $loop = @()
    for ($i = 0; $i -lt 5; $i++) {
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $doc.Range(0, 0).InsertAfter("x")
        $doc.ExportAsFixedFormat("$out\edit$i.pdf", 17, $false, 0, 3, 1, 1, 0, $false, $false, 0)
        $sw.Stop(); $loop += $sw.Elapsed.TotalMilliseconds
    }
    Rep "EDIT + 1-page re-export (mean)" ("{0:F0} ms" -f (($loop | Measure-Object -Average).Average))
    Rep "EDIT + 1-page re-export (min)" ("{0:F0} ms" -f (($loop | Measure-Object -Minimum).Minimum))

    $doc.Saved = $true
    $doc.Close(0)
}
catch { Rep "ERROR" $_.Exception.Message }
finally {
    # Quit(), never Quit(<arg>): under Windows PowerShell 5.1 -- the runtime every
    # .ps1 here runs under -- the argument form throws and the Word survives, and
    # process exit does not reap it either (probe-quit0-leak.ps1). The catch
    # reports rather than swallows; a silent swallow is what hid this for months.
    try { $word.Quit() } catch { Rep "Quit() FAILED (Word may leak)" $_.Exception.Message.Split([char]10)[0] }
    try { [Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch { }
    # Application.Quit() returns long before its process exits (measured 2.7-6.1 s
    # idle, longer under load -- ADR 0005), so a fixed 2 s sleep made this sweep
    # the actual reaper on every run and hid whether Quit() worked at all. Poll to
    # a generous deadline instead; on success that costs only the real exit time.
    #
    # One deadline for the whole set, not one per pid. A per-pid deadline would
    # multiply an already generous budget by N; reusing a single deadline *inside*
    # a per-pid loop -- which is what this did first -- is worse, because every pid
    # after the first inherits only what its predecessors did not spend and can be
    # killed after a fraction of it while the label claims the full budget. Poll
    # the set, record when each pid actually went, and derive every label from that
    # record: a probe that prints a duration it did not measure has fabricated a
    # measurement, which is the one thing an instrument may never do.
    $deadline = (Get-Date).AddSeconds(30)
    $started = Get-Date
    $exitedAfterMs = @{}
    while ($true) {
        foreach ($p in $appeared) {
            if (-not $exitedAfterMs.ContainsKey($p)) {
                # Name-checked, so a recycled pid cannot read as a survivor.
                $proc = Get-Process -Id $p -ErrorAction SilentlyContinue
                if (-not $proc -or $proc.ProcessName -ne 'WINWORD') {
                    $exitedAfterMs[$p] = [int]((Get-Date) - $started).TotalMilliseconds
                }
            }
        }
        if ($exitedAfterMs.Count -eq $appeared.Count -or (Get-Date) -ge $deadline) { break }
        Start-Sleep -Milliseconds 250
    }
    $waitedMs = [int]((Get-Date) - $started).TotalMilliseconds
    $survivors = @()
    foreach ($p in $appeared) {
        if ($exitedAfterMs.ContainsKey($p)) {
            Rep "exited on Quit()" ("pid {0} after {1} ms" -f $p, $exitedAfterMs[$p])
        }
        else { $survivors += $p }
    }
    # $appeared is a census DIFFERENCE (see its assignment above), and this used to
    # force-kill whatever survived the poll. That is measured unsound (#136):
    # probe-init-attribution.ps1 differenced 2 new pids for 1 instance created,
    # and a census control saw 2 strangers' WINWORDs appear in a 40 s window with
    # nothing launched -- so a survivor here can be another session's Word, and
    # Stop-Process -Force destroys unsaved work with no prompt. The poll stays,
    # because it is the instrument: without it a survivor count is a stopwatch
    # artefact rather than a leak. What follows it is now a report.
    if ($survivors.Count -gt 0) { Rep "still alive after" ("{0} ms" -f $waitedMs) }
    Write-CensusSurvivors $survivors
    Remove-Item $out -Recurse -Force -ErrorAction SilentlyContinue
    $leftover = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
    Rep "WINWORD pids after cleanup" ($(if ($leftover) { $leftover -join ',' } else { '(none)' }))
}
