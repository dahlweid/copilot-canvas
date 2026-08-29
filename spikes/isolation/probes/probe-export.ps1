# Probe: how expensive is a SINGLE-PAGE re-export via ExportAsFixedFormat?
# This is the gate between two very different architectures (plan phase 2b):
#   - if a one-page re-export is fast enough to hide behind an optimistic overlay,
#     option C wins (static colour-correct PDF display + COM edit channel)
#   - if it is slow, only pixel streaming can give a live editing feel
$ErrorActionPreference = 'Stop'

function Rep($l, $v) { Write-Output ("{0,-38} {1}" -f $l, $v) }

$out = "$env:TEMP\export-bench"
Remove-Item $out -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force -Path $out | Out-Null
$src = "$out\work.docx"
Copy-Item "$env:TEMP\desktop-probe.docx" $src

# Ownership detection: only ever kill a PID that did not exist before we started.
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
Rep "WINWORD pids before" ($(if ($before) { $before -join ',' } else { '(none)' }))

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$after = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$owned = @($after | Where-Object { $before -notcontains $_ })
Rep "owned pid" ($(if ($owned) { $owned -join ',' } else { '(attached - will NOT kill)' }))

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
    # The label states what was observed, never a cause the code cannot know.
    $deadline = (Get-Date).AddSeconds(30)
    foreach ($p in $owned) {
        while ((Get-Date) -lt $deadline -and (Get-Process -Id $p -ErrorAction SilentlyContinue).ProcessName -eq 'WINWORD') {
            Start-Sleep -Milliseconds 250
        }
        if ((Get-Process -Id $p -ErrorAction SilentlyContinue).ProcessName -eq 'WINWORD') {
            Stop-Process -Id $p -Force
            Rep "STILL ALIVE after 30 s, killed" $p
        }
        else { Rep "exited on Quit()" $p }
    }
    Remove-Item $out -Recurse -Force -ErrorAction SilentlyContinue
    $leftover = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
    Rep "WINWORD pids after cleanup" ($(if ($leftover) { $leftover -join ',' } else { '(none)' }))
}
