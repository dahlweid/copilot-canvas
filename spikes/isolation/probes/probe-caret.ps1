# Follow-up: can we get COLLAPSED-CARET geometry, not just paragraph bounds?
# The critique correctly points out probe-hittest measured an 811x78 rect, which is a
# paragraph, not a caret. Late-bound GetPoint also failed in probe-events. Settle both.
$ErrorActionPreference = 'Stop'
function Rep($l, $v) { Write-Output ("{0,-42} {1}" -f $l, $v) }

$out = "$env:TEMP\caret-probe"
Remove-Item $out -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force -Path $out | Out-Null
Copy-Item "$env:TEMP\desktop-probe.docx" "$out\work.docx"

$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$owned = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id | Where-Object { $before -notcontains $_ })

function Get-Pt($win, $obj) {
    $x = 0; $y = 0; $w = 0; $h = 0
    $win.GetPoint([ref]$x, [ref]$y, [ref]$w, [ref]$h, $obj)
    return @{ x = $x; y = $y; w = $w; h = $h }
}

try {
    $doc = $word.Documents.Open("$out\work.docx", $false, $true, $false)
    # A window must exist and be in print view for geometry to be meaningful.
    $win = $word.ActiveWindow
    $win.View.Type = 3          # wdPrintView
    Rep "view type" $win.View.Type

    foreach ($spec in @(@(120, 120), @(121, 121), @(200, 200), @(120, 140), @(120, 400))) {
        $r = $doc.Range($spec[0], $spec[1])
        try {
            $p = Get-Pt $win $r
            $kind = if ($spec[0] -eq $spec[1]) { "collapsed @$($spec[0])" } else { "span $($spec[0])-$($spec[1])" }
            Rep $kind ("x={0} y={1} w={2} h={3}" -f $p.x, $p.y, $p.w, $p.h)
        }
        catch { Rep "FAILED $($spec[0])-$($spec[1])" $_.Exception.Message.Split([char]10)[0] }
    }

    # Does a collapsed caret give zero width (a true caret) or the whole line?
    $a = Get-Pt $win $doc.Range(120, 120)
    $b = Get-Pt $win $doc.Range(121, 121)
    Rep "adjacent carets differ in x?" ("{0} -> {1}  (delta {2})" -f $a.x, $b.x, ($b.x - $a.x))
    Rep "collapsed caret width" $a.w
    Rep "collapsed caret height" $a.h

    # And the reverse mapping, which is what a click needs.
    try {
        $rp = $win.RangeFromPoint($a.x, $a.y)
        Rep "RangeFromPoint type" $rp.GetType().Name
        Rep "RangeFromPoint -> start" $rp.Start
        Rep "round-trip error (chars)" ($rp.Start - 120)
    }
    catch { Rep "RangeFromPoint FAILED" $_.Exception.Message.Split([char]10)[0] }

    $doc.Saved = $true; $doc.Close(0)
}
catch { Rep "ERROR" $_.Exception.Message.Split([char]10)[0] }
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
        foreach ($p in $owned) {
            if (-not $exitedAfterMs.ContainsKey($p)) {
                # Name-checked, so a recycled pid cannot read as a survivor.
                $proc = Get-Process -Id $p -ErrorAction SilentlyContinue
                if (-not $proc -or $proc.ProcessName -ne 'WINWORD') {
                    $exitedAfterMs[$p] = [int]((Get-Date) - $started).TotalMilliseconds
                }
            }
        }
        if ($exitedAfterMs.Count -eq $owned.Count -or (Get-Date) -ge $deadline) { break }
        Start-Sleep -Milliseconds 250
    }
    $waitedMs = [int]((Get-Date) - $started).TotalMilliseconds
    foreach ($p in $owned) {
        if ($exitedAfterMs.ContainsKey($p)) {
            Rep "exited on Quit()" ("pid {0} after {1} ms" -f $p, $exitedAfterMs[$p])
        }
        else {
            Stop-Process -Id $p -Force -ErrorAction SilentlyContinue
            Rep ("STILL ALIVE after {0} ms, killed" -f $waitedMs) $p
        }
    }
    Remove-Item $out -Recurse -Force -ErrorAction SilentlyContinue
    $left = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
    Rep "WINWORD after cleanup" ($(if ($left) { $left -join ',' } else { '(none)' }))
}
