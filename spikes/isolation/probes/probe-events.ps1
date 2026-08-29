# Probe: what does Word actually tell us when a document changes?
# The plan's change-notification design (and option C's re-export trigger) both depend on
# this. The critique asserts DocumentChange does NOT mean "content changed" and that no
# general scroll event exists. Settle it by measurement.
$ErrorActionPreference = 'Stop'

function Rep($l, $v) { Write-Output ("{0,-40} {1}" -f $l, $v) }

$out = "$env:TEMP\event-probe"
Remove-Item $out -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force -Path $out | Out-Null
$src = "$out\work.docx"
Copy-Item "$env:TEMP\desktop-probe.docx" $src

$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$owned = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id | Where-Object { $before -notcontains $_ })
Rep "owned pid" ($owned -join ',')

$script:log = [System.Collections.ArrayList]::new()
$subs = @()

try {
    # Which Application-level events does this Word even expose?
    $com = [Runtime.InteropServices.Marshal]::GetTypeFromCLSID
    $doc = $word.Documents.Open($src, $false, $false, $false)
    Rep "pages" $doc.Content.Information(4)

    foreach ($evt in @('DocumentChange', 'WindowSelectionChange', 'DocumentBeforeSave', 'WindowActivate')) {
        try {
            $s = Register-ObjectEvent -InputObject $word -EventName $evt -SourceIdentifier "w_$evt" -Action {
                [void]$script:log.Add("$($event.SourceIdentifier)")
            }
            $subs += "w_$evt"
            Rep "subscribed" $evt
        }
        catch { Rep "NOT AVAILABLE" "$evt -> $($_.Exception.Message.Split([char]10)[0])" }
    }
    # Events the plan/critique disputed the existence of:
    foreach ($evt in @('WindowScroll', 'DocumentChanged', 'ContentChange')) {
        try {
            Register-ObjectEvent -InputObject $word -EventName $evt -SourceIdentifier "x_$evt" -ErrorAction Stop | Out-Null
            $subs += "x_$evt"; Rep "EXISTS (disputed)" $evt
        }
        catch { Rep "does not exist (as critique said)" $evt }
    }

    function Drain($label) {
        Start-Sleep -Milliseconds 700
        $seen = $script:log -join ','
        Rep $label ($(if ($seen) { $seen } else { '(no events)' }))
        $script:log.Clear()
    }

    Drain "after open, baseline"

    Rep "--- typing into the document ---" ""
    Rep "doc.Saved before edit" $doc.Saved
    $doc.Range(0, 0).InsertAfter("hello")
    Rep "doc.Saved after edit" $doc.Saved
    Drain "events fired by a content edit"

    Rep "--- moving the selection ---" ""
    $word.Selection.SetRange(50, 60)
    Drain "events fired by a selection change"

    Rep "--- scrolling the window ---" ""
    $word.ActiveWindow.SmallScroll(3, 0, 0, 0)
    Drain "events fired by a scroll"

    # Cheap dirty-check candidates for polling, timed.
    $sw = [Diagnostics.Stopwatch]::StartNew(); for ($i = 0; $i -lt 50; $i++) { $null = $doc.Saved }; $sw.Stop()
    Rep "cost of reading doc.Saved" ("{0:F2} ms each" -f ($sw.Elapsed.TotalMilliseconds / 50))

    $sw = [Diagnostics.Stopwatch]::StartNew(); for ($i = 0; $i -lt 20; $i++) { $null = $doc.Content.Information(4) }; $sw.Stop()
    Rep "cost of reading page count" ("{0:F2} ms each" -f ($sw.Elapsed.TotalMilliseconds / 20))

    $sw = [Diagnostics.Stopwatch]::StartNew(); for ($i = 0; $i -lt 20; $i++) { $null = $word.ActiveWindow.VerticalPercentScrolled }; $sw.Stop()
    Rep "cost of VerticalPercentScrolled" ("{0:F2} ms each" -f ($sw.Elapsed.TotalMilliseconds / 20))

    # Does doc.Saved reset itself, i.e. is it a usable one-shot dirty flag?
    $doc.Saved = $true
    Rep "doc.Saved after manual reset" $doc.Saved
    $doc.Range(0, 0).InsertAfter("z")
    Rep "doc.Saved after another edit" $doc.Saved

    # Collapsed-caret geometry: the critique says probe-hittest measured a PARAGRAPH, not a caret.
    $r = $doc.Range(120, 120)   # collapsed range
    $x = 0; $y = 0; $w = 0; $h = 0
    $r.GetPoint([ref]$x, [ref]$y, [ref]$w, [ref]$h)
    Rep "collapsed range GetPoint (x,y,w,h)" "$x,$y,$w,$h"
    $r2 = $doc.Range(120, 140)  # short span on one line
    $x2 = 0; $y2 = 0; $w2 = 0; $h2 = 0
    $r2.GetPoint([ref]$x2, [ref]$y2, [ref]$w2, [ref]$h2)
    Rep "20-char range GetPoint (x,y,w,h)" "$x2,$y2,$w2,$h2"

    $doc.Saved = $true
    $doc.Close(0)
}
catch { Rep "ERROR" $_.Exception.Message }
finally {
    foreach ($s in $subs) { Unregister-Event -SourceIdentifier $s -ErrorAction SilentlyContinue }
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
