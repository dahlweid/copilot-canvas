# Probe: can we force print colours (a white page) on a Word instance running under a
# user profile that has Office dark mode enabled?
#
# CORRECTNESS NOTE. The first version of this probe was meaningless. It read
# `$word.Options.$name` inside a try/catch and printed "NOT PRESENT" on throw. Late-bound
# COM property access returns $null for a missing name rather than throwing, so the catch
# branch was unreachable and every result was indistinguishable from "present but null".
# Presence is now tested with InvokeMember, which raises DISP_E_UNKNOWNNAME (0x80020006)
# for a name the object does not expose.
#
# SEPARATE FINDING, recorded here because it supersedes probe-bench.ps1's theme flip:
# setting HKCU\...\Common\UI Theme = 5 and holding it across Word's ENTIRE startup still
# produces a dark page (brightness 49.0, 92 distinct colours). probe-bench.ps1 restored the
# value after 3 s while Word was still 8+ seconds from ready, so its negative result was
# contaminated. The dark page is not fixable through this registry value.
$ErrorActionPreference = 'Continue'

function Rep($l, $v) { Write-Output ("{0,-38} {1}" -f $l, $v) }

function Test-ComProperty($obj, $name) {
    try {
        $v = $obj.GetType().InvokeMember($name, 'GetProperty', $null, $obj, @())
        return @{ present = $true; value = $v }
    }
    catch {
        $inner = $_.Exception
        while ($inner.InnerException) { $inner = $inner.InnerException }
        if ($inner.HResult -eq 0x80020006) { return @{ present = $false; value = $null } }
        return @{ present = $true; value = "<error: $($inner.Message.Split([char]10)[0])>" }
    }
}

$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$owned = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id | Where-Object { $before -notcontains $_ })
Rep "owned pid" ($owned -join ',')

try {
    foreach ($name in @('DarkModeDocumentColor', 'DarkMode', 'DisableDarkMode', 'UseDarkModeDocumentColor')) {
        $r = Test-ComProperty $word.Options $name
        if ($r.present) { Rep "Application.Options.$name" ("PRESENT = " + $r.value) }
        else { Rep "Application.Options.$name" "absent (DISP_E_UNKNOWNNAME)" }
    }

    $doc = $word.Documents.Add()
    $win = $word.ActiveWindow
    Rep "ActiveWindow.View.Type" $win.View.Type
    foreach ($name in @('DisplayBackgrounds', 'DisplayPageBoundaries')) {
        $r = Test-ComProperty $win.View $name
        if ($r.present) { Rep "View.$name" ("PRESENT = " + $r.value) }
        else { Rep "View.$name" "absent (DISP_E_UNKNOWNNAME)" }
    }
    $doc.Saved = $true
    $doc.Close(0)
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
}

# The theme itself is a per-user registry value Word reads at startup. Report, never change.
$k = 'HKCU:\Software\Microsoft\Office\16.0\Common'
$v = (Get-ItemProperty $k -ErrorAction SilentlyContinue).'UI Theme'
Rep "HKCU Office Common 'UI Theme'" ($(if ($null -ne $v) { "$v  (4 = Black, 5 = White)" } else { "(not set)" }))
