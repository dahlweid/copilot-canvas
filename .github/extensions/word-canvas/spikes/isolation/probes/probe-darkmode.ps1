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
    try { $word.Quit(0) } catch { }
    try { [Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch { }
    Start-Sleep -Seconds 2
    foreach ($p in $owned) { if (Get-Process -Id $p -ErrorAction SilentlyContinue) { Stop-Process -Id $p -Force } }
}

# The theme itself is a per-user registry value Word reads at startup. Report, never change.
$k = 'HKCU:\Software\Microsoft\Office\16.0\Common'
$v = (Get-ItemProperty $k -ErrorAction SilentlyContinue).'UI Theme'
Rep "HKCU Office Common 'UI Theme'" ($(if ($null -ne $v) { "$v  (4 = Black, 5 = White)" } else { "(not set)" }))
