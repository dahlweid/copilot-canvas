# Shared helpers for the PowerPoint probes.
#
# Process hygiene is the reason this file exists. Other sessions on this machine
# drive Office concurrently, so a probe may only ever kill a POWERPNT.EXE that
# did not exist when it started. Snapshot first, kill only the difference.

function Say($s) { [Console]::WriteLine($s) }
function Rep($l, $v) { [Console]::WriteLine(("{0,-42} {1}" -f $l, $v)) }

function Get-PptPids {
    @(Get-Process POWERPNT -ErrorAction SilentlyContinue | ForEach-Object Id)
}

# Create a PowerPoint instance and report which PID (if any) we own.
function New-OwnedPowerPoint {
    $before = Get-PptPids
    $app = New-Object -ComObject PowerPoint.Application
    # TRAP: PowerPoint's alert enum is INVERTED relative to Word's.
    #   Word:       wdAlertsNone = 0, wdAlertsAll = -1
    #   PowerPoint: ppAlertsNone = 1, ppAlertsAll = 2
    # Porting `DisplayAlerts = 0` from the Word host would select an undefined
    # value, not "suppress everything".
    try { $app.DisplayAlerts = 1 } catch { }
    # PowerPoint refuses Visible = $false on most builds, so we do not set it
    # here; probe-hide.ps1 measures what is and is not possible.
    Start-Sleep -Milliseconds 300
    $after = Get-PptPids
    $owned = @($after | Where-Object { $before -notcontains $_ })
    [pscustomobject]@{ App = $app; Owned = $owned; Before = $before }
}

function Close-OwnedPowerPoint($ctx) {
    if ($null -eq $ctx) { return }
    # SAFETY: only quit an instance we actually created. PowerPoint hands back a
    # RUNNING instance from New-Object rather than starting a new process (see
    # probe-single-instance.ps1), so an unconditional Quit() here would close a
    # sibling session's -- or the user's -- PowerPoint, with their unsaved work.
    if ($ctx.App -and $ctx.Owned.Count -gt 0) {
        try { $ctx.App.Quit() } catch { }
    }
    elseif ($ctx.App) {
        Rep "NOT quitting" "attached to a pre-existing PowerPoint; leaving it alone"
    }
    try { if ($ctx.App) { [Runtime.InteropServices.Marshal]::ReleaseComObject($ctx.App) | Out-Null } } catch { }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Start-Sleep -Seconds 2
    foreach ($p in $ctx.Owned) {
        if (Get-Process -Id $p -ErrorAction SilentlyContinue) {
            Stop-Process -Id $p -Force -ErrorAction SilentlyContinue
            Rep "swept leaked owned pid" $p
        }
    }
    Start-Sleep -Milliseconds 800   # Stop-Process returns before the process is gone
}

# --- PDF inspection without a PDF library -------------------------------------
# Two facts are needed about an exported PDF: how many pages it has, and how big
# each page is. Both are needed to answer "one page per slide, page-accurately".
# PowerPoint emits /MediaBox and the page objects in the clear, so a byte scan is
# enough; these return $null / -1 rather than guessing if that stops holding.

function Get-PdfPageBoxes([string]$Path) {
    $bytes = [IO.File]::ReadAllBytes($Path)
    $text = [Text.Encoding]::GetEncoding('ISO-8859-1').GetString($bytes)
    $boxes = @()
    foreach ($m in [regex]::Matches($text, '/MediaBox\s*\[\s*([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s*\]')) {
        $boxes += [pscustomobject]@{
            Width  = [double]$m.Groups[3].Value - [double]$m.Groups[1].Value
            Height = [double]$m.Groups[4].Value - [double]$m.Groups[2].Value
        }
    }
    if ($boxes.Count -eq 0) { return $null }
    $boxes
}

function Get-PdfPageCount([string]$Path) {
    $bytes = [IO.File]::ReadAllBytes($Path)
    $text = [Text.Encoding]::GetEncoding('ISO-8859-1').GetString($bytes)
    # /Type /Page not followed by 's' -- excludes /Pages tree nodes.
    $n = ([regex]::Matches($text, '/Type\s*/Page(?![s])')).Count
    if ($n -gt 0) { return $n }
    $c = [regex]::Matches($text, '/Count\s+(\d+)')
    if ($c.Count -gt 0) { return [int](($c | ForEach-Object { [int]$_.Groups[1].Value } | Measure-Object -Maximum).Maximum) }
    return -1
}
