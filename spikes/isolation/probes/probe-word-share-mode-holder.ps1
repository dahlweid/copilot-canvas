# Worker half of probe-word-share-mode.ps1.
#
# Opens a document in its own Word instance, holds it, and signals the parent by
# creating the file named by -ReadyFile. Waits for -GoFile to appear, then shuts
# Word down and exits. Every wait is bounded.
#
# This lives in a separate file, and takes its paths as discrete -Parameter
# values, so that no path is ever interpolated into a command line. A path
# through a PowerShell single-quoted literal breaks on an apostrophe; the same
# path through cmd.exe is corrupted by &, ^ or a matched %VAR% pair. The two
# dangerous sets barely overlap, which is why escaping per site is the mistake
# and removing the parser is the fix.

param(
    [Parameter(Mandatory = $true)][string]$DocPath,
    [Parameter(Mandatory = $true)][string]$ReadyFile,
    [Parameter(Mandatory = $true)][string]$GoFile,
    [Parameter(Mandatory = $true)][string]$PidFile,
    [int]$HoldSeconds = 120
)

$ErrorActionPreference = 'Stop'

$before = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$word = $null
$ownedPid = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    $after = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $new = @($after | Where-Object { $before -notcontains $_ })
    # Differencing misattributes when another session creates a Word in the same
    # window -- measured elsewhere in this repo at 2 new pids for 1 instance
    # created. So this records the whole candidate set and claims ownership only
    # when it is unambiguous; the parent kills nothing it was not handed here.
    if ($new.Count -eq 1) {
        $ownedPid = $new[0]
        [IO.File]::WriteAllText($PidFile, [string]$ownedPid)
    } else {
        [IO.File]::WriteAllText($PidFile, '')
    }

    $doc = $word.Documents.Add()
    $doc.Content.Text = 'held open by Word'
    $doc.SaveAs2($DocPath, 16)   # 16 = wdFormatDocumentDefault (.docx)

    [IO.File]::WriteAllText($ReadyFile, 'ready')

    $deadline = (Get-Date).AddSeconds($HoldSeconds)
    while ((Get-Date) -lt $deadline -and -not (Test-Path -LiteralPath $GoFile)) {
        Start-Sleep -Milliseconds 100
    }

    $doc.Close(0)
    [Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
} finally {
    if ($null -ne $word) {
        try { $word.Quit(0) } catch { }
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch { }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        # Quit() returns in ~120ms and the process outlives it by seconds, longer
        # under load. Poll to a deadline; a flat sleep is a guess that is green on
        # a quiet machine and red on a busy one.
        if ($null -ne $ownedPid) {
            $d = (Get-Date).AddSeconds(90)
            while ((Get-Date) -lt $d -and $null -ne (Get-Process -Id $ownedPid -ErrorAction SilentlyContinue)) {
                Start-Sleep -Milliseconds 250
            }
        }
    }
}
