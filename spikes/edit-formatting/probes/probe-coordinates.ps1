# probe-coordinates.ps1 -- do character positions match the .NET string across
# the two hazards that pull oppositely?
#
#   table cell : Range.Text ends "\r" + chr(7); End-Start counts the end-of-cell
#                mark as ONE position but the string carries TWO chars, so
#                positions < string length. (Get-TextRange was written for this.)
#   field      : a HYPERLINK field's code occupies character positions absent
#                from Range.Text, so positions > string length.
#
# For each of ordinary / table-cell / hyperlink paragraphs, dump Range.Text
# length, Characters.Count, and End-Start; then compute a visible range TWO ways
# and print both, so we can see which (if either) is correct in all three:
#   (1) position-span trim: drop = (End-Start) - visible.Length; MoveEnd(-drop)
#   (2) Characters-based:   Range(chars[1].Start, chars[lastVisible].End)
#
# Cleanup kills at most one WINWORD: the instance this probe creates, attributed
# through its own window handle (ActiveWindow.Hwnd + GetWindowThreadProcessId)
# and torn down only through Stop-VerifiedWord, imported from
# spikes/isolation/probes/_common.ps1. A WINWORD census is differenced only to
# REPORT newcomers (Write-CensusSurvivors), never to authorise a kill: that
# inference is measured unsound on this shared machine (#136). A user's WINWORD
# is routinely alive here and must survive.
$ErrorActionPreference = 'Stop'
$WD_CHARACTER = 1
. (Join-Path $PSScriptRoot '..\..\isolation\probes\_common.ps1')
Add-Type -Namespace Win32 -Name Wnd -MemberDefinition @'
[DllImport("user32.dll", SetLastError = true)]
public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint lpdwProcessId);
'@
function Get-WordPids { @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }) | Sort-Object }
function Get-PidFromHwnd([IntPtr]$hwnd) {
    [uint32]$procId = 0
    [void][Win32.Wnd]::GetWindowThreadProcessId($hwnd, [ref]$procId)
    return [int]$procId
}
$pidsBefore = Get-WordPids
Write-Host "WINWORD before: $(if ($pidsBefore) { $pidsBefore -join ', ' } else { '(none)' })"

function Dump($label, $para, $doc) {
    $r = $para.Range
    $raw = [string]$r.Text
    $vis = $raw.TrimEnd("`r", [char]7, "`n")
    $chars = $r.Characters
    Write-Host "== $label =="
    Write-Host ("  Range.Text codes : {0}" -f (($raw.ToCharArray() | ForEach-Object { [int][char]$_ }) -join ','))
    Write-Host ("  Range.Text.Length={0}  visible.Length={1}  End-Start={2}  Characters.Count={3}" -f $raw.Length, $vis.Length, ($r.End - $r.Start), $chars.Count)

    # (1) position-span trim (the current Get-TextRange)
    $r1 = $para.Range
    $drop = ($r1.End - $r1.Start) - $vis.Length
    if ($drop -gt 0) { $r1.MoveEnd($WD_CHARACTER, -$drop) | Out-Null }
    Write-Host ("  (1) span-trim  -> Start={0} End={1} Text='{2}'" -f $r1.Start, $r1.End, (([string]$r1.Text) -replace "`r", '<CR>' -replace [string][char]7, '<BEL>'))

    # (2) Characters-based
    $count = $chars.Count
    $last = $count
    while ($last -ge 1) {
        $t = [string]$chars.Item($last).Text
        if ($t -eq "`r" -or $t -eq [string][char]7 -or $t -eq "`n") { $last-- } else { break }
    }
    if ($last -ge 1) {
        $r2 = $doc.Range($chars.Item(1).Start, $chars.Item($last).End)
        Write-Host ("  (2) chars-based-> Start={0} End={1} Text='{2}'" -f $r2.Start, $r2.End, (([string]$r2.Text) -replace "`r", '<CR>' -replace [string][char]7, '<BEL>'))
    } else {
        Write-Host "  (2) chars-based-> no visible characters"
    }
    Write-Host ''
}

$app = $null
$ownPid = 0
$ownStart = $null
try {
    $app = New-Object -ComObject Word.Application
    $app.Visible = $false

    # ordinary
    $doc = $app.Documents.Add()
    try {
        # Attribute this instance the moment a document makes ActiveWindow.Hwnd
        # resolvable, and record its StartTime while the pid is known live so
        # Stop-VerifiedWord can verify it at teardown. Never a census difference.
        try {
            $ownPid = Get-PidFromHwnd ([IntPtr][int64]$app.ActiveWindow.Hwnd)
            $ownStart = Get-WordStartTime $ownPid
            Write-Host "attributed this probe's WINWORD: pid $ownPid (start $ownStart)"
            Write-Host ''
        } catch {
            Write-Host "attribution FAILED -- $($_.Exception.Message.Split([char]10)[0])"
            Write-Host 'nothing will be killed; any survivor is reported instead.'
            Write-Host ''
        }
        $para = $doc.Paragraphs.Item(1)
        $para.Range.Text = 'alpha bravo charlie'
        Dump 'ordinary' $para $doc
    } finally { $doc.Close(0) }

    # table cell
    $doc = $app.Documents.Add()
    try {
        $tbl = $doc.Tables.Add($doc.Range(0, 0), 1, 2)
        $cell = $tbl.Cell(1, 1)
        $cellPara = $cell.Range.Paragraphs.Item(1)
        # Set cell text via the cell range trimmed of the end-of-cell mark.
        $cr = $cell.Range
        $cr.MoveEnd($WD_CHARACTER, -1) | Out-Null   # drop end-of-cell mark
        $cr.Text = 'cell 11'
        Dump 'table cell' ($cell.Range.Paragraphs.Item(1)) $doc
    } finally { $doc.Close(0) }

    # hyperlink
    $doc = $app.Documents.Add()
    try {
        $para = $doc.Paragraphs.Item(1)
        $para.Range.Text = 'alpha bravo charlie'
        $hs = $para.Range.Start + 'alpha '.Length
        $doc.Hyperlinks.Add($doc.Range($hs, $hs + 'bravo'.Length), 'https://example.com') | Out-Null
        Dump 'hyperlink' ($doc.Paragraphs.Item(1)) $doc
    } finally { $doc.Close(0) }

} catch {
    Write-Host "FAILED -- $($_.Exception.Message.Split([char]10)[0])"
} finally {
    if ($app) { try { $app.Quit() } catch {} }
}
# Poll for the attributed pid only (Quit returns before the process exits,
# measured 2.7-6.1 s idle and load-dependent), then decide through the shared
# helper -- which re-establishes identity (handle pin, name, StartTime) rather
# than trusting $ownPid, since the number could have been reused during the wait.
if ($ownPid -gt 0) {
    $deadline = (Get-Date).AddSeconds(60)
    while ((Get-Date) -lt $deadline) {
        if (-not (Get-Process -Id $ownPid -ErrorAction SilentlyContinue)) { break }
        Start-Sleep -Milliseconds 250
    }
    $outcome = Stop-VerifiedWord $ownPid $ownStart
    switch ($outcome) {
        'gone'   { Write-Host "cleanup: attributed WINWORD $ownPid exited on its own" }
        'killed' { Write-Host "cleanup: attributed WINWORD $ownPid was still up after 60 s; verified and killed" }
        default  { Write-Host "cleanup: Stop-VerifiedWord($ownPid) returned '$outcome' -- a Word may be left behind; confirm and close by hand" }
    }
} else {
    Write-Host 'cleanup: no pid was attributed, so nothing is killed'
}
# Reported, never acted on: a census difference may hold another session's Word.
Write-CensusSurvivors @(Get-WordPids | Where-Object { $pidsBefore -notcontains $_ -and $_ -ne $ownPid })
Write-Host "WINWORD after: $(if (Get-WordPids) { (Get-WordPids) -join ', ' } else { '(none)' })"
