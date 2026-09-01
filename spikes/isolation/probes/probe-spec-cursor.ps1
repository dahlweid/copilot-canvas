# Why does the first block of a create_document spec disappear?
#
# Observed: authoring [heading "A"; paragraph "B"; heading "C"] produced a
# document whose first paragraph was "B", with "A" gone entirely. That is a
# silent content loss -- the caller asked for three blocks, got two, and nothing
# raised.
#
# This replicates Add-SpecBlock's cursor in isolation and prints the document
# after every step, so the step that loses the paragraph is visible rather than
# inferred.
#
# Run: powershell -NoProfile -ExecutionPolicy Bypass -File probe-spec-cursor.ps1

$ErrorActionPreference = 'Stop'

$WD_CHARACTER = 1
$WD_DO_NOT_SAVE_CHANGES = 0
$WD_STYLE_NORMAL = -1

function Show($doc, $label) {
    $n = $doc.Paragraphs.Count
    Write-Host ("-- {0}: {1} paragraph(s)" -f $label, $n)
    for ($i = 1; $i -le $n; $i++) {
        $p = $doc.Paragraphs.Item($i)
        $t = ([string]$p.Range.Text).TrimEnd("`r", [char]7, "`n")
        Write-Host ("     {0}: style={1} text={2}" -f $i, $p.Style.NameLocal, ($t | ConvertTo-Json))
    }
}

function Get-TextRange($para) {
    $raw = [string]$para.Range.Text
    $visible = $raw.TrimEnd("`r", [char]7, "`n")
    $range = $para.Range
    $drop = ($range.End - $range.Start) - $visible.Length
    if ($drop -gt 0) { $range.MoveEnd($WD_CHARACTER, -$drop) | Out-Null }
    return $range
}

$before = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$app = New-Object -ComObject Word.Application
$app.Visible = $false
$app.DisplayAlerts = 0

try {
    $doc = $app.Documents.Add()
    Show $doc 'fresh document'

    $state = @{ fresh = $true }
    $next = {
        if ($state.fresh) { $state.fresh = $false; return $doc.Paragraphs.Item(1) }
        return $doc.Paragraphs.Add()
    }

    Write-Host "`n== cursor identity, which is the whole question =="
    $p1 = & $next
    Write-Host ("   call 1: state.fresh now = {0}; range = {1}..{2}" -f $state.fresh, $p1.Range.Start, $p1.Range.End)
    (Get-TextRange $p1).Text = 'A (heading)'
    Show $doc 'after writing block 1'

    $p2 = & $next
    Write-Host ("   call 2: state.fresh now = {0}; range = {1}..{2}" -f $state.fresh, $p2.Range.Start, $p2.Range.End)
    (Get-TextRange $p2).Text = 'B (paragraph)'
    Show $doc 'after writing block 2'

    $p3 = & $next
    Write-Host ("   call 3: state.fresh now = {0}; range = {1}..{2}" -f $state.fresh, $p3.Range.Start, $p3.Range.End)
    (Get-TextRange $p3).Text = 'C (heading)'
    Show $doc 'after writing block 3'

    Write-Host "`n== does a style applied after the text move it? =="
    $p1b = $doc.Paragraphs.Item(1)
    $p1b.Range.Style = -2
    Show $doc 'after styling paragraph 1 as Heading 1'

    Write-Host "`n== Paragraphs.Add() vs Add(Range) on the last paragraph =="
    $added = $doc.Paragraphs.Add()
    Write-Host ("   Add() returned a paragraph at {0}..{1}; document now has {2}" -f $added.Range.Start, $added.Range.End, $doc.Paragraphs.Count)
    (Get-TextRange $added).Text = 'D (appended)'
    Show $doc 'after writing into the appended paragraph'

    $doc.Close($WD_DO_NOT_SAVE_CHANGES)
} finally {
    try { $app.Quit() } catch { }
    try { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch { }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()

    # $before is a census taken before this probe started Word, so $mine/$leaked
    # below are census DIFFERENCES. The comment here used to read "Only processes
    # this probe started, identified by differencing ... Never kill a Word this
    # script cannot attribute to itself", which asserts an attribution the code
    # does not have and cannot get: probe-init-attribution.ps1 differenced 2 new
    # pids for 1 instance, and a census control saw 2 strangers' WINWORDs appear
    # in a 40 s window with nothing launched (#136). Nothing here kills, so this
    # is a reporting inaccuracy rather than a destructive one -- but a safety
    # comment read as evidence of soundness is how the differencing rule spread
    # through this tree in the first place.
    #
    # The poll is still worth its 30 s: Quit() returns in ~120 ms and the process
    # outlives it by seconds, so a survivor observed without it would be a
    # stopwatch artefact rather than a leak.
    $deadline = [Diagnostics.Stopwatch]::StartNew()
    while ($deadline.ElapsedMilliseconds -lt 30000) {
        $mine = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id })
        if ($mine.Count -eq 0) { break }
        Start-Sleep -Milliseconds 250
    }
    $leaked = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id })
    if ($leaked.Count -gt 0) { Write-Host ("`nWINWORD still up that was not up before: {0} -- may be ours, may be another session's" -f ($leaked.Id -join ', ')) }
}
