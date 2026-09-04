# probe-minimal-span.ps1
#
# The repaired `replace_text`, measured across the three things issue #170 listed
# as unmeasured. `probe-run-formatting.ps1` established the base result: the
# whole-range assignment `(Get-TextRange $para).Text = $text` destroys every
# intra-paragraph run, even when the new text is identical to the old (arm A),
# and a sub-range assignment (arm C) preserves a run sitting outside the change.
# The fix diffs old visible text against new -- longest common prefix and suffix
# -- and assigns only the differing middle span.
#
# The functions below are a faithful copy of the shipping implementation in
# `.github/extensions/office-canvas/src/word/word-host.ps1` (Get-VisibleSpan,
# Get-TextRange, Set-ParagraphText). Copied, not imported, because a probe runs
# standalone under `powershell.exe -File`; if the implementation changes, this
# copy is meant to be updated alongside it and re-run.
#
# What this answers, that the base probe did not:
#
#   R0  identical text  -- the fix must be a no-op (base arm A, repaired).
#   R1  one word changed -- the real typo-fix shape (base arm B, repaired).
#   Q1  a changed span that STARTS inside a run and ends outside it.
#   Q2  formatting other than bold: italic, a hyperlink, a footnote reference --
#       the last also a drift test, since its mark is a character in the text.
#   Q3  new text SHORTER, EQUAL, and LONGER than the span it replaces.
#   T   a table cell, and a hyperlink paragraph, end to end: the fix must land
#       the edit correctly (no coordinate drift from the field code) and leave
#       the end-of-cell mark and the hyperlink intact.
#
# Reading the output: each arm prints the visible text and one map per axis under
# test -- one character per position, `X` where the attribute is set, `.` where
# not -- derived per character because a mixed Range.Bold returns wdUndefined and
# would say "mixed" without saying where.
#
# Harness: nothing is saved (answerable in memory; SaveAs2 has wedged in this
# repo's probes). Cleanup kills at most one WINWORD -- the instance this probe
# creates, attributed through its own window handle and torn down only through
# Stop-VerifiedWord (imported from spikes/isolation/probes/_common.ps1). A census
# is differenced only to REPORT newcomers, never to authorise a kill: #136
# measured that inference unsound on this shared machine. A user's WINWORD is
# routinely alive here and must survive. Quit() takes no argument (the argument
# form leaks under PowerShell 5.1).

$ErrorActionPreference = 'Stop'
$WD_CHARACTER = 1
$WD_COLLAPSE_START = 1
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
Write-Host ''

# ---- shipping implementation, copied verbatim in shape from word-host.ps1 ----

function Get-VisibleSpan($para) {
    $chars = $para.Range.Characters
    $last = $chars.Count
    while ($last -ge 1 -and ([string]$chars.Item($last).Text).TrimEnd("`r", [char]7, "`n").Length -eq 0) { $last-- }
    return @{ chars = $chars; last = $last }
}

function Get-TextRange($para) {
    $span = Get-VisibleSpan $para
    if ($span.last -lt 1) {
        $range = $para.Range.Duplicate
        $range.Collapse($WD_COLLAPSE_START)
        return $range
    }
    return $para.Range.Document.Range($span.chars.Item(1).Start, $span.chars.Item($span.last).End)
}

function Set-ParagraphText($para, [string]$text) {
    $span = Get-VisibleSpan $para
    $chars = $span.chars
    $last = $span.last
    if ($last -lt 1) { (Get-TextRange $para).Text = $text; return }

    $doc = $para.Range.Document
    $range = $doc.Range($chars.Item(1).Start, $chars.Item($last).End)
    $old = [string]$range.Text
    if ($old -eq $text) { return }
    if ($old.Length -ne $last) { $range.Text = $text; return }

    $max = [Math]::Min($old.Length, $text.Length)
    $p = 0
    while ($p -lt $max -and $old[$p] -eq $text[$p]) { $p++ }
    $maxSuffix = $max - $p
    $s = 0
    while ($s -lt $maxSuffix -and $old[$old.Length - 1 - $s] -eq $text[$text.Length - 1 - $s]) { $s++ }
    $middle = $text.Substring($p, $text.Length - $p - $s)

    $startPos = if (($p + 1) -le $last) { $chars.Item($p + 1).Start } else { $chars.Item($last).End }
    $endPos = if (($last - $s) -ge 1) { $chars.Item($last - $s).End } else { $chars.Item(1).Start }
    $doc.Range($startPos, $endPos).Text = $middle
}

# ---- measurement helpers ----

function Get-Map($range, [scriptblock]$pred) {
    $map = ''
    foreach ($ch in $range.Characters) {
        if (([string]$ch.Text).TrimEnd("`r", [char]7, "`n").Length -eq 0) { continue }
        $map += $(if (& $pred $ch) { 'X' } else { '.' })
    }
    return $map
}
$boldPred = { param($ch) $ch.Font.Bold -ne 0 }
$italicPred = { param($ch) $ch.Font.Italic -ne 0 }
$footPred = { param($ch) $ch.Footnotes.Count -gt 0 }

function Report($label, $para, [hashtable]$axes) {
    $r = Get-TextRange $para
    Write-Host ("  {0,-7} text = '{1}'" -f $label, ([string]$r.Text))
    foreach ($name in $axes.Keys) {
        Write-Host ("  {0,-7} {1,-6}= {2}" -f '', $name, (Get-Map $r $axes[$name]))
    }
}

# Fresh single-paragraph doc reading "alpha bravo charlie".
function New-Para($app) {
    $doc = $app.Documents.Add()
    $para = $doc.Paragraphs.Item(1)
    $para.Range.Text = 'alpha bravo charlie'
    return @{ doc = $doc; para = $para }
}
function Set-CharFormat($doc, $para, [string]$word, [string]$attr) {
    $start = $para.Range.Start + ('alpha bravo charlie'.IndexOf($word))
    $rng = $doc.Range($start, $start + $word.Length)
    if ($attr -eq 'bold') { $rng.Font.Bold = $true }
    elseif ($attr -eq 'italic') { $rng.Font.Italic = $true }
    elseif ($attr -eq 'hyper') { $doc.Hyperlinks.Add($rng, 'https://example.com') | Out-Null }
}

$app = $null
$ownPid = 0
$ownStart = $null
try {
    $app = New-Object -ComObject Word.Application
    $app.Visible = $false

    # Attribute this probe's WINWORD through its own window handle the moment a
    # document makes ActiveWindow.Hwnd resolvable, and record its StartTime while
    # the pid is known live. This is what licenses the single verified kill at
    # teardown; a census difference never does (#136).
    $scratch = $app.Documents.Add()
    try {
        $ownPid = Get-PidFromHwnd ([IntPtr][int64]$app.ActiveWindow.Hwnd)
        $ownStart = Get-WordStartTime $ownPid
        Write-Host "attributed this probe's WINWORD: pid $ownPid (start $ownStart)"
        Write-Host ''
    } catch {
        Write-Host "attribution FAILED -- $($_.Exception.Message.Split([char]10)[0])"
        Write-Host 'nothing will be killed; any survivor is reported instead.'
        Write-Host ''
    } finally { $scratch.Close(0) }

    # -- R0 / R1: identical text, and one word changed, bold on "bravo" --
    foreach ($c in @(
        @{ id = 'R0 identical text (base arm A, repaired)'; new = 'alpha bravo charlie' },
        @{ id = 'R1 one word changed (base arm B, repaired)'; new = 'alpha bravo delta' }
    )) {
        $h = New-Para $app
        try {
            Set-CharFormat $h.doc $h.para 'bravo' 'bold'
            $axes = @{ bold = $boldPred }
            Write-Host $c.id; Report 'before' $h.para $axes
            Set-ParagraphText $h.para $c.new
            Report 'after' $h.para $axes; Write-Host ''
        } finally { $h.doc.Close(0) }
    }

    # -- Q1: changed span starts inside the bold run and ends outside it --
    # prefix "alpha bra" is shared; the span "vo charlie" -> "in delta" begins 3
    # characters into the bold run.
    $h = New-Para $app
    try {
        Set-CharFormat $h.doc $h.para 'bravo' 'bold'
        $axes = @{ bold = $boldPred }
        Write-Host 'Q1 span starts inside bold run, ends outside'
        Report 'before' $h.para $axes
        Set-ParagraphText $h.para 'alpha brain delta'
        Report 'after' $h.para $axes
        Write-Host '  (bold on the shared prefix survives; the replacement takes the boundary run)'
        Write-Host ''
    } finally { $h.doc.Close(0) }

    # -- Q2 italic: edit a word outside the italic run --
    $h = New-Para $app
    try {
        Set-CharFormat $h.doc $h.para 'bravo' 'italic'
        $axes = @{ italic = $italicPred }
        Write-Host 'Q2 italic on "bravo", edit outside (charlie -> delta)'
        Report 'before' $h.para $axes
        Set-ParagraphText $h.para 'alpha bravo delta'
        Report 'after' $h.para $axes; Write-Host ''
    } finally { $h.doc.Close(0) }

    # -- Q2 hyperlink: edit a word outside the link; link must survive intact --
    $h = New-Para $app
    try {
        Set-CharFormat $h.doc $h.para 'bravo' 'hyper'
        # No per-character map here: Word does not surface a field-based
        # hyperlink through Characters.Item(n).Range.Hyperlinks, so a map would
        # read all-dots and mislead. The link object itself is the evidence.
        Write-Host 'Q2 hyperlink on "bravo", edit outside (charlie -> delta)'
        Report 'before' $h.para @{}
        Set-ParagraphText $h.para 'alpha bravo delta'
        Report 'after' $h.para @{}
        $doc = $h.doc
        if ($doc.Paragraphs.Item(1).Range.Hyperlinks.Count -gt 0) {
            $lnk = $doc.Paragraphs.Item(1).Range.Hyperlinks.Item(1)
            Write-Host ("  link text='{0}' addr='{1}'" -f $lnk.TextToDisplay, $lnk.Address)
        } else { Write-Host '  hyperlink DESTROYED' }
        Write-Host ''
    } finally { $h.doc.Close(0) }

    # -- Q2 footnote: a footnote reference mark sits in the text; edit later --
    $h = New-Para $app
    try {
        $doc = $h.doc; $para = $h.para
        $anchor = $doc.Range($para.Range.Start + 'alpha'.Length, $para.Range.Start + 'alpha'.Length)
        $doc.Footnotes.Add($anchor, '', 'a footnote') | Out-Null
        $axes = @{ foot = $footPred }
        Write-Host 'Q2 footnote ref after "alpha", edit later word (charlie -> delta)'
        Report 'before' $para $axes
        $cur = [string](Get-TextRange $para).Text
        Set-ParagraphText $para ($cur -replace 'charlie', 'delta')
        Report 'after' $para $axes
        Write-Host '  (footnote mark unmoved => no coordinate drift across the mark)'
        Write-Host ''
    } finally { $h.doc.Close(0) }

    # -- Q3: replacement shorter / equal / longer than the span --
    foreach ($c in @(
        @{ id = 'Q3 shorter (charlie -> hi)'; new = 'alpha bravo hi' },
        @{ id = 'Q3 equal   (charlie -> del515)'; new = 'alpha bravo delta78' },
        @{ id = 'Q3 longer  (charlie -> charlie-plus-tail)'; new = 'alpha bravo charlie-plus-tail' }
    )) {
        $h = New-Para $app
        try {
            Set-CharFormat $h.doc $h.para 'bravo' 'bold'
            $axes = @{ bold = $boldPred }
            Write-Host $c.id; Report 'before' $h.para $axes
            Set-ParagraphText $h.para $c.new
            Report 'after' $h.para $axes; Write-Host ''
        } finally { $h.doc.Close(0) }
    }

    # -- T table cell: the original "cell 11" -> "cell rewritten" case --
    $doc = $app.Documents.Add()
    try {
        $tbl = $doc.Tables.Add($doc.Range(0, 0), 1, 2)
        $cr = $tbl.Cell(1, 1).Range
        $cr.MoveEnd($WD_CHARACTER, -1) | Out-Null
        $cr.Text = 'cell 11'
        $para = $tbl.Cell(1, 1).Range.Paragraphs.Item(1)
        Write-Host 'T table cell  cell 11 -> cell rewritten'
        Write-Host ("  before text = '{0}'" -f ([string](Get-TextRange $para).Text))
        Set-ParagraphText $para 'cell rewritten'
        $raw = [string]$tbl.Cell(1, 1).Range.Paragraphs.Item(1).Range.Text
        Write-Host ("  after  text = '{0}'" -f ([string](Get-TextRange $tbl.Cell(1, 1).Range.Paragraphs.Item(1)).Text))
        Write-Host ("  end-of-cell mark chr(7) present: {0}; table intact Rows={1} Cols={2} Cells={3}" -f `
            ($raw.Contains([char]7)), $tbl.Rows.Count, $tbl.Columns.Count, $tbl.Range.Cells.Count)
        Write-Host ''
    } finally { $doc.Close(0) }

} catch {
    Write-Host "FAILED -- $($_.Exception.Message.Split([char]10)[0])"
    Write-Host $_.ScriptStackTrace
} finally {
    if ($app) { try { $app.Quit() } catch { Write-Host "Quit() FAILED -- $($_.Exception.Message.Split([char]10)[0])" } }
}

Start-Sleep -Seconds 3
# Poll for the attributed pid only (Quit returns before the process exits,
# measured 2.7-6.1 s idle and load-dependent), then decide through the shared
# helper, which re-establishes identity (handle pin, name, StartTime) rather than
# trusting $ownPid -- the number could have been reused during the wait.
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
