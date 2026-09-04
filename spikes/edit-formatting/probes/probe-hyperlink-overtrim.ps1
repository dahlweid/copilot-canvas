# probe-hyperlink-overtrim.ps1 -- characterise the pre-#170 Get-TextRange
# over-trim on a paragraph carrying a HYPERLINK field, to decide whether it is a
# READ bug or a WRITE-only one.
#
# The shipped (main) helper trims a paragraph's Range down to its visible text by
# subtracting a *string length* from a *position span*:
#     $drop = ($range.End - $range.Start) - $visible.Length
#     if ($drop -gt 0) { $range.MoveEnd($WD_CHARACTER, -$drop) }
# A HYPERLINK field's code occupies character *positions* that never appear in
# Range.Text, so End-Start runs ahead of the string and $drop over-counts. This
# probe reproduces that helper verbatim (Get-TextRange-Old), alongside the
# shipping read expression and the #170 fix, and reports for each fixture:
#   - the coordinate facts (Range.Text.Length, End-Start, Characters.Count, drop)
#   - READ   : the text the read path yields ([string]Range.Text trimmed)
#   - OLD trim: Start/End/Text the over-trim collapses/truncates the range to
#   - OLD write: what (Get-TextRange-Old).Text = <new> does to the paragraph
#   - NEW write: what the #170 Set-ParagraphText does to the same paragraph
# across fixtures that vary the field-code length and the field's position, so we
# can see whether the over-count is fixture-specific and whether it collapses to
# empty or truncates by N.
#
# Cleanup kills at most one WINWORD: the instance this probe creates, attributed
# through its own window handle (ActiveWindow.Hwnd + GetWindowThreadProcessId) and
# torn down only through Stop-VerifiedWord, imported from
# spikes/isolation/probes/_common.ps1. A WINWORD census is differenced only to
# REPORT newcomers (Write-CensusSurvivors), never to authorise a kill -- that
# inference is measured unsound on this shared machine (#136). A user's WINWORD is
# routinely alive here and must survive.
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

# --- shipping expressions, copied so the probe measures the real code -----------

# main's OLD write helper -- the position-span trim under measurement.
function Get-TextRange-Old($para) {
    $raw = [string]$para.Range.Text
    $visible = $raw.TrimEnd("`r", [char]7, "`n")
    $range = $para.Range
    $drop = ($range.End - $range.Start) - $visible.Length
    if ($drop -gt 0) { $range.MoveEnd($WD_CHARACTER, -$drop) | Out-Null }
    return $range
}

# The read path's text source: reads never call Get-TextRange (it is write-only on
# main); they take [string]Range.Text and normalize it. This mirrors that source.
function Read-VisibleText($para) { ([string]$para.Range.Text).TrimEnd("`r", [char]7, "`n") }

# The #170 fix, copied verbatim from word-host.ps1 (Get-VisibleSpan / Get-TextRange
# / Set-ParagraphText), so NEW-write here is the shipping behaviour.
function Get-VisibleSpan($para) {
    $chars = $para.Range.Characters
    $last = $chars.Count
    while ($last -ge 1 -and ([string]$chars.Item($last).Text).TrimEnd("`r", [char]7, "`n").Length -eq 0) { $last-- }
    return @{ chars = $chars; last = $last }
}
function Get-TextRange-New($para) {
    $span = Get-VisibleSpan $para
    if ($span.last -lt 1) {
        $range = $para.Range.Duplicate
        $range.Collapse($WD_COLLAPSE_START)
        return $range
    }
    return $para.Range.Document.Range($span.chars.Item(1).Start, $span.chars.Item($span.last).End)
}
function Set-ParagraphText-New($para, [string]$text) {
    $span = Get-VisibleSpan $para
    $chars = $span.chars
    $last = $span.last
    if ($last -lt 1) { (Get-TextRange-New $para).Text = $text; return }
    $doc = $para.Range.Document
    $range = $doc.Range($chars.Item(1).Start, $chars.Item($last).End)
    $old = [string]$range.Text
    # Ordinal, here and in the scans below: -eq is case-insensitive and -ceq is
    # still culture-sensitive under de-DE (measured, probe-comparison-semantics.ps1),
    # either of which silently drops a real edit. Matches shipping Set-ParagraphText.
    if ([string]::Equals($old, $text, [System.StringComparison]::Ordinal)) { return }
    if ($old.Length -ne $last) { $range.Text = $text; return }
    $max = [Math]::Min($old.Length, $text.Length)
    $p = 0
    while ($p -lt $max -and [int]$old[$p] -eq [int]$text[$p]) { $p++ }
    $maxSuffix = $max - $p
    $s = 0
    while ($s -lt $maxSuffix -and [int]$old[$old.Length - 1 - $s] -eq [int]$text[$text.Length - 1 - $s]) { $s++ }
    $middle = $text.Substring($p, $text.Length - $p - $s)
    $startPos = if (($p + 1) -le $last) { $chars.Item($p + 1).Start } else { $chars.Item($last).End }
    $endPos = if (($last - $s) -ge 1) { $chars.Item($last - $s).End } else { $chars.Item(1).Start }
    $doc.Range($startPos, $endPos).Text = $middle
}

$SHOW = { param($t) (([string]$t) -replace "`r", '<CR>' -replace [string][char]7, '<BEL>') }

# Build a one-paragraph doc: text $before + $linked + $after, with a hyperlink
# whose address is $addr laid over $linked. Returns the doc (caller closes it).
function New-HyperlinkDoc($app, [string]$before, [string]$linked, [string]$after, [string]$addr) {
    $doc = $app.Documents.Add()
    $para = $doc.Paragraphs.Item(1)
    $para.Range.Text = "$before$linked$after"
    $ls = $para.Range.Start + $before.Length
    $doc.Hyperlinks.Add($doc.Range($ls, $ls + $linked.Length), $addr) | Out-Null
    return $doc
}

function Measure-Fixture($app, $label, $before, $linked, $after, $addr, $newText) {
    Write-Host "== $label =="
    Write-Host "  address = $addr"

    # facts + READ + OLD-trim coordinates: non-destructive (no assignment).
    $doc = New-HyperlinkDoc $app $before $linked $after $addr
    try {
        $para = $doc.Paragraphs.Item(1)
        $r = $para.Range
        $raw = [string]$r.Text
        $vis = $raw.TrimEnd("`r", [char]7, "`n")
        $drop = ($r.End - $r.Start) - $vis.Length
        Write-Host ("  Range.Text.Length={0}  End-Start={1}  Characters.Count={2}  drop=(End-Start)-visible.Length={3}" -f $raw.Length, ($r.End - $r.Start), $r.Characters.Count, $drop)
        Write-Host ("  READ  -> '{0}'  (len {1}; expected full visible text)" -f (& $SHOW (Read-VisibleText $para)), (Read-VisibleText $para).Length)
        $old = Get-TextRange-Old $para
        Write-Host ("  OLD trim -> Start={0} End={1} Text='{2}'" -f $old.Start, $old.End, (& $SHOW $old.Text))
    } finally { $doc.Close(0) }

    # OLD write: reproduce main's Set-ParagraphText on a fresh copy.
    $doc = New-HyperlinkDoc $app $before $linked $after $addr
    try {
        $para = $doc.Paragraphs.Item(1)
        try {
            (Get-TextRange-Old $para).Text = $newText
            $after1 = & $SHOW (([string]$doc.Paragraphs.Item(1).Range.Text).TrimEnd("`r", [char]7, "`n"))
            $links1 = $doc.Paragraphs.Item(1).Range.Hyperlinks.Count
            Write-Host ("  OLD write '{0}' -> paragraph now '{1}'  (hyperlinks left: {2})" -f $newText, $after1, $links1)
        } catch {
            Write-Host ("  OLD write '{0}' -> THREW: {1}" -f $newText, $_.Exception.Message.Split([char]10)[0])
        }
    } finally { $doc.Close(0) }

    # NEW write: the #170 fix on a fresh copy.
    $doc = New-HyperlinkDoc $app $before $linked $after $addr
    try {
        $para = $doc.Paragraphs.Item(1)
        try {
            Set-ParagraphText-New $para $newText
            $p2 = $doc.Paragraphs.Item(1)
            $after2 = & $SHOW (([string]$p2.Range.Text).TrimEnd("`r", [char]7, "`n"))
            $links2 = $p2.Range.Hyperlinks.Count
            $addrKept = if ($links2 -ge 1) { [string]$p2.Range.Hyperlinks.Item(1).Address } else { '(none)' }
            Write-Host ("  NEW write '{0}' -> paragraph now '{1}'  (hyperlinks left: {2}; addr: {3})" -f $newText, $after2, $links2, $addrKept)
        } catch {
            Write-Host ("  NEW write '{0}' -> THREW: {1}" -f $newText, $_.Exception.Message.Split([char]10)[0])
        }
    } finally { $doc.Close(0) }
    Write-Host ''
}

$app = $null
$ownPid = 0
$ownStart = $null
try {
    $app = New-Object -ComObject Word.Application
    $app.Visible = $false
    # Attribute this instance once a document makes ActiveWindow.Hwnd resolvable,
    # and record its StartTime while the pid is known live, so Stop-VerifiedWord
    # can verify it at teardown. Never a census difference.
    $seed = $app.Documents.Add()
    try {
        $ownPid = Get-PidFromHwnd ([IntPtr][int64]$app.ActiveWindow.Hwnd)
        $ownStart = Get-WordStartTime $ownPid
        Write-Host "attributed this probe's WINWORD: pid $ownPid (start $ownStart)"
        Write-Host ''
    } catch {
        Write-Host "attribution FAILED -- $($_.Exception.Message.Split([char]10)[0])"
        Write-Host 'nothing will be killed; any survivor is reported instead.'
        Write-Host ''
    } finally { $seed.Close(0) }

    # Edit changes only 'bravo' -> 'BRAVO' (equal length, entirely after the link
    # in F-end and spanning it elsewhere), so the reported paragraph text isolates
    # what the trim did, not what the replacement text was. A case-only edit is also
    # a test of the NEW column's ordinal comparison: under -eq it would be a no-op,
    # so a NEW paragraph that actually changed proves the compare is not -eq.
    Measure-Fixture $app 'F1 short address, link mid-paragraph' 'alpha ' 'bravo' ' charlie' 'https://a.co' 'alpha BRAVO charlie'
    Measure-Fixture $app 'F1b same fixture, edit OUTSIDE the link (alpha -> ALPHA)' 'alpha ' 'bravo' ' charlie' 'https://a.co' 'ALPHA bravo charlie'
    Measure-Fixture $app 'F2 long address, link mid-paragraph'  'alpha ' 'bravo' ' charlie' 'https://example.com/very/long/path?q=1&r=2' 'alpha BRAVO charlie'
    Measure-Fixture $app 'F3 link at paragraph end'             'alpha bravo ' 'charlie' '' 'https://example.com' 'alpha bravo CHARLIE'
    # F5: field code SHORTER than the visible text (drop < visible.Length), and the
    # edit lands outside the link. Shows what the over-trim does when it does not
    # overshoot to a collapse, and that the fix applies an outside-field edit.
    Measure-Fixture $app 'F5 long paragraph, short link, edit outside (the -> THE)' 'the quick brown ' 'fox' ' jumps over the lazy dog' 'https://a.co' 'THE quick brown fox jumps over the lazy dog'

    # F4: does the over-count scale with the *number* of fields? Report the drop
    # for one link vs two links over the same visible text. Facts only -- the
    # read/write behaviour is already characterised above. The rightmost link is
    # added first so the left field's inserted code cannot shift the right target.
    Write-Host '== F4 drop vs number of hyperlinks (same visible text) =='
    foreach ($n in 1, 2) {
        $doc = $app.Documents.Add()
        try {
            $para = $doc.Paragraphs.Item(1)
            $para.Range.Text = 'alpha bravo charlie'
            if ($n -eq 2) {
                $sc = $para.Range.Start + 'alpha bravo '.Length
                $doc.Hyperlinks.Add($doc.Range($sc, $sc + 'charlie'.Length), 'https://example.org') | Out-Null
            }
            $sb = $doc.Paragraphs.Item(1).Range.Start + 'alpha '.Length
            $doc.Hyperlinks.Add($doc.Range($sb, $sb + 'bravo'.Length), 'https://example.com') | Out-Null
            $r = $doc.Paragraphs.Item(1).Range
            $vis = ([string]$r.Text).TrimEnd("`r", [char]7, "`n")
            Write-Host ("  {0} link(s): End-Start={1}  visible.Length={2}  drop={3}  hyperlinks={4}" -f $n, ($r.End - $r.Start), $vis.Length, (($r.End - $r.Start) - $vis.Length), $r.Hyperlinks.Count)
        } finally { $doc.Close(0) }
    }
    Write-Host ''

} catch {
    Write-Host "FAILED -- $($_.Exception.Message.Split([char]10)[0])"
    Write-Host $_.ScriptStackTrace
} finally {
    if ($app) { try { $app.Quit() } catch { Write-Host "Quit() FAILED -- $($_.Exception.Message.Split([char]10)[0])" } }
}
# Poll for the attributed pid only (Quit returns before the process exits,
# measured 2.7-6.1 s idle and load-dependent, so a fixed sleep would guess against
# an unbounded tail -- poll instead), then decide through the shared helper, which
# re-establishes identity (handle pin, name, StartTime) rather than trusting the
# raw pid, since the number could have been reused during the wait.
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
