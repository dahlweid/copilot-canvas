# probe-run-formatting.ps1
#
# The question
# ------------
# `Set-ParagraphText` is a *whole-range* assignment:
#
#     function Set-ParagraphText($para, [string]$text) {
#         (Get-TextRange $para).Text = $text
#     }
#
# `replace_text` is built on it, so every text edit this repo performs replaces
# the whole paragraph even when the caller changed one word. Issue #166 asks
# whether the edit surface should gain finer operations. The argument against
# has been that `replace_text` already expresses the intent -- but that argument
# is only sound if the whole-range assignment costs nothing a finer operation
# would have kept.
#
# Intra-paragraph *run* formatting is exactly such a cost, and nothing in this
# repo measures it. Assigning `Range.Text` could plausibly
#
#   (a) preserve each run, since the character positions mostly line up; or
#   (b) collapse the paragraph to one run, taking the formatting of the first
#       character -- in which case bolding one word and then editing an
#       unrelated word in the same paragraph silently destroys the bold.
#
# (b) would make the whole-range assignment lossy, and a granular op would be
# buying something real rather than restating `replace_text`. So this is the
# load-bearing unknown behind #166, and it must be measured, not asserted.
#
# Arms
# ----
# All four run against one paragraph reading "alpha bravo charlie" with only
# "bravo" bold, so a collapse to the *first* character's formatting is visible
# as bold vanishing, and a collapse to bold-everything is equally visible. A
# probe that bolded the first word could not tell those apart.
#
#   A  identical text   -- assign exactly the text already there. If runs die
#                          even here, the loss is a property of the assignment
#                          and not of the text differing.
#   B  changed text     -- the real `replace_text` shape: one word altered.
#   C  sub-range        -- assign only the word being changed, leaving the rest
#                          of the paragraph untouched. This is the candidate a
#                          granular op would use, so it establishes whether the
#                          loss (if any) is *avoidable*.
#   D  Find/Replace     -- Word's own text-substitution route, as a second
#                          candidate that never assigns `.Text` at all.
#
# Reading the output
# -------------------
# Each arm prints a bold map: one character per character position, `B` where
# Font.Bold is set and `.` where it is not. The map is derived per character
# rather than from `Range.Bold`, because a mixed range returns wdUndefined
# (9999999) and would report "mixed" without saying where.
#
# Harness notes
# -------------
# Nothing is saved. `Document.SaveAs2` has wedged indefinitely in this repo's
# probes for reasons established as a script-shape artefact rather than a
# property of the COM calls (see probe-authoring.ps1), and this question is
# fully answerable in memory. Avoiding the save avoids that entire failure mode.
#
# Cleanup kills only WINWORD PIDs absent before the probe started. A WINWORD
# belonging to the user is routinely alive on this machine and must survive.

$ErrorActionPreference = 'Stop'

function Get-WordPids {
    @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }) | Sort-Object
}

$pidsBefore = Get-WordPids
Write-Host "WINWORD before: $(if ($pidsBefore) { $pidsBefore -join ', ' } else { '(none)' })"
Write-Host ''

# One character per character position: B where bold, . where not.
function Get-BoldMap($range) {
    $map = ''
    foreach ($ch in $range.Characters) {
        $t = [string]$ch.Text
        if ($t -eq "`r" -or $t -eq [string][char]7) { continue }
        $map += $(if ($ch.Font.Bold -ne 0) { 'B' } else { '.' })
    }
    return $map
}

# The repo's own trim: derived from the position span, never from string length,
# because inside a table End - Start counts the end-of-cell mark as one position
# while Range.Text carries it as two characters.
function Get-TextRange($para) {
    $range = $para.Range
    $visible = ([string]$range.Text).TrimEnd("`r", [char]7, "`n")
    $drop = ($range.End - $range.Start) - $visible.Length
    if ($drop -gt 0) { $range.MoveEnd(1, -$drop) | Out-Null }
    return $range
}

function Report($label, $range) {
    $text = ([string]$range.Text).TrimEnd("`r", [char]7, "`n")
    Write-Host ("  {0,-10} text = '{1}'" -f $label, $text)
    Write-Host ("  {0,-10} bold = {1}" -f '', (Get-BoldMap $range))
}

$app = $null
try {
    $app = New-Object -ComObject Word.Application
    $app.Visible = $false

    foreach ($arm in @('A', 'B', 'C', 'D')) {
        $doc = $app.Documents.Add()
        try {
            $para = $doc.Paragraphs.Item(1)
            $r = $para.Range
            $r.Text = 'alpha bravo charlie'

            # Bold exactly "bravo" -- the middle word, so neither a
            # first-character collapse nor a bold-everything collapse can be
            # mistaken for preservation.
            $start = $para.Range.Start + 'alpha '.Length
            $doc.Range($start, $start + 'bravo'.Length).Font.Bold = $true

            Write-Host "ARM $arm"
            Report 'before' (Get-TextRange $para)

            switch ($arm) {
                'A' { (Get-TextRange $para).Text = 'alpha bravo charlie' }
                'B' { (Get-TextRange $para).Text = 'alpha bravo delta' }
                'C' {
                    # Replace only "charlie", leaving the bold run untouched.
                    $s = $para.Range.Start + 'alpha bravo '.Length
                    $doc.Range($s, $s + 'charlie'.Length).Text = 'delta'
                }
                'D' {
                    # Replace must be passed explicitly. `Execute()` with no
                    # arguments defaults it to wdReplaceNone (0) and then
                    # returns True for the *find*, having changed nothing -- an
                    # arm that reported success while doing nothing. Position 11
                    # is Replace; position 10 is ReplaceWith.
                    $f = (Get-TextRange $para).Find
                    $f.ClearFormatting()
                    $f.Replacement.ClearFormatting()
                    $f.Forward = $true
                    $f.Wrap = 1          # wdFindContinue
                    $f.MatchCase = $false
                    $ok = $f.Execute('charlie', $false, $false, $false, $false,
                                     $false, $true, 1, $false, 'delta', 1)
                    Write-Host ("  {0,-10} Find.Execute returned {1}" -f 'note', $ok)
                }
            }

            Report 'after' (Get-TextRange $para)
            Write-Host ''
        } finally {
            $doc.Close(0)
        }
    }
} catch {
    Write-Host "FAILED -- $($_.Exception.Message.Split([char]10)[0])"
} finally {
    if ($app) {
        # Quit(), never Quit(<arg>): the argument form leaks under Windows
        # PowerShell 5.1 (probe-quit0-leak.ps1).
        try { $app.Quit() } catch { Write-Host "Quit() FAILED -- $($_.Exception.Message.Split([char]10)[0])" }
    }
}

Start-Sleep -Seconds 2
$leaked = @(Get-WordPids | Where-Object { $pidsBefore -notcontains $_ })
if ($leaked.Count -gt 0) {
    Write-Host "cleanup: killing WINWORD started by this probe: $($leaked -join ', ')"
    foreach ($p in $leaked) { try { Stop-Process -Id $p -Force } catch { } }
} else {
    Write-Host 'cleanup: no WINWORD left by this probe'
}
Write-Host "WINWORD after: $(if (Get-WordPids) { (Get-WordPids) -join ', ' } else { '(none)' })"
