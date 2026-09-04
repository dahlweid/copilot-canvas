# make-fixture.ps1 — generates a multi-page .docx fixture for the smoke test.
# Not committed output: writes to the path given by -Out.
param(
    [Parameter(Mandatory = $true)][string]$Out,
    # More chapters => a longer document, so a test can produce a genuinely
    # different file at the same path.
    [int]$Chapters = 3,
    # Adds paragraphs with verbatim-identical text, both repeated under one
    # heading and repeated under two different headings. Off by default so the
    # existing suites see the document they were written against.
    #
    # This exists because addressing is derived -- heading path + text +
    # occurrence index -- and the document that proved the scheme had no
    # duplicate text at all, so the occurrence index was never actually tested.
    [switch]$Duplicates,
    # Adds a 2x2 table. Off by default so the existing suites see the document
    # they were written against.
    #
    # This exists because Word's own Document.Paragraphs collection counts one
    # extra paragraph per table row -- the end-of-row mark, which has no w:p in
    # the markup. Measured on a fixture with a 2x2 table and a text box: the map
    # sees 8 paragraphs and Word reports 10. Without a table in the fixture, an
    # edit that addressed Word by document-order position would pass every test
    # and corrupt every real document containing a table.
    [switch]$Table,
    # Adds one paragraph with a bold run sitting between two non-bold runs, so a
    # test can edit outside the bold span and read the .docx bytes back to see
    # whether the bold survived. Off by default so the existing suites see the
    # document they were written against.
    #
    # This exists for issue #170: the whole-range `Range.Text =` assignment
    # collapsed every paragraph to a single unformatted run on every text edit,
    # destroying intra-paragraph formatting, and no fixture carried a run for a
    # test to lose. Only "keepbold" is bold; the token is unique so the oracle
    # (docx-zip.mjs `boldText`) can address it without ambiguity.
    [switch]$Formatted
)

$ErrorActionPreference = 'Stop'

# Real umlauts, built from codepoints rather than typed into this file.
#
# The line below used to read "Umlaut check: Grosse Aepfel, Strasse, Muenchen",
# which is pure ASCII -- a fixture named for the property it did not exercise.
# Every fixture in this suite being ASCII is exactly why issue #40 survived two
# merged layers: the host decoded agent text as the OEM codepage, and no test
# ever sent it a character that could show the difference.
#
# Codepoints rather than literals because a fixture must not inherit the
# ambiguity it is measuring: PowerShell 5.1 reads a BOM-less .ps1 as the ANSI
# codepage, so a typed umlaut here would be decoded by the same class of rule
# that is under test, and a failure could not be attributed.
#
# The names are `lo`/`up` rather than `$ae`/`$AE` because PowerShell variable
# names are **case-insensitive**: `$AE = ...` reassigns `$ae`. Written the
# obvious way, this fixture emitted "M<U+00DC>nchen" and an all-uppercase
# "<U+00C4><U+00D6><U+00DC>", so the lowercase umlauts -- the more common half
# of the alphabet this is here to exercise -- never appeared in the document at
# all. Measured by reading the bytes back out of the .docx zip.
$loA = [char]0x00E4
$loO = [char]0x00F6
$loU = [char]0x00FC
$upA = [char]0x00C4
$upO = [char]0x00D6
$upU = [char]0x00DC
$sz = [char]0x00DF
$umlautCheck = "Umlaut check: Gro${sz}e ${upA}pfel, Stra${sz}e, M${loU}nchen, ${loA}${loO}${loU}${upA}${upO}${upU}${sz}."
# One paragraph carrying a token that is unique *and* non-ASCII, so a test can
# address it deterministically and a search can send non-ASCII across the same
# boundary it is checking.
$umlautMarker = "Gr${loU}${sz}e aus M${loU}nchen -- the UMLAUTMARKER paragraph."

$WD_STYLE_NORMAL = -1
$WD_STYLE_HEADING1 = -2
$WD_STYLE_HEADING2 = -3
$WD_FORMAT_DOCX = 16
$WD_DO_NOT_SAVE_CHANGES = 0

$dir = Split-Path -Parent $Out
if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }

$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
try {
    $doc = $w.Documents.Add()
    $sel = $w.Selection

    foreach ($chapter in 1..$Chapters) {
        $sel.Style = $WD_STYLE_HEADING1
        $sel.TypeText("Chapter $chapter")
        $sel.TypeParagraph()

        foreach ($section in 1..2) {
            $sel.Style = $WD_STYLE_HEADING2
            $sel.TypeText("Section $chapter.$section")
            $sel.TypeParagraph()

            $sel.Style = $WD_STYLE_NORMAL
            foreach ($p in 1..12) {
                $sel.TypeText("Paragraph $chapter.$section.$p - the quick brown fox jumps over the lazy dog. $umlautCheck")
                $sel.TypeParagraph()
            }
        }
    }

    if ($Duplicates) {
        foreach ($twin in 1..2) {
            $sel.Style = $WD_STYLE_HEADING1
            $sel.TypeText("Twin Chapter $twin")
            $sel.TypeParagraph()

            $sel.Style = $WD_STYLE_NORMAL
            # Same text three times under one heading: only the occurrence index
            # can tell these apart.
            foreach ($repeat in 1..3) {
                $sel.TypeText("DUPLICATE LINE: this sentence appears verbatim more than once.")
                $sel.TypeParagraph()
            }
            # Empty paragraphs, which are the degenerate duplicate.
            $sel.TypeParagraph()
            $sel.TypeParagraph()
        }
    }

    if ($Table) {
        $sel.Style = $WD_STYLE_HEADING1
        $sel.TypeText("Table Chapter")
        $sel.TypeParagraph()

        $sel.Style = $WD_STYLE_NORMAL
        $sel.TypeText("The table below exists so the paragraph join is tested against row marks.")
        $sel.TypeParagraph()

        # Not `$table`: PowerShell variable names are case-insensitive, so that
        # would assign a COM object to the -Table switch parameter and fail.
        $grid = $doc.Tables.Add($sel.Range, 2, 2)
        foreach ($row in 1..2) {
            foreach ($col in 1..2) {
                $grid.Cell($row, $col).Range.Text = "cell $row$col"
            }
        }
        # Past the table, so anything appended afterwards is not inside it.
        $sel.EndKey(6) | Out-Null
        $sel.TypeParagraph()
        $sel.Style = $WD_STYLE_NORMAL
        $sel.TypeText("This paragraph follows the table and is offset by its row marks.")
        $sel.TypeParagraph()
    }

    $sel.Style = $WD_STYLE_HEADING1
    $sel.TypeText("Findable Marker Heading")
    $sel.TypeParagraph()
    $sel.Style = $WD_STYLE_NORMAL
    $sel.TypeText("This paragraph contains the unique token ZORBLAX for search testing.")
    $sel.TypeParagraph()
    $sel.TypeText($umlautMarker)
    $sel.TypeParagraph()

    if ($Formatted) {
        $sel.Style = $WD_STYLE_HEADING1
        $sel.TypeText("Formatting Marker Heading")
        $sel.TypeParagraph()
        $sel.Style = $WD_STYLE_NORMAL
        # Only the middle run is bold. Toggled explicitly around each run so the
        # leading and trailing runs cannot inherit it.
        $sel.Font.Bold = $false
        $sel.TypeText("spanmarker ")
        $sel.Font.Bold = $true
        $sel.TypeText("keepbold")
        $sel.Font.Bold = $false
        $sel.TypeText(" endword")
        $sel.TypeParagraph()
        $sel.Font.Bold = $false
    }

    # `.Item(...)` on DocumentProperties throws under PowerShell; use late binding.
    function Set-DocProp($document, [string]$name, [string]$value) {
        $props = $document.BuiltInDocumentProperties
        $prop = [System.__ComObject].InvokeMember('Item', 'GetProperty', $null, $props, @($name))
        [System.__ComObject].InvokeMember('Value', 'SetProperty', $null, $prop, @($value)) | Out-Null
    }
    Set-DocProp $doc "Title" "Word Canvas Fixture"
    Set-DocProp $doc "Author" "word-canvas tests"

    $doc.SaveAs2($Out, $WD_FORMAT_DOCX)
    $pages = $doc.ComputeStatistics(2)
    $doc.Close($WD_DO_NOT_SAVE_CHANGES)
    Write-Output "fixture: $Out ($pages pages)"
} finally {
    $w.Quit()
    [Runtime.InteropServices.Marshal]::ReleaseComObject($w) | Out-Null
}
