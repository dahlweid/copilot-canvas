# make-fixture.ps1 — generates a multi-page .docx fixture for the smoke test.
# Not committed output: writes to the path given by -Out.
param(
    [Parameter(Mandatory = $true)][string]$Out,
    # More chapters => a longer document, so a test can produce a genuinely
    # different file at the same path.
    [int]$Chapters = 3
)

$ErrorActionPreference = 'Stop'

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
                $sel.TypeText("Paragraph $chapter.$section.$p - the quick brown fox jumps over the lazy dog. Umlaut check: Grosse Aepfel, Strasse, Muenchen, aeoeue.")
                $sel.TypeParagraph()
            }
        }
    }

    $sel.Style = $WD_STYLE_HEADING1
    $sel.TypeText("Findable Marker Heading")
    $sel.TypeParagraph()
    $sel.Style = $WD_STYLE_NORMAL
    $sel.TypeText("This paragraph contains the unique token ZORBLAX for search testing.")
    $sel.TypeParagraph()

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
