# Measures a repeated-start failure in the format-only Word Find outline scan
# and compares its cost with an OutlineLevel paragraph walk.
#
# Run only when Word is available:
#   powershell.exe -NoProfile -ExecutionPolicy Bypass -File probe-find-reanchor.ps1

[CmdletBinding()]
param(
    [switch]$SkipLargeBenchmarks
)

$ErrorActionPreference = 'Stop'
$WD_FIND_STOP = 0
$WD_INFO_ACTIVE_END_PAGE_NUMBER = 3

. (Join-Path $PSScriptRoot '..\..\isolation\probes\_common.ps1')

Add-Type -Namespace Win32 -Name Wnd -MemberDefinition @'
[DllImport("user32.dll", SetLastError = true)]
public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint lpdwProcessId);
'@

function Get-PidFromHwnd([IntPtr]$hwnd) {
    [uint32]$procId = 0
    [void][Win32.Wnd]::GetWindowThreadProcessId($hwnd, [ref]$procId)
    return [int]$procId
}

function Get-PageOf($range) {
    try { return [int]$range.Information($WD_INFO_ACTIVE_END_PAGE_NUMBER) } catch { return 0 }
}

function Get-OutlineByUnanchoredFind($doc, [int]$limit) {
    # Derived from Cmd-Outline (word-host.ps1) at origin/main revision
    # 5d81ffca4756505cd5a094a6258c003ee4f9bde7. The only adaptations are
    # documented by Write-PreFixControlDerivation before this control runs.
    $found = New-Object System.Collections.ArrayList

    foreach ($level in 1..9) {
        if ($found.Count -ge $limit) { break }
        $styleId = -1 - $level
        try {
            $rng = $doc.Content
            $f = $rng.Find
            $f.ClearFormatting()
            $f.Text = ''
            $f.Forward = $true
            $f.Wrap = 0
            $f.Format = $true
            $f.Style = $styleId
            $guard = 0
            while ($guard -lt $limit -and $f.Execute()) {
                $guard++
                $text = ([string]$rng.Text).Trim([char]13, [char]7, [char]11, [char]12, ' ')
                if (-not [string]::IsNullOrWhiteSpace($text)) {
                    [void]$found.Add(@{
                            level = $level
                            text  = $text
                            page  = (Get-PageOf $rng)
                            start = [int]$rng.Start
                        })
                }
                if ($found.Count -ge $limit) { break }
            }
        } catch {
            # A document that does not define this heading style simply has no
            # headings at this level.
        }
    }

    $sorted = @($found | Sort-Object { $_.start })
    return @{ headings = $sorted; count = $sorted.Count }
}

function Write-PreFixControlDerivation {
    Write-Host 'Pre-fix control source: Cmd-Outline (word-host.ps1) from origin/main revision 5d81ffca4756505cd5a094a6258c003ee4f9bde7.'
    Write-Host 'Pre-fix control adaptations: receives $doc and $limit instead of resolving request arguments; derives built-in heading styles numerically instead of calling Get-HeadingStyleId.'
}

function Get-OutlineByParagraph($doc, [int]$limit) {
    $headings = New-Object System.Collections.ArrayList
    $count = [int]$doc.Paragraphs.Count
    for ($i = 1; $i -le $count -and $headings.Count -lt $limit; $i++) {
        $para = $doc.Paragraphs.Item($i)
        $level = [int]$para.OutlineLevel
        if ($level -lt 10) {
            [void]$headings.Add(@{
                level = $level
                start = [int]$para.Range.Start
                text = ([string]$para.Range.Text).Trim([char]13, [char]7, ' ')
                page = (Get-PageOf $para.Range)
            })
        }
    }
    return @($headings | Sort-Object { $_.start })
}

function Get-OutlineByParagraphEnumerator($doc, [int]$limit) {
    $headings = New-Object System.Collections.ArrayList
    foreach ($para in $doc.Paragraphs) {
        if ($headings.Count -ge $limit) { break }
        $level = [int]$para.OutlineLevel
        if ($level -lt 10) {
            [void]$headings.Add(@{
                    level = $level
                    start = [int]$para.Range.Start
                    text = ([string]$para.Range.Text).Trim([char]13, [char]7, ' ')
                    page = (Get-PageOf $para.Range)
                })
        }
    }
    return @($headings | Sort-Object { $_.start })
}

function Get-OutlineByXml($doc, [int]$limit) {
    $out = Join-Path ([IO.Path]::GetTempPath()) "outline-$([guid]::NewGuid().ToString('N')).xml"
    $parser = Join-Path $PSScriptRoot 'outline-from-xml.mjs'
    try {
        [IO.File]::WriteAllText($out, [string]$doc.Content.WordOpenXML, (New-Object System.Text.UTF8Encoding($false)))
        $records = @([IO.File]::ReadAllText($out) | & node $parser | ConvertFrom-Json)
        if ($LASTEXITCODE -ne 0) { throw 'The XML outline parser failed.' }

        $headings = New-Object System.Collections.ArrayList
        foreach ($record in $records) {
            if ($headings.Count -ge $limit) { break }
            $para = $doc.Paragraphs.Item([int]$record.wordIndex)
            [void]$headings.Add(@{
                    level = [int]$record.headingLevel
                    start = [int]$para.Range.Start
                    text = [string]$record.text
                    page = (Get-PageOf $para.Range)
                })
        }
        return @($headings | Sort-Object { $_.start })
    } finally {
        if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -Force }
    }
}

function Get-SearchStarts($doc, [string]$query, [int]$limit) {
    $rng = $doc.Content
    $find = $rng.Find
    $find.ClearFormatting()
    $find.Text = $query
    $find.Forward = $true
    $find.Wrap = 0
    $find.Format = $false
    $find.MatchCase = $false
    $find.MatchWholeWord = $false

    $starts = New-Object System.Collections.ArrayList
    while ($starts.Count -lt $limit -and $find.Execute()) {
        [void]$starts.Add([int]$rng.Start)
    }
    return @($starts)
}

function Test-StrictStarts($starts) {
    for ($i = 1; $i -lt $starts.Count; $i++) {
        if ($starts[$i] -le $starts[$i - 1]) { return $false }
    }
    return $true
}

function Measure-OutlineScan([string]$name, [scriptblock]$scan) {
    $stopwatch = [Diagnostics.Stopwatch]::StartNew()
    $headings = @(& $scan)
    $stopwatch.Stop()
    Write-Host "${name}: $($stopwatch.ElapsedMilliseconds) ms; headings: $($headings.Count)"
    return @{ headings = $headings; elapsedMilliseconds = $stopwatch.ElapsedMilliseconds }
}

function Assert-EquivalentOutlines([string]$name, $left, $right) {
    $leftStarts = @($left | ForEach-Object { $_.start })
    $rightStarts = @($right | ForEach-Object { $_.start })
    if ($leftStarts.Count -ne $rightStarts.Count -or ($leftStarts -join ',') -cne ($rightStarts -join ',')) {
        throw "$name returned different headings: left starts [$($leftStarts -join ', ')]; right starts [$($rightStarts -join ', ')]."
    }
}

function Add-HeadingFixture($doc) {
    $lines = @(
        @{ text = 'Plain heading'; level = 1 },
        @{ text = 'Plain body follows'; level = 10 },
        @{ text = 'Heading before table'; level = 1 },
        @{ text = 'Heading after table'; level = 1 },
        @{ text = 'needle alpha'; level = 10 },
        @{ text = 'needle beta'; level = 10 },
        @{ text = 'needle gamma'; level = 10 }
    )
    $doc.Content.Text = (($lines | ForEach-Object { $_.text }) -join "`r") + "`r"
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i].level -lt 10) {
            $doc.Paragraphs.Item($i + 1).Range.Style = -1 - $lines[$i].level
        }
    }

    # Insert the document's only table directly after the preceding heading.
    # The cell heading and post-table heading distinguish table entry from exit.
    $tableAt = [int]$doc.Paragraphs.Item(4).Range.Start
    $table = $doc.Tables.Add($doc.Range($tableAt, $tableAt), 2, 1)
    $table.Cell(1, 1).Range.Text = 'Heading in table cell'
    $table.Cell(1, 1).Range.Paragraphs.Item(1).Range.Style = -2
    $table.Cell(2, 1).Range.Text = 'Table body'
}

function Add-LargePlainFixture($doc, [int]$paragraphCount) {
    $lines = 1..$paragraphCount | ForEach-Object { "Body paragraph $_" }
    $doc.Content.Text = ($lines -join "`r") + "`r"

    $headingCount = 0
    for ($i = 250; $i -le $paragraphCount; $i += 250) {
        [void]($doc.Paragraphs.Item($i).Range.Style = -2)
        $headingCount++
    }
    return $headingCount
}
$app = $null
$doc = $null
$ownPid = 0
$ownStart = $null
try {
    $app = New-Object -ComObject Word.Application
    $app.Visible = $false
    $doc = $app.Documents.Add()
    $ownPid = Get-PidFromHwnd ([IntPtr][int64]$app.ActiveWindow.Hwnd)
    if ($ownPid -le 0) { throw 'Could not attribute the probe Word through ActiveWindow.Hwnd.' }
    $ownStart = Get-WordStartTime $ownPid
    Write-Host "Attributed WINWORD: pid $ownPid; start: $ownStart"

    Add-HeadingFixture $doc
    $preFixLimit = 20
    Write-PreFixControlDerivation
    $tableParagraph = Measure-OutlineScan 'Table fixture paragraph walk' { Get-OutlineByParagraph $doc $preFixLimit }
    $byParagraph = $tableParagraph.headings
    $tableFind = Measure-OutlineScan 'Table fixture format-only Find' { (Get-OutlineByUnanchoredFind $doc $preFixLimit).headings }
    $preFix = $tableFind.headings
    $preFixStarts = @($preFix | ForEach-Object { $_.start })
    $preFixStrict = Test-StrictStarts $preFixStarts
    $preFixRepeated = -not $preFixStrict

    Write-Host "Pre-fix Find headings: $($preFix.Count); paragraph headings: $($byParagraph.Count); cap: $preFixLimit; repeated start: $preFixRepeated"
    Write-Host "Pre-fix Find starts: $($preFixStarts -join ', ')"
    if (-not $preFixRepeated) {
        throw 'The unanchored Find control did not reproduce a repeated start. Strengthen the fixture before treating this probe as evidence for the fix.'
    }
    Write-Host 'PASS: The unanchored Find control reproduced the repeated-start failure.'

    $doc.Saved = $true
    $savedBeforeMarkup = [bool]$doc.Saved
    $tableXml = Measure-OutlineScan 'Table fixture XML outline' { Get-OutlineByXml $doc $preFixLimit }
    $savedAfterMarkup = [bool]$doc.Saved
    Write-Host "Document saved before XML outline: $savedBeforeMarkup; after: $savedAfterMarkup"
    if (-not $savedBeforeMarkup -or -not $savedAfterMarkup) {
        throw 'Reading and writing the XML outline marked the probe document unsaved.'
    }
    Assert-EquivalentOutlines 'Table fixture XML comparison' $byParagraph $tableXml.headings
    Write-Host "PASS: Table fixture XML outline matches the paragraph oracle on $($tableXml.headings.Count) headings."

    $searchStarts = Get-SearchStarts $doc 'needle' 20
    $searchStrict = Test-StrictStarts $searchStarts
    Write-Host "Text-Find starts for 'needle': $($searchStarts -join ', ')"
    Write-Host "Text-Find hits: $($searchStarts.Count); expected: 3; strictly increasing and unique: $searchStrict"
    if ($searchStarts.Count -ne 3 -or -not $searchStrict) {
        throw 'The repeated-text Find control did not return three distinct, strictly increasing starts.'
    }
    Write-Host 'PASS: The repeated-text Find control advanced through every occurrence.'

    if (-not $SkipLargeBenchmarks) {
        foreach ($paragraphCount in 2000, 3000) {
            $doc.Content.Text = ''
            $expectedHeadingCount = Add-LargePlainFixture $doc $paragraphCount
            $largeFind = Measure-OutlineScan "Plain $paragraphCount-paragraph format-only Find" { (Get-OutlineByUnanchoredFind $doc 2000).headings }
            $largeIndexed = Measure-OutlineScan "Plain $paragraphCount-paragraph indexed OutlineLevel walk" { Get-OutlineByParagraph $doc 2000 }
            $largeEnumerated = Measure-OutlineScan "Plain $paragraphCount-paragraph enumerated OutlineLevel walk" { Get-OutlineByParagraphEnumerator $doc 2000 }
            $largeXml = Measure-OutlineScan "Plain $paragraphCount-paragraph XML outline" { Get-OutlineByXml $doc 2000 }
            Assert-EquivalentOutlines "Plain $paragraphCount-paragraph indexed comparison" $largeFind.headings $largeIndexed.headings
            Assert-EquivalentOutlines "Plain $paragraphCount-paragraph enumerated comparison" $largeFind.headings $largeEnumerated.headings
            Assert-EquivalentOutlines "Plain $paragraphCount-paragraph XML comparison" $largeFind.headings $largeXml.headings
            if ($largeFind.headings.Count -ne $expectedHeadingCount) {
                throw "Plain $paragraphCount-paragraph Find returned $($largeFind.headings.Count) headings; expected $expectedHeadingCount."
            }
            Write-Host "PASS: Plain $paragraphCount-paragraph Find, both OutlineLevel walks, and XML outline agree on $expectedHeadingCount headings."
        }
    }
} finally {
    if ($doc) { try { $doc.Close(0) } catch { Write-Host "Document close failed: $($_.Exception.Message)" } }
    if ($app) { try { $app.Quit() } catch { Write-Host "Word quit failed: $($_.Exception.Message)" } }
    if ($ownPid -gt 0) {
        $deadline = (Get-Date).AddSeconds(60)
        while ((Get-Date) -lt $deadline -and (Get-Process -Id $ownPid -ErrorAction SilentlyContinue)) {
            Start-Sleep -Milliseconds 250
        }
        Write-Host "Cleanup for attributed WINWORD ${ownPid}: $(Stop-VerifiedWord $ownPid $ownStart)"
    }
}
