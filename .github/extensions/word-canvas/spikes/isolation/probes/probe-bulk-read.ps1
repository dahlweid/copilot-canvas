# Probe: the 4144 ms structural read is a per-property COM round-trip problem.
# Each $p.Range.Text / $p.OutlineLevel is a cross-process call. Test whether
# bulk extraction removes it.
#
#   A. naive     - iterate Paragraphs, touch properties per paragraph
#   B. bulk text - one Content.Text call, split on CR
#   C. bulk both - one Content.Text call + one pass for outline levels
#   D. XML       - single WordOpenXML call, parsed outside Word

$ErrorActionPreference = 'Stop'
$src = $args[0]
$root = Join-Path $env:TEMP ("bulk-" + [guid]::NewGuid().ToString('N').Substring(0,8))
New-Item -ItemType Directory -Path $root | Out-Null
$doc = Join-Path $root 'original.docx'
Copy-Item $src $doc

$word = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false; $word.DisplayAlerts = 0
    $d = $word.Documents.Open($doc, $false, $true)

    $null = $d.Content.Text   # warm

    "== A: naive per-paragraph property access =="
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $n = 0
    foreach ($p in $d.Paragraphs) { $null = $p.Range.Text; $null = $p.OutlineLevel; $n++ }
    $sw.Stop()
    "  $n paragraphs, $($sw.ElapsedMilliseconds) ms"

    "== B: one Content.Text call, split locally =="
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $text = $d.Content.Text
    $lines = $text -split "`r"
    $sw.Stop()
    "  $($lines.Count) lines, $($sw.ElapsedMilliseconds) ms"

    "== C: bulk text + outline levels via ListParagraphs-free single pass =="
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $text = $d.Content.Text
    $lines = $text -split "`r"
    # outline level only for paragraphs that are headings: ask Word once per heading
    $levels = @{}
    $hdrs = $d.Paragraphs
    $sw.Stop()
    "  (text) $($sw.ElapsedMilliseconds) ms - outline still needs a per-para call, see D"

    "== D: single WordOpenXML call, parsed outside Word =="
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $xml = $d.Content.WordOpenXML
    $sw.Stop()
    $xmlMs = $sw.ElapsedMilliseconds
    "  fetched $([math]::Round($xml.Length/1024)) KB of WordprocessingML in $xmlMs ms"

    $sw = [Diagnostics.Stopwatch]::StartNew()
    $x = [xml]$xml
    $ns = New-Object Xml.XmlNamespaceManager($x.NameTable)
    $ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
    $nodes = $x.SelectNodes('//w:body/w:p', $ns)
    $parsed = foreach ($node in $nodes) {
        $style = $node.SelectSingleNode('./w:pPr/w:pStyle/@w:val', $ns)
        $texts = $node.SelectNodes('.//w:t', $ns) | ForEach-Object { $_.InnerText }
        [pscustomobject]@{ Style = $(if ($style) { $style.Value } else { '' }); Text = ($texts -join '') }
    }
    $sw.Stop()
    "  parsed $($parsed.Count) paragraphs in $($sw.ElapsedMilliseconds) ms"
    "  TOTAL for D: $($xmlMs + $sw.ElapsedMilliseconds) ms"
    $styled = @($parsed | Where-Object { $_.Style -ne '' }).Count
    "  paragraphs carrying an explicit style: $styled"
    $h = @($parsed | Where-Object { $_.Style -like 'berschrift*' -or $_.Style -like 'Heading*' }) 
    "  heading-styled paragraphs: $($h.Count)"
    if ($h.Count) { "  first heading style name: '$($h[0].Style)'  text: '$($h[0].Text.Substring(0,[Math]::Min(40,$h[0].Text.Length)))'" }

    $d.Close(0)
}
finally {
    if ($word) { try { $word.Quit() } catch {}; [Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null }
    Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 1
    Remove-Item $root -Recurse -Force -ErrorAction SilentlyContinue
    Get-ChildItem $env:TEMP -Directory -Filter 'bulk-*' -ErrorAction SilentlyContinue |
        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    "swept"
}
