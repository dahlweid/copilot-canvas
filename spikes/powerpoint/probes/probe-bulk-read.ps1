# Probe: is there a bulk structural read, the way Word has Content.WordOpenXML?
#
# In Word the naive walk cost 3724 ms for 219 paragraphs and one bulk
# WordOpenXML call cost 289 ms -- a 13x defect caused entirely by per-property
# COM round trips. PowerPoint's object model is deeper (Presentation > Slide >
# Shapes > TextFrame > TextRange), so the same trap is more likely, not less.
#
# PowerPoint has no WordOpenXML equivalent on any object. So the arms are:
#
#   A  naive     per-shape walk: Shapes.Item(i).TextFrame.TextRange.Text
#   B  per-slide Shapes.Range().Text -- one call per slide instead of per shape
#   C  OOXML     read the .pptx as a ZIP and parse ppt/slides/slideN.xml,
#                with no COM at all
#
# C is the interesting one. A .pptx is a zip of XML; if it is fast enough, the
# structural read needs no PowerPoint instance, which matters a great deal given
# that probe-single-instance.ps1 shows there is no such thing as a private one.
#
# C is measured twice: against a file PowerPoint is holding open, and against a
# free file, because reading the zip is only useful if it works in both states.

param([string]$Fixture)

$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot '_common.ps1')

if (-not $Fixture) { $Fixture = Join-Path (Split-Path $PSScriptRoot -Parent) '.fixtures\deck.pptx' }
if (-not (Test-Path $Fixture)) { throw "fixture missing: $Fixture (run make-fixture.ps1 first)" }

Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue

$root = Join-Path $env:TEMP ("pptbulk-" + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Force -Path $root | Out-Null
$deck = Join-Path $root 'deck.pptx'
Copy-Item $Fixture $deck

# C: parse the package directly. No PowerPoint involved.
function Read-DeckXml([string]$Path) {
    $items = @()
    # Copy first: PowerPoint holds a read lock, and ZipFile wants its own handle.
    $tmp = Join-Path $root ("zip-" + [guid]::NewGuid().ToString('N').Substring(0, 6) + ".pptx")
    [IO.File]::Copy($Path, $tmp, $true)
    $zip = [IO.Compression.ZipFile]::OpenRead($tmp)
    try {
        $slideEntries = @($zip.Entries | Where-Object { $_.FullName -match '^ppt/slides/slide\d+\.xml$' } |
                Sort-Object { [int]([regex]::Match($_.FullName, '(\d+)').Value) })
        foreach ($e in $slideEntries) {
            $sr = New-Object IO.StreamReader($e.Open())
            $xmlText = $sr.ReadToEnd(); $sr.Close()
            $x = [xml]$xmlText
            $ns = New-Object Xml.XmlNamespaceManager($x.NameTable)
            $ns.AddNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
            $ns.AddNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main')
            foreach ($sp in $x.SelectNodes('//p:sp', $ns)) {
                $ph = $sp.SelectSingleNode('.//p:nvSpPr/p:nvPr/p:ph/@type', $ns)
                $nm = $sp.SelectSingleNode('.//p:nvSpPr/p:cNvPr/@name', $ns)
                $texts = @($sp.SelectNodes('.//a:t', $ns) | ForEach-Object { $_.InnerText })
                $items += [pscustomobject]@{
                    Slide           = $e.FullName
                    PlaceholderType = $(if ($ph) { $ph.Value } else { '' })
                    ShapeName       = $(if ($nm) { $nm.Value } else { '' })
                    Text            = ($texts -join '')
                }
            }
        }
    }
    finally { $zip.Dispose(); Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
    , $items
}

$ctx = $null
try {
    Rep "POWERPNT pids before" ($(if (Get-PptPids) { (Get-PptPids) -join ',' } else { '(none)' }))
    $ctx = New-OwnedPowerPoint
    Rep "owned pid" ($(if ($ctx.Owned) { $ctx.Owned -join ',' } else { '(ATTACHED - will not quit)' }))
    $pres = $ctx.App.Presentations.Open($deck, 0, 0, 0)
    $slides = $pres.Slides.Count
    $null = $pres.Slides.Item(1).Shapes.Count   # warm the proxy

    Say "== A: naive per-shape walk =="
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $nShapes = 0; $textA = @()
    for ($i = 1; $i -le $slides; $i++) {
        $shapes = $pres.Slides.Item($i).Shapes
        for ($j = 1; $j -le $shapes.Count; $j++) {
            $sh = $shapes.Item($j)
            $nShapes++
            if ($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1) { $textA += $sh.TextFrame.TextRange.Text }
        }
    }
    $sw.Stop(); $msA = $sw.ElapsedMilliseconds
    Rep "  $slides slides, $nShapes shapes" ("$msA ms")

    Say "== B: one Shapes.Range().Text call per slide =="
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $textB = @()
    for ($i = 1; $i -le $slides; $i++) {
        try { $textB += $pres.Slides.Item($i).Shapes.Range().TextFrame.TextRange.Text } catch { }
    }
    $sw.Stop(); $msB = $sw.ElapsedMilliseconds
    Rep "  $slides calls" ("$msB ms")

    Say "== C: read the .pptx package directly, no COM (file HELD open) =="
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $itemsHeld = Read-DeckXml $deck
    $sw.Stop(); $msC1 = $sw.ElapsedMilliseconds
    Rep "  $($itemsHeld.Count) shapes parsed" ("$msC1 ms")
    Rep "  readable while PowerPoint holds it" ($(if ($itemsHeld.Count -gt 0) { 'YES' } else { 'NO' }))

    try { $pres.Saved = -1; $pres.Close() } catch { }

    Say "== C2: same read with the file FREE (control) =="
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $itemsFree = Read-DeckXml $deck
    $sw.Stop(); $msC2 = $sw.ElapsedMilliseconds
    Rep "  $($itemsFree.Count) shapes parsed" ("$msC2 ms")

    Say "== comparison =="
    Rep "  A naive per-shape walk" ("$msA ms")
    Rep "  B per-slide Range().Text" ("$msB ms" + $(if ($msB -gt 0) { "  ({0:F1}x faster than A)" -f ($msA / [double]$msB) } else { '' }))
    Rep "  C OOXML, no COM (held)" ("$msC1 ms" + $(if ($msC1 -gt 0) { "  ({0:F1}x faster than A)" -f ($msA / [double]$msC1) } else { '' }))
    Rep "  C OOXML, no COM (free)" ("$msC2 ms")
    Rep "  text extracted A / C" ("{0} / {1} non-empty shapes" -f @($textA | Where-Object { $_ }).Count, @($itemsFree | Where-Object { $_.Text }).Count)

    Say "== what the OOXML path gives you that COM does not =="
    $ph = @($itemsFree | Where-Object { $_.PlaceholderType })
    Rep "  shapes with a placeholder type" $ph.Count
    Rep "  distinct placeholder types" ((@($ph | Select-Object -ExpandProperty PlaceholderType -Unique)) -join ', ')
    if ($ph.Count) {
        Rep "  example: type / shape name" ("'{0}' / '{1}'" -f $ph[0].PlaceholderType, $ph[0].ShapeName)
    }
}
catch { Rep "ERROR" $_.Exception.Message.Split([char]10)[0] }
finally {
    Close-OwnedPowerPoint $ctx
    Remove-Item $root -Recurse -Force -ErrorAction SilentlyContinue
    Rep "POWERPNT pids after" ($(if (Get-PptPids) { (Get-PptPids) -join ',' } else { '(none)' }))
}
