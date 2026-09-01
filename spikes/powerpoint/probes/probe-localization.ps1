# Probe: are PowerPoint's layout, master and placeholder names localized?
#
# Word on this machine is German, and it has cost this repo twice:
# w:pStyle/@w:val returns "Uberschrift1" rather than "Heading1", and
# Selection.Style = "Heading 1" throws. Any PowerPoint code that addresses a
# placeholder or a layout BY NAME is exposed to the same class of failure, so
# the question is not whether names differ but whether there is a
# language-independent way to address the same things.
#
#   L1  What language is this Office, per the object model?
#   L2  Master and custom-layout names as PowerPoint reports them.
#   L3  Shape names of the placeholders on a real slide.
#   L4  THE TRAP: address a placeholder by its English name, then by the name
#       PowerPoint actually reports. Word's equivalent throws.
#   L5  The language-independent alternatives: Shapes.Placeholders[i] and
#       PlaceholderFormat.Type (an enum), plus p:ph/@type in the OOXML.
#
# L4 and L5 together are the useful result: if names are localized but the enum
# and the XML type are not, then "never address by name" is a rule the host can
# actually follow.

param([string]$Fixture)

$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot '_common.ps1')

if (-not $Fixture) { $Fixture = Join-Path (Split-Path $PSScriptRoot -Parent) '.fixtures\deck.pptx' }
if (-not (Test-Path $Fixture)) { throw "fixture missing: $Fixture (run make-fixture.ps1 first)" }

Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue

$root = Join-Path $env:TEMP ("pptloc-" + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Force -Path $root | Out-Null
$deck = Join-Path $root 'deck.pptx'
Copy-Item $Fixture $deck
$pristine = Join-Path $root 'pristine.pptx'   # L5b must not see L6's rename
Copy-Item $Fixture $pristine

$ctx = $null
$pres = $null
try {
    Rep "POWERPNT pids before" ($(if (Get-PptPids) { (Get-PptPids) -join ',' } else { '(none)' }))
    $ctx = New-PowerPointInstance
    Rep "new POWERPNT pids seen" ($(if ($ctx.NewPids) { $ctx.NewPids -join ',' } else { '(none appeared - attached)' }))
    $app = $ctx.App

    Say "== L1: which language is this Office? =="
    # msoLanguageIDUI = 2, msoLanguageIDInstall = 1
    foreach ($pair in @(@(2, 'UI'), @(1, 'Install'))) {
        try {
            $id = $app.LanguageSettings.LanguageID($pair[0])
            $name = try { ([Globalization.CultureInfo]::GetCultureInfo([int]$id)).Name } catch { '?' }
            Rep "  LanguageID($($pair[1]))" "$id ($name)"
        }
        catch { Rep "  LanguageID($($pair[1]))" 'unreadable' }
    }

    $pres = $app.Presentations.Open($deck, 0, 0, 0)

    Say "== L2: master and custom layout names =="
    try { Rep "  SlideMaster.Name" $pres.SlideMaster.Name } catch { Rep "  SlideMaster.Name" 'unreadable' }
    try { Rep "  Designs.Item(1).Name" $pres.Designs.Item(1).Name } catch { }
    try {
        $layouts = $pres.SlideMaster.CustomLayouts
        Rep "  custom layouts" $layouts.Count
        for ($i = 1; $i -le [Math]::Min(6, $layouts.Count); $i++) {
            Say ("      [{0}] '{1}'" -f $i, $layouts.Item($i).Name)
        }
    }
    catch { Rep "  custom layouts" ('unreadable -> ' + $_.Exception.Message.Split([char]10)[0]) }

    Say "== L3: placeholder shape names on slide 2 =="
    $slide = $pres.Slides.Item(2)
    $names = @()
    for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
        $sh = $slide.Shapes.Item($j)
        $names += $sh.Name
        $phType = 'n/a'
        try { $phType = $sh.PlaceholderFormat.Type } catch { }
        Say ("      shape {0}: name='{1}'  PlaceholderFormat.Type={2}" -f $j, $sh.Name, $phType)
    }
    try { Rep "  slide 2 CustomLayout.Name" $slide.CustomLayout.Name } catch { }

    Say "== L4: THE TRAP - addressing a placeholder by name =="
    foreach ($candidate in @('Title 1', 'Titel 1', $names[0])) {
        try {
            $s = $slide.Shapes.Item($candidate)
            Rep "  Shapes.Item('$candidate')" ("OK -> '" + $s.Name + "'")
        }
        catch { Rep "  Shapes.Item('$candidate')" ('THROWS -> ' + $_.Exception.Message.Split([char]10)[0]) }
    }

    Say "== L5: the language-independent alternatives =="
    try {
        $ph = $slide.Shapes.Placeholders
        Rep "  Shapes.Placeholders.Count" $ph.Count
        for ($j = 1; $j -le $ph.Count; $j++) {
            Say ("      Placeholders[{0}] type={1} name='{2}'" -f $j, $ph.Item($j).PlaceholderFormat.Type, $ph.Item($j).Name)
        }
        Rep "  index + enum are language-neutral" 'YES (no string involved)'
    }
    catch { Rep "  Shapes.Placeholders" ('unreadable -> ' + $_.Exception.Message.Split([char]10)[0]) }

    Say "== L6: CONTROL - is the COM name a live property, or a translation? =="
    # If COM merely reports what the file stores, a custom name set via COM must
    # come back verbatim from both sides. If COM instead translates DEFAULT
    # placeholder names into English, a custom name will match on both sides
    # while the untouched shapes keep disagreeing. That is the discriminator.
    $custom = 'ZZ_probe_custom_name'
    $slide.Shapes.Item(1).Name = $custom
    Rep "  set shape 1 name via COM" $custom
    Rep "  COM reads back" $slide.Shapes.Item(1).Name
    $pres.Save()
    $renamed = Join-Path $root 'renamed.pptx'
    Copy-Item $deck $renamed -Force
    try { $pres.Saved = -1; $pres.Close(); $pres = $null } catch { }

    $zc = [IO.Compression.ZipFile]::OpenRead($renamed)
    try {
        $e = $zc.Entries | Where-Object { $_.FullName -eq 'ppt/slides/slide2.xml' } | Select-Object -First 1
        $sr = New-Object IO.StreamReader($e.Open()); $cx = [xml]$sr.ReadToEnd(); $sr.Close()
        $cns = New-Object Xml.XmlNamespaceManager($cx.NameTable)
        $cns.AddNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main')
        $stored = @($cx.SelectNodes('//p:nvSpPr/p:cNvPr/@name', $cns) | ForEach-Object { $_.Value })
        Rep "  names stored in the file now" ($stored -join ' | ')
        Rep "  custom name survived to disk" $(if ($stored -contains $custom) { 'YES - COM name is a live property' } else { 'NO' })
        $stillGerman = @($stored | Where-Object { $_ -match '^(Titel|Textplatzhalter|Textfeld)' })
        Rep "  untouched shapes still German" $(if ($stillGerman.Count) { 'YES -> ' + ($stillGerman -join ', ') } else { 'NO' })
    }
    finally { $zc.Dispose() }

    Say "== L5b: and in the OOXML, which needs no PowerPoint at all =="
    $tmp = Join-Path $root 'x.pptx'; [IO.File]::Copy($pristine, $tmp, $true)
    $zip = [IO.Compression.ZipFile]::OpenRead($tmp)
    try {
        $e = $zip.Entries | Where-Object { $_.FullName -eq 'ppt/slides/slide2.xml' } | Select-Object -First 1
        $sr = New-Object IO.StreamReader($e.Open()); $xmlText = $sr.ReadToEnd(); $sr.Close()
        $x = [xml]$xmlText
        $ns = New-Object Xml.XmlNamespaceManager($x.NameTable)
        $ns.AddNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
        $ns.AddNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main')
        foreach ($sp in $x.SelectNodes('//p:sp', $ns)) {
            $t = $sp.SelectSingleNode('.//p:nvSpPr/p:nvPr/p:ph/@type', $ns)
            $n = $sp.SelectSingleNode('.//p:nvSpPr/p:cNvPr/@name', $ns)
            Say ("      p:ph/@type='{0}'   cNvPr/@name='{1}'" -f $(if ($t) { $t.Value } else { '(none)' }), $(if ($n) { $n.Value } else { '' }))
        }
        # The layout names as STORED, versus as PowerPoint reports them.
        $lay = $zip.Entries | Where-Object { $_.FullName -eq 'ppt/slideLayouts/slideLayout2.xml' } | Select-Object -First 1
        if ($lay) {
            $sr2 = New-Object IO.StreamReader($lay.Open()); $lx = [xml]$sr2.ReadToEnd(); $sr2.Close()
            $lns = New-Object Xml.XmlNamespaceManager($lx.NameTable)
            $lns.AddNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main')
            $t = $lx.SelectSingleNode('/p:sldLayout/@type', $lns)
            $nm = $lx.SelectSingleNode('//p:cSld/@name', $lns)
            Rep "  slideLayout2 @type (stored)" $(if ($t) { $t.Value } else { '(none)' })
            Rep "  slideLayout2 cSld/@name (stored)" $(if ($nm) { $nm.Value } else { '(none)' })
        }
    }
    finally { $zip.Dispose() }
}
catch { Rep "ERROR" $_.Exception.Message.Split([char]10)[0] }
finally {
    # Ours, opened from our own temp root -- close it before releasing the
    # application, which may be one we merely attached to.
    try { if ($pres) { $pres.Saved = -1; $pres.Close() } } catch { }
    $pres = $null
    Close-PowerPointInstance $ctx
    Remove-Item $root -Recurse -Force -ErrorAction SilentlyContinue
    Rep "POWERPNT pids after" ($(if (Get-PptPids) { (Get-PptPids) -join ',' } else { '(none)' }))
}
