# Builds the .pptx fixture the other probes measure against.
#
# The fixture is generated rather than committed because spikes/ must not carry
# binaries into a gist-shared extension, and because a generated deck is
# reproducible: same slide count, same shapes, same text on every machine.
#
#   -Slides   how many slides (default 13, matching the 13-page Word fixture)
#   -Path     where to write it (default <spike>/.fixtures/deck.pptx)
#
# Fixtures land inside this worktree on purpose. Sibling sessions run in their
# own worktrees, so no two sessions can ever contend for the same file path.

param(
    [int]$Slides = 13,
    [string]$Path
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '_common.ps1')

if (-not $Path) {
    $Path = Join-Path (Split-Path $PSScriptRoot -Parent) '.fixtures\deck.pptx'
}
$dir = Split-Path $Path -Parent
New-Item -ItemType Directory -Force -Path $dir | Out-Null
Remove-Item $Path -Force -ErrorAction SilentlyContinue

$ctx = $null
$pres = $null
try {
    $ctx = New-PowerPointInstance
    Rep "new POWERPNT pids seen" ($(if ($ctx.NewPids) { $ctx.NewPids -join ',' } else { '(none appeared - attached)' }))
    $app = $ctx.App

    # WithWindow := msoFalse (0) keeps the deck off screen; see probe-hide.ps1.
    $pres = $app.Presentations.Add(0)

    # ppLayoutTitle = 1, ppLayoutText = 2
    for ($i = 1; $i -le $Slides; $i++) {
        $layout = if ($i -eq 1) { 1 } else { 2 }
        $slide = $pres.Slides.Add($i, $layout)
        $slide.Shapes.Item(1).TextFrame.TextRange.Text = "Slide $i title"
        if ($slide.Shapes.Count -ge 2) {
            $body = 1..6 | ForEach-Object { "Bullet $_ on slide $i with enough text to wrap onto a second line when it is rendered" }
            $slide.Shapes.Item(2).TextFrame.TextRange.Text = ($body -join "`r")
        }
        # A second text box per slide, so the per-shape walk in probe-bulk-read
        # has a realistic number of shapes to traverse.
        $tb = $slide.Shapes.AddTextbox(1, 40, 460, 600, 40)
        $tb.TextFrame.TextRange.Text = "Footnote for slide $i"
    }

    $pres.SaveAs($Path)
    $pres.Close()
    $pres = $null

    Rep "fixture" $Path
    Rep "slides" $Slides
    Rep "size" ("{0:N0} bytes" -f (Get-Item $Path).Length)
}
finally {
    # Ours: we added it. Close it before releasing an application that may be
    # one we merely attached to.
    try { if ($pres) { $pres.Saved = -1; $pres.Close() } } catch { }
    $pres = $null
    Close-PowerPointInstance $ctx
    Rep "POWERPNT pids after" ($(if (Get-PptPids) { (Get-PptPids) -join ',' } else { '(none)' }))
}
