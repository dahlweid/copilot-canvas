# Probe: is the transient-lock design actually implementable?
#
# Two things it depends on that have never been measured:
#
#   M1. What does a full open-edit-save-close round trip cost?
#   M2. Can we detect "already open" WITHOUT calling Documents.Open,
#       which is the call that hangs?
#
# M2 is load-bearing. If detection is unreliable, every operation risks the
# indefinite hang measured in section 16.

$ErrorActionPreference = 'Stop'
$src = $args[0]
$root = Join-Path $env:TEMP ("txlock-" + [guid]::NewGuid().ToString('N').Substring(0,8))
New-Item -ItemType Directory -Path $root | Out-Null
$doc = Join-Path $root 'original.docx'
Copy-Item $src $doc

# Detection candidate: try to take a write handle. Must be fast and must not block.
function Test-Locked([string]$Path) {
    try {
        $fs = [IO.File]::Open($Path, 'Open', 'Write', 'None')
        $fs.Close()
        return $false
    } catch { return $true }
}

# Word's owner file: "~$" + basename with the first two characters dropped.
function Get-OwnerFile([string]$Path) {
    $dir = Split-Path $Path -Parent
    $name = Split-Path $Path -Leaf
    Join-Path $dir ('~$' + $name.Substring(2))
}

$word = $null
try {
    "== M1: open-edit-save-close round trip (file free) =="
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false; $word.DisplayAlerts = 0

    # discard first iteration: one-off engine warmup
    $d = $word.Documents.Open($doc, $false, $false)
    $d.Content.InsertAfter("warmup`r`n"); $d.Save(); $d.Close(0)

    $times = @()
    for ($i = 1; $i -le 5; $i++) {
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $d = $word.Documents.Open($doc, $false, $false)
        $d.Content.InsertAfter("edit $i`r`n")
        $d.Save()
        $d.Close(0)
        $sw.Stop()
        $times += $sw.ElapsedMilliseconds
    }
    "  per-iteration ms: $($times -join ', ')"
    "  mean: $([math]::Round(($times | Measure-Object -Average).Average)) ms"

    "== M2a: detection when the file is FREE (false-positive check) =="
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $locked = Test-Locked $doc
    $sw.Stop()
    "  write-handle test -> locked=$locked in $($sw.ElapsedMilliseconds) ms (expect False)"
    "  owner-file test   -> exists=$(Test-Path (Get-OwnerFile $doc)) (expect False)"

    "== M2b: detection when ANOTHER Word holds it =="
    $marker = Join-Path $root 'held.flag'
    $holder = Start-Job -ArgumentList $doc, $marker -ScriptBlock {
        param($doc, $marker)
        $w = New-Object -ComObject Word.Application
        $w.Visible = $false; $w.DisplayAlerts = 0
        $d = $w.Documents.Open($doc, $false, $false)
        New-Item -ItemType File -Path $marker -Force | Out-Null
        Start-Sleep -Seconds 40
        try { $d.Close(0); $w.Quit() } catch {}
    }
    $waited = 0
    while (-not (Test-Path $marker) -and $waited -lt 60) { Start-Sleep -Milliseconds 200; $waited += 0.2 }

    if (Test-Path $marker) {
        Start-Sleep -Milliseconds 500
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $locked = Test-Locked $doc
        $sw.Stop()
        "  write-handle test -> locked=$locked in $($sw.ElapsedMilliseconds) ms (expect True, fast)"
        $owner = Get-OwnerFile $doc
        "  owner-file test   -> exists=$(Test-Path $owner) at '$(Split-Path $owner -Leaf)' (expect True)"
    } else {
        "  holder job never signalled; M2b inconclusive"
    }

    Stop-Job $holder -ErrorAction SilentlyContinue
    Remove-Job $holder -Force -ErrorAction SilentlyContinue
}
finally {
    if ($word) { try { $word.Quit() } catch {}; [Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null }
    Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 1
    Remove-Item $root -Recurse -Force -ErrorAction SilentlyContinue
    Get-ChildItem $env:TEMP -Directory -Filter 'txlock-*' -ErrorAction SilentlyContinue |
        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    "swept"
}
