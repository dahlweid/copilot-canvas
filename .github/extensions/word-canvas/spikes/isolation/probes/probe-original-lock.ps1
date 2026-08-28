# Probe: what does holding the ORIGINAL open read-write actually cost?
#
# Q6 proposes the agent edits the original directly and lets Word arbitrate
# file locking. Three things need to be true for that to be clean:
#
#   T1. Word's lock artefacts don't disturb the user's folder or our watcher.
#   T2. An external script can still regenerate the file while we hold it.
#   T3. The user's own Word can still open it.
#
# T2 is the one that matters most: v1's headline feature is auto-refresh when a
# script rewrites the document.

$ErrorActionPreference = 'Stop'
$root = Join-Path $env:TEMP ("lockprobe-" + [guid]::NewGuid().ToString('N').Substring(0,8))
New-Item -ItemType Directory -Path $root | Out-Null
$doc = Join-Path $root 'original.docx'
Copy-Item $args[0] $doc
$regen = Join-Path $root 'regen-source.docx'
Copy-Item $args[0] $regen

function Show-Dir($label) {
    $names = Get-ChildItem $root -Force | Select-Object -ExpandProperty Name
    "  [$label] " + ($names -join ', ')
}

$word = $null
try {
    "== T0: baseline =="
    Show-Dir 'before open'

    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    "Word pid-ish check: Visible=$($word.Visible)"

    $sw = [Diagnostics.Stopwatch]::StartNew()
    $d = $word.Documents.Open($doc, $false, $false)   # ReadOnly = $false
    $sw.Stop()
    "Opened read-write in $($sw.ElapsedMilliseconds) ms; ReadOnly=$($d.ReadOnly)"

    "== T1: lock artefacts =="
    Show-Dir 'while open'

    "== T2: can an external writer replace the file? =="
    # (a) direct overwrite, the way a generator script would
    try {
        Copy-Item $regen $doc -Force
        "  (a) Copy-Item overwrite: SUCCEEDED"
    } catch {
        "  (a) Copy-Item overwrite: FAILED -> $($_.Exception.Message.Split([char]10)[0])"
    }
    # (b) replace-and-rename, the pattern python-docx / pandoc actually use
    try {
        $tmp = Join-Path $root 'new.tmp'
        Copy-Item $regen $tmp -Force
        Move-Item $tmp $doc -Force
        "  (b) write-temp-then-rename: SUCCEEDED"
    } catch {
        "  (b) write-temp-then-rename: FAILED -> $($_.Exception.Message.Split([char]10)[0])"
    }
    # (c) open for exclusive write, what a real generator does internally
    try {
        $fs = [IO.File]::Open($doc, 'Open', 'Write', 'None')
        $fs.Close()
        "  (c) exclusive write handle: SUCCEEDED"
    } catch {
        "  (c) exclusive write handle: FAILED -> $($_.Exception.Message.Split([char]10)[0])"
    }

    "== T3: can a second Word open it? =="
    $w2 = New-Object -ComObject Word.Application
    $w2.Visible = $false
    $w2.DisplayAlerts = 0
    try {
        $d2 = $w2.Documents.Open($doc, $false, $false)
        "  second instance: opened, ReadOnly=$($d2.ReadOnly)"
        $d2.Close(0)
    } catch {
        "  second instance: FAILED -> $($_.Exception.Message.Split([char]10)[0])"
    }
    $w2.Quit()
    [Runtime.InteropServices.Marshal]::ReleaseComObject($w2) | Out-Null

    "== T4: transient-lock alternative =="
    $d.Close(0)
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $d3 = $word.Documents.Open($doc, $false, $false)
    $d3.Content.InsertAfter("probe edit")
    $d3.Save()
    $d3.Close(0)
    $sw.Stop()
    "  open + edit + save + close round trip: $($sw.ElapsedMilliseconds) ms"
    Show-Dir 'after close'
}
finally {
    if ($word) {
        try { $word.Quit() } catch {}
        [Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
    Remove-Item $root -Recurse -Force -ErrorAction SilentlyContinue
}
