# Probe: is "read-then-address" implementable against Word's object model?
#
# The model needs three things that have not been checked:
#
#   S1. A structural read has to be cheap enough to do routinely.
#   S2. Paragraphs need a stable identity. Word exposes none natively, so any
#       ID must be derived - and derived IDs collide when paragraphs repeat.
#   S3. A revision token has to distinguish "the agent changed it" from
#       "something else changed it".

$ErrorActionPreference = 'Stop'
$src = $args[0]
$root = Join-Path $env:TEMP ("addr-" + [guid]::NewGuid().ToString('N').Substring(0,8))
New-Item -ItemType Directory -Path $root | Out-Null
$doc = Join-Path $root 'original.docx'
Copy-Item $src $doc

function Get-Token([string]$Path) {
    $h = Get-FileHash -Path $Path -Algorithm SHA256
    return $h.Hash.Substring(0, 16)
}

$word = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false; $word.DisplayAlerts = 0

    $d = $word.Documents.Open($doc, $false, $true)   # read-only for the survey
    $d.Content.InsertAfter("") | Out-Null            # no-op to warm the object model

    "== S1: cost of a full structural read =="
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $paras = @()
    foreach ($p in $d.Paragraphs) {
        $t = $p.Range.Text -replace "[`r`a]", ""
        $paras += [pscustomobject]@{
            Text    = $t
            Outline = $p.OutlineLevel
            Start   = $p.Range.Start
        }
    }
    $sw.Stop()
    "  $($paras.Count) paragraphs read in $($sw.ElapsedMilliseconds) ms"

    "== S2: is derived identity viable? =="
    $empty = @($paras | Where-Object { $_.Text.Trim() -eq '' }).Count
    "  empty paragraphs: $empty of $($paras.Count)"

    $dupGroups = $paras | Group-Object Text | Where-Object { $_.Count -gt 1 }
    $dupTotal = ($dupGroups | Measure-Object -Property Count -Sum).Sum
    "  paragraphs whose text is not unique: $dupTotal in $($dupGroups.Count) groups"
    if ($dupGroups) {
        $top = $dupGroups | Sort-Object Count -Descending | Select-Object -First 3
        foreach ($g in $top) {
            $label = if ($g.Name.Trim() -eq '') { '<empty>' } else { '"' + $g.Name.Substring(0, [Math]::Min(30, $g.Name.Length)) + '"' }
            "    x$($g.Count)  $label"
        }
    }

    # heading path + text: does that disambiguate the duplicates?
    $headingPath = ''
    $keys = @()
    foreach ($p in $paras) {
        if ($p.Outline -lt 10) { $headingPath = $p.Text }
        $keys += "$headingPath|$($p.Text)"
    }
    $keyDup = ($keys | Group-Object | Where-Object { $_.Count -gt 1 } | Measure-Object -Property Count -Sum).Sum
    "  still colliding after adding heading path: $keyDup"
    "  -> a per-key occurrence index is $(if ($keyDup) { 'REQUIRED' } else { 'not needed here' })"

    $d.Close(0)

    "== S3: revision token behaviour =="
    $t0 = Get-Token $doc
    "  token before          : $t0"

    # our own edit
    $d = $word.Documents.Open($doc, $false, $false)
    $d.Content.InsertAfter("agent edit`r`n")
    $d.Save(); $d.Close(0)
    $t1 = Get-Token $doc
    "  token after our edit  : $t1  (changed: $($t1 -ne $t0))"

    # a save with no content change - does Word rewrite the bytes anyway?
    $d = $word.Documents.Open($doc, $false, $false)
    $d.Save(); $d.Close(0)
    $t2 = Get-Token $doc
    "  token after no-op save: $t2  (changed: $($t2 -ne $t1))"

    # an external regeneration
    Copy-Item $src $doc -Force
    $t3 = Get-Token $doc
    "  token after external  : $t3  (changed: $($t3 -ne $t2))"

    $sw = [Diagnostics.Stopwatch]::StartNew(); Get-Token $doc | Out-Null; $sw.Stop()
    "  cost of computing a token: $($sw.ElapsedMilliseconds) ms"
}
finally {
    if ($word) { try { $word.Quit() } catch {}; [Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null }
    Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 1
    Remove-Item $root -Recurse -Force -ErrorAction SilentlyContinue
    Get-ChildItem $env:TEMP -Directory -Filter 'addr-*' -ErrorAction SilentlyContinue |
        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    "swept"
}
