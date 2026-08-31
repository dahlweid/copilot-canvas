# Probe: is "read-then-address" implementable against Word's object model?
#
# The model needs three things that have not been checked:
#
#   S1. A structural read has to be cheap enough to do routinely.
#   S2. Paragraphs need a stable identity. Word exposes none natively, so any
#       ID must be derived - and derived IDs collide when paragraphs repeat.
#   S3. A revision token has to distinguish "the agent changed it" from
#       "something else changed it".
#
# WHAT THIS PROBE DOES AND DOES NOT BACK. S1 measures ONE strategy: the
# per-paragraph walk, touching `Range.Text` and `OutlineLevel` on every
# paragraph. It has no `Content.WordOpenXML` arm and no `Content.Text` arm, so
# it backs the cost of the walk (3724 ms on 219 paragraphs) and NOT the 289 ms
# WordOpenXML figure or the 6 ms `Content.Text` figure it was compared against.
# Those two came from a bulk-read comparison that is no longer committed. Cite
# this file for the walk; do not cite it for the alternatives.
#
# Retired with the rest of the streaming archive (#103) and restored in the same
# PR: it is the only committed instrument behind the S1 walk cost quoted in
# `word-host.ps1` and the S3 revision-token behaviour quoted in
# `revision-token.mjs`. Its cleanup was rewritten on restore -- see the note in
# `finally`, which is the reason it is safe to run on a shared machine now and
# was not before.

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
    # There is deliberately NO `Get-Process WINWORD | Stop-Process -Force` here.
    #
    # There was, and it killed every Word on the machine -- other sessions',
    # and the developer's own open document. That is the cardinal rule in this
    # repo: never kill a process we did not start.
    #
    # The obvious repair -- "kill only the pids that appeared after we started"
    # -- is the rule this repo has MEASURED TO BE FALSE. With one concurrent
    # Word started from a separate process, differencing reported **2 new pids
    # for the 1 instance created** (`probe-word-ownership.ps1`, arm A; the
    # finding is carried in `word-pids.mjs`). A cold `New-Object
    # Word.Application` is ~4.5 s here, so the window is wide enough to catch a
    # stranger routinely, and parentage carries no signal either -- every
    # WINWORD parents to the DCOM launcher, not to its creator. Swapping a
    # blanket kill for a differenced kill would trade a visible hazard for a
    # subtle one.
    #
    # An over-broad assertion fails loudly; an over-broad kill destroys
    # silently. Only the destructive operation needs provable attribution --
    # and this probe holds the RCW it created, so quitting *that* is
    # attribution by construction and needs no pid at all.
    #
    # `Quit()` takes no argument on purpose: under Windows PowerShell 5.1
    # `Quit(0)` throws, because the parameters bind as `VARIANT*`. The throw
    # used to be swallowed by an empty `catch {}` while the process survived,
    # which is the "evidence code lies rather than fails" trap -- the probe
    # printed a tidy cleanup line and the leak became somebody else's problem.
    # So a refusal is now reported, and the instance is left running rather
    # than hunted: a visible leak is recoverable, a wrong kill is not.
    if ($word) {
        try { $word.Quit() }
        catch { "  cleanup: Quit threw -- $($_.Exception.Message.Split([char]10)[0]) -- instance left ALIVE deliberately, not hunted by pid" }
        try {
            $rc = [Runtime.InteropServices.Marshal]::ReleaseComObject($word)
            if ($rc -ne 0) { "  cleanup: $rc reference(s) still held after release" }
        } catch { "  cleanup: release threw -- $($_.Exception.Message.Split([char]10)[0])" }
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    # Word exits asynchronously -- `Quit` returns in 3-28 ms but the process
    # lingers seconds -- so this settle is not long enough to guarantee the
    # document handle is released, and deleting $root may fail. That is why the
    # `addr-*` sweep below exists and runs unconditionally: whatever this run
    # cannot remove, the next run collects. Force-killing to free the handle
    # sooner is exactly the behaviour removed above.
    Start-Sleep -Seconds 1
    Remove-Item $root -Recurse -Force -ErrorAction SilentlyContinue
    Get-ChildItem $env:TEMP -Directory -Filter 'addr-*' -ErrorAction SilentlyContinue |
        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    "swept"
}
