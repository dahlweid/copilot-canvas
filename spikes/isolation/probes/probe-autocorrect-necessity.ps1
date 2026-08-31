# probe-autocorrect-necessity.ps1
#
# The question
# ------------
# `word-host.ps1` switches five settings off around authoring so that text goes
# in verbatim, and switching them off edits the USER's Word -- they persist, and
# whichever Word exits cleanly last decides the stored value
# (probe-autocorrect-concurrency.mjs). That is real shared mutable state, so it
# has to be buying something.
#
# probe-autocorrect.ps1 arm F says it buys nothing: 0 of 6 bait lines rewritten
# with every setting ON. But its own positive control, arm H, rewrote only
# 3 of 6 -- so for up to three baits "not rewritten" may mean THE BAIT WAS
# INERT rather than the feature did not fire. A German Word ships a German
# AutoCorrect list, and `teh` -> `the` is an English entry.
#
# An instrument that cannot observe the property it is cited for is the exact
# defect that produced the retracted per-process claim. So this probe does not
# re-run arm F. It asks the prior question:
#
#   WHICH BAITS ARE LIVE ON THIS MACHINE, and therefore which of the five
#   features does the 0-of-N result actually cover?
#
# The five settings are not one mechanism. Two are list-driven and therefore
# locale-dependent; three are algorithmic and are not:
#
#   AutoCorrect.ReplaceText                     -- the replacement LIST. A bait
#                                                  is live only if its trigger
#                                                  is in AutoCorrect.Entries on
#                                                  THIS machine. Checkable.
#   AutoCorrect.CorrectSentenceCaps             -- algorithmic.
#   AutoCorrect.CorrectInitialCaps              -- algorithmic.
#   Options.AutoFormatAsYouTypeReplaceQuotes    -- algorithmic.
#   Options.AutoFormatAsYouTypeReplaceSymbols   -- algorithmic.
#
# So the liveness doubt arm H raises applies to the LIST baits specifically, and
# this probe resolves those against the machine's own list rather than against
# an assumption about which locale is installed.
#
# What this probe can and cannot establish, stated up front so the result is not
# read as more than it is:
#
#   * It CAN establish that a list bait is live: the entry is present and its
#     replacement value is known. If a live trigger is not rewritten, the
#     feature demonstrably did not fire.
#   * It CAN establish that the quote and symbol baits are rewritable at all, by
#     the batch Content.AutoFormat() path (arm D) -- the same positive control
#     arm H used.
#   * It CANNOT positively control the two capitalisation features. There is no
#     programmatic trigger for them: they are keystroke handlers, and this host
#     never types. That is a gap in the evidence and is reported as one, not
#     argued away. It is also, note, the same fact that makes them unable to
#     fire on our path -- but that is reasoning, and this file reports readings.
#
# The host's insertion path is mirrored exactly, not approximated: every
# character `create_document` and `edit_document` write reaches the document
# through Set-ParagraphText -> (Get-TextRange $para).Text = $text
# (word-host.ps1:1717-1734). There is no Selection.TypeText anywhere in the
# host. Arm C therefore assigns Range.Text, because that is the only way text
# actually gets in.
#
# Housekeeping, from the rules this repo learned the hard way:
#   * Argument-less Quit() only. Quit(<arg>) leaks a WINWORD that process exit
#     does not reap.
#   * WINWORD is censused before and after. Only pids absent beforehand are ever
#     touched -- the maintainer's own Word is routinely alive here and killing
#     it destroys their work.
#   * The five settings are captured and restored, because this probe has to
#     turn them ON to measure them and they persist for the user.
#   * SaveAs2 hangs indefinitely after Content.AutoFormat() has run (measured
#     twice, recorded in probe-autocorrect.ps1). Arm D therefore runs last and
#     compares live rather than from disk.

$ErrorActionPreference = 'Stop'

$WD_DO_NOT_SAVE = 0

function Get-WordPids {
    @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }) | Sort-Object
}

$failures = 0
function Check([string] $what, [bool] $ok, [string] $detail) {
    if ($ok) { Write-Host "ok    $what -- $detail" }
    else { $script:failures += 1; Write-Host "FAIL  $what -- $detail" }
}

$pidsBefore = Get-WordPids
Write-Host "WINWORD before: [$($pidsBefore -join ', ')]"

$root = Join-Path $env:TEMP ("acn-" + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $root | Out-Null

# The algorithmic baits. Each targets one named feature, and none of them
# depends on a list entry existing, so "the bait was inert" is not available as
# an explanation for these four the way it is for a list bait.
#
# The quote bait is checked for ANY change, not for a specific curly pair: a
# German Word replaces with low-9/high-6 quotes rather than the English pair, so
# asserting a particular replacement would silently pass a rewrite it did not
# expect.
$algorithmic = @(
    @{ feature = 'Options.AutoFormatAsYouTypeReplaceQuotes'; text = 'She said "hello" and left.' }
    @{ feature = 'Options.AutoFormatAsYouTypeReplaceSymbols'; text = 'width -- height' }
    @{ feature = 'AutoCorrect.CorrectSentenceCaps'; text = 'one. two three' }
    @{ feature = 'AutoCorrect.CorrectInitialCaps'; text = 'THis is fine' }
)

# The triggers arm F used for the list feature. Both are checked against the
# machine's own Entries rather than assumed, which is the whole point of arm A.
$historicalListBaits = @('teh', '(c)')

$app = $null
$prior = $null
$priorFormat = $null
try {
    $app = New-Object -ComObject Word.Application
    $app.Visible = $false
    $app.DisplayAlerts = 0

    $prior = @{
        'ReplaceText'                              = [bool]$app.AutoCorrect.ReplaceText
        'CorrectSentenceCaps'                      = [bool]$app.AutoCorrect.CorrectSentenceCaps
        'CorrectInitialCaps'                       = [bool]$app.AutoCorrect.CorrectInitialCaps
        'Options.AutoFormatAsYouTypeReplaceQuotes' = [bool]$app.Options.AutoFormatAsYouTypeReplaceQuotes
        'Options.AutoFormatAsYouTypeReplaceSymbols' = [bool]$app.Options.AutoFormatAsYouTypeReplaceSymbols
    }
    Write-Host "found state: $(($prior.GetEnumerator() | Sort-Object Name | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join ' ')"

    # Arm D has to switch the BATCH autoformat options on, and those persist for
    # the user exactly as the five under test do. Captured here so the probe
    # cannot commit, in passing, the very fault it exists to investigate.
    $priorFormat = @{
        'AutoFormatReplaceQuotes'    = [bool]$app.Options.AutoFormatReplaceQuotes
        'AutoFormatReplaceSymbols'   = [bool]$app.Options.AutoFormatReplaceSymbols
        'AutoFormatApplyHeadings'    = [bool]$app.Options.AutoFormatApplyHeadings
        'AutoFormatApplyLists'       = [bool]$app.Options.AutoFormatApplyLists
        'AutoFormatApplyBulletedLists' = [bool]$app.Options.AutoFormatApplyBulletedLists
    }

    # ---------------------------------------------------------------- arm A
    # Which list baits are live here? An entry that is not in this Word's list
    # can never be rewritten by any path, so a bait built on one measures
    # nothing and must not be counted as evidence either way.
    Write-Host "`narm A -- which AutoCorrect list entries exist on THIS machine?"

    $entryCount = $app.AutoCorrect.Entries.Count
    Check 'the machine has an AutoCorrect replacement list at all' ($entryCount -gt 0) "$entryCount entries"

    $liveBaits = @()
    foreach ($needle in $historicalListBaits) {
        $value = $null
        try { $value = [string]$app.AutoCorrect.Entries.Item($needle).Value } catch { $value = $null }
        if ($null -ne $value -and $value -ne '') {
            $liveBaits += @{ feature = 'AutoCorrect.ReplaceText'; trigger = $needle; value = $value }
            Write-Host "      '$needle' IS in the list -> '$value' (live bait)"
        }
        else {
            Write-Host "      '$needle' is NOT in the list -- any 'not rewritten' result for it measures nothing"
        }
    }

    # Whatever the historical baits turn out to be, take baits from the machine's
    # own list too. Short pure-ASCII triggers only: a trigger carrying an umlaut
    # cannot be distinguished from an encoding fault on the way back out of the
    # file, and one containing a space is not a single word.
    $sampled = @()
    for ($k = 1; $k -le $entryCount -and $sampled.Count -lt 3; $k++) {
        $entry = $app.AutoCorrect.Entries.Item($k)
        $name = [string]$entry.Name
        $value = [string]$entry.Value
        if ($name -cmatch '^[a-z]{3,12}$' -and $value -cmatch '^[A-Za-z]{2,20}$' -and $value -ne $name) {
            $sampled += @{ feature = 'AutoCorrect.ReplaceText'; trigger = $name; value = $value }
        }
    }
    foreach ($s in $sampled) {
        Write-Host "      sampled from this machine's list: '$($s.trigger)' -> '$($s.value)'"
    }
    $liveBaits += $sampled

    Check 'at least one PROVABLY LIVE list bait was found' ($liveBaits.Count -gt 0) "$($liveBaits.Count) live"

    # ---------------------------------------------------------------- arm B
    # Turn everything ON. This is the condition arm F ran under and the
    # condition the host would be in if the suppression were deleted.
    Write-Host "`narm B -- all five settings ON, which is the condition under test"
    $app.AutoCorrect.ReplaceText = $true
    $app.AutoCorrect.CorrectSentenceCaps = $true
    $app.AutoCorrect.CorrectInitialCaps = $true
    $app.Options.AutoFormatAsYouTypeReplaceQuotes = $true
    $app.Options.AutoFormatAsYouTypeReplaceSymbols = $true

    $readBack = @(
        [bool]$app.AutoCorrect.ReplaceText
        [bool]$app.AutoCorrect.CorrectSentenceCaps
        [bool]$app.AutoCorrect.CorrectInitialCaps
        [bool]$app.Options.AutoFormatAsYouTypeReplaceQuotes
        [bool]$app.Options.AutoFormatAsYouTypeReplaceSymbols
    )
    # Read back rather than trusting the assignments: a misspelled COM property
    # reads $null and throws only on assignment, and "five assignments did not
    # throw" is not the same observation as "Word reports them on".
    Check 'Word reports all five ON' (($readBack | Where-Object { -not $_ }).Count -eq 0) ($readBack -join ',')

    # ---------------------------------------------------------------- arm C
    # The measurement. Every bait goes in the way the host puts text in, and the
    # text is compared to what was asked for.
    #
    # Compared LIVE rather than saved and reopened, and that is a harness
    # limitation worth stating: SaveAs2 hangs indefinitely in this probe --
    # measured twice here, >140 s with the five settings ON and again with them
    # OFF, so it is not the settings -- while every other call returns in
    # milliseconds. probe-autocorrect.ps1 records the same call hanging under
    # Start-Job. The product's own save is not affected and is not being taken on
    # trust: `create_document` saves through this exact host in
    # probe-autocorrect-concurrency.mjs arm 6 and in create-smoke.mjs, both of
    # which ran green today. What this arm therefore cannot see is a rewrite that
    # happens during the save rather than at insertion -- which is not what
    # autocorrect is, but is named here rather than left for someone to discover.
    Write-Host "`narm C -- do live baits survive the host's own insertion path verbatim?"

    $baits = @()
    foreach ($a in $algorithmic) { $baits += @{ feature = $a.feature; text = $a.text; live = 'algorithmic -- no list entry needed' } }
    foreach ($b in $liveBaits) { $baits += @{ feature = $b.feature; text = "$($b.trigger) danach"; live = "entry present -> '$($b.value)'" } }

    $doc = $app.Documents.Add()
    $written = @()
    try {
        foreach ($b in $baits) {
            $para = $doc.Paragraphs.Add()
            # Set-ParagraphText's exact mechanism: trim the paragraph mark off
            # the range by POSITION SPAN, never by string length, then assign.
            $range = $para.Range
            $raw = [string]$range.Text
            $visible = $raw.TrimEnd("`r", [char]7, "`n")
            $drop = ($range.End - $range.Start) - $visible.Length
            if ($drop -gt 0) { $range.MoveEnd(1, -$drop) | Out-Null }
            $range.Text = $b.text
        }
        # A rewrite that arrives on a background pass rather than on the
        # assignment would be missed by reading back immediately, and "we read it
        # too early" is not a difference this probe could otherwise tell from
        # "it never fired".
        Start-Sleep -Seconds 2
        for ($k = 1; $k -le $doc.Paragraphs.Count; $k++) {
            $t = ([string]$doc.Paragraphs.Item($k).Range.Text).TrimEnd("`r", [char]7, "`n")
            if ($t -ne '') { $written += $t }
        }
    }
    finally {
        $doc.Close($WD_DO_NOT_SAVE)
    }

    foreach ($b in $baits) {
        $hit = $written | Where-Object { $_ -eq $b.text }
        $wasRewritten = -not $hit
        $detail = if ($wasRewritten) {
            "REWRITTEN: asked for '$($b.text)', document holds '$(($written | Where-Object { $_ -ne $b.text }) -join ' | ')'"
        }
        else {
            "verbatim: '$($b.text)'  [$($b.live)]"
        }
        Check "$($b.feature) did not rewrite its bait" (-not $wasRewritten) $detail
    }

    # ---------------------------------------------------------------- arm D
    # The positive control, and the reason arm C's result is worth anything:
    # if nothing here rewrites either, arm C only proves the baits are inert.
    # Runs LAST because SaveAs2 hangs indefinitely once AutoFormat has run, so
    # this arm compares live text and its document is never saved.
    Write-Host "`narm D -- positive control: can these baits be rewritten at all?"

    $app.Options.AutoFormatReplaceQuotes = $true
    $app.Options.AutoFormatReplaceSymbols = $true
    $app.Options.AutoFormatApplyHeadings = $false
    $app.Options.AutoFormatApplyLists = $false
    $app.Options.AutoFormatApplyBulletedLists = $false

    $control = $app.Documents.Add()
    try {
        foreach ($b in $baits) {
            $para = $control.Paragraphs.Add()
            $range = $para.Range
            $raw = [string]$range.Text
            $visible = $raw.TrimEnd("`r", [char]7, "`n")
            $drop = ($range.End - $range.Start) - $visible.Length
            if ($drop -gt 0) { $range.MoveEnd(1, -$drop) | Out-Null }
            $range.Text = $b.text
        }
        $control.Content.AutoFormat()

        $live = @()
        for ($k = 1; $k -le $control.Paragraphs.Count; $k++) {
            $t = ([string]$control.Paragraphs.Item($k).Range.Text).TrimEnd("`r", [char]7, "`n")
            if ($t -ne '') { $live += $t }
        }

        $rewritten = 0
        foreach ($b in $baits) {
            $survived = $live | Where-Object { $_ -eq $b.text }
            if ($survived) {
                Write-Host "      $($b.feature): '$($b.text)' NOT rewritten even by AutoFormat -- this bait is INERT here, and arm C's clean result for it proves nothing"
            }
            else {
                $rewritten += 1
                Write-Host "      $($b.feature): '$($b.text)' WAS rewritten by AutoFormat -- bait demonstrably live, so arm C's clean result for it is real"
            }
        }
        Check 'AutoFormat rewrote at least one bait, so the instrument can detect a rewrite' ($rewritten -gt 0) "$rewritten of $($baits.Count) rewritten"
    }
    finally {
        $control.Close($WD_DO_NOT_SAVE)
    }
}
finally {
    if ($null -ne $app) {
        if ($null -ne $prior) {
            # Put the user's settings back. This probe had to turn them ON to
            # measure them, and they persist -- leaving them changed is the very
            # damage the work around this probe exists to undo.
            try {
                $app.AutoCorrect.ReplaceText = $prior['ReplaceText']
                $app.AutoCorrect.CorrectSentenceCaps = $prior['CorrectSentenceCaps']
                $app.AutoCorrect.CorrectInitialCaps = $prior['CorrectInitialCaps']
                $app.Options.AutoFormatAsYouTypeReplaceQuotes = $prior['Options.AutoFormatAsYouTypeReplaceQuotes']
                $app.Options.AutoFormatAsYouTypeReplaceSymbols = $prior['Options.AutoFormatAsYouTypeReplaceSymbols']
                Write-Host "`nrestored the five settings to the values found at start"
            }
            catch {
                $failures += 1
                Write-Host "FAIL  could not restore the five settings: $($_.Exception.Message)"
            }
        }
        if ($null -ne $priorFormat) {
            try {
                $app.Options.AutoFormatReplaceQuotes = $priorFormat['AutoFormatReplaceQuotes']
                $app.Options.AutoFormatReplaceSymbols = $priorFormat['AutoFormatReplaceSymbols']
                $app.Options.AutoFormatApplyHeadings = $priorFormat['AutoFormatApplyHeadings']
                $app.Options.AutoFormatApplyLists = $priorFormat['AutoFormatApplyLists']
                $app.Options.AutoFormatApplyBulletedLists = $priorFormat['AutoFormatApplyBulletedLists']
                Write-Host "restored the batch autoformat options arm D changed"
            }
            catch {
                $failures += 1
                Write-Host "FAIL  could not restore the autoformat options: $($_.Exception.Message)"
            }
        }
        # Argument-less. Quit(<arg>) leaks a WINWORD that process exit does not reap.
        try { $app.Quit() } catch { }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
        $app = $null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue

    Start-Sleep -Seconds 3
    $pidsAfter = Get-WordPids
    $leaked = @($pidsAfter | Where-Object { $pidsBefore -notcontains $_ })
    Write-Host "WINWORD after: [$($pidsAfter -join ', ')]"
    foreach ($leakedPid in $leaked) {
        Write-Host "      reaping this probe's own Word, pid $leakedPid"
        try { Stop-Process -Id $leakedPid -Force -ErrorAction Stop } catch { }
    }
    Check 'no WINWORD was left behind' ($leaked.Count -eq 0) $(if ($leaked.Count) { "leaked $($leaked -join ', ') (reaped)" } else { '' })
}

if ($failures -gt 0) {
    Write-Host "`n$failures check(s) failed"
    exit 1
}
Write-Host "`nall checks passed"
exit 0
