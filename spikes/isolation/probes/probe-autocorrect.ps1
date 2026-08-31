# probe-autocorrect.ps1
#
# The question
# ------------
# `create_document` writes text the agent was asked to write. Word's autocorrect
# and autoformat-as-you-type rewrite text *without raising anything*, so a
# document can silently end up containing something other than what was asked
# for. That is a correctness bug, not a cosmetic one, and it has to be measured
# rather than assumed.
#
# Three things need settling before any authoring code exists:
#
#   1. What are these settings actually called? A COM property that does not
#      exist reads as $null, which is falsy -- so a probe that *guesses* a name
#      reports "the setting is off" for a setting that was never there. Three
#      names guessed from memory turned out not to exist on this Word
#      (Options.AutoFormatAsYouTypeReplaceHyphens, ...ReplaceHyphensWithDash,
#      AutoCorrect.CorrectTwoInitialCapitals). Arm D enumerates instead.
#
#   2. Does assigning `Range.Text` trigger autocorrect at all? Autocorrect and
#      autoformat-as-you-type are *typing* features. If assignment does not run
#      them, `create_document` never has to touch a setting, which is the
#      strongest possible form of "the user's own Word is unaffected".
#
#   3. If we did have to switch them off, is that scoped to our instance or does
#      it edit the user's Word? Word is multi-instance, but that says nothing
#      about whether Application.AutoCorrect is per-process or per-user.
#
# Harness note -- this cost two runs
# ----------------------------------
# The first two versions of this probe ran each arm inside `Start-Job`. Every
# arm that reached `Document.SaveAs2` hung there indefinitely, on both the
# Range.Text and the Selection.TypeText path, with DisplayAlerts = 0 and the
# target directory already created. probe-saveas-apartment.ps1 then ran the
# identical body as a real `powershell.exe` and SaveAs2 returned in 149 ms (STA)
# and 138 ms (MTA) -- so the hang is an artefact of `Start-Job`, and it is *not*
# apartment state, since a real MTA process saves fine.
#
# The finding to keep: a Word COM call that pumps a message loop can wedge on a
# PowerShell job runspace thread. Probes here launch a real powershell.exe.
#
# Every value the worker needs is passed as a discrete argv element via
# Start-Process -ArgumentList. Nothing is interpolated into a command string, so
# a workdir path containing an apostrophe or an '&' cannot reach a parser.
#
# Cleanup kills only PIDs absent before the probe started -- several WINWORD.EXE
# belonging to other sessions are routinely alive here, and killing one destroys
# someone's unsaved work.

param(
    # Optional: run a subset of the arms. Each arm starts a Word and the full
    # set takes minutes, so re-examining one arm should not cost a full run.
    [string[]] $Only
)

$ErrorActionPreference = 'Stop'

function Get-WordPids {
    @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }) | Sort-Object
}

$pidsBefore = Get-WordPids
Write-Host "WINWORD before: $($pidsBefore -join ', ')"

$root = Join-Path $env:TEMP ("ac-" + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $root | Out-Null
$worker = Join-Path $root 'worker.ps1'

@'
param(
    [Parameter(Mandatory = $true)][ValidateSet('D', 'A', 'B', 'E', 'F', 'G', 'H', 'C')][string] $Arm,
    [Parameter(Mandatory = $true)][string] $Root,
    [Parameter(Mandatory = $true)][string] $Trace,
    [Parameter(Mandatory = $true)][string] $InstanceBScript
)

# Parameters are deliberately named so that no COM handle later in this script
# shares a name with one of them. A param() type constraint follows the variable
# for the rest of the scope and PowerShell names are case-insensitive, so
# `$doc = $w.Documents.Add()` under a `[string] $Doc` parameter *coerces the
# Document to a string* instead of failing -- and the error then surfaces on a
# later, correct line as "the property Text was not found on this object".
# That cost a run of probe-saveas-apartment.ps1.

$ErrorActionPreference = 'Stop'
$sw = [Diagnostics.Stopwatch]::StartNew()
function Step([string] $what) {
    Add-Content -LiteralPath $Trace -Value ("{0,7}ms  {1}" -f $sw.ElapsedMilliseconds, $what)
}

$WD_DO_NOT_SAVE = 0
$WD_FORMAT_DOCX = 16   # wdFormatDocumentDefault
$WD_CHARACTER   = 1

# The bait. Each line targets one named feature that is ON on this machine, so
# a "nothing was rewritten" result is only meaningful if these can in fact be
# rewritten -- which is what arm E exists to demonstrate.
$bait = @(
    'She said "hello" and left.'   # AutoFormatAsYouTypeReplaceQuotes -> curly quotes
    'width -- height'              # AutoFormatAsYouTypeReplaceSymbols -> en/em dash
    'teh result'                   # AutoCorrect.ReplaceText          -> 'the result'
    'one. two three'               # AutoCorrect.CorrectSentenceCaps  -> 'one. Two three'
    'THis is fine'                 # AutoCorrect.CorrectInitialCaps   -> 'This is fine'
    '(c) 2026'                     # AutoCorrect.ReplaceText          -> a copyright glyph
)

function New-Word {
    $a = New-Object -ComObject Word.Application
    $a.Visible = $false
    $a.DisplayAlerts = 0
    return $a
}

function Read-Settings($a) {
    # Names taken from arm D's enumeration, never from memory.
    [ordered]@{
        'AutoCorrect.ReplaceText'                   = $a.AutoCorrect.ReplaceText
        'AutoCorrect.CorrectSentenceCaps'           = $a.AutoCorrect.CorrectSentenceCaps
        'AutoCorrect.CorrectInitialCaps'            = $a.AutoCorrect.CorrectInitialCaps
        'Options.AutoFormatAsYouTypeReplaceQuotes'  = $a.Options.AutoFormatAsYouTypeReplaceQuotes
        'Options.AutoFormatAsYouTypeReplaceSymbols' = $a.Options.AutoFormatAsYouTypeReplaceSymbols
    }
}

function Show-Settings($a, [string] $label) {
    "${label}:"
    (Read-Settings $a).GetEnumerator() | ForEach-Object { "  $($_.Key) = $($_.Value)" }
}

# What matters is what the *file* says, not what the live object model reports:
# a rewrite could in principle happen at save. So every arm reopens from disk.
function Compare-Saved($a, [string] $file, $expected) {
    $check = $a.Documents.Open($file, $false, $true)
    $rewritten = 0
    $j = 1
    foreach ($text in $expected) {
        $got = ([string]$check.Paragraphs.Item($j).Range.Text).TrimEnd("`r", [char]7, "`n")
        if ($got -ceq $text) {
            "  [verbatim ] $(Show-Escaped $text)"
        } else {
            "  [REWRITTEN] asked: $(Show-Escaped $text)"
            "              got:   $(Show-Escaped $got)"
            $rewritten++
        }
        $j++
    }
    "  paragraphs rewritten: $rewritten of $($expected.Count)"
    $check.Close($WD_DO_NOT_SAVE)
}

# Render a string so a non-ASCII character survives being printed. THIS IS NOT
# COSMETIC. Everything below exists to find out which character Word substituted,
# and the console cannot show one: stdout here is CP850, and reading it back
# elsewhere as Windows-1252 turns U+00A9 into a cedilla. Worse, a curly quote and
# an en dash have no CP850 mapping at all and best-fit to '"' and '-', so a
# rewritten line prints identical to the line it replaced and the substitution
# this probe was written to catch reads as a no-op.
#
# The comparisons themselves are sound -- they are -ceq against the live COM
# string and never went near the console. It is only the report that was lossy,
# which is the more dangerous of the two: an earlier revision of this probe
# printed the copyright sign as a cedilla, and a comment in word-host.ps1 came from
# that output, describing the result as "(c) as a symbol" because the author
# could not see what it was. Measuring correctly and reporting illegibly is
# still a wrong answer at the far end. (That earlier output rendered the
# copyright sign as a cedilla; this comment does not reproduce it, because this
# file must stay ASCII-only -- Windows PowerShell reads a BOM-less UTF-8 .ps1 as
# ANSI, so the character would be corrupted here too.)
function Show-Escaped([string] $s) {
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $s.ToCharArray()) {
        $code = [int][char]$ch
        if ($code -ge 32 -and $code -le 126) { [void]$sb.Append($ch) }
        else { [void]$sb.AppendFormat('<U+{0:X4}>', $code) }
    }
    return $sb.ToString()
}

# The same comparison against a live document, without going through disk.
# Arm H needs this because `SaveAs2` hangs indefinitely after
# `Content.AutoFormat()` has run -- measured, twice. Nothing on the
# `create_document` path calls AutoFormat, so that hang is not a constraint on
# the product, but it does mean the positive control cannot save.
function Compare-Live($document, $expected) {
    $rewritten = 0
    $j = 1
    foreach ($text in $expected) {
        $got = ([string]$document.Paragraphs.Item($j).Range.Text).TrimEnd("`r", [char]7, "`n")
        if ($got -ceq $text) {
            "  [verbatim ] $(Show-Escaped $text)"
        } else {
            "  [REWRITTEN] asked: $(Show-Escaped $text)"
            "              got:   $(Show-Escaped $got)"
            $rewritten++
        }
        $j++
    }
    "  paragraphs rewritten: $rewritten of $($expected.Count)"
}

# Assign paragraph text the way the shipping code does: trim the paragraph mark
# off the range first, deriving the trim from the *position span* rather than
# from string length. Inside a table a cell paragraph's Range.Text ends with
# "`r" plus chr(7) -- two characters in the string -- while End - Start counts
# the end-of-cell mark as one position.
function Set-ParagraphText($para, [string] $text) {
    $r = $para.Range
    $visible = ([string]$r.Text).TrimEnd("`r", [char]7, "`n")
    $drop = ($r.End - $r.Start) - $visible.Length
    if ($drop -gt 0) { $r.MoveEnd($WD_CHARACTER, -$drop) | Out-Null }
    $r.Text = $text
}

function Write-ByRangeText($a, [string] $file) {
    $document = $a.Documents.Add()
    Step 'Documents.Add returned'
    $i = 1
    foreach ($text in $bait) {
        if ($i -gt 1) { $document.Paragraphs.Add() | Out-Null }
        Set-ParagraphText $document.Paragraphs.Item($i) $text
        Step "wrote paragraph $i"
        $i++
    }
    Step 'calling SaveAs2'
    $document.SaveAs2($file, $WD_FORMAT_DOCX)
    Step 'SaveAs2 returned'
    $document.Close($WD_DO_NOT_SAVE)
}

function Write-BySelection($a, [string] $file) {
    $document = $a.Documents.Add()
    Step 'Documents.Add returned'
    $sel = $a.Selection
    $i = 1
    foreach ($text in $bait) {
        if ($i -gt 1) { $sel.TypeParagraph() }
        $sel.TypeText($text)
        Step "typed paragraph $i"
        $i++
    }
    Step 'calling SaveAs2'
    $document.SaveAs2($file, $WD_FORMAT_DOCX)
    Step 'SaveAs2 returned'
    $document.Close($WD_DO_NOT_SAVE)
}

# The same text, one character per TypeText call.
#
# This exists because the whole-string version could not go dirty. AutoCorrect
# fires on a *terminator* -- a space, a punctuation mark, a paragraph mark --
# after the word it is going to replace. Handing Word the entire string in one
# call gives it the word and its terminator in a single operation, so there may
# never be a moment where 'teh' stands finished at the insertion point. Feeding
# it a character at a time is the closest the object model gets to typing, and
# is the strongest chance of provoking a rewrite.
function Write-BySelectionCharwise($a, [string] $file) {
    $document = $a.Documents.Add()
    Step 'Documents.Add returned'
    $sel = $a.Selection
    $i = 1
    foreach ($text in $bait) {
        if ($i -gt 1) { $sel.TypeParagraph() }
        foreach ($ch in $text.ToCharArray()) { $sel.TypeText([string]$ch) }
        Step "typed paragraph $i charwise"
        $i++
    }
    # A trailing paragraph mark is itself a terminator, so give the last line one
    # before saving; otherwise the final word never gets the event that would
    # trigger a replacement.
    $sel.TypeParagraph()
    Step 'calling SaveAs2'
    $document.SaveAs2($file, $WD_FORMAT_DOCX)
    Step 'SaveAs2 returned'
    $document.Close($WD_DO_NOT_SAVE)
}

function Set-AutocorrectOff($a) {
    $a.AutoCorrect.ReplaceText = $false
    $a.AutoCorrect.CorrectSentenceCaps = $false
    $a.AutoCorrect.CorrectInitialCaps = $false
    $a.Options.AutoFormatAsYouTypeReplaceQuotes = $false
    $a.Options.AutoFormatAsYouTypeReplaceSymbols = $false
}

$app = $null
try {
    switch ($Arm) {

        # --- D: what are these properties actually called? -------------------
        'D' {
            $app = New-Word
            Step 'ready'
            'Application.Options -- AutoFormatAsYouType*:'
            @($app.Options | Get-Member -MemberType Property -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -like 'AutoFormatAsYouType*' } |
                ForEach-Object { "  $($_.Name) = $($app.Options.$($_.Name))" })
            'Application.Options -- AutoFormat* (the batch command, not as-you-type):'
            @($app.Options | Get-Member -MemberType Property -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -like 'AutoFormat*' -and $_.Name -notlike 'AutoFormatAsYouType*' } |
                ForEach-Object { "  $($_.Name) = $($app.Options.$($_.Name))" })
            'Application.AutoCorrect -- Correct*/Replace*:'
            @($app.AutoCorrect | Get-Member -MemberType Property -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -like 'Correct*' -or $_.Name -like 'Replace*' } |
                ForEach-Object { "  $($_.Name) = $($app.AutoCorrect.$($_.Name))" })
            Step 'enumerated'
        }

        # --- A: Range.Text, settings left exactly as the user has them -------
        'A' {
            $app = New-Word
            Step 'ready'
            Show-Settings $app 'settings as found'
            $file = Join-Path $Root 'a.docx'
            Write-ByRangeText $app $file
            'reopened from disk:'
            Compare-Saved $app $file $bait
            Step 'compared'
        }

        # --- B: Selection.TypeText, settings left as the user has them -------
        # This is the control. If B rewrites nothing either, the bait is inert
        # on this machine and arm A proves nothing -- arm A's clean result only
        # means something if a dirty result was reachable.
        'B' {
            $app = New-Word
            Step 'ready'
            Show-Settings $app 'settings as found'
            $file = Join-Path $Root 'b.docx'
            Write-BySelection $app $file
            'reopened from disk:'
            Compare-Saved $app $file $bait
            Step 'compared'
        }

        # --- E: Selection.TypeText with the settings switched off ------------
        # Shows the switches actually govern the rewriting seen in arm B, so
        # they are a real fallback if arm A ever comes out dirty.
        'E' {
            $app = New-Word
            Step 'ready'
            Set-AutocorrectOff $app
            Show-Settings $app 'settings after switching off'
            Step 'switched off'
            $file = Join-Path $Root 'e.docx'
            Write-BySelection $app $file
            'reopened from disk:'
            Compare-Saved $app $file $bait
            Step 'compared'
        }

        # --- F: Selection.TypeText one character at a time, settings on ------
        # The real control. Arms A, B and E all came out verbatim, which means
        # nothing unless *some* arm can come out dirty. If F is also verbatim,
        # the conclusion is not "we suppressed autocorrect" but "this bait
        # cannot be rewritten through the object model at all", and the probe
        # has to say so rather than claim a pass.
        'F' {
            $app = New-Word
            Step 'ready'
            Show-Settings $app 'settings as found'
            $file = Join-Path $Root 'f.docx'
            Write-BySelectionCharwise $app $file
            'reopened from disk:'
            Compare-Saved $app $file $bait
            Step 'compared'
        }

        # --- G: the same, with the settings switched off ---------------------
        'G' {
            $app = New-Word
            Step 'ready'
            Set-AutocorrectOff $app
            Show-Settings $app 'settings after switching off'
            Step 'switched off'
            $file = Join-Path $Root 'g.docx'
            Write-BySelectionCharwise $app $file
            'reopened from disk:'
            Compare-Saved $app $file $bait
            Step 'compared'
        }

        # --- H: the positive control ----------------------------------------
        # Arms A, B, E, F and G all came out verbatim -- every one of them,
        # including the charwise one with every setting ON. "0 of 6 rewritten"
        # is worth nothing until something demonstrates this bait *can* be
        # rewritten by this Word, on this machine, in this language. Otherwise
        # the clean result might only mean the bait was inert -- a German Word
        # ships a German AutoCorrect list, and 'teh' -> 'the' is an English
        # entry.
        #
        # Two independent checks:
        #   - is the bait actually in the AutoCorrect replacement list?
        #   - does Range.AutoFormat() -- the batch command, which is explicitly
        #     invoked rather than triggered by typing -- rewrite it?
        # If H comes out dirty, the other arms' verbatim results are real.
        'H' {
            $app = New-Word
            Step 'ready'

            'AutoCorrect replacement entries for the bait:'
            foreach ($needle in 'teh', '(c)', '(C)') {
                $hit = $null
                try { $hit = $app.AutoCorrect.Entries.Item($needle) } catch { }
                if ($null -ne $hit) { "  '$needle' -> '$(Show-Escaped $hit.Value)'" } else { "  '$needle' -> <no entry>" }
            }
            "AutoCorrect entry count: $($app.AutoCorrect.Entries.Count)"
            # A sample of what this Word's list actually contains, so a future
            # revision can pick bait that is live *here* rather than bait that
            # is live in an English Word.
            'a sample of short ASCII entries on this machine:'
            $shown = 0
            for ($k = 1; $k -le $app.AutoCorrect.Entries.Count -and $shown -lt 20; $k++) {
                $e = $app.AutoCorrect.Entries.Item($k)
                if ($e.Name -cmatch '^[ -~]{1,10}$') { "  '$($e.Name)' -> '$(Show-Escaped $e.Value)'"; $shown++ }
            }
            Step 'entries read'

            # Options.AutoFormatReplaceQuotes and friends are a *separate* set
            # from the AsYouType ones, and they govern the batch command.
            $app.Options.AutoFormatReplaceQuotes = $true
            $app.Options.AutoFormatReplaceSymbols = $true
            $app.Options.AutoFormatApplyHeadings = $false
            $app.Options.AutoFormatApplyLists = $false
            $app.Options.AutoFormatApplyBulletedLists = $false
            Step 'autoformat options set'

            $document = $app.Documents.Add()
            $i = 1
            foreach ($text in $bait) {
                if ($i -gt 1) { $document.Paragraphs.Add() | Out-Null }
                Set-ParagraphText $document.Paragraphs.Item($i) $text
                $i++
            }
            Step 'wrote bait by Range.Text'

            'before AutoFormat:'
            Compare-Live $document $bait

            $document.Content.AutoFormat()
            Step 'Content.AutoFormat() returned'

            # Compared live, not from disk: SaveAs2 hangs after AutoFormat.
            'after AutoFormat (REWRITTEN lines here are the point of this arm):'
            Compare-Live $document $bait
            Step 'compared'

            $document.Close($WD_DO_NOT_SAVE)
            Step 'closed'
        }

        # --- C: is switching them off scoped to one instance? ----------------
        #
        # RETRACTED. This arm concluded "per-process": A switches them off, B
        # starts and still reads them on, therefore our suppression cannot reach
        # the user's Word. The observation reproduces. The conclusion is false.
        #
        # The defect is one line: B is started **while A is still alive**. A
        # concurrent reader sees the pre-write value whether these are
        # per-process or per-user. Isolation and persistence-with-lag are the
        # *same observation* here, and this arm cannot separate them -- it was cited
        # for exactly the property its own construction made invisible.
        #
        # Re-measured sequentially (set in A, QUIT A, then read a fresh
        # instance): all five come back changed. They are per-user and they
        # persist. Suppression therefore edits the user's own Word permanently,
        # which is why word-host.ps1 now captures and restores around the
        # authoring call rather than switching them off at startup.
        #
        # The registry readings below are also evidence of nothing. The HKCU
        # Word\Options key genuinely is absent, but that is simply not where
        # these live, so its absence never supported the isolation reading.
        #
        # Arm F's finding stands and is untouched: the baits rewrite 0 of 6 with
        # every setting ON, so a verbatim bait is a forward guard, not proof of
        # suppression.
        #
        # The sequential measurement lives in probe-autocorrect-untouched.mjs,
        # which also proves the tool leaves the user's next Word unchanged.
        'C' {
            $key = 'HKCU:\Software\Microsoft\Office\16.0\Word\Options'
            function Show-Registry([string] $label) {
                "registry ${label}:"
                foreach ($name in 'AutoFormatAsYouTypeReplaceQuotes', 'AutoFormatAsYouTypeReplaceSymbols') {
                    $v = try { (Get-ItemProperty -Path $key -Name $name -ErrorAction Stop).$name } catch { '<absent>' }
                    "    ${name}: $v"
                }
            }

            $app = New-Word
            Step 'instance A ready'
            Show-Settings $app 'instance A, as found'
            $original = Read-Settings $app
            Show-Registry 'before'

            $app.AutoCorrect.ReplaceText = $false
            $app.AutoCorrect.CorrectSentenceCaps = $false
            $app.AutoCorrect.CorrectInitialCaps = $false
            $app.Options.AutoFormatAsYouTypeReplaceQuotes = $false
            $app.Options.AutoFormatAsYouTypeReplaceSymbols = $false
            Step 'switched off in A'
            Show-Settings $app 'instance A, after switching off'
            Show-Registry 'while A holds them off'

            # A second Word.Application in the *same process* fails with
            # CO_E_SERVER_EXEC_FAILURE, so instance B is started as its own
            # powershell.exe -- discrete argv again, no command line built. The
            # script itself is written by the parent and handed over as a path,
            # because a here-string cannot be nested: the inner terminator sits
            # at column 0 and would close the outer one.
            $bOut = Join-Path $Root 'c-instance-b.txt'

            $bProc = Start-Process -FilePath 'powershell.exe' -PassThru -WindowStyle Hidden -Wait -ArgumentList @(
                '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass'
                '-File', $InstanceBScript
                '-ResultPath', $bOut
            )
            Step "instance B exited $($bProc.ExitCode)"
            'instance B, freshly started while A holds them off:'
            if (Test-Path -LiteralPath $bOut) { Get-Content -LiteralPath $bOut } else { '  <no result>' }

            foreach ($entry in $original.GetEnumerator()) {
                switch ($entry.Key) {
                    'AutoCorrect.ReplaceText'                   { $app.AutoCorrect.ReplaceText = $entry.Value }
                    'AutoCorrect.CorrectSentenceCaps'           { $app.AutoCorrect.CorrectSentenceCaps = $entry.Value }
                    'AutoCorrect.CorrectInitialCaps'            { $app.AutoCorrect.CorrectInitialCaps = $entry.Value }
                    'Options.AutoFormatAsYouTypeReplaceQuotes'  { $app.Options.AutoFormatAsYouTypeReplaceQuotes = $entry.Value }
                    'Options.AutoFormatAsYouTypeReplaceSymbols' { $app.Options.AutoFormatAsYouTypeReplaceSymbols = $entry.Value }
                }
            }
            Step 'restored A'
            Show-Settings $app 'instance A, restored'
            Show-Registry 'after'
        }
    }
} catch {
    Step "ERROR $($_.Exception.GetType().FullName): $($_.Exception.Message)"
    "ERROR: $($_.Exception.Message)"
} finally {
    if ($null -ne $app) {
        # Quit(), never Quit(<arg>). Under Windows PowerShell 5.1 the argument
        # form does not bind: it throws and the Word survives, and process exit
        # does not reap it (PLAN.md 20, probe-quit0-leak.ps1). The no-argument
        # form takes the same default. Reporting via Step rather than swallowing,
        # because the empty catch is what hid the leak.
        try { $app.Quit() } catch { Step "Quit() FAILED (Word may leak) -- $($_.Exception.Message.Split([char]10)[0])" }
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch { }
    }
    Step 'quit'
}
'@ | Set-Content -LiteralPath $worker -Encoding UTF8

# Written by the parent, not by the worker: a here-string cannot be nested,
# because the inner terminator sits at column 0 and closes the outer one.
$instanceB = Join-Path $root 'instance-b.ps1'
@'
param([Parameter(Mandatory = $true)][string] $ResultPath)
$b = New-Object -ComObject Word.Application
$b.Visible = $false
$b.DisplayAlerts = 0
$lines = @(
    "  AutoCorrect.ReplaceText                   = $($b.AutoCorrect.ReplaceText)"
    "  AutoCorrect.CorrectSentenceCaps           = $($b.AutoCorrect.CorrectSentenceCaps)"
    "  AutoCorrect.CorrectInitialCaps            = $($b.AutoCorrect.CorrectInitialCaps)"
    "  Options.AutoFormatAsYouTypeReplaceQuotes  = $($b.Options.AutoFormatAsYouTypeReplaceQuotes)"
    "  Options.AutoFormatAsYouTypeReplaceSymbols = $($b.Options.AutoFormatAsYouTypeReplaceSymbols)"
)
Set-Content -LiteralPath $ResultPath -Value $lines -Encoding UTF8
# Quit(), never Quit(<arg>). Under Windows PowerShell 5.1 the argument form does
# not bind: it throws and the Word survives, and process exit does not reap it
# (PLAN.md 20, probe-quit0-leak.ps1).
#
# The report goes into the result file because that is the only channel anyone
# reads: the parent starts this with `Start-Process -WindowStyle Hidden`, whose
# console goes nowhere, so a catch writing to stdout or stderr here would be a
# swallow wearing a report's clothes.
$quitNote = 'ok'
try { $b.Quit() } catch { $quitNote = 'FAILED (Word may leak) -- ' + $_.Exception.Message.Split([char]10)[0] }
Add-Content -LiteralPath $ResultPath -Value "  instance B Quit(): $quitNote" -Encoding UTF8
'@ | Set-Content -LiteralPath $instanceB -Encoding UTF8

function Invoke-Arm([string] $arm, [string] $title) {
    if ($Only -and $Only -notcontains $arm) { return }
    $out = Join-Path $root "$arm.out"
    $err = Join-Path $root "$arm.err"
    $trace = Join-Path $root "$arm.trace"
    New-Item -ItemType File -Path $trace | Out-Null

    Write-Host ""
    Write-Host "=== Arm ${arm}: $title ==="

    $p = Start-Process -FilePath 'powershell.exe' -PassThru -WindowStyle Hidden `
        -RedirectStandardOutput $out -RedirectStandardError $err -ArgumentList @(
        '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass'
        '-File', $worker
        '-Arm', $arm
        '-Root', $root
        '-Trace', $trace
        '-InstanceBScript', $instanceB
    )

    # Poll to a deadline. Word's teardown is load-dependent, so a flat sleep
    # would let a busy machine decide the outcome.
    $deadline = (Get-Date).AddSeconds(120)
    while (-not $p.HasExited -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 250 }
    if (-not $p.HasExited) {
        Write-Host "  HUNG (no exit within 120s)"
        try { Stop-Process -Id $p.Id -Force } catch { }
    }

    if (Test-Path -LiteralPath $out) { Get-Content -LiteralPath $out | ForEach-Object { Write-Host $_ } }
    $errText = if (Test-Path -LiteralPath $err) { (Get-Content -LiteralPath $err -Raw) } else { '' }
    if ($errText -and $errText.Trim()) { Write-Host "  stderr: $($errText.Trim())" }
    Write-Host "  --- step trace ---"
    Get-Content -LiteralPath $trace | ForEach-Object { Write-Host "  $_" }
}

Invoke-Arm 'D' 'what are these properties actually called?'
Invoke-Arm 'A' 'Range.Text, settings untouched'
Invoke-Arm 'B' 'Selection.TypeText whole string, settings untouched'
Invoke-Arm 'E' 'Selection.TypeText whole string, settings switched off'
Invoke-Arm 'F' 'Selection.TypeText charwise, settings untouched (the control)'
Invoke-Arm 'G' 'Selection.TypeText charwise, settings switched off'
Invoke-Arm 'H' 'positive control: can this bait be rewritten at all?'
Invoke-Arm 'C' 'is switching them off scoped to one instance?'

$leaked = @(Get-WordPids | Where-Object { $pidsBefore -notcontains $_ })
Write-Host ""
if ($leaked.Count -gt 0) {
    Write-Host "cleaning up WINWORD started by this probe: $($leaked -join ', ')"
    foreach ($p in $leaked) { try { Stop-Process -Id $p -Force -ErrorAction SilentlyContinue } catch { } }
} else {
    Write-Host "no WINWORD left behind"
}

Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
