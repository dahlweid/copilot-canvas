# probe-authoring.ps1
#
# The question
# ------------
# `create_document` has to turn an intent spec into Word calls. Four things it
# needs are not knowable from the existing code, because nothing in this repo
# authors a document -- it only opens and edits ones that already exist.
#
#   1. Which numeric wdStyle* constants exist, and what are they? Heading levels
#      are `-1 - level`, which is already established. Lists are not. Values
#      remembered from documentation (wdStyleListBullet = -49, wdStyleListNumber
#      = -50) are *not* evidence, and a wrong negative constant either throws or
#      silently selects a different style. So this arm enumerates the range and
#      reports NameLocal for every id that resolves, instead of guessing.
#
#   2. Does ListFormat.ApplyBulletDefault() work on a range, and does it survive
#      a save/reopen? A list applied by direct formatting rather than by style
#      is invisible to a style-based reader.
#
#   3. What `styleId` actually lands in the saved file for each of these? That
#      is what `read_document` reports and what the round-trip acceptance
#      criterion compares against. Word mints ids from the *localized* style
#      name with non-ASCII dropped, so on this German Word "Überschrift 1"
#      becomes `berschrift1`. Ids are copied verbatim, never constructed and
#      never matched -- so the only way to know one is to read it back.
#
#   4. Tables: does Tables.Add plus per-cell text work, and what does a cell
#      paragraph's range look like? A cell paragraph's Range.Text ends with "`r"
#      plus chr(7) -- two characters in the string -- while End - Start counts
#      the end-of-cell mark as *one* position. Any trim derived from string
#      length is therefore wrong inside a table. This arm prints both numbers so
#      the discrepancy is on the record rather than in a comment.
#
# Harness
# -------
# Each arm is a real `powershell.exe`, not a Start-Job: probe-autocorrect.ps1
# lost two runs to `Document.SaveAs2` hanging indefinitely on a job runspace
# thread, while the identical body in a real process saves in ~150 ms. Values
# reach the worker as discrete argv elements, never interpolated into a command
# string.
#
# Cleanup kills only PIDs absent before the probe started -- several WINWORD.EXE
# belonging to other sessions are routinely alive here.
#
# Status of the individual arms
# -----------------------------
# Arm S answered question 1 and its table of constants stands. Arms R, L and T
# produced their diagnostics but then wedged at `Document.SaveAs2`, so they never
# reached the reopen that answers questions 2 and 3.
#
# That wedge is **not** a property of styles, lists or tables. It is an
# unexplained script-shape artefact: a second worker doing the same COM work
# saves in ~120 ms, and running the two alternately -- same directory, same file
# name, same live Word population -- alternates OK/HANG perfectly. Machine load,
# Word instance count, file name, paragraph text and stdout redirection were each
# mutated separately and none of them changed the outcome. The mechanism is still
# unidentified and is deliberately not guessed at here.
#
# `probe-authoring-save.ps1` answers questions 2, 3 and 4 on a harness that does
# save, and its control arm is one that can fail. Prefer it. These arms are kept
# because arm S and arm R's exception *types* are still the evidence for the
# style-constant table and for style-name assignment throwing.

param(
    [string[]] $Only
)

$ErrorActionPreference = 'Stop'

function Get-WordPids {
    @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }) | Sort-Object
}

$pidsBefore = Get-WordPids
Write-Host "WINWORD before: $($pidsBefore -join ', ')"

$root = Join-Path $env:TEMP ("auth-" + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $root | Out-Null
$worker = Join-Path $root 'worker.ps1'

@'
param(
    [Parameter(Mandatory = $true)][ValidateSet('S', 'L', 'T', 'R')][string] $Arm,
    [Parameter(Mandatory = $true)][string] $Root,
    [Parameter(Mandatory = $true)][string] $Trace
)

# No COM handle below shares a name with a parameter above: a param() type
# constraint follows the variable for the rest of the scope and PowerShell names
# are case-insensitive, so `$doc = ...Add()` under a `[string] $Doc` parameter
# silently coerces the Document to a string and the failure surfaces later, on a
# correct line, as "the property Text was not found on this object".

$ErrorActionPreference = 'Stop'
$sw = [Diagnostics.Stopwatch]::StartNew()
function Step([string] $what) {
    Add-Content -LiteralPath $Trace -Value ("{0,7}ms  {1}" -f $sw.ElapsedMilliseconds, $what)
}

$WD_DO_NOT_SAVE = 0
$WD_FORMAT_DOCX = 16   # wdFormatDocumentDefault
$WD_CHARACTER   = 1

function New-Word {
    $a = New-Object -ComObject Word.Application
    $a.Visible = $false
    $a.DisplayAlerts = 0
    return $a
}

# Trim derived from the *position span*, never from string length.
function Set-ParagraphText($para, [string] $text) {
    $r = $para.Range
    $visible = ([string]$r.Text).TrimEnd("`r", [char]7, "`n")
    $drop = ($r.End - $r.Start) - $visible.Length
    if ($drop -gt 0) { $r.MoveEnd($WD_CHARACTER, -$drop) | Out-Null }
    $r.Text = $text
}

$app = $null
try {
    switch ($Arm) {

        # --- S: which negative style constants actually exist? ---------------
        'S' {
            $app = New-Word
            $document = $app.Documents.Add()
            Step 'ready'
            'negative style ids that resolve on this Word:'
            for ($id = -1; $id -ge -80; $id--) {
                $name = $null
                try { $name = $document.Styles.Item($id).NameLocal } catch { }
                if ($null -ne $name) { "  {0,4} = {1}" -f $id, $name }
            }
            Step 'enumerated'
            $document.Close($WD_DO_NOT_SAVE)
        }

        # --- L: lists, by style constant and by ListFormat -------------------
        'L' {
            $app = New-Word
            $document = $app.Documents.Add()
            Step 'ready'

            # Applied by style id. The ids come from arm S, never from memory.
            $bulletId = -49   # candidate wdStyleListBullet, confirmed by arm S
            $numberId = -50   # candidate wdStyleListNumber, confirmed by arm S

            $lines = @('alpha', 'beta', 'gamma', 'one', 'two', 'three')
            $i = 1
            foreach ($t in $lines) {
                if ($i -gt 1) { $document.Paragraphs.Add() | Out-Null }
                Set-ParagraphText $document.Paragraphs.Item($i) $t
                $i++
            }
            Step 'wrote 6 paragraphs'

            for ($j = 1; $j -le 3; $j++) {
                $document.Paragraphs.Item($j).Range.Style = $bulletId
            }
            Step "applied style $bulletId to 1-3"
            for ($j = 4; $j -le 6; $j++) {
                $document.Paragraphs.Item($j).Range.Style = $numberId
            }
            Step "applied style $numberId to 4-6"

            $file = Join-Path $Root 'l.docx'
            $document.SaveAs2($file, $WD_FORMAT_DOCX)
            $document.Close($WD_DO_NOT_SAVE)
            Step 'saved'

            $check = $app.Documents.Open($file, $false, $true)
            'reopened -- what each paragraph reports:'
            for ($j = 1; $j -le $check.Paragraphs.Count; $j++) {
                $p = $check.Paragraphs.Item($j)
                $txt = ([string]$p.Range.Text).TrimEnd("`r", [char]7, "`n")
                $styleName = $p.Range.Style.NameLocal
                $listString = ''
                try { $listString = $p.Range.ListFormat.ListString } catch { }
                $listType = ''
                try { $listType = $p.Range.ListFormat.ListType } catch { }
                "  [{0}] '{1}' style='{2}' listString='{3}' listType={4}" -f $j, $txt, $styleName, $listString, $listType
            }
            $check.Close($WD_DO_NOT_SAVE)
            Step 'inspected'
            "saved file: $file"
        }

        # --- T: tables, and the chr(7) position/length discrepancy -----------
        'T' {
            $app = New-Word
            $document = $app.Documents.Add()
            Step 'ready'

            Set-ParagraphText $document.Paragraphs.Item(1) 'before the table'
            $document.Paragraphs.Add() | Out-Null
            $anchor = $document.Paragraphs.Item(2).Range
            $table = $document.Tables.Add($anchor, 2, 3)
            Step 'Tables.Add(2,3) returned'

            $rows = @(@('h1', 'h2', 'h3'), @('a', 'b', 'c'))
            for ($r = 1; $r -le 2; $r++) {
                for ($c = 1; $c -le 3; $c++) {
                    $cell = $table.Cell($r, $c)
                    $cr = $cell.Range
                    $visible = ([string]$cr.Text).TrimEnd("`r", [char]7, "`n")
                    $span = $cr.End - $cr.Start
                    if ($r -eq 1 -and $c -eq 1) {
                        "  empty cell: Range.Text length = $(([string]$cr.Text).Length), End-Start = $span"
                        "  empty cell: chars = $((([string]$cr.Text).ToCharArray() | ForEach-Object { [int]$_ }) -join ',')"
                    }
                    $drop = $span - $visible.Length
                    if ($drop -gt 0) { $cr.MoveEnd($WD_CHARACTER, -$drop) | Out-Null }
                    $cr.Text = $rows[$r - 1][$c - 1]
                }
            }
            Step 'filled cells'

            $file = Join-Path $Root 't.docx'
            $document.SaveAs2($file, $WD_FORMAT_DOCX)
            $document.Close($WD_DO_NOT_SAVE)
            Step 'saved'

            $check = $app.Documents.Open($file, $false, $true)
            "reopened -- tables: $($check.Tables.Count), paragraphs: $($check.Paragraphs.Count)"
            'every paragraph as read_document would walk them:'
            for ($j = 1; $j -le $check.Paragraphs.Count; $j++) {
                $p = $check.Paragraphs.Item($j)
                $raw = [string]$p.Range.Text
                $codes = ($raw.ToCharArray() | ForEach-Object { [int]$_ }) -join ','
                $span = $p.Range.End - $p.Range.Start
                "  [{0}] len={1} span={2} codes={3}" -f $j, $raw.Length, $span, $codes
            }
            $check.Close($WD_DO_NOT_SAVE)
            Step 'inspected'
            "saved file: $file"
        }

        # --- R: headings by numeric constant, and the styleId on disk --------
        'R' {
            $app = New-Word
            $document = $app.Documents.Add()
            Step 'ready'

            $blocks = @(
                @{ level = 1; text = 'Title one' }
                @{ level = 0; text = 'Body text under it.' }
                @{ level = 2; text = 'Sub two' }
                @{ level = 0; text = 'More body.' }
                @{ level = 3; text = 'Deep three' }
            )
            $i = 1
            foreach ($b in $blocks) {
                if ($i -gt 1) { $document.Paragraphs.Add() | Out-Null }
                $para = $document.Paragraphs.Item($i)
                Set-ParagraphText $para $b.text
                if ($b.level -ge 1) {
                    # wdStyleHeading1 = -2 ... wdStyleHeading9 = -10
                    $para.Range.Style = (-1 - $b.level)
                } else {
                    $para.Range.Style = -1   # wdStyleNormal
                }
                $i++
            }
            Step 'wrote and styled'

            # Does naming a style really throw here? Asserted repeatedly in this
            # repo; measured here so the claim has a probe behind it.
            foreach ($name in 'Heading 1', 'berschrift1', "$([char]0xDC)berschrift 1") {
                try {
                    $document.Paragraphs.Item(1).Range.Style = $name
                    "  Range.Style = '$name' -> ACCEPTED"
                } catch {
                    "  Range.Style = '$name' -> threw: $($_.Exception.GetType().Name)"
                }
            }
            # Put it back by numeric constant.
            $document.Paragraphs.Item(1).Range.Style = -2
            Step 'style-name assignment tested'

            $file = Join-Path $Root 'r.docx'
            $document.SaveAs2($file, $WD_FORMAT_DOCX)
            $document.Close($WD_DO_NOT_SAVE)
            Step 'saved'

            $check = $app.Documents.Open($file, $false, $true)
            'reopened -- NameLocal, OutlineLevel, and the id in the file:'
            for ($j = 1; $j -le $check.Paragraphs.Count; $j++) {
                $p = $check.Paragraphs.Item($j)
                $txt = ([string]$p.Range.Text).TrimEnd("`r", [char]7, "`n")
                "  [{0}] '{1}' NameLocal='{2}' OutlineLevel={3}" -f $j, $txt, $p.Range.Style.NameLocal, $p.OutlineLevel
            }
            $check.Close($WD_DO_NOT_SAVE)
            Step 'inspected'
            "saved file: $file"
        }
    }
} catch {
    Step "ERROR $($_.Exception.GetType().FullName): $($_.Exception.Message)"
    "ERROR: $($_.Exception.Message)"
} finally {
    if ($null -ne $app) {
        # Quit(), never Quit(<arg>). Under Windows PowerShell 5.1 -- the runtime
        # every .ps1 here runs under -- the argument form does not bind, so it
        # throws and the Word survives; process exit does not reap it either
        # (probe-quit0-leak.ps1). The no-argument form takes the same
        # default, since wdSaveChanges is only consulted for a dirty document.
        # Reporting, not swallowing: an empty catch is what turned this from a
        # hard failure into a leak nobody could see. Step writes the trace file
        # the parent prints, so this lands on a channel that is read.
        try { $app.Quit() } catch { Step "Quit() FAILED (Word may leak) -- $($_.Exception.Message.Split([char]10)[0])" }
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null } catch { }
    }
    Step 'quit'
}
'@ | Set-Content -LiteralPath $worker -Encoding UTF8

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

Invoke-Arm 'S' 'which negative style constants exist?'
Invoke-Arm 'R' 'headings by numeric constant, and style names on disk'
Invoke-Arm 'L' 'lists by style constant'
Invoke-Arm 'T' 'tables, and the chr(7) position/length discrepancy'

$leaked = @(Get-WordPids | Where-Object { $pidsBefore -notcontains $_ })
Write-Host ""
if ($leaked.Count -gt 0) {
    Write-Host "cleaning up WINWORD started by this probe: $($leaked -join ', ')"
    foreach ($p in $leaked) { try { Stop-Process -Id $p -Force -ErrorAction SilentlyContinue } catch { } }
} else {
    Write-Host "no WINWORD left behind"
}

# The saved files are left in place only if -Only was used, so a single arm can
# be re-examined by hand; a full run cleans up after itself.
if (-not $Only) { Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue }
else { Write-Host "workdir kept: $root" }
