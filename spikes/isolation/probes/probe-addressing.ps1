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
# ---------------------------------------------------------------------------
# TEARDOWN. This probe ends only the Word instance it created, and it ends it
# through the COM handle it holds rather than by pid -- it kills nothing. An
# instance that turns out to predate the probe is released rather than quit.
#
# It used to end its `finally` with an unfiltered
# `Get-Process WINWORD | Stop-Process -Force`, which killed every Word on the
# machine including ones it never started -- another session's, or your own open
# document. That is fixed (#114). The measured arms below are untouched, so the
# figures this file backs still stand.
# ---------------------------------------------------------------------------
#
# WHAT THIS PROBE DOES AND DOES NOT BACK. S1 measures ONE strategy: the
# per-paragraph walk, touching `Range.Text` and `OutlineLevel` on every
# paragraph. It has no `Content.WordOpenXML` arm and no `Content.Text` arm, so
# it backs the cost of that walk and NOT the 289 ms WordOpenXML figure or the
# 6 ms `Content.Text` figure the walk was compared against. Those came from a
# bulk-read comparison that is no longer committed. Cite this file for the
# walk; do not cite it for the alternatives.
#
# Retired with the rest of the streaming archive (#103) and restored in the same
# PR: it is the only committed instrument behind the S1 walk cost quoted in
# `word-host.ps1` and the S3 revision-token behaviour quoted in
# `revision-token.mjs`. Both now cite this file directly, because the chain that
# used to reach it ran through a document rather than a citation and so was
# invisible to `tools/check-citations.mjs`.

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

# --- ownership: attribute by window handle, never by differencing (#114) ----
# `Add-Type` is hoisted here, outside the `try`, so its one-time ~800 ms compile
# can never land between a warm-up and a stopwatch.
Add-Type -Namespace Win32 -Name Wnd -MemberDefinition @'
[DllImport("user32.dll", SetLastError = true)]
public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint lpdwProcessId);
'@

function Get-WordPids { @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id) }

# Which process owns the window belonging to THIS RCW. Nothing is differenced
# and nothing is guessed: every rejection returns $null, and the caller then
# refuses to name a pid at all rather than picking a plausible one. Differencing
# the WINWORD set around `New-Object` is measured unsound here -- with one
# concurrent Word from a separate process it reported 2 new pids for the 1
# instance created, and agreed only by luck
# (probe-init-attribution.ps1:250-251, test/integration/word-pids.mjs:29-38).
function Get-AttributedWordPid($app) {
    try {
        $hwnd = [IntPtr][int64]$app.ActiveWindow.Hwnd
        if ($hwnd -eq [IntPtr]::Zero) { return $null }
        [uint32]$procId = 0
        [void][Win32.Wnd]::GetWindowThreadProcessId($hwnd, [ref]$procId)
        $candidate = [int]$procId
        if ($candidate -le 4) { return $null }
        $p = Get-Process -Id $candidate -ErrorAction SilentlyContinue
        if ($null -eq $p -or $p.ProcessName -ne 'WINWORD') { return $null }
        return $candidate
    } catch { return $null }
}

function Get-WordStartTime([int]$candidate) {
    try {
        $p = Get-Process -Id $candidate -ErrorAction SilentlyContinue
        if ($null -eq $p -or $p.ProcessName -ne 'WINWORD') { return $null }
        return $p.StartTime
    } catch { return $null }
}

# 'gone' | 'alive' | 'unknown'. The expected StartTime is recorded once at
# attribution and never overwritten, but it IS re-read and compared on every
# poll: a pid recycled onto another process would otherwise be reported as our
# Word surviving. Only a matching name AND start time is a survivor.
function Get-WordLiveness([int]$candidate, $expectedStart) {
    $p = Get-Process -Id $candidate -ErrorAction SilentlyContinue
    if ($null -eq $p) { return 'gone' }
    if ($p.ProcessName -ne 'WINWORD') { return 'gone' }
    $actual = try { $p.StartTime } catch { $null }
    if ($null -eq $actual) { return 'unknown' }
    if ($actual -ne $expectedStart) { return 'gone' }
    return 'alive'
}

$ownedPid = $null      # created here and attributed -- the only pid ever polled
$ownedStart = $null
$attachedPid = $null   # attributed but pre-existing: someone else's, never quit
$verdict = 0           # 0 ok | 1 our Word survived | 2 teardown unverified

try {
    # Census before New-Object, used ONLY as a negative: a pid that already
    # existed cannot be one we created. It never selects a pid.
    $before = Get-WordPids
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false; $word.DisplayAlerts = 0

    $d = $word.Documents.Open($doc, $false, $true)   # read-only for the survey

    # Attribution happens here: after the document exists (ActiveWindow needs
    # one) but AHEAD of the warm-up, so the warm-up remains the last operation
    # before the stopwatch and the measured arm is unperturbed. Reading early
    # also matters for correctness: a late ActiveWindow read can name a
    # Protected View sandbox WINWORD the caller never bound to
    # (probe-init-attribution.ps1:36-44).
    $attributed = Get-AttributedWordPid $word
    if ($null -eq $attributed) {
        "  ownership: could not attribute this instance to a pid."
    }
    elseif ($before -contains $attributed) {
        $attachedPid = $attributed
        "  ownership: attached to a pre-existing Word (pid $attachedPid) -- it will NOT be ended."
    }
    else {
        $ownedPid = $attributed
        $ownedStart = Get-WordStartTime $ownedPid
        "  ownership: this probe created pid $ownedPid."
    }

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
    # Teardown acts only on what this probe owns, and it kills nothing. `Quit`
    # is addressed to the RCW we created, so it can only ever end the instance
    # we are bound to; a kill is addressed to a coordinate and can land on a
    # Word we never started. Destroy by handle is safe where destroy by
    # coordinate is not (word-host.ps1:804-813).
    #
    # Every step below is individually wrapped: an error escaping this block
    # would replace a real failure from the probe body.
    if ($word) {
        if ($attachedPid) {
            "  pid $attachedPid predates this probe -- released, NOT quit."
        }
        else {
            # Quit() takes no argument. Under Windows PowerShell 5.1 -- this
            # file's runtime -- Quit(0) throws AND leaves the process alive
            # (probe-quit0-leak.ps1, test/unit/quit-argument.test.mjs). An
            # unattributed instance is still quit: the handle is ours even when
            # the pid is unknown, and not quitting is how a Word gets stranded.
            try { $word.Quit() } catch { "  Quit threw -- $($_.Exception.Message.Split([char]10)[0])" }
        }
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch { }
    }

    # Only an owned pid is polled. An attached instance is deliberately left
    # running, so polling it would spend the budget confirming the expected and
    # then report someone else's Word as a leak.
    if ($ownedPid) {
        # 90 s is a generous observation budget, not a measured boundary: a Word
        # has been seen to outlive a 30 s poll under concurrent load and then
        # exit on its own (test/integration/word-pids.mjs:130-137). Polling
        # makes the budget free on success.
        $deadline = (Get-Date).AddSeconds(90)
        $state = Get-WordLiveness $ownedPid $ownedStart
        while ($state -eq 'alive' -and (Get-Date) -lt $deadline) {
            Start-Sleep -Milliseconds 250
            $state = Get-WordLiveness $ownedPid $ownedStart
        }
        if ($state -eq 'gone') {
            "  our Word (pid $ownedPid) exited."
        }
        elseif ($state -eq 'unknown') {
            $verdict = 2
            "  our Word (pid $ownedPid) is still present but its identity could not be confirmed -- teardown UNVERIFIED."
        }
        else {
            $verdict = 1
            "  *** our Word (pid $ownedPid) is STILL RUNNING after 90 s. Nothing is killed here."
            "  *** End it by hand: Stop-Process -Id $ownedPid"
        }
    }
    elseif ($word -and -not $attachedPid) {
        # An instance existed and was quit through its handle, but we never
        # learned its pid, so we cannot confirm it went. That is not a clean
        # teardown and must not be reported as one.
        $verdict = 2
        "  teardown UNVERIFIED: this instance was quit through its own handle, but it was never attributed to a pid, so its exit could not be confirmed. Nothing is killed on a guess."
    }

    # Only this run's own directory. The wildcard $env:TEMP\addr-* sweep that
    # used to be here deleted concurrently-running siblings' working
    # directories, which is the same "act on what you own" defect in miniature.
    Remove-Item $root -Recurse -Force -ErrorAction SilentlyContinue
}

# Placed after the finally so cleanup always runs first: an exit that skipped
# cleanup would trade a reported failure for a leaked WINWORD
# (probe-fileshare-algebra.ps1:243-248). If the probe body threw, that exception
# propagates past this line and the original failure is preserved.
if ($verdict -ne 0) { exit $verdict }
