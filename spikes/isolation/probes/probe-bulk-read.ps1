# Probe: the 4144 ms structural read is a per-property COM round-trip problem.
# Each $p.Range.Text / $p.OutlineLevel is a cross-process call. Test whether
# bulk extraction removes it.
#
#   A. naive     - iterate Paragraphs, touch properties per paragraph
#   B. bulk text - one Content.Text call, split on CR
#   C. bulk both - one Content.Text call + one pass for outline levels
#   D. XML       - single WordOpenXML call, parsed outside Word

$ErrorActionPreference = 'Stop'
$src = $args[0]
$root = Join-Path $env:TEMP ("bulk-" + [guid]::NewGuid().ToString('N').Substring(0,8))
New-Item -ItemType Directory -Path $root | Out-Null
$doc = Join-Path $root 'original.docx'
Copy-Item $src $doc

$word = $null
$d = $null             # the document this probe opens; closed in `finally`
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
# instance created, and agreed only by luck -- see the "ambiguous condition was
# reached" branch of probe-init-attribution.ps1, which reports that `$new[0]` was
# ours "by luck", and the "Attribution is now sound at its source" note in
# test/integration/word-pids.mjs.
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
#
# The whole observation is wrapped, not just the StartTime read: a process can
# exit between `Get-Process` returning and a property being fetched, which
# throws rather than reporting absence. This runs inside `finally`, where an
# escaping error would displace a real failure from the probe body.
function Get-WordLiveness([int]$candidate, $expectedStart) {
    try {
        $p = Get-Process -Id $candidate -ErrorAction SilentlyContinue
        if ($null -eq $p) { return 'gone' }
        if ($p.ProcessName -ne 'WINWORD') { return 'gone' }
        # No recorded start time means identity cannot be tested at all. Falling
        # through would compare a real DateTime against $null, which is -ne, and
        # so report a RUNNING Word as 'gone' -- a clean teardown nobody observed.
        # Note the asymmetry this preserves: absence, checked above, is sound
        # without a start time, because no process holds the pid. Presence is
        # not, because the pid may have been recycled. Only the sound half is
        # allowed to conclude.
        if ($null -eq $expectedStart) { return 'unknown' }
        $actual = $p.StartTime
        if ($null -eq $actual) { return 'unknown' }
        if ($actual -ne $expectedStart) { return 'gone' }
        return 'alive'
    } catch { return 'unknown' }
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
    $d = $word.Documents.Open($doc, $false, $true)

    # Attribution happens here: after the document exists (ActiveWindow needs
    # one) but AHEAD of the warm-up, so the warm-up remains the last operation
    # before the stopwatch and arm A is unperturbed. Reading early also matters
    # for correctness: a late ActiveWindow read can name a Protected View
    # sandbox WINWORD the caller never bound to -- the "What this deliberately
    # does NOT measure" note in probe-init-attribution.ps1 records why the read
    # is early and what moving it would cost.
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
        if ($null -eq $ownedStart) {
            "  ownership: its start time could not be read, so its exit can be confirmed only by the pid going absent."
        }
    }

    $null = $d.Content.Text   # warm

    "== A: naive per-paragraph property access =="
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $n = 0
    foreach ($p in $d.Paragraphs) { $null = $p.Range.Text; $null = $p.OutlineLevel; $n++ }
    $sw.Stop()
    "  $n paragraphs, $($sw.ElapsedMilliseconds) ms"

    "== B: one Content.Text call, split locally =="
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $text = $d.Content.Text
    $lines = $text -split "`r"
    $sw.Stop()
    "  $($lines.Count) lines, $($sw.ElapsedMilliseconds) ms"

    "== C: bulk text + outline levels via ListParagraphs-free single pass =="
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $text = $d.Content.Text
    $lines = $text -split "`r"
    # outline level only for paragraphs that are headings: ask Word once per heading
    $levels = @{}
    $hdrs = $d.Paragraphs
    $sw.Stop()
    "  (text) $($sw.ElapsedMilliseconds) ms - outline still needs a per-para call, see D"

    "== D: single WordOpenXML call, parsed outside Word =="
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $xml = $d.Content.WordOpenXML
    $sw.Stop()
    $xmlMs = $sw.ElapsedMilliseconds
    "  fetched $([math]::Round($xml.Length/1024)) KB of WordprocessingML in $xmlMs ms"

    $sw = [Diagnostics.Stopwatch]::StartNew()
    $x = [xml]$xml
    $ns = New-Object Xml.XmlNamespaceManager($x.NameTable)
    $ns.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
    $nodes = $x.SelectNodes('//w:body/w:p', $ns)
    $parsed = foreach ($node in $nodes) {
        $style = $node.SelectSingleNode('./w:pPr/w:pStyle/@w:val', $ns)
        $texts = $node.SelectNodes('.//w:t', $ns) | ForEach-Object { $_.InnerText }
        [pscustomobject]@{ Style = $(if ($style) { $style.Value } else { '' }); Text = ($texts -join '') }
    }
    $sw.Stop()
    "  parsed $($parsed.Count) paragraphs in $($sw.ElapsedMilliseconds) ms"
    "  TOTAL for D: $($xmlMs + $sw.ElapsedMilliseconds) ms"
    $styled = @($parsed | Where-Object { $_.Style -ne '' }).Count
    "  paragraphs carrying an explicit style: $styled"
    $h = @($parsed | Where-Object { $_.Style -like 'berschrift*' -or $_.Style -like 'Heading*' }) 
    "  heading-styled paragraphs: $($h.Count)"
    if ($h.Count) { "  first heading style name: '$($h[0].Style)'  text: '$($h[0].Text.Substring(0,[Math]::Min(40,$h[0].Text.Length)))'" }

    $d.Close(0)
}
finally {
    # Teardown kills nothing, and acts through handles rather than pids. `Quit`
    # is addressed to the RCW we created, so it can only ever end the instance
    # we are bound to; a kill is addressed to a coordinate and can land on a
    # Word we never started. Destroy by handle is safe where destroy by
    # coordinate is not -- the same argument the `Quit` gate in `Stop-Word`
    # (word-host.ps1) makes, in those words.
    #
    # Every step below is individually wrapped: an error escaping this block
    # would replace a real failure from the probe body.
    #
    # Close this probe's document first. On the attached path nothing else
    # would: `Quit` is skipped there, so without this our temp document is left
    # open in someone else's Word -- and the directory holding it is removed a
    # few lines below, leaving them a document backed by a deleted file. Close(0)
    # rather than Close(): argument-less Close prompts on a dirty document, and a
    # prompt on a hidden instance is a hang. Already closed is the normal case
    # and throws, so the result is ignored.
    if ($d) {
        try { $d.Close(0) } catch { }
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($d) | Out-Null } catch { }
    }

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
        # exit on its own -- the note on `assertNoLeakedWord`'s 90 s deadline in
        # test/integration/word-pids.mjs records that run. Polling
        # makes the budget free on success.
        $deadline = (Get-Date).AddSeconds(90)
        # Poll until the pid is absent, not merely until it is identified. An
        # 'unknown' keeps watching rather than concluding: it may still resolve
        # to a sound 'gone', and a transient failure to read a start time no
        # longer ends the observation early. The deadline still bounds it.
        $state = Get-WordLiveness $ownedPid $ownedStart
        while ($state -ne 'gone' -and (Get-Date) -lt $deadline) {
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
            "  *** our Word (pid $ownedPid, started $ownedStart) is STILL RUNNING after 90 s."
            "  *** Nothing is killed here, and no kill command is offered on purpose:"
            "  *** a pid is a coordinate, and by the time anyone reads this it may"
            "  *** belong to a different process. End it by hand only after confirming"
            "  *** that pid $ownedPid is still a WINWORD started at $ownedStart."
        }
    }
    elseif ($word -and -not $attachedPid) {
        # An instance existed and was quit through its handle, but we never
        # learned its pid, so we cannot confirm it went. That is not a clean
        # teardown and must not be reported as one.
        $verdict = 2
        "  teardown UNVERIFIED: this instance was quit through its own handle, but it was never attributed to a pid, so its exit could not be confirmed. Nothing is killed on a guess."
    }

    # Only this run's own directory. The wildcard $env:TEMP\bulk-* sweep that
    # used to be here deleted concurrently-running siblings' working
    # directories, which is the same "act on what you own" defect in miniature.
    Remove-Item $root -Recurse -Force -ErrorAction SilentlyContinue
}

# Placed after the finally so cleanup always runs first: an exit that skipped
# cleanup would trade a reported failure for a leaked WINWORD -- probe-fileshare-algebra.ps1
# places its own exit gate after its `finally` for that reason, "so PART B's Word
# is always reaped first". If the probe body threw, that exception
# propagates past this line and the original failure is preserved.
if ($verdict -ne 0) { exit $verdict }
