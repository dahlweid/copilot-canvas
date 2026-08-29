# word-host.ps1
#
# Long-lived PowerShell process that owns a hidden Microsoft Word instance and
# acts as the rendering/query engine for the office-canvas extension.
#
# Protocol: newline-delimited JSON over stdio.
#   in  -> {"id":1,"cmd":"open","args":{...}}
#   out -> {"id":1,"ok":true,"result":{...}}
#       -> {"id":1,"ok":false,"error":{"code":"...","message":"..."}}
#
# Hard rules (see plan):
#   * Numeric wd*/mso* constants only. Localized Word UIs do not have
#     "Heading 1" and string style names throw.
#   * Macros are force-disabled before any document is opened.
#   * Reads only ever open an unblocked temp copy, never the user's original.
#     Edits are the deliberate exception: they must touch the original, so they
#     open it directly, one operation at a time, and close it before returning
#     (ADR 0005). This line used to say the original is *never* opened, which
#     `edit_document` made false -- a safety rule stating something the file's
#     own code does not do is worse than no rule, because it is the one a
#     reader trusts without checking.
#   * A Word instance we did not create is never quit and never hidden.
#   * Global Application.Options are never modified. The reason is *not* that
#     they persist for the user: measured, they are per-process -- toggled off
#     in one instance, a fresh second process still read the original values,
#     and the HKCU Word\Options key stayed absent throughout, including after
#     both instances quit. So our hidden Word cannot reach the user's settings
#     by construction, and this rule is belt-and-braces rather than the thing
#     standing between them and a changed Word. Keep the rule; it costs nothing
#     and the isolation is a measured property of Word, not a guarantee we own.

param(
    # Directory used to record which WINWORD process this host owns, so an
    # instance orphaned by a crash can be reaped by the next host that starts.
    [string]$PidDir
)

$ErrorActionPreference = 'Stop'

[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)

# --- Word / Office constants -------------------------------------------------
$WD_ALERTS_NONE = 0
$MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3
$WD_DO_NOT_SAVE_CHANGES = 0
$WD_EXPORT_FORMAT_PDF = 17
$WD_EXPORT_ALL_DOCUMENT = 0
$WD_EXPORT_FROM_TO = 3
$WD_EXPORT_OPTIMIZE_FOR_PRINT = 0
$WD_EXPORT_DOCUMENT_CONTENT = 0
$WD_EXPORT_CREATE_HEADING_BOOKMARKS = 1
$WD_INFO_ACTIVE_END_PAGE_NUMBER = 3
$WD_INFO_NUMBER_OF_PAGES = 4
$WD_STATISTIC_WORDS = 0
$WD_STATISTIC_PAGES = 2
$WD_FIND_STOP = 0
$WD_GOTO_PAGE = 1
$WD_GOTO_ABSOLUTE = 1
$WD_OUTLINE_BODY_TEXT = 10

# wdStyleHeading1 = -2 ... wdStyleHeading9 = -10
function Get-HeadingStyleId([int]$level) { return -1 - $level }

# A password that will never match, so an encrypted document fails fast with an
# error instead of blocking automation on a modal password dialog.
$FAIL_FAST_PASSWORD = '#word-canvas-no-password#'

$script:App = $null
$script:OwnedPid = $null
# The start time of the Word at $script:OwnedPid, captured when ownership is
# learned. A pid identifies a process only as long as that process lives; this
# is what lets a later kill prove it is acting on the same one.
$script:OwnedStart = $null
$script:Docs = @{}
# Remembers how each document was opened so a dead COM connection can be
# recovered without the caller noticing.
$script:DocArgs = @{}

# --- plumbing ----------------------------------------------------------------

function Send-Message($obj) {
    $json = $obj | ConvertTo-Json -Compress -Depth 12
    [Console]::Out.WriteLine($json)
    [Console]::Out.Flush()
}

function Send-Ok($id, $result) {
    Send-Message @{ id = $id; ok = $true; result = $result }
}

function Send-Fail($id, $code, $message, $data = $null) {
    $err = @{ code = $code; message = $message }
    if ($null -ne $data) { $err.data = $data }
    Send-Message @{ id = $id; ok = $false; error = $err }
}

# A typed failure.
#
# Every failure used to reach the caller as `word_error`, which meant the only
# way to tell "the file is missing" from "the file is locked" from "Word itself
# died" was to pattern-match `$_.Exception.Message` -- and those messages are
# localized, so matching them on this German machine is matching a translation.
# `throw` records the thrown object on `$_.TargetObject`, so the dispatch loop
# can recover the code without parsing anything.
#
# `data` carries facts the caller needs *on the failure path*, which is where
# they matter most: `writable` is the obvious one, since "the original is locked"
# is precisely the case where a read fails and the caller wants to know why.
function New-HostError([string]$code, [string]$message, $data = $null) {
    return @{ __hostError = $true; code = $code; message = $message; data = $data }
}

function Get-WordPids {
    return @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
}

# --- owned-process registry --------------------------------------------------
#
# Each host records the WINWORD process it created in $PidDir\<hostPid>.pid.
# A host that dies without running its teardown leaves both the file and the
# Word process behind; the next host to start finds the file, sees that the
# recording host is gone, and ends the orphan. Without this, a crash leaks a
# hidden Word that nothing will ever clean up -- and worse, a later CreateObject
# can attach to that orphan, at which point ownership detection sees no new
# process and (correctly, but unhelpfully) refuses to quit it.

function Get-PidFilePath {
    if ([string]::IsNullOrWhiteSpace($PidDir)) { return $null }
    return (Join-Path $PidDir "$PID.pid")
}

# A pid is a handle, valid at a moment -- not an identity. Windows recycles pids,
# so `is there a WINWORD at this pid` and `is this the WINWORD we started` are
# different questions, and only the second one licenses a Kill. StartTime is the
# cheapest stable discriminator available: measured, it is readable on our own
# instance and identical across separate Get-Process calls.
#
# Returns $null when the start time cannot be read. Callers must treat that as
# "ownership unproven" and refuse to kill -- never as "probably ours". Note that
# a non-elevated probe on this machine could NOT produce an unreadable
# StartTime, even against SYSTEM-owned pids 0 and 4, so this branch is written
# from the documented failure rather than from an observed one. That is
# precisely why it must not be a silent catch: an instrument that has never
# produced a positive cannot tell us the branch is dead.
function Get-WordStartTime([int]$candidate) {
    try {
        $p = Get-Process -Id $candidate -ErrorAction SilentlyContinue
        if ($null -eq $p -or $p.ProcessName -ne 'WINWORD') { return $null }
        return $p.StartTime
    } catch { return $null }
}

# Verifies identity and kills through ONE pinned handle, and is the only
# sanctioned way to kill a Word in this host.
#
# Checking identity and then killing is not enough even when both go through a
# single `Process` object, and that is measured rather than reasoned. On
# Windows PowerShell 5.1, with a process allowed to exit between the two calls:
#
#   one Process object, .Handle never touched -> Kill() throws "Zugriff
#     verweigert" (access denied), i.e. it re-opened the pid and found
#     something it may not terminate
#   one Process object, .Handle touched first -> Kill() throws "the process
#     (NNNN) has exited", i.e. it consulted the handle it already held
#
# Two different failures from the same sequence is the discriminator: without
# the pin, `Kill()` performs its own OpenProcess by pid, so the pid can be
# recycled between the StartTime read and the terminate and the two land on
# different processes. Touching `.Handle` caches an open handle on the object;
# Windows will not reuse a pid while a handle to it is open, so from that point
# the identity read and the kill are provably about the same process.
#
# `StartTime` routing through the same cached handle is inferred rather than
# measured -- both go through .NET's GetProcessHandle -- but the ordering makes
# that immaterial: the pin is taken before the read, so anything the read sees
# is what the kill will terminate.
#
# Returns one of: 'killed', 'gone' (nothing of ours is at this pid -- either no
# process, or one that is not Word), or 'declined:<reason>' (a WINWORD is there
# and we could not prove it is ours). Callers must report a decline rather than
# swallow it: it means a leaked Word, a pid collision, or both.
function Stop-VerifiedWord([int]$candidate, $expectedStart) {
    $p = Get-Process -Id $candidate -ErrorAction SilentlyContinue
    if ($null -eq $p) { return 'gone' }
    # Pin first. Every check below is only meaningful once the pid is held.
    try { $null = $p.Handle } catch { return 'declined:the process handle could not be opened, so the pid cannot be pinned' }
    if ($p.ProcessName -ne 'WINWORD') { return 'gone' }
    if ($null -eq $expectedStart) { return 'declined:no start time was ever recorded for this pid' }
    $actual = try { $p.StartTime } catch { $null }
    if ($null -eq $actual) { return 'declined:the start time could not be read, so ownership is unproven' }
    if ($actual -ne $expectedStart) { return "declined:start time $actual does not match the recorded $expectedStart" }
    try { $p.Kill() } catch { return "declined:the terminate failed ($($_.Exception.Message.Split([char]10)[0]))" }
    return 'killed'
}

# stdout is the JSON-RPC channel and writing to it corrupts the protocol, so
# diagnostics go to stderr. That stderr reaches the extension's log because
# word-host.mjs line-splits it and calls `log` as it arrives -- which is a
# statement about the far side of a boundary, so it is measured
# (spikes/isolation/probes/probe-decline-diagnostic-reach.mjs) and covered by
# the orphan-sweep test, not asserted from this side. It was false when first
# written: stderr went into a ring buffer read only on exit, and a refusal to
# kill -- emitted during the startup sweep -- was discarded at a clean dispose.
function Write-HostDiagnostic([string]$text) {
    try { [Console]::Error.WriteLine($text) } catch { }
}

function Clear-OrphanedWord {
    if ([string]::IsNullOrWhiteSpace($PidDir)) { return }
    try {
        if (-not (Test-Path -LiteralPath $PidDir)) {
            New-Item -ItemType Directory -Force -Path $PidDir | Out-Null
            return
        }
        foreach ($file in Get-ChildItem -LiteralPath $PidDir -Filter '*.pid' -ErrorAction SilentlyContinue) {
            $hostPid = 0
            if (-not [int]::TryParse([IO.Path]::GetFileNameWithoutExtension($file.Name), [ref]$hostPid)) { continue }
            if ($hostPid -eq $PID) { continue }
            if ($null -ne (Get-Process -Id $hostPid -ErrorAction SilentlyContinue)) { continue }  # still running
            # The ledger records the start time alongside the pid because this
            # reap is the most exposed kill in the host: the file outlives a
            # crashed host by an unbounded interval, so `there is a WINWORD at
            # this pid` says even less here than it does in Stop-Word's 10 s
            # window. A legacy single-field file has no recorded identity, so it
            # is unproven by construction and the Word is reported, not killed.
            $recorded = (Get-Content -LiteralPath $file.FullName -Raw -ErrorAction SilentlyContinue)
            $wordPid = 0
            $expectedStart = $null
            if ($null -ne $recorded) {
                $parts = $recorded.Trim() -split '\s+'
                if ($parts.Count -ge 1) { [void][int]::TryParse($parts[0], [ref]$wordPid) }
                $ticks = [long]0
                if ($parts.Count -ge 2 -and [long]::TryParse($parts[1], [ref]$ticks)) {
                    $expectedStart = [datetime]::new($ticks)
                }
            }
            if ($wordPid -gt 0) {
                $outcome = Stop-VerifiedWord $wordPid $expectedStart
                if ($outcome -notin @('killed', 'gone')) {
                    Write-HostDiagnostic "[word-host] refusing to reap pid ${wordPid} recorded by host ${hostPid}: $($outcome -replace '^declined:', ''). Not killed; it is either a leaked Word or an unrelated process that inherited the pid."
                }
            }
            Remove-Item -LiteralPath $file.FullName -Force -ErrorAction SilentlyContinue
        }
    } catch { }
}

function Register-OwnedWord([int]$wordPid, $startTime) {
    $file = Get-PidFilePath
    if ($null -eq $file) { return }
    try {
        if (-not (Test-Path -LiteralPath $PidDir)) { New-Item -ItemType Directory -Force -Path $PidDir | Out-Null }
        # pid alone is a coordinate; the start ticks are what make it an identity
        # a later reaper can verify. Written as one line, whitespace separated,
        # so a file from an older host still parses as "pid, identity unknown".
        $ticks = if ($null -ne $startTime) { $startTime.Ticks } else { '' }
        Set-Content -LiteralPath $file -Value "$wordPid $ticks".Trim() -Encoding ascii
    } catch { }
}

function Unregister-OwnedWord {
    $file = Get-PidFilePath
    if ($null -eq $file) { return }
    try { Remove-Item -LiteralPath $file -Force -ErrorAction SilentlyContinue } catch { }
}

# PowerShell cannot call the parameterized `Item` property on Office's
# DocumentProperties collection directly -- `$doc.BuiltInDocumentProperties.Item($name)`
# throws NullReferenceException. Late-bound InvokeMember is the working path.
function Get-DocProp($doc, [string]$name) {
    try {
        $props = $doc.BuiltInDocumentProperties
        $prop = [System.__ComObject].InvokeMember('Item', 'GetProperty', $null, $props, @($name))
        return [System.__ComObject].InvokeMember('Value', 'GetProperty', $null, $prop, @())
    } catch {
        return $null
    }
}

function Get-PageOf($range) {
    try { return [int]$range.Information($WD_INFO_ACTIVE_END_PAGE_NUMBER) } catch { return 0 }
}

# The Word-side twin of `normalizeText` in structure-map.mjs, so a paragraph's
# text read over COM can be compared with the same paragraph's text parsed out
# of the markup.
#
# The extra step here is stripping control characters: Range.Text carries
# Word's own in-band marks -- \r for the paragraph mark, \a (7) for an
# end-of-cell mark, \v (11) for a line break -- which have no counterpart in
# the XML the map was built from. Whitespace collapsing then matches, because
# Word splits runs invisibly and the map already normalizes for that.
function Get-NormalizedText([string]$value) {
    if ([string]::IsNullOrEmpty($value)) { return '' }

    # This must mirror paragraphText() in structure-map.mjs character for
    # character. Word renders layout marks as control characters in Range.Text;
    # the map renders the same marks from the markup. Any disagreement makes the
    # affected paragraph *permanently uneditable*: the pre-mutation text check
    # fails, and the failure is reported as "the file changed between the read
    # and the edit", which is false and sends the caller into an infinite
    # re-read loop, because re-reading mints the identical address.
    #
    # Measured end to end on a fixture carrying each mark. Before this mapping,
    # w:br and w:noBreakHyphen paragraphs failed every edit; w:tab and
    # w:softHyphen passed, the first by luck and the second because both sides
    # happen to drop it.
    #
    #   mark                     markup            JS gives   Word gives
    #   line/page/column break   w:br              "\n"->" "  \u000B \u000C \u000E
    #   non-breaking hyphen      w:noBreakHyphen   "-"        \u001E
    #   optional (soft) hyphen   w:softHyphen      dropped    \u001F
    #   tab                      w:tab             "\t"->" "  \u0009
    $clean = [regex]::Replace($value, '[\u000B\u000C\u000E]', ' ')
    $clean = [regex]::Replace($clean, '\u001E', '-')
    $clean = [regex]::Replace($clean, '[\u0000-\u0008\u000F-\u001F]', '')
    return ([regex]::Replace($clean, '\s+', ' ')).Trim()
}

# Whether the file could be opened for writing right now -- 4 ms when another
# process holds it, 9 ms when free, correct in both directions.
#
# ADR 0005: this is the *only* safe way to ask. Probing by asking Word to open
# the document is precisely the call that hangs indefinitely on a held file,
# with DisplayAlerts already off, leaving two processes to be killed by hand.
# True when a *write* handle can be taken. Note what that does and does not
# mean, because the answer is reported to callers: it is false for a sharing
# violation, but equally for an ACL that denies write and for the read-only
# attribute. Measured: the read-only attribute and Word's own lock -- a handle
# with *write access* granting FileShare::Read -- both still allow *reading*
# and copying, so `writable = $false` must never be reported as "another
# process has it open" -- that is one of three causes. Note also that this
# function grants FileShare::None *and* requests write access, so it conflicts
# with any existing handle on both of Windows' checks -- whatever that handle
# shares, and whatever access it holds: a document open in Word reports
# writable = $false, correctly, while still being perfectly readable.
function Test-FileWritable([string]$path) {
    try {
        $stream = [IO.File]::Open($path, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        $stream.Close()
        $stream.Dispose()
        return $true
    } catch {
        return $false
    }
}

# ComputeStatistics(wdStatisticPages) returns 1 on a hidden, window-less Word
# instance because the document is never laid out for the screen. Information()
# reports the real paginated count, so it is the primary source here.
function Get-PageCount($doc) {
    try {
        $n = [int]$doc.Content.Information($WD_INFO_NUMBER_OF_PAGES)
        if ($n -gt 0) { return $n }
    } catch { }
    try { return [int]$doc.ComputeStatistics($WD_STATISTIC_PAGES) } catch { return 0 }
}

function Get-WordCount($doc) {
    try { return [int]$doc.ComputeStatistics($WD_STATISTIC_WORDS) } catch { return 0 }
}

function Test-AppAlive {
    if ($null -eq $script:App) { return $false }
    try {
        # PowerShell does not reliably throw when reading a property off a
        # disconnected COM proxy -- it can hand back $null instead. Probing the
        # *value* is therefore the only trustworthy liveness check.
        $version = $script:App.Version
        return -not [string]::IsNullOrEmpty([string]$version)
    } catch {
        return $false
    }
}

function Test-DocAlive($doc) {
    if ($null -eq $doc) { return $false }
    try {
        $name = $doc.Name
        return -not [string]::IsNullOrEmpty([string]$name)
    } catch {
        return $false
    }
}

function Resolve-Doc($docId) {
    # Re-open transparently if Word died underneath us (crash, stuck dialog that
    # was killed, user ending the process). The caller should not have to care.
    if ($script:Docs.ContainsKey($docId)) {
        $doc = $script:Docs[$docId]
        if (Test-DocAlive $doc) { return $doc }
        $script:Docs.Remove($docId) | Out-Null
    }
    if ($script:DocArgs.ContainsKey($docId)) {
        $saved = $script:DocArgs[$docId]
        Open-DocInternal $docId $saved.path $saved.workDir | Out-Null
        return $script:Docs[$docId]
    }
    throw (New-HostError 'no_such_document' "No open document with id '$docId'.")
}

# --- Word lifecycle ----------------------------------------------------------

function Initialize-Word {
    if (Test-AppAlive) { return }

    if ($null -ne $script:App) {
        # The instance died underneath us. Drop every stale handle before
        # creating a replacement, or later calls fail with RPC_E_DISCONNECTED.
        $script:App = $null
        $script:OwnedPid = $null
        $script:OwnedStart = $null
        $script:Docs = @{}
    }

    $before = Get-WordPids
    try {
        $script:App = New-Object -ComObject Word.Application
    } catch {
        throw (New-HostError 'word_unavailable' `
            "Microsoft Word could not be started. Word must be installed to use this canvas. ($($_.Exception.Message))")
    }

    # Ownership detection: if a brand new WINWORD process appeared we created it
    # and may hide and later quit it. If we merely attached to a Word the user
    # already had running, leave its visibility alone and never quit it out from
    # under them.
    #
    # The process does not always show up in the process list by the time
    # CreateObject returns -- especially right after a previous instance was
    # killed -- so poll briefly rather than taking a single snapshot. Getting
    # this wrong leaks a Word process, so the retry is worth the latency.
    $new = @()
    for ($attempt = 0; $attempt -lt 15; $attempt++) {
        $new = @(Get-WordPids | Where-Object { $before -notcontains $_ })
        if ($new.Count -ge 1) { break }
        Start-Sleep -Milliseconds 100
    }
    if ($new.Count -ge 1) {
        $script:OwnedPid = [int]$new[0]
        # Captured here, at the moment ownership is learned, and never re-read at
        # kill time: reading it just before the Kill would compare an impostor
        # against itself and always match.
        $script:OwnedStart = Get-WordStartTime $script:OwnedPid
        Register-OwnedWord $script:OwnedPid $script:OwnedStart
    } else {
        $script:OwnedPid = $null
        $script:OwnedStart = $null
    }

    if ($null -ne $script:OwnedPid) {
        try { $script:App.Visible = $false } catch { }
    }
    try { $script:App.DisplayAlerts = $WD_ALERTS_NONE } catch { }
    # Force-disable macros before any document is opened. We render documents
    # the user may not have authored.
    try { $script:App.AutomationSecurity = $MSO_AUTOMATION_SECURITY_FORCE_DISABLE } catch { }
}

function Close-Doc($docId) {
    if (-not $script:Docs.ContainsKey($docId)) { return }
    $doc = $script:Docs[$docId]
    $script:Docs.Remove($docId)
    try { $doc.Close($WD_DO_NOT_SAVE_CHANGES) } catch { }
    try { [Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null } catch { }
}

function Stop-Word {
    foreach ($docId in @($script:Docs.Keys)) { Close-Doc $docId }
    $script:DocArgs = @{}
    if ($null -ne $script:App) {
        if ($null -ne $script:OwnedPid) {
            # No argument, deliberately. `Application.Quit` takes its parameters
            # as `VARIANT*`, and Windows PowerShell 5.1 -- which is what
            # `powershell.exe` gives, and what this host runs under -- refuses to
            # bind a by-value argument to one:
            #
            #   Argument "1" must be System.Management.Automation.PSReference.
            #   Use [ref].
            #
            # Measured, on both the literal `0` and a variable holding it: it
            # throws every time, and the process is still alive 21 s later. There
            # is no literal-versus-variable distinction, which is the trap --
            # under PowerShell 7.6.5 *none* of the three forms throw, so a
            # reduction run in a 7.x shell clears this and is worthless.
            # `Quit()` and `Quit([ref]$WD_DO_NOT_SAVE_CHANGES)` both bind under
            # 5.1 and exit within 3 s; the no-argument form takes the same
            # default, so it is the smaller of the two.
            #
            # Not to be generalised to `$doc.Close($WD_DO_NOT_SAVE_CHANGES)`
            # above: that takes a by-value `VARIANT` too, and was *measured* not
            # to throw under 5.1. Why the two differ is unknown and is not worth
            # a guess -- which is exactly why neither may be inferred from the
            # other. Do not "consistency-fix" `Close` to take no argument: with
            # none, Word prompts on a dirty document, and a modal prompt in a
            # hidden instance is the silent hang this host exists to avoid.
            try { $script:App.Quit() } catch { }
        }
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($script:App) | Out-Null } catch { }
        $script:App = $null
    }
    # The instance we own, if Quit did not take it. Until the call above was
    # fixed this was not a fallback but the entire teardown: the Quit threw into
    # its swallowing catch on every single run, so no instance this host ever
    # started got a graceful exit -- every one was killed here. The leak
    # assertion in edit-smoke stayed green throughout, rescued by the very line
    # whose comment called it redundant. A green teardown assertion is weak
    # evidence while the teardown swallows.
    #
    # The wait is a poll rather than the fixed 300 ms it used to be, and that is
    # not tidying: measured through the full suite, `Quit()` returns before Word
    # has gone, so with the old fixed wait the corrected call still ended in a
    # kill and the fix had *no observable effect at all*. Isolated in both
    # directions -- `Quit()` with the poll exits on its own on an idle machine
    # and this branch does not fire; the old by-value call with the same poll
    # still throws and still needs the kill. Both halves are load-bearing.
    #
    # That "does not fire" is an observation on an idle machine, not a
    # guarantee, and the distinction is #26's: exit latency is load-dependent,
    # so on a busy machine this branch fires and the graceful path is pre-empted.
    # Nothing here can promise otherwise -- see the ceiling below for why the
    # remedy is not simply a bigger number.
    #
    # The wait is bounded by elapsed time, not by an iteration count. Measured
    # by probe-quit-exit-gap.ps1, `Quit()` returns in 3-28 ms and the process
    # exits 3039-3702 ms after that, so the count-based version this replaced --
    # 30 x 100 ms, nominally 3000 ms -- budgeted less than the thing it waits for.
    #
    # It worked anyway, and the reason is worth stating because it is the actual
    # defect: the loop's real budget was 4014-4375 ms, and *none of that margin
    # was chosen*. Its control run separates the two sources -- `Start-Sleep
    # -Milliseconds 100` overshoots, so the sleeps alone come to ~3292 ms, and
    # `Get-Process` adds a further 24-46 ms per call depending on load. Both are
    # incidental costs of unrelated APIs, both are load-dependent, and nothing
    # asserts either. Had they narrowed -- or had Word been slower, which under
    # load it measurably is -- the kill below would have resumed firing and this
    # fix would have silently reverted to the behaviour it exists to correct,
    # with every test still green.
    #
    # Timing out here is not a failure -- it falls through to the kill -- so the
    # bound is deliberately generous rather than fitted to the numbers above.
    #
    # It is *not*, however, freely chosen, and that is the part a future reader
    # will otherwise re-derive. This function runs inside the `quit` JSON-RPC
    # command, and word-host.mjs's `dispose` sends that command under a 20 s
    # client timeout whose expiry kills the host process outright (`#send`,
    # "restarting host"). So the client's timeout is a hard ceiling on this
    # wait, and a bound above it does not buy Word more time to exit -- it
    # spends the difference getting the host killed mid-wait, destroying the
    # graceful exit that raising the number was meant to protect.
    #
    # Measured, probe-quit-rpc-ceiling.mjs, wait made unconditional so the bound
    # is what the quit costs:
    #
    #   bound 10000 ms (under the ceiling): quit returned cleanly at 10279 ms
    #   bound 30000 ms (over the ceiling):  client timed out at 20007 ms, host killed
    #
    # This is why #26's suggestion of 30 s -- reasonable from the exit
    # distribution alone, and matching what the test-side waiters use -- is not
    # portable into this function. word-pids.mjs can wait 90 s because nothing
    # is waiting on *it*.
    #
    # Argued from the spread rather than a mean, since the mean is the figure
    # that moved: across every run anyone has taken here, samples span 779 ms to
    # 3702 ms, and word-pids.mjs records a Word that survived a 30 s poll under
    # concurrent load. The tail is not bounded by anything we have measured, so
    # **no reachable bound is outside the distribution** -- 30 s would not be
    # either, even if the 20 s ceiling allowed it. That kills the idea that the
    # right number is one the distribution cannot cross, and with it the
    # temptation to nudge 10 s to 15 s, which would look measured and change
    # nothing.
    #
    # So this is a budget, not a guarantee, and the design has to be safe when
    # it expires rather than tuned so that it does not. It is: expiry falls
    # through to a kill that must first prove identity. 10 s is ~2.7x the
    # slowest exit yet observed and half the ceiling, leaving the rest of the
    # teardown room inside the same 20 s.
    #
    # The two numbers are coupled and nothing enforces the coupling, so the
    # 20 s in word-host.mjs carries the reciprocal note. Raising either alone
    # is a silent change to the other's meaning.
    if ($null -ne $script:OwnedPid) {
        try {
            $p = $null
            $waited = [Diagnostics.Stopwatch]::StartNew()
            while ($waited.ElapsedMilliseconds -lt 10000) {
                Start-Sleep -Milliseconds 100
                $p = Get-Process -Id $script:OwnedPid -ErrorAction SilentlyContinue
                if ($null -eq $p -or $p.ProcessName -ne 'WINWORD') { break }
            }
            # `a WINWORD exists at this pid` was an adequate guard while the gap
            # between looking and killing was the old incidental ~300 ms. This
            # branch now waits up to 10 s, so the same guard spans a window
            # thirty times longer, and a pid recycled onto another WINWORD in
            # that gap would be killed as ours -- taking a user's unsaved
            # documents with it. Widening the wait is what obliged the guard to
            # become an identity check rather than an existence check.
            #
            # The loop above is deliberately *not* the guard. It observes with a
            # bare Get-Process because an observation may be stale without harm;
            # only the terminate needs a pinned handle, and Stop-VerifiedWord
            # takes its own. Nothing read in the loop is carried into the kill.
            $outcome = Stop-VerifiedWord $script:OwnedPid $script:OwnedStart
            if ($outcome -notin @('killed', 'gone')) {
                # Refusing to kill is the safe half. Saying nothing is not:
                # this is either a Word we leaked or a pid collision, and both
                # are worth knowing. #33 replaced every bare `catch {}` in this
                # function for the same reason, and a silent `else` would
                # reintroduce it in a different shape.
                Write-HostDiagnostic "[word-host] declining to kill pid $($script:OwnedPid): $($outcome -replace '^declined:', '') (recorded start $($script:OwnedStart)). Either our Word exited and the pid was reused, or its identity could not be proved. Not killed."
            }
        } catch { }
        $script:OwnedPid = $null
        $script:OwnedStart = $null
    }
    Unregister-OwnedWord
}

# --- commands ----------------------------------------------------------------

function Cmd-Ping($a) {
    Initialize-Word
    return @{
        ready       = $true
        wordVersion = [string]$script:App.Version
        ownedPid    = $script:OwnedPid
        owned       = ($null -ne $script:OwnedPid)
    }
}

function Open-DocInternal([string]$docId, [string]$path, [string]$workDir, [bool]$withStats = $true) {
    # Validate before Initialize-Word. These checks are string and filesystem
    # work costing microseconds; Initialize-Word costs up to ~4.5 s cold.
    #
    # Honest scope, because the obvious claim for this ordering is wrong here:
    # it saves no Word startup today, and measurement said so. Word is started
    # by the bridge, not by this command -- word-host.mjs pings as soon as the
    # host process exists, so that `ownedPid` is known on the tools-only path
    # and the reap net is not inert. By the time any command is dispatched,
    # cold or warm, Word is already up; reordering inside this function cannot
    # reach that. And the shipping read path never gets here with a bad path at
    # all: DocumentReader.read() stats the file and throws first.
    #
    # It is kept because the ordering is correct on its own terms and free, and
    # because it is what makes the eager ping the *only* thing standing between
    # a doomed request and a Word startup, rather than one of two.
    if ([string]::IsNullOrWhiteSpace($path)) { throw (New-HostError 'invalid_request' "No document path supplied.") }
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw (New-HostError 'file_not_found' "File not found: $path")
    }

    if ([string]::IsNullOrWhiteSpace($workDir)) {
        throw (New-HostError 'invalid_request' "No working directory supplied.")
    }
    if (-not (Test-Path -LiteralPath $workDir)) {
        try {
            New-Item -ItemType Directory -Force -Path $workDir -ErrorAction Stop | Out-Null
        } catch {
            throw (New-HostError 'write_failed' "Could not create the working directory. ($($_.Exception.Message))")
        }
    }

    Initialize-Word

    # Never open the user's original file: copy it, strip the mark-of-the-web
    # (Protected View refuses automation) and open the copy read-only. This is
    # what makes "read-only" structurally true rather than merely intended.
    $ext = [IO.Path]::GetExtension($path)
    $copy = Join-Path $workDir ("source" + $ext)
    try {
        Copy-Item -LiteralPath $path -Destination $copy -Force -ErrorAction Stop
    } catch {
        # The copy is the first thing that touches the original, so a failure to
        # read it surfaces here rather than at Documents.Open -- which is the
        # good outcome, because that is the call that hangs.
        #
        # Branch on the exception *type*, never the message: messages are
        # localized, and matching them is the trap that has already cost this
        # project twice. Measured on this machine:
        #
        #   FileShare::None (exclusive)      -> System.IO.IOException
        #   ACL denying read                 -> System.UnauthorizedAccessException
        #   Word's own lock (write access,   -> copy succeeds
        #     granting FileShare::Read)
        #   read-only attribute              -> copy succeeds
        #
        # The last two are why a read works against a document the user has
        # open, and why "read-only" is not a failure at all. Copy-Item requests
        # read and grants ReadWrite, so it passes both of Windows' checks
        # against Word's handle: Word's Read share admits a reader, and
        # Copy-Item's ReadWrite share admits Word's writer. A reader granting
        # only FileShare::Read fails the second check and would error here on a
        # file that copies fine.
        $ex = $_.Exception
        $name = [IO.Path]::GetFileName($path)
        if ($ex -is [System.UnauthorizedAccessException]) {
            throw (New-HostError 'permission_denied' `
                "Not allowed to read $name. This is a permissions problem, not another program holding the file." `
                    @{ writable = $false })
        }
        if ($ex -is [System.IO.IOException]) {
            throw (New-HostError 'file_locked' `
                "Another process is holding $name open more strictly than Word does. A document merely open in Word can still be read." `
                    @{ writable = $false })
        }
        throw (New-HostError 'copy_failed' "Could not make a working copy of the document. ($($ex.Message))")
    }
    try { Unblock-File -LiteralPath $copy -ErrorAction SilentlyContinue } catch { }

    Close-Doc $docId

    try {
        # FileName, ConfirmConversions, ReadOnly, AddToRecentFiles,
        # PasswordDocument, PasswordTemplate, Revert
        $doc = $script:App.Documents.Open($copy, $false, $true, $false, $FAIL_FAST_PASSWORD, $FAIL_FAST_PASSWORD, $false)
    } catch {
        throw (New-HostError 'document_unreadable' `
            "Word could not open the document. It may be password-protected or corrupt. ($($_.Exception.Message))")
    }

    $script:Docs[$docId] = $doc
    $script:DocArgs[$docId] = @{ path = $path; workDir = $workDir }

    $item = Get-Item -LiteralPath $path
    # Pagination and word count are a full pass over the document. The renderer
    # needs them; a structure read does not, and pays for them in latency.
    $pageCount = 0
    $wordCount = 0
    if ($withStats) {
        $pageCount = (Get-PageCount $doc)
        $wordCount = (Get-WordCount $doc)
    }
    return @{
        docId       = $docId
        path        = $path
        name        = $item.Name
        pageCount   = $pageCount
        wordCount   = $wordCount
        sizeBytes   = [int64]$item.Length
        modifiedIso = $item.LastWriteTimeUtc.ToString('o')
        title       = [string](Get-DocProp $doc 'Title')
        author      = [string](Get-DocProp $doc 'Author')
        workingCopy = $copy
    }
}

function Cmd-Open($a) {
    return Open-DocInternal ([string]$a.docId) ([string]$a.path) ([string]$a.workDir)
}

# One `Content.WordOpenXML` call, written to a file, document closed at once.
#
# Measured (PLAN.md §18.1) on a 219-paragraph document: walking Paragraphs and
# touching Range.Text / OutlineLevel per paragraph cost 3724 ms, because every
# property touch is a cross-process COM call; one WordOpenXML call cost 289 ms
# and returned strictly more -- text, style and full markup. A per-paragraph
# property walk is a defect, not a slow path.
#
# The markup goes to a file rather than back through the protocol: it is
# routinely 100 KB and can be megabytes, and JSON-escaping that onto a single
# stdio line is pure overhead when the reader is on the same machine.
function Cmd-Structure($a) {
    $docId = [string]$a.docId
    $path = [string]$a.path
    $out = [string]$a.out
    if ([string]::IsNullOrWhiteSpace($out)) {
        throw (New-HostError 'invalid_request' "No output path supplied for the structure read.")
    }
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw (New-HostError 'file_not_found' "File not found: $path")
    }

    # Reported, not enforced: a read is served from a copy and so never needs
    # the original. It tells the caller whether an *edit* would collide -- and it
    # is attached to failures too, because a failed read is exactly when the
    # caller most needs to know the original is held by someone else.
    $writable = Test-FileWritable $path

    try {
        $meta = Open-DocInternal $docId $path ([string]$a.workDir) $false
    } catch {
        $thrown = $_.TargetObject
        if ($thrown -is [hashtable] -and $thrown.__hostError) {
            if ($null -eq $thrown.data) { $thrown.data = @{} }
            # Only fill it in when the failure did not already say so. A deeper
            # error knows more than this pre-flight does: it observed the actual
            # open, whereas $writable was sampled before the attempt and the
            # file can change hands in between. Overwriting it discarded the
            # better answer in favour of the staler one.
            if (-not $thrown.data.ContainsKey('writable')) {
                $thrown.data.writable = $writable
            }
        }
        throw
    }
    try {
        $xml = [string](Resolve-Doc $docId).Content.WordOpenXML
    } finally {
        # The lock window ends the moment the markup is in hand. Nothing below
        # this point needs Word.
        $script:DocArgs.Remove($docId) | Out-Null
        Close-Doc $docId
    }

    $outDir = Split-Path -Parent $out
    if (-not [string]::IsNullOrWhiteSpace($outDir) -and -not (Test-Path -LiteralPath $outDir)) {
        New-Item -ItemType Directory -Force -Path $outDir | Out-Null
    }
    # No BOM: the caller parses this as XML, and a BOM ahead of the declaration
    # is a parse error in stricter readers.
    try {
        [IO.File]::WriteAllText($out, $xml, (New-Object System.Text.UTF8Encoding($false)))
    } catch {
        # A short write here is what produces truncated markup downstream, so it
        # gets its own code rather than arriving as a parse failure later.
        throw (New-HostError 'write_failed' `
            "Could not write the document markup to $out. ($($_.Exception.Message))" @{ writable = $writable })
    }

    return @{
        out         = $out
        bytes       = [int64](Get-Item -LiteralPath $out).Length
        writable    = $writable
        name        = $meta.name
        sizeBytes   = $meta.sizeBytes
        modifiedIso = $meta.modifiedIso
        title       = $meta.title
        author      = $meta.author
    }
}

function Cmd-Export($a) {
    $doc = Resolve-Doc ([string]$a.docId)
    $out = [string]$a.out
    $outDir = Split-Path -Parent $out
    if (-not (Test-Path -LiteralPath $outDir)) {
        New-Item -ItemType Directory -Force -Path $outDir | Out-Null
    }

    $from = 0
    $to = 0
    $range = $WD_EXPORT_ALL_DOCUMENT
    if ($null -ne $a.from -and $null -ne $a.to) {
        $from = [int]$a.from
        $to = [int]$a.to
        if ($from -gt 0 -and $to -ge $from) { $range = $WD_EXPORT_FROM_TO }
    }

    # OutputFileName, ExportFormat, OpenAfterExport, OptimizeFor, Range, From,
    # To, Item, IncludeDocProps, KeepIRM, CreateBookmarks, DocStructureTags,
    # BitmapMissingFonts, UseISO19005_1
    $doc.ExportAsFixedFormat(
        $out,
        $WD_EXPORT_FORMAT_PDF,
        $false,
        $WD_EXPORT_OPTIMIZE_FOR_PRINT,
        $range,
        $from,
        $to,
        $WD_EXPORT_DOCUMENT_CONTENT,
        $true,
        $true,
        $WD_EXPORT_CREATE_HEADING_BOOKMARKS,
        $true,
        $true,
        $false)

    return @{
        out       = $out
        pageCount = (Get-PageCount $doc)
        sizeBytes = [int64](Get-Item -LiteralPath $out).Length
    }
}

function Cmd-Outline($a) {
    $doc = Resolve-Doc ([string]$a.docId)
    $limit = 2000
    if ($null -ne $a.limit) { $limit = [int]$a.limit }

    $found = New-Object System.Collections.ArrayList

    # Style-based Find costs one COM round trip per heading. Walking every
    # paragraph would cost one per paragraph, which is unusably slow on long
    # documents.
    foreach ($level in 1..9) {
        if ($found.Count -ge $limit) { break }
        $styleId = Get-HeadingStyleId $level
        try {
            $rng = $doc.Content
            $f = $rng.Find
            $f.ClearFormatting()
            $f.Text = ''
            $f.Forward = $true
            $f.Wrap = $WD_FIND_STOP
            $f.Format = $true
            $f.Style = $styleId
            $guard = 0
            while ($guard -lt $limit -and $f.Execute()) {
                $guard++
                $text = ([string]$rng.Text).Trim([char]13, [char]7, [char]11, [char]12, ' ')
                if (-not [string]::IsNullOrWhiteSpace($text)) {
                    [void]$found.Add(@{
                            level = $level
                            text  = $text
                            page  = (Get-PageOf $rng)
                            start = [int]$rng.Start
                        })
                }
                if ($found.Count -ge $limit) { break }
            }
        } catch {
            # A document that does not define this heading style simply has no
            # headings at this level.
        }
    }

    # Fallback for documents using custom heading styles: scan outline levels
    # directly, but only when the document is small enough for that to be cheap.
    if ($found.Count -eq 0) {
        try {
            $paraCount = [int]$doc.Paragraphs.Count
            if ($paraCount -le 2000) {
                for ($i = 1; $i -le $paraCount; $i++) {
                    $p = $doc.Paragraphs.Item($i)
                    $lvl = [int]$p.OutlineLevel
                    if ($lvl -lt $WD_OUTLINE_BODY_TEXT) {
                        $text = ([string]$p.Range.Text).Trim([char]13, [char]7, ' ')
                        if (-not [string]::IsNullOrWhiteSpace($text)) {
                            [void]$found.Add(@{
                                    level = $lvl
                                    text  = $text
                                    page  = (Get-PageOf $p.Range)
                                    start = [int]$p.Range.Start
                                })
                        }
                    }
                    if ($found.Count -ge $limit) { break }
                }
            }
        } catch { }
    }

    $sorted = @($found | Sort-Object { $_.start })
    return @{ headings = $sorted; count = $sorted.Count }
}

function Cmd-Search($a) {
    $doc = Resolve-Doc ([string]$a.docId)
    $query = [string]$a.query
    if ([string]::IsNullOrWhiteSpace($query)) { throw (New-HostError 'invalid_request' "Search query is empty.") }

    $limit = 200
    if ($null -ne $a.limit) { $limit = [int]$a.limit }
    $matchCase = $false
    if ($null -ne $a.matchCase) { $matchCase = [bool]$a.matchCase }
    $wholeWord = $false
    if ($null -ne $a.wholeWord) { $wholeWord = [bool]$a.wholeWord }

    $docEnd = [int]$doc.Content.End
    $hits = New-Object System.Collections.ArrayList

    $rng = $doc.Content
    $f = $rng.Find
    $f.ClearFormatting()
    $f.Text = $query
    $f.Forward = $true
    $f.Wrap = $WD_FIND_STOP
    $f.Format = $false
    $f.MatchCase = $matchCase
    $f.MatchWholeWord = $wholeWord

    $guard = 0
    while ($guard -lt $limit -and $f.Execute()) {
        $guard++
        $start = [int]$rng.Start
        $end = [int]$rng.End
        $ctxStart = [Math]::Max(0, $start - 60)
        $ctxEnd = [Math]::Min($docEnd - 1, $end + 60)
        $snippet = ''
        try {
            $snippet = ([string]$doc.Range($ctxStart, $ctxEnd).Text) -replace '[\r\n\a\v\f\t]+', ' '
            $snippet = $snippet.Trim()
        } catch { }
        [void]$hits.Add(@{
                page    = (Get-PageOf $rng)
                start   = $start
                end     = $end
                snippet = $snippet
            })
    }

    $result = @($hits)
    return @{ hits = $result; count = $result.Count; query = $query }
}

function Cmd-Text($a) {
    $doc = Resolve-Doc ([string]$a.docId)

    $fromPage = 0
    if ($null -ne $a.fromPage) { $fromPage = [int]$a.fromPage }
    $toPage = 0
    if ($null -ne $a.toPage) { $toPage = [int]$a.toPage }

    $totalPages = Get-PageCount $doc

    if ($fromPage -le 0) {
        return @{ text = [string]$doc.Content.Text; fromPage = 1; toPage = $totalPages }
    }

    if ($toPage -le 0 -or $toPage -gt $totalPages) { $toPage = $totalPages }
    if ($fromPage -gt $totalPages) { throw (New-HostError 'page_out_of_range' "Page $fromPage is beyond the end of the document ($totalPages pages).") }

    $startRange = $doc.GoTo($WD_GOTO_PAGE, $WD_GOTO_ABSOLUTE, $fromPage)
    $start = [int]$startRange.Start
    if ($toPage -ge $totalPages) {
        $end = [int]$doc.Content.End
    } else {
        $nextRange = $doc.GoTo($WD_GOTO_PAGE, $WD_GOTO_ABSOLUTE, $toPage + 1)
        $end = [int]$nextRange.Start
    }
    if ($end -lt $start) { $end = $start }

    return @{ text = [string]$doc.Range($start, $end).Text; fromPage = $fromPage; toPage = $toPage }
}

function Cmd-Info($a) {
    $doc = Resolve-Doc ([string]$a.docId)
    return @{
        name       = [string]$doc.Name
        pageCount  = (Get-PageCount $doc)
        wordCount  = (Get-WordCount $doc)
        paragraphs = [int]$doc.Paragraphs.Count
        title      = [string](Get-DocProp $doc 'Title')
        author     = [string](Get-DocProp $doc 'Author')
        subject    = [string](Get-DocProp $doc 'Subject')
        company    = [string](Get-DocProp $doc 'Company')
        created    = [string](Get-DocProp $doc 'Creation date')
        revision   = [string](Get-DocProp $doc 'Revision number')
    }
}

# --- editing the original ----------------------------------------------------
#
# Everything above this line serves reads, and a read never touches the user's
# file: Open-DocInternal works on an unblocked temp copy. An edit cannot. It
# must open the original, which drags in two failure modes that a read never
# meets, and *both of them hang rather than fail*:
#
#   1. Another process holds the file. Measured: Documents.Open never returns,
#      with DisplayAlerts already off; both processes needed external kills.
#   2. The file carries the mark of the web. Measured on this machine with a
#      Zone.Identifier of ZoneId=3: Documents.Open also never returns. Word
#      wants to route the file into Protected View, and the automation call
#      simply blocks. It does not raise "this document is in Protected View".
#
# So neither can be discovered by trying. Both are checked before the open, and
# the open itself stays inside the caller's timeout regardless (ADR 0007).

$WD_STYLE_NORMAL = -1
$WD_CHARACTER = 1
$WD_COLLAPSE_START = 1

# Reads the Zone.Identifier alternate data stream. Sub-millisecond, and it
# cannot hang, which is the entire point: the obvious test -- ask Word to open
# it and see what happens -- is the call that wedges.
function Test-MarkOfTheWeb([string]$path) {
    try {
        $zone = Get-Content -LiteralPath $path -Stream 'Zone.Identifier' -Raw -ErrorAction Stop
    } catch {
        return $false
    }
    if ($zone -match 'ZoneId\s*=\s*(\d+)') { return ([int]$Matches[1] -ge 3) }
    return $false
}

# Opens the user's own document for writing, and returns the Document.
#
# For a marked file this takes the Protected View route, which is the
# automation equivalent of a human clicking "Enable Editing":
# ProtectedViewWindows.Open() loads the file in the sandbox and .Edit() promotes
# it to an ordinary editable Document. Measured: the edit lands on disk and the
# file's Zone.Identifier is still there afterwards.
#
# The alternative, Unblock-File, was rejected. It deletes a security marker from
# a file the user owns, silently and permanently, to work around a limitation of
# ours. The Protected View path leaves the marker alone; its cost is a per-file
# trust record under HKCU -- exactly the record clicking Enable Editing creates,
# per-file, and clearable from the Trust Center. ADR 0007 records the trade.
function Open-OriginalForEdit([string]$path) {
    Initialize-Word

    if (Test-MarkOfTheWeb $path) {
        # FileName, Password, AddToRecentFiles, Repair
        $window = $script:App.ProtectedViewWindows.Open($path, $FAIL_FAST_PASSWORD, $false, $false)
        try {
            $doc = $window.Edit()
        } catch {
            # Edit() failing leaves the Protected View window open, and that
            # window holds the user's file. Our own hidden Word would then be
            # the "other program" every later edit tells the user to close --
            # recreated on every retry, until idle shutdown. Close it here.
            #
            # Closed but deliberately not released -- see the note below the
            # try/catch. Releasing this RCW leaks WINWORD processes across
            # later operations, measured.
            try { $window.Close() } catch { }
            throw
        }
        # Edit() transitions the window into a normal document, so closing it is
        # a no-op or throws; either way the Protected View window is done.
        try { $window.Close() } catch { }
        # DO NOT add ReleaseComObject($window) here. It is the obvious tidy-up,
        # every other COM object in this file gets one, and a review asked for it
        # on exactly that reasoning -- consistency, plus a stated risk of
        # "leaking COM references and prolonging file/Word resource retention".
        # Adding it *causes* that leak instead of preventing it. Measured on the
        # full edit suite, one variable, twice from opposite directions:
        #
        #   with ReleaseComObject($window)     19/20, 7 WINWORD alive 90 s after
        #                                      teardown, +7 on the machine
        #   without it (everything else same)  20/20, zero left, count unchanged
        #
        # The single-operation probe that ought to have caught it says the
        # opposite -- file free in 16 ms and Word exiting in ~2.8 s either way,
        # against a control that shows the probe can see retention. Both are
        # true. The release costs nothing *within* an operation and breaks Word's
        # lifecycle *across* them, so a probe that runs one operation and quits
        # Word itself measures the wrong axis, however careful its controls are.
        # Protected View also spawns a second WINWORD we never hold a handle to
        # (measured: one per open), which is the likely route, but the mechanism
        # is not pinned and this note deliberately does not invent one.
        #
        # The concern behind the review is real, so it is now asserted rather
        # than argued: the Protected View case in edit-smoke.mjs checks
        # `lockReleased`, which the host sets by polling for a write handle.
        return @{ doc = $doc; protectedView = $true }
    }

    # FileName, ConfirmConversions, ReadOnly, AddToRecentFiles,
    # PasswordDocument, PasswordTemplate, Revert
    $doc = $script:App.Documents.Open($path, $false, $false, $false, $FAIL_FAST_PASSWORD, $FAIL_FAST_PASSWORD, $false)
    return @{ doc = $doc; protectedView = $false }
}

# A paragraph's Range ends with its paragraph mark, and inside a table cell with
# an end-of-cell mark as well. Assigning to a Range that includes those marks
# deletes them, welding the paragraph onto the next one. Trimming them off is
# what makes "replace the text of this paragraph" mean only that.
function Get-TextRange($para) {
    # Word's string and its character *model* disagree inside a table. A cell
    # paragraph's Range.Text ends "`r" + chr(7) -- two characters in the string --
    # but End-Start counts the end-of-cell mark as one position. Measured: a
    # blind MoveEnd(-2) ate a real character, rewriting "cell 11" as
    # "cell rewritten1". So derive the trim from the position span rather than
    # from the string length, which is correct in both models.
    $raw = [string]$para.Range.Text
    $visible = $raw.TrimEnd("`r", [char]7, "`n")
    $range = $para.Range
    $drop = ($range.End - $range.Start) - $visible.Length
    if ($drop -gt 0) { $range.MoveEnd($WD_CHARACTER, -$drop) | Out-Null }
    return $range
}

function Set-ParagraphText($para, [string]$text) {
    (Get-TextRange $para).Text = $text
}

# Style assignment, and the only two mechanisms that work here.
#
# Measured on this German Word: Range.Style = 'berschrift1' (the style id Word
# actually writes into the file) throws, and so does Range.Style = 'Heading 1'.
# Assigning the numeric wd* constant works, and so does assigning another
# paragraph's Style *object*. The write side therefore never names a style.
function Set-ParagraphHeadingLevel($para, [int]$level) {
    if ($level -le 0) { $para.Range.Style = $WD_STYLE_NORMAL }
    else { $para.Range.Style = (Get-HeadingStyleId $level) }
}

# What a new paragraph should look like when the caller did not say.
#
# Word's own behaviour when you press Enter at the end of a heading is to start
# the next paragraph in that style's follow-on style -- body text, not another
# heading. NextParagraphStyle is that rule, as an object, so it carries across
# localizations without naming anything.
function Set-InheritedStyle($target, $reference) {
    try {
        $next = $reference.Style.NextParagraphStyle
        if ($null -ne $next) { $target.Range.Style = $next; return }
    } catch { }
    try { $target.Range.Style = $reference.Style } catch { }
}

# Applies one intent to an already-open document. Split out so the open/close
# lock window in Cmd-Edit reads as a single unbroken sequence.
function Invoke-EditOperation($doc, [int]$wordIndex, $a) {
    $op = [string]$a.op
    $para = $doc.Paragraphs.Item($wordIndex)

    switch ($op) {
        'replace_text' {
            Set-ParagraphText $para ([string]$a.text)
            return $wordIndex
        }
        'delete_paragraph' {
            # Constraint 6's bug class at a second site. A cell paragraph's
            # Range runs to the end-of-cell mark, and deleting that mark either
            # throws or quietly damages the table -- a cell must keep at least
            # one paragraph. Trim the mark off the range, derived from the
            # position span rather than the string length for the reason
            # Get-TextRange documents, and delete what is left.
            $range = $para.Range
            $raw = [string]$range.Text
            if ($raw.Contains([char]7)) {
                $visible = $raw.TrimEnd("`r", [char]7, "`n")
                $drop = ($range.End - $range.Start) - $visible.Length
                if ($drop -gt 0) { $range.MoveEnd($WD_CHARACTER, -$drop) | Out-Null }
            }
            $range.Delete() | Out-Null
            return 0
        }
        'set_heading_level' {
            Set-ParagraphHeadingLevel $para ([int]$a.headingLevel)
            return $wordIndex
        }
        'insert_paragraph_after' {
            $para.Range.InsertParagraphAfter()
            $new = $doc.Paragraphs.Item($wordIndex + 1)
            Set-ParagraphText $new ([string]$a.text)
            if ($null -ne $a.headingLevel) { Set-ParagraphHeadingLevel $new ([int]$a.headingLevel) }
            else { Set-InheritedStyle $new $para }
            return ($wordIndex + 1)
        }
        'insert_paragraph_before' {
            $range = $para.Range
            $range.Collapse($WD_COLLAPSE_START)
            $range.InsertParagraphBefore()
            $new = $doc.Paragraphs.Item($wordIndex)
            Set-ParagraphText $new ([string]$a.text)
            if ($null -ne $a.headingLevel) { Set-ParagraphHeadingLevel $new ([int]$a.headingLevel) }
            else { Set-InheritedStyle $new $doc.Paragraphs.Item($wordIndex + 1) }
            return $wordIndex
        }
        default { throw "Unsupported edit operation '$op'." }
    }
}

# One edit: open the original, change one paragraph, save, close. The lock is
# held for that sequence and no longer (ADR 0005), and the document is never
# registered in $script:Docs, so no later command can resolve it and silently
# reacquire the lock.
#
# Failures come back as a `status` on the success channel rather than as
# exceptions. The dispatch loop collapses every throw to a single `word_error`,
# which would leave the caller string-matching localized German exception text
# to tell "the file is locked" from "that paragraph is not where you think".
function Cmd-Edit($a) {
    $path = [string]$a.path
    $started = [Diagnostics.Stopwatch]::StartNew()

    if ([string]::IsNullOrWhiteSpace($path)) { return @{ status = 'file_not_found'; path = $path } }
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return @{ status = 'file_not_found'; path = $path } }

    # Pre-flight, in this order: the lock check is 4 ms and the MOTW check is
    # sub-millisecond, and either one hanging Word is a two-process kill.
    if (-not (Test-FileWritable $path)) { return @{ status = 'document_locked'; path = $path } }
    $marked = Test-MarkOfTheWeb $path

    $openMs = 0
    $lockHeldMs = 0
    $opened = $null

    # The lock window starts where the file is actually acquired, not where the
    # command was entered. The pre-flight above touches the file but never holds
    # it, and timing both from one stopwatch is what made the reported figure
    # overstate the window it is named for -- in a design whose entire claim is
    # that the window is short, so the error ran in the flattering direction.
    $lockStarted = [Diagnostics.Stopwatch]::StartNew()
    try {
        $opened = Open-OriginalForEdit $path
        $openMs = [int]$lockStarted.ElapsedMilliseconds
    } catch {
        return @{ status = 'open_failed'; path = $path; protectedView = $marked; detail = $_.Exception.Message }
    }

    $doc = $opened.doc
    $result = $null
    try {
        $count = [int]$doc.Paragraphs.Count
        $wordIndex = [int]$a.wordIndex

        if ($wordIndex -lt 1 -or $wordIndex -gt $count) {
            return @{ status = 'address_not_resolvable'; reason = 'out_of_range'; wordIndex = $wordIndex; paragraphCount = $count }
        }

        # The map said this paragraph is at this position with this text. One
        # property read confirms it before anything is mutated. Without this an
        # address that has drifted -- a document edited outside the session
        # between the read and the edit -- silently rewrites the wrong paragraph.
        $expected = Get-NormalizedText ([string]$a.expectedText)
        $actual = Get-NormalizedText ([string]$doc.Paragraphs.Item($wordIndex).Range.Text)
        if ($expected -ne $actual) {
            return @{
                status         = 'address_not_resolvable'
                reason         = 'text_mismatch'
                wordIndex      = $wordIndex
                paragraphCount = $count
                expectedText   = $expected
                actualText     = $actual
            }
        }

        $editStarted = [Diagnostics.Stopwatch]::StartNew()
        $touched = Invoke-EditOperation $doc $wordIndex $a
        $editMs = [int]$editStarted.ElapsedMilliseconds

        # A delete is the one operation whose success cannot be confirmed by
        # looking at the paragraph it names, because that paragraph is meant to
        # be gone. So confirm it by the count, before saving rather than after:
        # inside a table a cell must keep one paragraph, so a delete there can
        # quietly do nothing instead of throwing, and an unverified 'edited'
        # would then be a lie the caller has no way to detect.
        if ([string]$a.op -eq 'delete_paragraph') {
            $afterCount = [int]$doc.Paragraphs.Count
            if ($afterCount -ne ($count - 1)) {
                return @{
                    status         = 'edit_had_no_effect'
                    reason         = 'paragraph_count_unchanged'
                    wordIndex      = $wordIndex
                    paragraphCount = $afterCount
                    expectedCount  = $count - 1
                }
            }
        }

        $saveStarted = [Diagnostics.Stopwatch]::StartNew()
        $doc.Save()
        $saveMs = [int]$saveStarted.ElapsedMilliseconds

        # Everything past this point is reporting, not mutation: the file on
        # disk is already changed. A COM throw here must not surface as a failed
        # edit, because the caller treats a throw as "outcome unknown" and the
        # truth is "the edit succeeded and we failed to describe it".
        $page = 0
        $paragraphCount = 0
        try {
            if ($touched -gt 0 -and $touched -le [int]$doc.Paragraphs.Count) {
                $page = Get-PageOf $doc.Paragraphs.Item($touched).Range
            }
            $paragraphCount = [int]$doc.Paragraphs.Count
        } catch {
            $page = 0
        }

        $result = @{
            status         = 'edited'
            protectedView  = [bool]$opened.protectedView
            markOfTheWeb   = $marked
            wordIndex      = $touched
            paragraphCount = $paragraphCount
            page           = $page
            openMs         = $openMs
            editMs         = $editMs
            saveMs         = $saveMs
        }
    } finally {
        # The lock window ends here whatever happened above, including a throw
        # in the middle of a mutation. wdDoNotSaveChanges: anything worth
        # keeping was saved explicitly, and a half-applied edit must not be.
        try { $doc.Close($WD_DO_NOT_SAVE_CHANGES) } catch { }
        # Stopped here, immediately after the close that releases the file, so
        # the figure covers acquire -> release and nothing else. Measured at
        # 0-1 ms, `Close()` is what ends the window; the release poll below only
        # confirms it and must not be counted inside it.
        $lockHeldMs = [int]$lockStarted.ElapsedMilliseconds
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null } catch { }
    }

    if ($null -ne $result) {
        # Close() returning is not proof the file is released. Quit() is cited
        # here only as precedent for "a call returning is not the work being
        # done" -- measured, it returns 3.0-3.7 s before its process exits. It is
        # not evidence about Close, which is a different call with measurably
        # different behaviour, and nothing about either may be inferred from the
        # other; the poll below is what establishes the fact for Close. (This
        # comment read "~120 ms" until that number was actually measured. It was
        # wrong by a factor of ~25-30, and Stop-Word had sized a wait on it.)
        # Measure it rather than assume, because the next thing the caller does
        # is read the file to confirm the edit, and a re-open into a still-held
        # file is the hang this whole command is built to avoid.
        $releaseStarted = [Diagnostics.Stopwatch]::StartNew()
        $released = $false
        while ($releaseStarted.ElapsedMilliseconds -lt 5000) {
            if (Test-FileWritable $path) { $released = $true; break }
            Start-Sleep -Milliseconds 20
        }
        $result.releaseMs = [int]$releaseStarted.ElapsedMilliseconds
        $result.released = $released
        $result.lockHeldMs = $lockHeldMs
        # The whole command, pre-flight and release poll included. Reported
        # beside `lockHeldMs` rather than as it, because they answer different
        # questions and conflating them is what this pair exists to prevent.
        $result.totalMs = [int]$started.ElapsedMilliseconds
    }

    return $result
}

function Cmd-Close($a) {
    $docId = [string]$a.docId
    $script:DocArgs.Remove($docId) | Out-Null
    Close-Doc $docId
    return @{ closed = $true }
}

# --- dispatch loop -----------------------------------------------------------

# Clean up after any previous host that died without tearing its Word down.
Clear-OrphanedWord

$exitRequested = $false
try {
    while (-not $exitRequested) {
        $line = [Console]::In.ReadLine()
        if ($null -eq $line) { break }
        if ([string]::IsNullOrWhiteSpace($line)) { continue }

        $id = $null
        try {
            $req = $line | ConvertFrom-Json
            $id = $req.id
            $cmdArgs = $req.args
            switch ([string]$req.cmd) {
                'ping' { Send-Ok $id (Cmd-Ping $cmdArgs) }
                'open' { Send-Ok $id (Cmd-Open $cmdArgs) }
                'structure' { Send-Ok $id (Cmd-Structure $cmdArgs) }
                'edit' { Send-Ok $id (Cmd-Edit $cmdArgs) }
                'export' { Send-Ok $id (Cmd-Export $cmdArgs) }
                'outline' { Send-Ok $id (Cmd-Outline $cmdArgs) }
                'search' { Send-Ok $id (Cmd-Search $cmdArgs) }
                'text' { Send-Ok $id (Cmd-Text $cmdArgs) }
                'info' { Send-Ok $id (Cmd-Info $cmdArgs) }
                'close' { Send-Ok $id (Cmd-Close $cmdArgs) }
                'quit' {
                    Stop-Word
                    Send-Ok $id @{ stopped = $true }
                    $exitRequested = $true
                }
                default { Send-Fail $id 'unknown_command' "Unknown command '$([string]$req.cmd)'." }
            }
        } catch {
            # A typed failure carries its own code; anything else is genuinely
            # unclassified and stays `word_error`.
            $thrown = $_.TargetObject
            if ($thrown -is [hashtable] -and $thrown.__hostError) {
                Send-Fail $id ([string]$thrown.code) ([string]$thrown.message) $thrown.data
            } else {
                Send-Fail $id 'word_error' $_.Exception.Message
            }
        }
    }
} finally {
    Stop-Word
}
