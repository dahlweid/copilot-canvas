# Shared helpers for the PowerPoint probes.
#
# Process hygiene is the reason this file exists, and the rule it enforces is
# narrower than it used to be.
#
# This file used to snapshot the POWERPNT pid set, attach with New-Object, and
# treat the difference as "the instance we own" -- then Quit() and force-kill it.
# Every part of that is unsound here:
#
#   * New-Object -ComObject PowerPoint.Application ATTACHES to a running
#     instance rather than starting one (FINDINGS.md, single-instance). So the
#     object we hold is routinely the USER'S PowerPoint, with their unsaved
#     decks in it.
#   * Differencing a pid census over-reports: probe-init-attribution.ps1
#     measured 2 new pids for 1 instance created. And the difference is
#     non-empty by construction whenever anyone else starts PowerPoint during
#     our census window.
#
# Those combine into the exact inversion the old guard was written to prevent:
# `if ($ctx.Owned.Count -gt 0) { Quit }` passed PRECISELY in the race where we
# had attached to somebody else's instance.
#
# The rule now:
#
#   * The census is a REPORT, never an authorisation. An empty difference is
#     evidence we started nothing; a non-empty one is evidence of nothing at
#     all, and must never be used as permission to act.
#   * A COM-obtained PowerPoint is never quit and never killed. There is no
#     signal available at this layer that would make that safe, so the honest
#     classification is "unproven" and the honest action is to release and
#     report.
#   * The only process this tree may kill is one whose pid came back from
#     CreateProcess, through Stop-VerifiedPpt below.

function Say($s) { [Console]::WriteLine($s) }
function Rep($l, $v) { [Console]::WriteLine(("{0,-42} {1}" -f $l, $v)) }

function Get-PptPids {
    @(Get-Process POWERPNT -ErrorAction SilentlyContinue | ForEach-Object Id)
}

# Attach to (or start) a PowerPoint instance and report the census difference.
#
# NewPids is NOT ownership. It is what the census happened to show, and it is
# reported so probes can say what they saw -- nothing may be killed or quit on
# the strength of it. See the header.
function New-PowerPointInstance {
    $before = Get-PptPids
    $app = New-Object -ComObject PowerPoint.Application
    # TRAP: PowerPoint's alert enum is INVERTED relative to Word's.
    #   Word:       wdAlertsNone = 0, wdAlertsAll = -1
    #   PowerPoint: ppAlertsNone = 1, ppAlertsAll = 2
    # Porting `DisplayAlerts = 0` from the Word host would select an undefined
    # value, not "suppress everything".
    #
    # NOTE: this is an application-level write, so when we have attached it
    # changes a setting in the user's live session. It is kept because the
    # probes need alerts suppressed to run unattended, and because it is not
    # destructive -- but it is a real side effect and is recorded as one in
    # README.md rather than described as harmless.
    try { $app.DisplayAlerts = 1 } catch { }
    # PowerPoint refuses Visible = $false on most builds, so we do not set it
    # here; probe-hide.ps1 measures what is and is not possible.
    Start-Sleep -Milliseconds 300
    $after = Get-PptPids
    $newPids = @($after | Where-Object { $before -notcontains $_ })
    [pscustomobject]@{ App = $app; NewPids = $newPids; Before = $before }
}

# Release the RCW and report. Never quits, never kills -- see the header.
function Close-PowerPointInstance($ctx) {
    if ($null -eq $ctx) { return }
    try { if ($ctx.App) { [Runtime.InteropServices.Marshal]::ReleaseComObject($ctx.App) | Out-Null } } catch { }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Start-Sleep -Seconds 2
    # Report anything the census saw appear, so a leak is visible rather than
    # silent. Deliberately no pid is offered as a kill target: a pid is a
    # coordinate, and handing one out invites the caller to do what this
    # function has just refused to do.
    $still = @(Get-PptPids | Where-Object { $ctx.NewPids -contains $_ })
    if ($still.Count -gt 0) {
        Rep "POSSIBLY left running" "$($still.Count) POWERPNT appeared during this probe and is still up -- confirm and close it by hand"
    }
}

# The only sanctioned kill in this tree.
#
# Takes a pid that came back from CreateProcess plus the StartTime recorded for
# it at launch, and refuses unless the live process still matches both. Mirrors
# Stop-VerifiedWord in spikes/isolation/probes/_common.ps1, step for step and
# state for state.
#
# NOT a mirror of Stop-VerifiedWord in
# .github/extensions/office-canvas/src/word/word-host.ps1, which this comment
# used to claim. That one terminates with $p.Kill() inside a try/catch -- which
# THROWS where Stop-Process -ErrorAction SilentlyContinue swallows -- and
# returns 'gone' for a pid that is not Word where both spikes helpers decline.
# Whether the shipped host should converge is a question about shipped
# behaviour with its own callers, and is filed separately.
#
# Order matters. The handle is pinned FIRST, because everything after it is a
# read of a process that could otherwise exit and have its pid reused
# underneath the checks. Absence is distinguished from non-verification: a pid
# that is gone is 'gone' and fine, whereas a pid we cannot verify is declined
# and said out loud.
#
# The ProcessName read is wrapped in a try/catch whose catch CANNOT CURRENTLY
# FIRE, and the justification that used to sit here was false. It claimed
# ProcessName on an exited process can throw into this caller, and that an
# uncaught throw would propagate into the caller's finally and skip everything
# below the call. The second half describes the callers correctly --
# Stop-IsolatedPowerPoint (spikes/powerpoint/probes/_isolated.ps1) runs from
# spikes/powerpoint/probes/probe-cross-instance-lock.ps1's finally and closes a
# desktop after this call, and spikes/powerpoint/probes/probe-second-process.ps1
# calls this from a finally that then releases the RCW, closes a desktop and
# prints the leak census -- but the first half is not true, so the consequence
# never arises. Measured, spikes/isolation/probes/probe-processname-after-exit.ps1
# (the behaviour is System.Diagnostics.Process's, not Word's or PowerPoint's):
# the .NET getter does throw, but PowerShell's property adapter converts it to
# $null before any caller sees it, with `$Error` growing by zero even under
# 'Stop' (arms A1/A2); and via `Get-Process -Id`, as on line 133, the name is
# materialized at acquisition and that copy outlives the process anyway (A1).
# So 'declined:unreadable-name' is UNREACHABLE here: an unreadable name would
# arrive below as $null and be refused as 'declined:name'. The try/catch is
# kept as defence in depth should the acquisition route ever change; nothing
# should be written that depends on the state existing.
#
# 'killed' vs 'killed:survived': both mean every guard passed and the terminate
# was issued. They differ only in what was then OBSERVED. Only 'gone' and
# 'killed' mean no PowerPoint is left behind.
#
# HYPOTHESIS, not a measured fact: the PROCESS_INFORMATION.hProcess handle from
# CreateProcess is never closed by these probes (no CloseHandle anywhere in
# this tree), and Windows will not recycle a pid while a handle to the process
# object is open -- so the pid is probably pinned already. That inference has
# NOT been probed here, and probing it means launching a process. The checks
# below therefore do not rely on it. Anyone adding CloseHandle should know they
# are changing something whose effect is unverified in both directions.
function Stop-VerifiedPpt([int]$ProcessId, $ExpectedStart) {
    if (-not $ProcessId) { return 'declined:nopid' }
    $p = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
    if (-not $p) { return 'gone' }
    try { $null = $p.Handle } catch { return 'declined:handle' }
    try { $name = $p.ProcessName } catch { return 'declined:unreadable-name' }
    if ($name -ne 'POWERPNT') { return 'declined:name' }
    if ($null -eq $ExpectedStart) { return 'declined:unverified' }
    try { $actual = $p.StartTime } catch { return 'declined:unreadable' }
    if ($actual -ne $ExpectedStart) { return 'declined:start' }
    Stop-Process -Id $ProcessId -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 800   # the call returns before the process is gone
    # Report the OBSERVED outcome, not the attempted one. Stop-Process here
    # swallows its errors, so without this check 'killed' would be a claim about
    # a call rather than about a process, and a caller that treats 'killed' as
    # "reaped" would hide the leak. Both states mean the guards passed and the
    # kill was issued; they differ only in what the machine did next.
    if (Get-Process -Id $ProcessId -ErrorAction SilentlyContinue) { return 'killed:survived' }
    return 'killed'
}

# --- PDF inspection without a PDF library -------------------------------------
# Two facts are needed about an exported PDF: how many pages it has, and how big
# each page is. Both are needed to answer "one page per slide, page-accurately".
# PowerPoint emits /MediaBox and the page objects in the clear, so a byte scan is
# enough; these return $null / -1 rather than guessing if that stops holding.

function Get-PdfPageBoxes([string]$Path) {
    $bytes = [IO.File]::ReadAllBytes($Path)
    $text = [Text.Encoding]::GetEncoding('ISO-8859-1').GetString($bytes)
    $boxes = @()
    foreach ($m in [regex]::Matches($text, '/MediaBox\s*\[\s*([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s*\]')) {
        $boxes += [pscustomobject]@{
            Width  = [double]$m.Groups[3].Value - [double]$m.Groups[1].Value
            Height = [double]$m.Groups[4].Value - [double]$m.Groups[2].Value
        }
    }
    if ($boxes.Count -eq 0) { return $null }
    $boxes
}

function Get-PdfPageCount([string]$Path) {
    $bytes = [IO.File]::ReadAllBytes($Path)
    $text = [Text.Encoding]::GetEncoding('ISO-8859-1').GetString($bytes)
    # /Type /Page not followed by 's' -- excludes /Pages tree nodes.
    $n = ([regex]::Matches($text, '/Type\s*/Page(?![s])')).Count
    if ($n -gt 0) { return $n }
    $c = [regex]::Matches($text, '/Count\s+(\d+)')
    if ($c.Count -gt 0) { return [int](($c | ForEach-Object { [int]$_.Groups[1].Value } | Measure-Object -Maximum).Maximum) }
    return -1
}
