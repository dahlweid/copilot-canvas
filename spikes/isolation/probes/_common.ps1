# Shared helpers for the Word isolation probes.
#
# This file exists for one reason: there is exactly one sound way to kill a
# WINWORD from this tree, and it should exist in exactly one place.
#
# The unsound way, which every probe here used to use, is to snapshot the
# WINWORD pid set, create a Word, snapshot again, and treat the difference as
# "the instance we own". This repository has measured that false twice:
#
#   * spikes/isolation/probes/probe-init-attribution.ps1 differenced 2 new pids
#     for the 1 instance it created.
#   * A census control in the same probe saw 2 strangers' WINWORDs appear in a
#     40-second window with nothing launched. This is a shared machine and a
#     cold New-Object costs seconds, so the window routinely catches strangers.
#
# A census difference is therefore an inference over a POPULATION, and the thing
# a kill needs is a fact about ONE PROCESS. The rule this file enforces:
#
#   * The census is a REPORT, never an authorisation. An empty difference is
#     weak evidence we started nothing; a non-empty one is evidence of nothing
#     at all, and must never be used as permission to act.
#   * The only WINWORD this tree may kill is one whose pid came back from
#     CreateProcess, or one resolved from a window handle we obtained off our
#     own RCW -- and then only through Stop-VerifiedWord below.
#
# Note the asymmetry with spikes/powerpoint/probes/_common.ps1, which forbids
# killing or quitting a COM-obtained instance outright. That is right there and
# would be wrong here: PowerPoint is single-instance, so New-Object ATTACHES to
# the user's session, whereas spikes/isolation/probes/probe-newobject-attach.ps1
# measured that New-Object Word.Application never attached to a running Word in
# either arm. Quitting a Word we created is sound. Killing a pid we merely
# observed appearing is not, and that distinction is the whole of issue #136.

# Shares its SHAPE with Stop-VerifiedWord in
# .github/extensions/office-canvas/src/word/word-host.ps1 and Stop-VerifiedPpt in
# spikes/powerpoint/probes/_common.ps1: all three ATTEMPT to pin a handle, prove
# identity against a recorded StartTime, and refuse rather than kill when the
# proof fails. Attempt, not achieve: only the shipped host checks that the pin
# succeeded. `try { $null = $p.Handle } catch { ... }` below cannot tell, because
# .Handle answers $null WITHOUT throwing on a process that has exited or refuses
# the open (spikes/isolation/probes/probe-processname-after-exit.ps1, arms
# E2/F1/F2/F3), so 'declined:handle' is unreachable and control continues on an
# object that may never have been pinned. That matters more here than it did
# there, because line 107 kills by PID: the pin is the only thing standing
# between a recycled pid and Stop-Process. Left as it is deliberately, and filed
# rather than fixed in passing -- see issue #155.
# They are NOT a mirror, and this comment used to say they were.
# Concretely they differ in the state vocabulary they return (the shipped host
# returns prose, not these short tokens, and has no 'declined:nopid'), in the
# kill primitive, and in whether the terminate is followed by a re-check --
# which is not an inconsistency to be tidied away but a consequence of the
# primitive: `Stop-Process -ErrorAction SilentlyContinue` below swallows its
# failure, so only a re-check can see it, whereas `.Kill()` in the shipped host
# throws and is caught at the call. Symmetry is not evidence, and issue #145
# settled that a divergence is not on its own a defect.
#
# Takes a pid that came back from CreateProcess -- or from the hwnd route, which
# is sound for a different and stronger reason: it reads the pid off a window we
# reached through our own RCW, so no stranger can enter it by racing our census
# -- plus the StartTime recorded for that pid at launch. Refuses unless the live
# process still matches both.
#
# ORDER IS LOAD-BEARING, and each step earns its place:
#
#   1. The handle is pinned FIRST. Everything after it is a read of a process
#      that could otherwise exit and have its pid reused underneath the checks.
#      Windows will not recycle a pid while a handle to the process object is
#      open, so from that point the identity read and the terminate are provably
#      about the same process.
#   2. ProcessName, so a pid already recycled onto something else is refused
#      before we look at times at all. The read is wrapped in a try/catch, and
#      that catch CANNOT CURRENTLY FIRE -- measured, and the justification that
#      used to sit here was false. It claimed `ProcessName` on an exited process
#      can throw into this caller, and that an uncaught throw would propagate
#      into the caller's `finally` and skip the cleanup below it. Neither half
#      survives measurement, for two independent reasons
#      (spikes/isolation/probes/probe-processname-after-exit.ps1):
#      the .NET getter does throw, but PowerShell's property adapter converts it
#      to $null before any caller sees it -- with `$Error` growing by zero, even
#      under 'Stop' (arms A1/A2); and the object never reaches that state here
#      anyway, because `Get-Process -Id` materializes the name and that copy
#      outlives the process (arm A1). So 'declined:unreadable-name' is
#      UNREACHABLE via this acquisition route: an unreadable name would arrive
#      below as $null and be refused by step 2's own comparison as
#      'declined:name'. The try/catch is kept as defence in depth should the
#      acquisition route ever change -- deleting it is a separate decision --
#      but nothing should be written that depends on the state existing.
#   3. `$null -eq $ExpectedStart` BEFORE any comparison against it. This is not
#      style. Under PowerShell, `$someDateTime -ne $null` evaluates to $true, so
#      a missing expected time compared directly would pass the mismatch test and
#      reach the kill. Issue #114 shipped exactly that inversion.
#   4. Only then the StartTime match.
#
# 'killed' vs 'killed:survived': both mean every guard passed and the terminate
# was issued. They differ in what was then OBSERVED. Only 'gone' and 'killed'
# mean no Word is left behind; everything else is a leak the caller must say out
# loud.
#
# Absence is distinguished from non-verification. A pid that is gone is 'gone'
# and fine -- no process holds it, so nothing we launched is running under it,
# and that is true whether or not a StartTime was ever recorded. A pid we cannot
# verify is declined and said out loud. Callers must report a decline rather
# than swallow it: it means a leaked Word, a pid collision, or both.
#
# Every failure path returns. None falls through to the kill.
function Stop-VerifiedWord([int]$ProcessId, $ExpectedStart) {
    if (-not $ProcessId) { return 'declined:nopid' }
    $p = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
    if (-not $p) { return 'gone' }
    try { $null = $p.Handle } catch { return 'declined:handle' }
    try { $name = $p.ProcessName } catch { return 'declined:unreadable-name' }
    if ($name -ne 'WINWORD') { return 'declined:name' }
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

# The launch-time half of the contract above.
#
# Read the StartTime immediately after the pid is minted. Without it
# Stop-VerifiedWord declines every time and the teardown silently becomes a
# leak; with a failed read it returns $null, which declines HONESTLY rather than
# killing on the pid alone. Never synthesise a value here.
function Get-WordStartTime([int]$ProcessId) {
    if (-not $ProcessId) { return $null }
    $p = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
    if (-not $p) { return $null }
    try { return $p.StartTime } catch { return $null }
}

# What a probe says when its census saw a WINWORD appear and it is still up.
#
# Deliberately offers no pid as a kill target and names no cause: this set is
# differenced, so it may hold a Word another session started, and the code
# cannot tell which. `WorkerKilled` is the one extra thing a probe can honestly
# say -- if we force-killed a worker powershell.exe, we orphaned whatever Word
# that worker was driving, and that IS a cause this code knows.
function Write-CensusSurvivors($Survivors, [switch]$WorkerKilled) {
    $s = @($Survivors)
    if ($s.Count -eq 0) { return }
    Write-Host "  $($s.Count) WINWORD appeared during this probe and is still up: $($s -join ', ')"
    if ($WorkerKilled) {
        Write-Host "  This probe force-killed its own worker process, which orphans the Word that"
        Write-Host "  worker was driving -- so at least one of these is probably that orphan."
    }
    Write-Host "  Not killed: this set is a census difference, which is measured unsound here"
    Write-Host "  (2 strangers' WINWORDs appeared in a 40 s window with nothing launched), so it"
    Write-Host "  may hold another session's Word. Confirm and close by hand."
}
