# Probe: does PowerPoint expose Application.Hwnd, and does Application.ActiveWindow
# throw when no presentation is open?
#
# Both questions are UNMEASURED for PowerPoint in this tree. Every ActiveWindow /
# Hwnd probe under spikes/ is a WORD probe (spikes/isolation, spikes/live-word);
# there is no ActiveWindow reference anywhere under spikes/powerpoint. The claim
# that anchors the design discussion --
#
#   "Application.Hwnd does not exist on Word (PowerPoint has it, Word does not),
#    but Application.ActiveWindow.Hwnd does -- once a document is open"
#
# -- lives only as a parenthetical in the header of the Word probe
# spikes/isolation/probes/probe-word-ownership-hwnd.ps1, with no PowerPoint
# measurement behind it. Its sibling claim, that PowerPoint's ActiveWindow throws
# without a presentation, is a Word result (probe-init-attribution.ps1 arm B,
# "the arm that can sink the design") imported by analogy. This measures both on
# PowerPoint directly.
#
#   A  Application.Hwnd -- does the property resolve at all, and does the handle
#      it returns map (via GetWindowThreadProcessId) to the pid CreateProcess
#      handed us? If so, attribution needs no presentation and no window walk.
#   B  Application.ActiveWindow -- read it with a deck open (B1), then with none
#      open (B2, repeated across fresh isolated instances so the deckless result
#      is reproduced across processes, not sampled once). B2 is the PowerPoint
#      analogue of the Word arm above.
#
# Scope note, so this is not over-read: attribution is NOT exclusivity. Even if A
# holds, it does not make writing to a New-Object instance safe, because
# probe-single-instance.ps1 measures PowerPoint as single-instance -- a New-Object
# instance is the user's whenever they have one open, whatever its pid is. So this
# probe informs a shipped comment and issue #141's "Note", not the choice of the
# CreateProcess route for H1/H3.
#
# SAFETY: no arm writes to the Application object, so there is nothing to restore.
# The only PowerPoint touched is instances obtained through Start-IsolatedPowerPoint
# (each a CreateProcess pid) -- the top instance for arms A/B1, plus one fresh
# instance per B2 cycle -- and every one is torn down on every path in a finally via
# Stop-IsolatedPowerPoint. Measured exceptions are
# DISCRIMINATED by .NET type, never by message -- this Office is German and the
# message is localized. A message appears only as terminal diagnostic text in the
# top-level catch, where the type alone often will not say what went wrong; it is
# never branched on.

param([string]$Fixture, [int]$B2Cycles = 5)

$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot '_common.ps1')
. (Join-Path $PSScriptRoot '_isolated.ps1')

if (-not $Fixture) { $Fixture = Join-Path (Split-Path $PSScriptRoot -Parent) '.fixtures\deck.pptx' }
if (-not (Test-Path $Fixture)) { throw "fixture missing: $Fixture (run make-fixture.ps1 first)" }

$root = Join-Path $env:TEMP ("ppthwnd-" + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Force -Path $root | Out-Null
$src = Join-Path $root 'deck.pptx'; Copy-Item $Fixture $src

# The root cause under any COM wrapper: a method/property that throws comes back
# wrapped in MethodInvocationException, so the interesting type is the innermost.
function Root-Type($ex) {
    $r = $ex; while ($r.InnerException) { $r = $r.InnerException }
    "outer {0} / root {1}" -f $ex.GetType().Name, $r.GetType().Name
}

# Owning pid of a window handle, so an Application.Hwnd value is checked against
# the CreateProcess pid rather than trusted. GetWindowThreadProcessId is already
# P/Invoked on PptIso in _isolated.ps1; reuse it rather than redeclaring.
function Get-PidFromHwnd([IntPtr]$hwnd) {
    [uint32]$owner = 0
    [void][PptIso]::GetWindowThreadProcessId($hwnd, [ref]$owner)
    [int]$owner
}

Rep "POWERPNT pids before" ($(if (Get-PptPids) { (Get-PptPids) -join ',' } else { '(none)' }))

$iso = $null
try {
    $iso = Start-IsolatedPowerPoint -File $src
    if (-not $iso.App) {
        Rep "isolated instance" "NOT BOUND -- $($iso.Diag)"
    }
    else {
        $app = $iso.App
        Rep "launched pid (CreateProcess)" $iso.Pid
        Rep "bound in" "$($iso.BoundMs) ms"

        Say ""
        Say "== A: does Application.HWND exist on PowerPoint, and yield a usable handle? =="
        # Three outcomes look alike at a glance: the member is absent (IDispatch
        # does not know the name), the member exists but returns null, or it
        # returns a handle. Reflection over the IDispatch separates absent from
        # null; a bare property read cannot.
        $memberExists = $null
        try {
            $null = $app.GetType().InvokeMember('HWND',
                [System.Reflection.BindingFlags]::GetProperty, $null, $app, @())
            $memberExists = $true
        }
        catch [System.MissingMemberException] { $memberExists = $false }
        catch { $memberExists = 'read threw -- ' + (Root-Type $_.Exception) }
        Rep "  HWND is a member of the OM object" $memberExists

        $hwndVal = '<unset>'; $hwndEx = $null
        try { $hwndVal = $app.HWND } catch { $hwndEx = $_.Exception }
        if ($hwndEx) {
            Rep "  Application.HWND value" ("throws -- " + (Root-Type $hwndEx))
        }
        elseif ($null -eq $hwndVal) {
            # Report only what was observed: a null value came back through the
            # bind. Do NOT claim "no exception" -- $ErrorActionPreference is
            # 'Continue' here, so a non-terminating error would never reach the
            # catch above. $Error growth is not leaned on either: PowerShell's .NET
            # property adapter is measured (probe-processname-after-exit.ps1) to
            # convert a throw to $null without $Error growing, and whether the COM
            # adapter does the same is unmeasured in this tree -- either way it is
            # not a detector to rely on.
            Rep "  Application.HWND value" 'null through the OBJID_NATIVEOM bind (no usable handle marshalled)'
        }
        else {
            Rep "  Application.HWND value" "$hwndVal"
            $ownerPid = -1
            try { $ownerPid = Get-PidFromHwnd ([IntPtr][int64]$hwndVal) } catch { }
            Rep "  GetWindowThreadProcessId(HWND) -> pid" $ownerPid
            Rep "  matches CreateProcess pid ($($iso.Pid))" ($ownerPid -eq $iso.Pid)
        }

        # Contrast: the frame window plainly EXISTS -- Start-IsolatedPowerPoint
        # located it by class to bind. If the OM's HWND is null while the real
        # handle is right here and attributes our pid, then attribution through
        # this binding comes from external window enumeration, not from asking the
        # object -- which is exactly what the isolated route already does.
        $frame = [PptIso]::FindClass($iso.Desk, [uint32]$iso.Pid, 'PPTFrameClass')
        if ($frame -ne [IntPtr]::Zero) {
            $framePid = Get-PidFromHwnd $frame
            Rep "  external PPTFrameClass handle" ("$frame (pid $framePid; ours = $($framePid -eq $iso.Pid))")
        }
        else {
            Rep "  external PPTFrameClass handle" 'not found'
        }

        $usable = (($null -ne $hwndVal) -and ($hwndVal -ne '<unset>') -and ($null -eq $hwndEx))
        Rep "  VERDICT (A)" ($(if ($usable) {
                    'Application.HWND yields a usable handle on PowerPoint'
                }
                else {
                    'Application.HWND does NOT yield a usable handle through this binding (member/value above)'
                }))

        Say ""
        Say "== B1: Application.ActiveWindow with a deck open =="
        # Same three-way split as B2 below: a $null must not be reported as a
        # window just because $null.Hwnd does not throw.
        $aw = $null; $awEx = $null
        try { $aw = $app.ActiveWindow } catch { $awEx = $_.Exception }
        if ($awEx) { Rep "  B1 ActiveWindow (deck open)" ("throws -- " + (Root-Type $awEx)) }
        elseif ($null -eq $aw) { Rep "  B1 ActiveWindow (deck open)" 'returned $null -- no window object, no exception caught' }
        else {
            $awHwnd = '<unset>'; $awHwndEx = $null
            try { $awHwnd = $aw.Hwnd } catch { $awHwndEx = $_.Exception }
            $awHwndRep = $(if ($awHwndEx) { 'read threw -- ' + (Root-Type $awHwndEx) }
                elseif ($null -eq $awHwnd) { 'null' } else { "$awHwnd" })
            Rep "  B1 ActiveWindow (deck open)" "returns a live window object; Hwnd = $awHwndRep"
        }

        Say ""
        Say "== B2: Application.ActiveWindow (and Hwnd) with NO deck open, across fresh instances =="
        # The claim this arm feeds FINDINGS is about PowerPoint, not about one
        # process: "ActiveWindow returns $null when Presentations.Count = 0 and the
        # process is alive". Reading ONE deckless instance N times would only show
        # that instance is stable -- one observation sampled N times, which could be
        # an artifact of how THIS process closed its deck or tore down. So the loop is
        # over CYCLES, each a fresh Start-IsolatedPowerPoint -> close deck -> read ->
        # sweep, which gives reproduction ACROSS instances -- what FINDINGS asserts.
        #
        # The per-cycle sweep is in a finally, and that is load-bearing, not tidiness:
        # Quit() is measured NOT to reap (probe-stability.ps1, 15/15 survived), so
        # every cycle leaves a live PowerPoint that Stop-IsolatedPowerPoint must kill.
        # A finally runs that sweep on the continue-on-NOT-BOUND path, the catch path
        # and the normal path alike, so no cycle leaks even when one throws.
        #
        # Liveness is the discriminator between "no window because deckless" and "the
        # RCW is dead because the process exited", and it must STRADDLE each read, not
        # merely precede it: a read on an RCW whose process exited mid-call can itself
        # surface as $null. So each read is bracketed by a liveness sample on both
        # sides; a $null not bracketed by a live process is INCONCLUSIVE, never counted
        # as a windowless read.
        $b2Null = 0; $b2Window = 0; $b2Threw = 0; $b2Incon = 0
        for ($cy = 1; $cy -le $B2Cycles; $cy++) {
            $cySrc = Join-Path $root ("b2-$cy.pptx"); Copy-Item $Fixture $cySrc
            $cyIso = $null
            try {
                $cyIso = Start-IsolatedPowerPoint -File $cySrc
                if (-not $cyIso.App) {
                    $b2Incon++
                    Rep "  B2 [$cy/$B2Cycles]" "INCONCLUSIVE -- isolated instance NOT BOUND ($($cyIso.Diag))"
                    continue
                }
                $cyApp = $cyIso.App
                # Close the only presentation so this fresh instance is deckless.
                try { foreach ($p in @($cyApp.Presentations)) { try { $p.Saved = -1 } catch { }; $p.Close() } } catch { }
                Start-Sleep -Milliseconds 900
                # presCount: tell a null Presentations collection apart from a genuine
                # 0 -- $null.Count is 0 in this shell (no Set-StrictMode under
                # spikes/powerpoint/), so a bare 0 is not on its own evidence of deckless.
                $presObj = '<unset>'; try { $presObj = $cyApp.Presentations } catch { }
                $presCount = $(if ($presObj -eq '<unset>') { 'unreadable (read threw)' }
                    elseif ($null -eq $presObj) { 'Presentations object is null' }
                    else { $pc = 'unreadable'; try { $pc = $presObj.Count } catch { }; "$pc" })
                # Bracketed ActiveWindow read (the headline).
                $ab = [bool](Get-Process -Id $cyIso.Pid -ErrorAction SilentlyContinue)
                $aw2 = $null; $aw2Ex = $null
                try { $aw2 = $cyApp.ActiveWindow } catch { $aw2Ex = $_.Exception }
                $aa = [bool](Get-Process -Id $cyIso.Pid -ErrorAction SilentlyContinue)
                # Bracketed deckless Hwnd read (secondary; arm A on the deck-open
                # instance is the primary Application.Hwnd measurement).
                $hb = [bool](Get-Process -Id $cyIso.Pid -ErrorAction SilentlyContinue)
                $hwnd2 = $null; $hwnd2Ex = $null
                try { $hwnd2 = $cyApp.Hwnd } catch { $hwnd2Ex = $_.Exception }
                $ha = [bool](Get-Process -Id $cyIso.Pid -ErrorAction SilentlyContinue)
                $hwndRep = $(if (-not ($hb -and $ha)) { 'INCONCLUSIVE (process not alive across read)' }
                    elseif ($hwnd2Ex) { 'throws -- ' + (Root-Type $hwnd2Ex) }
                    elseif ($null -eq $hwnd2) { 'null through the bind' }
                    else { "$hwnd2" })
                # Classify the ActiveWindow read. A $null that is not alive-bracketed
                # is a dead-RCW artifact, not a windowless read -- INCONCLUSIVE.
                if (-not ($ab -and $aa)) {
                    $b2Incon++
                    Rep "  B2 [$cy/$B2Cycles] pres=$presCount Hwnd=$hwndRep ActiveWindow" 'INCONCLUSIVE -- process not alive across the read'
                }
                elseif ($aw2Ex) {
                    $b2Threw++
                    Rep "  B2 [$cy/$B2Cycles] pres=$presCount Hwnd=$hwndRep ActiveWindow" ('throws -- ' + (Root-Type $aw2Ex))
                }
                elseif ($null -eq $aw2) {
                    $b2Null++
                    Rep "  B2 [$cy/$B2Cycles] pres=$presCount Hwnd=$hwndRep ActiveWindow" 'returned $null -- no window object, no exception caught, process alive across the read'
                }
                else {
                    $b2Window++
                    $aw2Hwnd = '<unset>'; $aw2HwndEx = $null
                    try { $aw2Hwnd = $aw2.Hwnd } catch { $aw2HwndEx = $_.Exception }
                    $awHwndRep = $(if ($aw2HwndEx) { 'read threw -- ' + (Root-Type $aw2HwndEx) }
                        elseif ($null -eq $aw2Hwnd) { 'null' } else { "$aw2Hwnd" })
                    Rep "  B2 [$cy/$B2Cycles] pres=$presCount Hwnd=$hwndRep ActiveWindow" "returns a live window object; Hwnd = $awHwndRep"
                }
            }
            catch {
                # One cycle's failure must not lose the others: record it and go on.
                # Type, never message (Office is German) -- Root-Type unwraps the wrap.
                $b2Incon++
                Rep "  B2 [$cy/$B2Cycles]" ("INCONCLUSIVE -- cycle threw: " + (Root-Type $_.Exception))
            }
            finally {
                # Load-bearing (see arm header): Quit() does not reap, so each cycle
                # leaves a live PowerPoint. Runs on the continue and throw paths too.
                # Stop-IsolatedPowerPoint no-ops on $null.
                Stop-IsolatedPowerPoint $cyIso
            }
        }
        Rep "  B2 ActiveWindow tally across instances (null/window/threw/inconclusive)" ("{0}/{1}/{2}/{3} of {4}" -f $b2Null, $b2Window, $b2Threw, $b2Incon, $B2Cycles)
        if ($b2Threw -gt 0 -and $b2Null -eq 0 -and $b2Window -eq 0) {
            Rep "  VERDICT (B)" "ActiveWindow THROWS without a presentation on PowerPoint (matches the Word behaviour); $b2Threw/$B2Cycles fresh instances"
        }
        elseif ($b2Null -gt 0 -and $b2Window -eq 0 -and $b2Threw -eq 0) {
            Rep "  VERDICT (B)" ('ActiveWindow returns $null (no window, no throw caught) without a presentation on PowerPoint; reproduced ' + "$b2Null/$B2Cycles fresh instances")
        }
        elseif ($b2Window -gt 0 -and $b2Null -eq 0 -and $b2Threw -eq 0) {
            Rep "  VERDICT (B)" "ActiveWindow returns a live window (does not throw) without a presentation on PowerPoint; $b2Window/$B2Cycles fresh instances"
        }
        else {
            Rep "  VERDICT (B)" "MIXED / INCONCLUSIVE -- null=$b2Null window=$b2Window threw=$b2Threw inconclusive=$b2Incon of $B2Cycles"
        }
    }
}
catch { Rep "ERROR" $_.Exception.Message.Split([char]10)[0] }
finally {
    # No application-level write was made, so there is nothing to restore. The
    # instance came from CreateProcess, so Stop-IsolatedPowerPoint owns it:
    # graceful Quit first, then Stop-VerifiedPpt (verified pid + StartTime) if it
    # does not reap, then CloseDesktop. Runs on the failure path too.
    Stop-IsolatedPowerPoint $iso
    Remove-Item $root -Recurse -Force -ErrorAction SilentlyContinue
    Rep "POWERPNT pids after" ($(if (Get-PptPids) { (Get-PptPids) -join ',' } else { '(none)' }))
}
