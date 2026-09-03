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
#      open (B2). B2 is the PowerPoint analogue of the Word arm above.
#
# Scope note, so this is not over-read: attribution is NOT exclusivity. Even if A
# holds, it does not make writing to a New-Object instance safe, because
# probe-single-instance.ps1 measures PowerPoint as single-instance -- a New-Object
# instance is the user's whenever they have one open, whatever its pid is. So this
# probe informs a shipped comment and issue #141's "Note", not the choice of the
# CreateProcess route for H1/H3.
#
# SAFETY: neither arm writes to the Application object, so there is nothing to
# restore. The only PowerPoint touched is one obtained through
# Start-IsolatedPowerPoint (a CreateProcess pid), and it is torn down on every
# path in a finally via Stop-IsolatedPowerPoint. Exceptions are reported by
# .NET TYPE, never by message -- this Office is German and the message is
# localized.

param([string]$Fixture)

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
            # catch above, and $Error growth is not a reliable detector either
            # (the adapter can convert a throw to $null without $Error growing).
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
        Say "== B2: Application.ActiveWindow (and Hwnd) with NO deck open =="
        # Close the only presentation so the instance is genuinely windowless.
        # Closing the last deck MAY make PowerPoint exit (FINDINGS Q4 measured idle
        # survival only for instances that still had a deck open), so the process
        # is checked for life first, and "ActiveWindow threw because there is no
        # window" is reported apart from "the RCW is dead because the process
        # exited" -- distinguished by whether the pid is still alive.
        try { foreach ($p in @($app.Presentations)) { try { $p.Saved = -1 } catch { }; $p.Close() } } catch { }
        Start-Sleep -Milliseconds 900
        # Distinguish "the Presentations collection is null" from "the collection
        # holds 0 decks": $null.Count is 0 in this shell (no Set-StrictMode under
        # spikes/powerpoint/), so a bare .Count of 0 cannot tell them apart and is
        # not on its own evidence that the instance is deckless.
        $presObj = '<unset>'; try { $presObj = $app.Presentations } catch { }
        $presCount = $(if ($presObj -eq '<unset>') { 'unreadable (read threw)' }
            elseif ($null -eq $presObj) { 'Presentations object is null' }
            else { $c = 'unreadable'; try { $c = $presObj.Count } catch { }; "$c" })
        $alive = [bool](Get-Process -Id $iso.Pid -ErrorAction SilentlyContinue)
        Rep "  after closing the deck: process alive" $alive
        Rep "  Presentations.Count" $presCount
        if ($alive) {
            $hwnd2 = $null; $hwnd2Ex = $null
            try { $hwnd2 = $app.Hwnd } catch { $hwnd2Ex = $_.Exception }
            if ($hwnd2Ex) { Rep "  B2 Application.Hwnd (no deck)" ("throws -- " + (Root-Type $hwnd2Ex)) }
            elseif ($null -eq $hwnd2) { Rep "  B2 Application.Hwnd (no deck)" 'null through the bind' }
            else { Rep "  B2 Application.Hwnd (no deck)" "$hwnd2 (Application.Hwnd is presentation-independent)" }

            # Three outcomes, kept apart -- the whole point of arm B. A bare
            # try/catch would fold "returned $null" into the success branch,
            # because $null.Hwnd does not throw in this shell, so the verdict would
            # print "returns a window" for a $null just as readily as for a real
            # one. Separate: threw (caught below), returned $null, returned a
            # window object. "no exception caught" is stated rather than "no
            # exception", for the same EAP='Continue' reason as arm A.
            $aw2 = $null; $aw2Ex = $null
            try { $aw2 = $app.ActiveWindow } catch { $aw2Ex = $_.Exception }
            if ($aw2Ex) {
                Rep "  B2 ActiveWindow (no deck)" ("throws -- " + (Root-Type $aw2Ex))
                Rep "  VERDICT (B)" 'ActiveWindow THROWS without a presentation on PowerPoint (matches the Word behaviour)'
            }
            elseif ($null -eq $aw2) {
                Rep "  B2 ActiveWindow (no deck)" 'returned $null -- no window object, and no exception caught'
                Rep "  VERDICT (B)" 'ActiveWindow returns $null (no window, no throw caught) without a presentation on PowerPoint'
            }
            else {
                $aw2Hwnd = '<unset>'; $aw2HwndEx = $null
                try { $aw2Hwnd = $aw2.Hwnd } catch { $aw2HwndEx = $_.Exception }
                $hwndRep = $(if ($aw2HwndEx) { 'read threw -- ' + (Root-Type $aw2HwndEx) }
                    elseif ($null -eq $aw2Hwnd) { 'null' } else { "$aw2Hwnd" })
                Rep "  B2 ActiveWindow (no deck)" "returns a live window object; Hwnd = $hwndRep"
                Rep "  VERDICT (B)" 'ActiveWindow returns a live window (does not throw) without a presentation on PowerPoint'
            }
        }
        else {
            Rep "  VERDICT (B)" 'INCONCLUSIVE -- closing the last deck exited the process, so a live windowless read was not possible'
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
