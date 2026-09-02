# Probe: can PowerPoint run genuinely hidden, and does a hidden instance stay
# alive?
#
# The hidden instance is load-bearing for the whole design: v1's promise is that
# the user's Office never visibly moves. Word honours Application.Visible =
# $false. PowerPoint has historically not, so this measures what is actually
# available and how invisible it really is.
#
#   H1  Application.Visible = msoFalse  -- accepted, or refused?
#   H2  Presentations.Open(WithWindow := msoFalse) -- the documented alternative.
#       Counts top-level windows belonging to our PID, and how many are visible.
#   H3  Application.WindowState = ppWindowMinimized, the fallback if H1 fails.
#
# H1 and H3 are WRITES to the Application object. #139 removed both rather than
# risk landing them on the user's PowerPoint, which New-Object attaches to
# (single-instance, probe-single-instance.ps1). They return here on the
# CreateProcess instance from _isolated.ps1 -- a pid the kernel handed us, not a
# census guess -- and every write is captured first and put back in a finally
# that runs on the failure path too, because "never restored" was half of the
# original defect. H2 and H4 make no application-level write, so they stay on the
# New-Object route they were always measured on.
#
#   H4  SURVIVAL, with an A/B control. A windowless instance has no window whose
#       closing keeps it alive, and PowerPoint is known to shut itself down when
#       it believes it has nothing left to show. Three probes here died on the
#       COM call after an export with "RPC server is unavailable", so this is not
#       hypothetical. Hold an instance idle and poll:
#         A (test)    presentation opened with WithWindow := msoFalse
#         B (control) presentation opened with WithWindow := msoTrue
#       If A dies and B does not, "genuinely hidden" and "stays alive" are in
#       direct conflict and the design has to pick one.
#
#   -IdleSec       how long to hold each survival arm (default 45)
#   -InjectFailure force a throw between the H3 write and its inline restore, so
#                  the only thing that can put the window back is the finally.
#                  This is how the failure-path restore is demonstrated rather
#                  than merely asserted; off by default, the normal measurement
#                  never sets it.

param([string]$Fixture, [int]$IdleSec = 45, [switch]$InjectFailure)

$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot '_common.ps1')
. (Join-Path $PSScriptRoot '_isolated.ps1')

if (-not $Fixture) { $Fixture = Join-Path (Split-Path $PSScriptRoot -Parent) '.fixtures\deck.pptx' }
if (-not (Test-Path $Fixture)) { throw "fixture missing: $Fixture (run make-fixture.ps1 first)" }

Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public class Win {
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr p);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] static extern int GetWindowTextW(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] static extern int GetWindowLongW(IntPtr h, int i);
    delegate bool EnumProc(IntPtr h, IntPtr p);
    public static string[] Windows(uint target) {
        var list = new List<string>();
        EnumWindows((h, p) => {
            uint pid; GetWindowThreadProcessId(h, out pid);
            if (pid != target) return true;
            var sb = new StringBuilder(256); GetWindowTextW(h, sb, 256);
            int ex = GetWindowLongW(h, -20);          // GWL_EXSTYLE
            bool tool = (ex & 0x00000080) != 0;       // WS_EX_TOOLWINDOW
            bool app  = (ex & 0x00040000) != 0;       // WS_EX_APPWINDOW
            list.Add(string.Format("{0}|{1}|tool={2}|app={3}|{4}",
                h.ToInt64(), IsWindowVisible(h) ? "VISIBLE" : "hidden", tool, app, sb.ToString()));
            return true;
        }, IntPtr.Zero);
        return list.ToArray();
    }
}
'@ -ErrorAction SilentlyContinue

function Show-Windows($ppid, $label) {
    $w = @(); try { $w = [Win]::Windows([uint32]$ppid) } catch { }
    $vis = @($w | Where-Object { $_ -match '\|VISIBLE\|' })
    Rep "  [$label] top-level windows / visible" ("{0} / {1}" -f $w.Count, $vis.Count)
    foreach ($v in ($vis | Select-Object -First 4)) { Say "      $v" }
}

$root = Join-Path $env:TEMP ("ppthide-" + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Force -Path $root | Out-Null

Rep "POWERPNT pids before" ($(if (Get-PptPids) { (Get-PptPids) -join ',' } else { '(none)' }))

# --- H1 & H3: application-level writes, on an instance we PROVE we started -----
# The CreateProcess instance from _isolated.ps1 is the only one whose pid the
# code did not have to guess, so it is the only one it is safe to write to. Each
# write is captured before it is made and restored twice: inline for a clean
# measurement of the next arm, and again in the finally so the restore survives a
# throw between the write and its inline restore -- a restore that only runs when
# nothing threw is the #141 defect wearing a finally.
$h13 = $null
$visOrig = $null      # captured Application.Visible; $null once restored or never captured
$wsOrig = $null       # captured Application.WindowState; same convention
try {
    $src13 = Join-Path $root 'h13.pptx'; Copy-Item $Fixture $src13
    $h13 = Start-IsolatedPowerPoint -File $src13
    if (-not $h13.App) {
        Rep "H1/H3 isolated instance" "NOT BOUND -- $($h13.Diag)"
    }
    else {
        $app = $h13.App
        Rep "H1/H3 isolated pid (CreateProcess)" $h13.Pid

        Say "== H1: Application.Visible = msoFalse (isolated instance) =="
        try { $visOrig = $app.Visible } catch { }
        Rep "  Visible (before write)" ($(if ($null -ne $visOrig) { $visOrig } else { 'unreadable' }))
        $h1 = ''
        try { $app.Visible = 0; $h1 = 'accepted' }
        catch {
            $r = $_.Exception; while ($r.InnerException) { $r = $r.InnerException }
            $h1 = 'refused (' + $r.GetType().Name + ')'   # TYPE, not message -- Office is German
        }
        $visAfter = 'unreadable'; try { $visAfter = $app.Visible } catch { }
        Rep "  H1 Application.Visible = 0" $h1
        Rep "  Visible (after write)" $visAfter
        # Inline restore so H3 is measured at the instance's original visibility.
        if ($null -ne $visOrig) {
            try { $app.Visible = $visOrig } catch { }
            $vr = 'unreadable'; try { $vr = $app.Visible } catch { }
            Rep "  Visible (after restore)" $vr
            $visOrig = $null
        }

        Say "== H3: Application.WindowState = ppWindowMinimized (isolated instance) =="
        $wc = 'unreadable'; try { $wc = $app.Windows.Count } catch { }
        Rep "  Application.Windows.Count" $wc
        if (($wc -is [int]) -and $wc -gt 0) {
            try { $wsOrig = $app.WindowState } catch { }
            Rep "  WindowState (before write)" ($(if ($null -ne $wsOrig) { $wsOrig } else { 'unreadable' }))
            $h3 = ''
            try { $app.WindowState = 2; $h3 = 'accepted' }
            catch {
                $r = $_.Exception; while ($r.InnerException) { $r = $r.InnerException }
                $h3 = 'refused (' + $r.GetType().Name + ')'
            }
            $wsAfter = 'unreadable'; try { $wsAfter = $app.WindowState } catch { }
            Rep "  H3 WindowState = ppWindowMinimized" $h3
            Rep "  WindowState (after write)" $wsAfter
            # Failure-path demonstration: throw AFTER the write has landed but
            # BEFORE the inline restore below, so the only thing that can put the
            # window back is the finally. Off unless -InjectFailure is passed.
            if ($InjectFailure) { throw 'injected failure (test): thrown between the H3 write and its inline restore' }
            if ($null -ne $wsOrig) {
                try { $app.WindowState = $wsOrig } catch { }
                $wr = 'unreadable'; try { $wr = $app.WindowState } catch { }
                Rep "  WindowState (after restore)" $wr
                $wsOrig = $null
            }
        }
        else {
            Rep "  H3 WindowState = ppWindowMinimized" 'no window to minimise (Windows.Count = 0)'
        }
    }
}
catch { Rep "ERROR (H1/H3 isolated)" $_.Exception.Message.Split([char]10)[0] }
finally {
    # The restore that cannot be skipped. If the try threw after a write but
    # before its inline restore, $visOrig / $wsOrig are still set and are put back
    # here; if the inline restore already ran they are $null and this is a no-op.
    # The read-back is taken while the instance is still alive, before teardown,
    # so a failure-path restore leaves visible proof and is not merely asserted.
    if ($h13 -and $h13.App) {
        $restored = $false
        if ($null -ne $visOrig) {
            try { $h13.App.Visible = $visOrig } catch { }
            $vb = 'unreadable'; try { $vb = $h13.App.Visible } catch { }
            Rep "  [finally] Visible restored to" ("{0} (read-back {1})" -f $visOrig, $vb)
            $restored = $true
        }
        if ($null -ne $wsOrig) {
            try { $h13.App.WindowState = $wsOrig } catch { }
            $wb = 'unreadable'; try { $wb = $h13.App.WindowState } catch { }
            Rep "  [finally] WindowState restored to" ("{0} (read-back {1})" -f $wsOrig, $wb)
            $restored = $true
        }
        if (-not $restored) { Say "  [finally] nothing outstanding -- the inline restores already ran" }
    }
    Stop-IsolatedPowerPoint $h13
}

# --- H2: Presentations.Open(WithWindow := msoFalse) ---------------------------
# The documented windowless-open alternative. It makes no application-level
# write -- it opens a deck from our own temp copy and closes it in the finally --
# so it stays on the New-Object route it was always measured on.
$ctx = $null
$pres = $null
try {
    $ctx = New-PowerPointInstance
    $app = $ctx.App
    $ppid = $ctx.NewPids | Select-Object -First 1
    Rep "new POWERPNT pids seen" ($(if ($ppid) { $ppid } else { '(none appeared - attached)' }))

    Say "== H2: Presentations.Open(WithWindow := msoFalse) =="
    $srcH2 = Join-Path $root 'h2.pptx'; Copy-Item $Fixture $srcH2
    $pres = $app.Presentations.Open($srcH2, 0, 0, 0)
    Start-Sleep -Milliseconds 700
    Rep "  opened, slides" $pres.Slides.Count
    try { Rep "  Presentation.Windows.Count" $pres.Windows.Count } catch { Rep "  Presentation.Windows.Count" 'n/a' }
    try { Rep "  Application.Windows.Count" $app.Windows.Count } catch { Rep "  Application.Windows.Count" 'n/a' }
    if ($ppid) { Show-Windows $ppid 'H2 windowless open' }
}
catch { Rep "ERROR (H2)" $_.Exception.Message }
finally {
    # Ours: opened from our own temp root above. Close it here so a failure cannot
    # leave it open in an instance we merely attached to.
    try { if ($pres) { $pres.Saved = -1; $pres.Close() } } catch { }
    $pres = $null
    Close-PowerPointInstance $ctx
}

# --- H4: survival A/B ---------------------------------------------------------
function Invoke-SurvivalArm([string]$Label, [int]$WithWindow) {
    $src = Join-Path $root "$Label.pptx"; Copy-Item $Fixture $src
    $ctx = $null
    $pres = $null
    try {
        $ctx = New-PowerPointInstance
        $ppid = $ctx.NewPids | Select-Object -First 1
        $pres = $ctx.App.Presentations.Open($src, 0, 0, $WithWindow)
        $diedAt = $null
        $sw = [Diagnostics.Stopwatch]::StartNew()
        while ($sw.Elapsed.TotalSeconds -lt $IdleSec) {
            Start-Sleep -Seconds 3
            try { $null = $pres.Slides.Count }
            catch { $diedAt = [int]$sw.Elapsed.TotalSeconds; break }
            if ($ppid -and -not (Get-Process -Id $ppid -ErrorAction SilentlyContinue)) {
                $diedAt = [int]$sw.Elapsed.TotalSeconds; break
            }
        }
        $sw.Stop()
        if ($null -ne $diedAt) { Rep "  [$Label] DIED after" "$diedAt s idle" }
        else { Rep "  [$Label] survived" "$IdleSec s idle, handle still valid" }
        try { $pres.Saved = -1; $pres.Close(); $pres = $null } catch { }
    }
    catch { Rep "  [$Label] ERROR" $_.Exception.Message.Split([char]10)[0] }
    finally {
        try { if ($pres) { $pres.Saved = -1; $pres.Close() } } catch { }
        $pres = $null
        Close-PowerPointInstance $ctx
    }
}

Say "== H4: does a windowless instance stay alive while idle? =="
Say "  A (test):    WithWindow = msoFalse"
Invoke-SurvivalArm 'A-windowless' 0
Say "  B (control): WithWindow = msoTrue"
Invoke-SurvivalArm 'B-windowed' -1

Remove-Item $root -Recurse -Force -ErrorAction SilentlyContinue
Rep "POWERPNT pids after" ($(if (Get-PptPids) { (Get-PptPids) -join ',' } else { '(none)' }))
