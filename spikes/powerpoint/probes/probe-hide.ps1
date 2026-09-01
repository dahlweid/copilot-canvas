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
#   -IdleSec  how long to hold each survival arm (default 45)

param([string]$Fixture, [int]$IdleSec = 45)

$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot '_common.ps1')

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

# --- H1..H3 -------------------------------------------------------------------
$ctx = $null
$pres = $null
try {
    $ctx = New-PowerPointInstance
    $app = $ctx.App
    $ppid = $ctx.NewPids | Select-Object -First 1
    Rep "new POWERPNT pids seen" ($(if ($ppid) { $ppid } else { '(none appeared - attached)' }))

    Say "== H1: Application.Visible = msoFalse =="
    # NOT MEASURED ANY MORE, and the reason is the point of issue #139.
    #
    # This arm used to assign $app.Visible = 0. That is a WRITE to the
    # Application object, and New-Object attaches to a running PowerPoint, so on
    # any machine where the user has PowerPoint open the write lands on THEIR
    # instance -- and nothing here ever set it back. The probe knew: :81 above
    # prints "(none appeared - attached)" and the teardown declines to kill on
    # exactly that basis. It then wrote to the instance anyway.
    #
    # There is no gate that fixes this. Gating on the census difference being
    # non-empty is the same unsound inference the kills were removed for, in the
    # other direction: a non-empty difference is evidence of nothing (it
    # over-reports 2-for-1, probe-init-attribution.ps1) and is non-empty by
    # construction in the race where somebody else starts PowerPoint mid-census.
    #
    # Measuring this properly needs an instance we can PROVE we started, which
    # means the CreateProcess route in _isolated.ps1 rather than New-Object.
    # That is a redesign of this probe, not a change to it.
    Rep "  H1 Application.Visible = 0" 'NOT MEASURED - would write to a possibly-attached instance (#139)'
    try { Rep "  Application.Visible reads" $app.Visible } catch { Rep "  Application.Visible reads" 'unreadable' }
    if ($ppid) { Show-Windows $ppid 'H1 no presentation' }

    Say "== H2: Presentations.Open(WithWindow := msoFalse) =="
    $src = Join-Path $root 'h2.pptx'; Copy-Item $Fixture $src
    $pres = $app.Presentations.Open($src, 0, 0, 0)
    Start-Sleep -Milliseconds 700
    Rep "  opened, slides" $pres.Slides.Count
    try { Rep "  Presentation.Windows.Count" $pres.Windows.Count } catch { Rep "  Presentation.Windows.Count" 'n/a' }
    try { Rep "  Application.Windows.Count" $app.Windows.Count } catch { Rep "  Application.Windows.Count" 'n/a' }
    if ($ppid) { Show-Windows $ppid 'H2 windowless open' }

    Say "== H3: WindowState fallback (only meaningful if a window exists) =="
    # Same reason as H1: WindowState = 2 minimises the Application, and the
    # instance may be the user's. Worse than H1, in fact -- its old guard was
    # `$app.Windows.Count -gt 0`, which is true precisely when somebody has a
    # deck open. Not measured (#139); H2 above already reports the window count.
    Rep "  H3 WindowState = ppWindowMinimized" 'NOT MEASURED - would write to a possibly-attached instance (#139)'
}
catch { Rep "ERROR (H1-H3)" $_.Exception.Message }
finally {
    # Ours: opened from our own temp root at H2. Close it here so a failure
    # above cannot leave it open in an instance we merely attached to.
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
