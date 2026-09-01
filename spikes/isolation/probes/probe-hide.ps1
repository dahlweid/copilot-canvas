# Probe: can a live Word window be made invisible to the user (no taskbar, no
# Alt+Tab, off-screen) while still being capturable with PrintWindow?
#
# Throwaway. Runs entirely in TEMP, kills only the Word PID it started.

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot '_common.ps1')

Add-Type -TypeDefinition @"
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

public static class Probe
{
    [DllImport("user32.dll")] public static extern int GetWindowLong(IntPtr h, int i);
    [DllImport("user32.dll")] public static extern int SetWindowLong(IntPtr h, int i, int v);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int c);
    [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr a, int x, int y, int cx, int cy, uint f);
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr dc, uint f);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern int GetWindowThreadProcessId(IntPtr h, out int pid);

    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }

    public const int GWL_EXSTYLE = -20;
    public const int WS_EX_TOOLWINDOW = 0x00000080;
    public const int WS_EX_APPWINDOW  = 0x00040000;

    public static int ExStyle(IntPtr h) { return GetWindowLong(h, GWL_EXSTYLE); }

    // Taskbar/Alt+Tab presence is only re-evaluated when the window is re-shown,
    // so the style change has to be bracketed by hide/show.
    public static int MakeToolWindow(IntPtr h)
    {
        ShowWindow(h, 0); // SW_HIDE
        int ex = GetWindowLong(h, GWL_EXSTYLE);
        ex = (ex | WS_EX_TOOLWINDOW) & ~WS_EX_APPWINDOW;
        SetWindowLong(h, GWL_EXSTYLE, ex);
        ShowWindow(h, 8); // SW_SHOWNOACTIVATE
        return GetWindowLong(h, GWL_EXSTYLE);
    }

    // Returns mean pixel brightness and the count of distinct colours, so a
    // blank/black capture is distinguishable from a real render.
    public static string Capture(IntPtr h, string outPath)
    {
        RECT r; GetWindowRect(h, out r);
        int w = r.R - r.L, ht = r.B - r.T;
        if (w <= 0 || ht <= 0) return "empty-rect";
        using (Bitmap bmp = new Bitmap(w, ht, PixelFormat.Format32bppArgb))
        using (Graphics g = Graphics.FromImage(bmp))
        {
            IntPtr dc = g.GetHdc();
            bool ok;
            try { ok = PrintWindow(h, dc, 2); } finally { g.ReleaseHdc(dc); }

            long sum = 0; var seen = new System.Collections.Generic.HashSet<int>();
            for (int y = 0; y < ht; y += 4)
                for (int x = 0; x < w; x += 4)
                {
                    Color c = bmp.GetPixel(x, y);
                    sum += (c.R + c.G + c.B) / 3;
                    if (seen.Count < 5000) seen.Add(c.ToArgb());
                }
            int n = ((ht + 3) / 4) * ((w + 3) / 4);
            if (outPath != null) bmp.Save(outPath, ImageFormat.Jpeg);
            return string.Format("ok={0} size={1}x{2} brightness={3:F1} colours={4}",
                ok, w, ht, (double)sum / n, seen.Count);
        }
    }
}
"@ -ReferencedAssemblies System.Drawing, System.Drawing.Primitives

function Report($label, $value) { Write-Output ("{0,-38} {1}" -f $label, $value) }

$word = New-Object -ComObject Word.Application
$ownPid = $null
try {
    $word.AutomationSecurity = 3
    $word.Visible = $true
    $doc = $word.Documents.Add()
    $doc.Content.Text = "Probe document. " * 200

    $win = $word.ActiveWindow
    $hwnd = [IntPtr][int64]$win.Hwnd

    # Attribute by hwnd, not by pid differencing. The differencing loop that used
    # to sit here required *exactly one* new WINWORD ($new.Count -eq 1); measured
    # on this machine an external producer mints Words concurrently, $new.Count
    # came back 2, and $ownPid was therefore never assigned. That is not a noisy
    # answer, it is no answer -- and because the whole teardown below is guarded
    # by `if ($ownPid)`, the probe then skipped its own leak check in silence and
    # reported an empty "owned pid". A Word it started could outlive it with
    # nothing said. The hwnd was already being computed on the line above and
    # simply was not joined to the attribution it bore on.
    #
    # Guarded the same way as the other probes: an unresolved hwnd yields pid 0
    # (the Idle process, which never exits) and a recycled pid can name something
    # that is not Word. Either way, refuse to name it rather than act on it.
    #
    # Note what makes this route sound, because it is NOT the reason the
    # CreateProcess probes are sound (#136). Those hold a pid the kernel returned
    # for a process it made for them. This one never consults the WINWORD
    # population at all: the hwnd comes off our own RCW's ActiveWindow, so there
    # is no window in which a stranger's Word could enter the answer. That is the
    # stronger of the two properties, and it is why this repo calls the hwnd route
    # the sound instrument.
    $wp = 0
    [void][Probe]::GetWindowThreadProcessId($hwnd, [ref]$wp)
    $wproc = Get-Process -Id $wp -ErrorAction SilentlyContinue
    if ($wp -gt 4 -and $wproc -and $wproc.ProcessName -eq 'WINWORD') { $ownPid = $wp }
    else { Report "attribution FAILED (teardown cannot verify)" "hwnd resolved to pid $wp ($($wproc.ProcessName))" }
    # Recorded here, while the pid is certainly still the process the hwnd named.
    # Stop-VerifiedWord declines a pid whose StartTime was never captured.
    $ownStart = Get-WordStartTime $ownPid
    Report "owned pid" $ownPid

    Report "ex-style before" ("0x{0:X8}" -f [Probe]::ExStyle($hwnd))
    Report "visible before" ([Probe]::IsWindowVisible($hwnd))

    # 1. Baseline: visible, on-screen.
    [Probe]::SetWindowPos($hwnd, [IntPtr]::Zero, 100, 100, 800, 600, 0x0014) | Out-Null
    Start-Sleep -Milliseconds 600
    Report "capture visible on-screen" ([Probe]::Capture($hwnd, "$env:TEMP\probe-1-onscreen.jpg"))

    # 2. Off-screen (what the spike does today).
    [Probe]::SetWindowPos($hwnd, [IntPtr]::Zero, 30000, 0, 800, 600, 0x0014) | Out-Null
    Start-Sleep -Milliseconds 400
    Report "capture off-screen" ([Probe]::Capture($hwnd, "$env:TEMP\probe-2-offscreen.jpg"))

    # 3. Cross-process WS_EX_TOOLWINDOW: does it stick, and does capture survive?
    $ex = [Probe]::MakeToolWindow($hwnd)
    Start-Sleep -Milliseconds 600
    Report "ex-style after toolwindow" ("0x{0:X8}" -f $ex)
    Report "toolwindow bit set" (($ex -band 0x80) -ne 0)
    Report "appwindow bit cleared" (($ex -band 0x40000) -eq 0)
    Report "visible after" ([Probe]::IsWindowVisible($hwnd))
    Report "capture as toolwindow" ([Probe]::Capture($hwnd, "$env:TEMP\probe-3-toolwindow.jpg"))

    # 4. Does Word still work after we mutated its window styles?
    $doc.Content.InsertAfter("Still alive after style change.")
    Report "word responsive after restyle" ("paragraphs=" + $doc.Paragraphs.Count)

    # 5. Truly hidden: does PrintWindow return anything at all?
    [Probe]::ShowWindow($hwnd, 0) | Out-Null   # SW_HIDE
    Start-Sleep -Milliseconds 400
    Report "visible when hidden" ([Probe]::IsWindowVisible($hwnd))
    Report "capture while SW_HIDE" ([Probe]::Capture($hwnd, "$env:TEMP\probe-4-hidden.jpg"))

    # 6. Application.Visible = false (the production renderer's state).
    [Probe]::ShowWindow($hwnd, 8) | Out-Null
    $word.Visible = $false
    Start-Sleep -Milliseconds 400
    Report "capture with App.Visible=false" ([Probe]::Capture($hwnd, "$env:TEMP\probe-5-appinvisible.jpg"))
}
finally {
    try { $doc.Saved = $true } catch { }
    # Quit(), never Quit(<arg>): under Windows PowerShell 5.1 -- the runtime every
    # .ps1 here runs under -- the argument form throws and the Word survives, and
    # process exit does not reap it either (probe-quit0-leak.ps1). The catch
    # reports rather than swallows; a silent swallow is what hid this for months.
    try { $word.Quit() } catch { Report "Quit() FAILED (Word may leak)" $_.Exception.Message.Split([char]10)[0] }
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch { }
    # Application.Quit() returns long before its process exits (measured 2.7-6.1 s
    # idle, longer under load -- ADR 0005), so the fixed 1.2 s wait made this kill
    # the actual reaper on every run and hid whether Quit() worked at all. Poll to
    # a generous deadline instead; on success that costs only the real exit time.
    # The label states what was observed, never a cause the code cannot know.
    if ($ownPid) {
        $deadline = (Get-Date).AddSeconds(30)
        while ((Get-Date) -lt $deadline -and (Get-Process -Id $ownPid -ErrorAction SilentlyContinue).ProcessName -eq 'WINWORD') {
            Start-Sleep -Milliseconds 250
        }
        if ((Get-Process -Id $ownPid -ErrorAction SilentlyContinue).ProcessName -eq 'WINWORD') {
            # Quit() did not take it. The kill stays -- $ownPid came off our own
            # RCW's window handle, never off a census difference (#136) -- but it
            # is routed through Stop-VerifiedWord, which pins the handle and
            # re-verifies name and StartTime first. That also settles the race
            # this comment used to be about: the pid can exit between the test
            # above and the call, and under $ErrorActionPreference = 'Stop' an
            # unguarded terminate throws inside this finally and silently
            # truncates the teardown report. The helper swallows that race and
            # returns 'gone'. Report the outcome; never label a kill unobserved.
            $outcome = Stop-VerifiedWord $ownPid $ownStart
            Report "STILL ALIVE after 30 s, teardown" "$ownPid -> $outcome"
            if ($outcome -ne 'killed' -and $outcome -ne 'gone') {
                # Not `-like 'declined:*'`: 'killed:survived' means the guards
                # passed and the terminate was issued and the process is STILL
                # THERE, which is the same leak reported as a success.
                Report "pid $ownPid NOT terminated -- this probe's own Word has leaked" 'close it by hand'
            }
        }
        else { Report "exited on Quit()" $ownPid }
    }
}
