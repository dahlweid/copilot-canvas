# Probe: can a live Word window be made invisible to the user (no taskbar, no
# Alt+Tab, off-screen) while still being capturable with PrintWindow?
#
# Throwaway. Runs entirely in TEMP, kills only the Word PID it started.

$ErrorActionPreference = 'Stop'

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

$before = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$word = New-Object -ComObject Word.Application
$ownPid = $null
try {
    $word.AutomationSecurity = 3
    $word.Visible = $true
    $doc = $word.Documents.Add()
    $doc.Content.Text = "Probe document. " * 200

    for ($i = 0; $i -lt 20 -and -not $ownPid; $i++) {
        $now = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
        $new = @($now | Where-Object { $before -notcontains $_ })
        if ($new.Count -eq 1) { $ownPid = $new[0] } else { Start-Sleep -Milliseconds 100 }
    }
    Report "owned pid" $ownPid

    $win = $word.ActiveWindow
    $hwnd = [IntPtr][int64]$win.Hwnd

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
            Stop-Process -Id $ownPid -Force
            Report "STILL ALIVE after 30 s, killed" $ownPid
        }
        else { Report "exited on Quit()" $ownPid }
    }
}
