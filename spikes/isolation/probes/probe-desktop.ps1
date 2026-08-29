# Probe: can Word run on a SEPARATE DESKTOP OBJECT (the mechanism UAC's secure
# desktop uses) and still be captured?
#
# The decisive question is whether a window on a non-input desktop renders at
# all, or whether capture comes back black because DWM only composes the
# desktop that is currently receiving input.
#
# Throwaway. Runs in TEMP; kills only the Word PID it starts.

$ErrorActionPreference = 'Stop'

Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Text;

public static class DesktopProbe
{
    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern IntPtr CreateDesktop(string name, IntPtr dev, IntPtr mode, int flags, uint access, IntPtr sa);
    [DllImport("user32.dll", SetLastError = true)] public static extern bool CloseDesktop(IntPtr h);
    [DllImport("user32.dll", SetLastError = true)] public static extern bool SetThreadDesktop(IntPtr h);
    [DllImport("user32.dll", SetLastError = true)] public static extern IntPtr GetThreadDesktop(int threadId);
    [DllImport("kernel32.dll")] public static extern int GetCurrentThreadId();
    [DllImport("user32.dll")] public static extern bool EnumDesktopWindows(IntPtr desk, EnumProc cb, IntPtr p);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int max);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int max);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr dc, uint f);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr a, int x, int y, int cx, int cy, uint f);

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern bool CreateProcess(IntPtr app, string cmd, IntPtr pa, IntPtr ta, bool inherit,
        uint flags, IntPtr env, IntPtr dir, ref STARTUPINFO si, out PROCESS_INFORMATION pi);

    public delegate bool EnumProc(IntPtr hwnd, IntPtr param);

    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct STARTUPINFO
    {
        public int cb; public string reserved; public string desktop; public string title;
        public int x, y, xSize, ySize, xCount, yCount, fill, flags;
        public short showWindow, cbReserved2; public IntPtr reserved2, hStdIn, hStdOut, hStdErr;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct PROCESS_INFORMATION { public IntPtr hProcess, hThread; public int pid, tid; }

    public const uint GENERIC_ALL = 0x10000000;

    public static List<IntPtr> WindowsOn(IntPtr desk, uint wantPid)
    {
        var found = new List<IntPtr>();
        EnumDesktopWindows(desk, delegate(IntPtr h, IntPtr p) {
            uint pid; GetWindowThreadProcessId(h, out pid);
            if (pid == wantPid) found.Add(h);
            return true;
        }, IntPtr.Zero);
        return found;
    }

    public static string Describe(IntPtr h)
    {
        var cls = new StringBuilder(256); GetClassName(h, cls, 256);
        var txt = new StringBuilder(256); GetWindowText(h, txt, 256);
        RECT r; GetWindowRect(h, out r);
        return string.Format("class={0} visible={1} rect={2}x{3} title='{4}'",
            cls, IsWindowVisible(h), r.R - r.L, r.B - r.T, txt);
    }

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
            long sum = 0; var seen = new HashSet<int>();
            for (int y = 0; y < ht; y += 4)
                for (int x = 0; x < w; x += 4)
                { Color c = bmp.GetPixel(x, y); sum += (c.R + c.G + c.B) / 3; if (seen.Count < 5000) seen.Add(c.ToArgb()); }
            int n = ((ht + 3) / 4) * ((w + 3) / 4);
            if (outPath != null) bmp.Save(outPath, ImageFormat.Jpeg);
            return string.Format("ok={0} size={1}x{2} brightness={3:F1} colours={4}", ok, w, ht, (double)sum / n, seen.Count);
        }
    }
}
"@ -ReferencedAssemblies System.Drawing, System.Drawing.Primitives

function Report($label, $value) { Write-Output ("{0,-34} {1}" -f $label, $value) }

$deskName = "CopilotWordRender"
$desk = [DesktopProbe]::CreateDesktop($deskName, [IntPtr]::Zero, [IntPtr]::Zero, 0, [DesktopProbe]::GENERIC_ALL, [IntPtr]::Zero)
Report "CreateDesktop handle" $desk
if ($desk -eq [IntPtr]::Zero) { throw "CreateDesktop failed: $([ComponentModel.Win32Exception]::new([Runtime.InteropServices.Marshal]::GetLastWin32Error()).Message)" }

$wordExe = @(
    "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
    "C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE"
) | Where-Object { Test-Path $_ } | Select-Object -First 1
Report "winword.exe" $wordExe

$docPath = Join-Path $env:TEMP "desktop-probe.docx"
if (-not (Test-Path $docPath)) {
    $w = New-Object -ComObject Word.Application
    $before = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $w.Visible = $false
    $d = $w.Documents.Add()
    $d.Content.Text = ("Desktop probe paragraph. " * 60)
    $d.SaveAs2($docPath, 16)
    $d.Close(0)
    $seedPid = (Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | Select-Object -First 1).Id
    # Quit(), never Quit(<arg>): under Windows PowerShell 5.1 -- the runtime every
    # .ps1 here runs under -- the argument form throws and the Word survives, and
    # process exit does not reap it either (probe-quit0-leak.ps1). Unswallowed,
    # this threw under $ErrorActionPreference = 'Stop' and aborted the probe
    # before the sweep below could run.
    #
    # That sweep is also not a safety net at this site: $before is snapshotted
    # *after* New-Object, so $seedPid is not something this probe can rely on.
    # Sound attribution is the hwnd route (#25); it is deliberately not
    # retrofitted here, because a differencing sweep that starts firing could
    # kill a Word belonging to another session.
    try { $w.Quit() } catch { Report "Quit() FAILED (Word may leak)" $_.Exception.Message.Split([char]10)[0] }
    [Runtime.InteropServices.Marshal]::ReleaseComObject($w) | Out-Null
    Start-Sleep -Milliseconds 1000
    if ($seedPid -and (Get-Process -Id $seedPid -ErrorAction SilentlyContinue)) { Stop-Process -Id $seedPid -Force }
}
Report "fixture" (Test-Path $docPath)

$si = New-Object DesktopProbe+STARTUPINFO
$si.cb = [Runtime.InteropServices.Marshal]::SizeOf($si)
$si.desktop = $deskName
$pi = New-Object DesktopProbe+PROCESS_INFORMATION

$cmd = '"' + $wordExe + '" /w /q "' + $docPath + '"'
$ok = [DesktopProbe]::CreateProcess([IntPtr]::Zero, $cmd, [IntPtr]::Zero, [IntPtr]::Zero, $false, 0, [IntPtr]::Zero, [IntPtr]::Zero, [ref]$si, [ref]$pi)
$err = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
Report "CreateProcess on desktop" "$ok pid=$($pi.pid) err=$err ($([ComponentModel.Win32Exception]::new($err).Message))"
Report "STARTUPINFO cb" $si.cb
Report "desktop string" $si.desktop
if (-not $ok) { [DesktopProbe]::CloseDesktop($desk) | Out-Null; throw "CreateProcess failed" }

$ownPid = $pi.pid
try {
    Start-Sleep -Seconds 12

    $wins = [DesktopProbe]::WindowsOn($desk, [uint32]$ownPid)
    Report "windows on that desktop" $wins.Count

    $target = $null
    foreach ($h in $wins) {
        $desc = [DesktopProbe]::Describe($h)
        Write-Output "    $desc"
        if ($desc -match 'class=OpusApp') { $target = $h }
    }

    if ($target) {
        [DesktopProbe]::SetWindowPos($target, [IntPtr]::Zero, 0, 0, 900, 700, 0x0014) | Out-Null
        Start-Sleep -Milliseconds 800
        Report "CAPTURE on alt desktop" ([DesktopProbe]::Capture($target, "$env:TEMP\probe-desktop.jpg"))
    }
    else {
        Report "CAPTURE on alt desktop" "no OpusApp window found"
    }
}
finally {
    if ($ownPid -and (Get-Process -Id $ownPid -ErrorAction SilentlyContinue)) {
        Stop-Process -Id $ownPid -Force
        Write-Output "killed pid $ownPid"
    }
    Start-Sleep -Milliseconds 500
    [DesktopProbe]::CloseDesktop($desk) | Out-Null
}
