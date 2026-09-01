# Probe: can Word run on a SEPARATE DESKTOP OBJECT (the mechanism UAC's secure
# desktop uses) and still be captured?
#
# The decisive question is whether a window on a non-input desktop renders at
# all, or whether capture comes back black because DWM only composes the
# desktop that is currently receiving input.
#
# Throwaway. Runs in TEMP; kills only the Word PID it starts.

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot '_common.ps1')

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
    # before the cleanup below could run.
    try { $w.Quit() } catch { Report "Quit() FAILED (Word may leak)" $_.Exception.Message.Split([char]10)[0] }
    [Runtime.InteropServices.Marshal]::ReleaseComObject($w) | Out-Null
    # $seedPid is NOT sound attribution: $before is snapshotted *after*
    # New-Object, so the instance this block created is usually already in it and
    # $seedPid picks up whatever WINWORD appeared next -- which on this machine is
    # measurably often a Word belonging to a concurrent session (census control:
    # 2 appeared in a 40 s window with nothing launched). This used to
    # Stop-Process -Force that pid, which destroys unsaved work with no prompt.
    # It is now polled and reported, never killed: #25's rule is that a Word which
    # cannot be attributed must not be killed, and a reported orphan beats a
    # destroyed document. The sound instrument is the hwnd route; retrofitting it
    # here would mean this fixture-seed block starts differencing-sweeping, which
    # is the change #25 forbids.
    if ($seedPid) {
        $seedDeadline = (Get-Date).AddSeconds(30)
        while ((Get-Date) -lt $seedDeadline -and (Get-Process -Id $seedPid -ErrorAction SilentlyContinue).ProcessName -eq 'WINWORD') {
            Start-Sleep -Milliseconds 250
        }
        if ((Get-Process -Id $seedPid -ErrorAction SilentlyContinue).ProcessName -eq 'WINWORD') {
            Report "seed-adjacent WINWORD pid $seedPid still alive, NOT killed (unattributable)" 'reported only'
        }
    }
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
# Recorded immediately: Stop-VerifiedWord declines a pid whose StartTime was
# never captured, and a silent decline in teardown is a leak.
$ownStart = Get-WordStartTime $ownPid
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
    # $ownPid comes from CreateProcess, so it is genuinely ours -- a kernel fact
    # about a process made for us, not the census difference the seed block above
    # deliberately refuses to act on (#136). It is nevertheless routed through
    # Stop-VerifiedWord, which pins the handle and re-verifies name and StartTime:
    # the pid can exit between any test and the call, and measured under
    # $ErrorActionPreference = 'Stop' an unguarded terminate throws inside this
    # finally, which would skip CloseDesktop below and leak the desktop handle.
    # Outcome observed rather than asserted.
    $outcome = Stop-VerifiedWord $ownPid $ownStart
    Write-Output "teardown pid ${ownPid}: $outcome"
    if ($outcome -ne 'killed' -and $outcome -ne 'gone') {
        # Not `-like 'declined:*'`: 'killed:survived' means the guards passed and
        # the terminate was issued and the process is STILL THERE, which is the
        # same leak and would have been reported as a success.
        Write-Output "pid $ownPid NOT terminated -- this probe's own Word has leaked; close it by hand"
    }
    Start-Sleep -Milliseconds 500
    [DesktopProbe]::CloseDesktop($desk) | Out-Null
}
