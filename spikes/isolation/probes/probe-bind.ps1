# Probe: bind COM to a Word instance we launched OURSELVES on an isolated
# desktop (via AccessibleObjectFromWindow / OBJID_NATIVEOM), then settle the
# open fidelity question: does the page render dark (following the user's Office
# theme) or white (print colours)?
#
# This also validates the technique that lets us attach to a specific Word
# process without ever touching the user's own running Word.

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot '_common.ps1')

Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Text;

public static class Bind
{
    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern IntPtr CreateDesktop(string name, IntPtr dev, IntPtr mode, int flags, uint access, IntPtr sa);
    [DllImport("user32.dll")] public static extern bool CloseDesktop(IntPtr h);
    [DllImport("user32.dll")] public static extern bool EnumDesktopWindows(IntPtr desk, EnumProc cb, IntPtr p);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent, EnumProc cb, IntPtr p);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int max);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr dc, uint f);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr a, int x, int y, int cx, int cy, uint f);
    [DllImport("user32.dll")] public static extern uint GetDpiForWindow(IntPtr h);

    [DllImport("oleacc.dll")]
    public static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint objId, ref Guid iid,
        [MarshalAs(UnmanagedType.IDispatch)] out object obj);

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

    public static IntPtr FindClass(IntPtr desk, uint pid, string wanted)
    {
        IntPtr hit = IntPtr.Zero;
        EnumDesktopWindows(desk, delegate(IntPtr h, IntPtr p) {
            uint got; GetWindowThreadProcessId(h, out got);
            if (got != pid) return true;
            var cls = new StringBuilder(256); GetClassName(h, cls, 256);
            if (cls.ToString() == wanted) { hit = h; return false; }
            return true;
        }, IntPtr.Zero);
        return hit;
    }

    // The document surface is a child of the frame; OBJID_NATIVEOM on it yields
    // a Word.Window, which is the entry point into that instance's object model.
    public static IntPtr FindChildClass(IntPtr parent, string wanted)
    {
        IntPtr hit = IntPtr.Zero;
        EnumChildWindows(parent, delegate(IntPtr h, IntPtr p) {
            var cls = new StringBuilder(256); GetClassName(h, cls, 256);
            if (cls.ToString() == wanted) { hit = h; return false; }
            return true;
        }, IntPtr.Zero);
        return hit;
    }

    public static object NativeOm(IntPtr hwnd)
    {
        Guid iid = new Guid("00020400-0000-0000-C000-000000000046"); // IID_IDispatch
        object obj;
        int hr = AccessibleObjectFromWindow(hwnd, 0xFFFFFFF0, ref iid, out obj); // OBJID_NATIVEOM
        if (hr != 0) throw new COMException("AccessibleObjectFromWindow failed", hr);
        return obj;
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
            // Sample the middle of the document area, well below the ribbon, to
            // judge page colour rather than chrome colour.
            long sum = 0; int n = 0;
            for (int y = ht / 2; y < ht / 2 + 100 && y < ht; y += 2)
                for (int x = w / 2; x < w / 2 + 100 && x < w; x += 2)
                { Color c = bmp.GetPixel(x, y); sum += (c.R + c.G + c.B) / 3; n++; }
            if (outPath != null) bmp.Save(outPath, ImageFormat.Jpeg);
            return string.Format("ok={0} size={1}x{2} pageBrightness={3:F1}", ok, w, ht, n > 0 ? (double)sum / n : -1);
        }
    }
}
"@ -ReferencedAssemblies System.Drawing, System.Drawing.Primitives

function Report($l, $v) { Write-Output ("{0,-32} {1}" -f $l, $v) }

$deskName = "CopilotWordBind"
$desk = [Bind]::CreateDesktop($deskName, [IntPtr]::Zero, [IntPtr]::Zero, 0, 0x10000000, [IntPtr]::Zero)
$exe = "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
$docPath = "$env:TEMP\desktop-probe.docx"

$si = New-Object Bind+STARTUPINFO
$si.cb = [Runtime.InteropServices.Marshal]::SizeOf($si)
$si.desktop = $deskName
$pi = New-Object Bind+PROCESS_INFORMATION

$cmd = '"' + $exe + '" /w /q "' + $docPath + '"'
$ok = [Bind]::CreateProcess([IntPtr]::Zero, $cmd, [IntPtr]::Zero, [IntPtr]::Zero, $false, 0, [IntPtr]::Zero, [IntPtr]::Zero, [ref]$si, [ref]$pi)
Report "launched" "$ok pid=$($pi.pid)"
$ownPid = $pi.pid
# Recorded immediately, while the pid is certainly still this process.
# Stop-VerifiedWord declines without it, which would silently become a leak.
$ownStart = Get-WordStartTime $ownPid

try {
    Start-Sleep -Seconds 12
    $frame = [Bind]::FindClass($desk, [uint32]$ownPid, "OpusApp")
    Report "frame hwnd" $frame
    $docWin = [Bind]::FindChildClass($frame, "_WwG")
    Report "doc surface hwnd" $docWin

    $wdWindow = [Bind]::NativeOm($docWin)
    Report "bound object" $wdWindow.GetType().FullName
    $app = $wdWindow.Application
    Report "app version" $app.Version
    Report "app pid matches" ($app.Documents.Count)
    Report "document" $wdWindow.Document.Name

    [Bind]::SetWindowPos($frame, [IntPtr]::Zero, 0, 0, 900, 1200, 0x0014) | Out-Null
    Start-Sleep -Milliseconds 800
    Report "dpiScale" ([Math]::Round([Bind]::GetDpiForWindow($frame) / 96.0, 2))

    Report "A default view/zoom" ([Bind]::Capture($frame, "$env:TEMP\bind-a-default.jpg"))
    Report "   view.Type" $wdWindow.View.Type
    Report "   zoom" $wdWindow.View.Zoom.Percentage

    # Force print view -- the spike did this, and it is the only view whose
    # colours are supposed to match the printed page.
    $wdWindow.View.Type = 3
    Start-Sleep -Milliseconds 600
    Report "B print view" ([Bind]::Capture($frame, "$env:TEMP\bind-b-printview.jpg"))
    Report "   zoom" $wdWindow.View.Zoom.Percentage

    # Then fit the page the way the spike does.
    $wdWindow.View.Zoom.PageFit = 2
    Start-Sleep -Milliseconds 400
    $fitted = [int]$wdWindow.View.Zoom.Percentage
    $scale = [Bind]::GetDpiForWindow($frame) / 96.0
    $wdWindow.View.Zoom.Percentage = [int][Math]::Round(($fitted / $scale) * 0.92)
    Start-Sleep -Milliseconds 600
    Report "C fitted zoom" ([Bind]::Capture($frame, "$env:TEMP\bind-c-fitted.jpg"))
    Report "   zoom" $wdWindow.View.Zoom.Percentage
}
finally {
    # $ownPid came back from CreateProcess -- a kernel fact about a process made
    # for us, not a census difference (#136) -- so the kill is kept, and routed
    # through Stop-VerifiedWord, which re-verifies name and StartTime against a
    # pinned handle first.
    Report "teardown pid $ownPid" (Stop-VerifiedWord $ownPid $ownStart)
    Start-Sleep -Milliseconds 500
    [Bind]::CloseDesktop($desk) | Out-Null
}
