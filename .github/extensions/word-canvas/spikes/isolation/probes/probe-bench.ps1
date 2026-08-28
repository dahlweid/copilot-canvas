# Combined probe on an isolated desktop object. Answers, in one run:
#   1. Can we force print colours (white page) despite the user's dark Office theme?
#   2. What frame rate does PrintWindow sustain on a NON-INPUT desktop?
#   3. Does COM Documents.Open work across the desktop boundary (DDE hang risk)?
#   4. What is the scroll -> frame latency via COM?

$ErrorActionPreference = 'Stop'

Add-Type -TypeDefinition @"
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Text;

public static class P2
{
    [DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
    public static extern IntPtr CreateDesktop(string n, IntPtr d, IntPtr m, int f, uint a, IntPtr s);
    [DllImport("user32.dll")] public static extern bool CloseDesktop(IntPtr h);
    [DllImport("user32.dll")] public static extern bool EnumDesktopWindows(IntPtr d, EnumProc cb, IntPtr p);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int m);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr dc, uint f);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr a, int x, int y, int cx, int cy, uint f);
    [DllImport("oleacc.dll")] public static extern int AccessibleObjectFromWindow(IntPtr h, uint o, ref Guid i,
        [MarshalAs(UnmanagedType.IDispatch)] out object obj);
    [DllImport("kernel32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
    public static extern bool CreateProcess(IntPtr a, string c, IntPtr pa, IntPtr ta, bool inh, uint f,
        IntPtr env, IntPtr dir, ref STARTUPINFO si, out PROCESS_INFORMATION pi);

    public delegate bool EnumProc(IntPtr h, IntPtr p);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L,T,R,B; }
    [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)]
    public struct STARTUPINFO { public int cb; public string r1, desktop, title;
        public int x,y,xs,ys,xc,yc,fill,flags; public short show, cb2; public IntPtr r2, si, so, se; }
    [StructLayout(LayoutKind.Sequential)]
    public struct PROCESS_INFORMATION { public IntPtr hP, hT; public int pid, tid; }

    public static IntPtr Find(IntPtr desk, uint pid, string want) {
        IntPtr hit = IntPtr.Zero;
        EnumDesktopWindows(desk, delegate(IntPtr h, IntPtr p) {
            uint g; GetWindowThreadProcessId(h, out g); if (g != pid) return true;
            var c = new StringBuilder(256); GetClassName(h, c, 256);
            if (c.ToString() == want) { hit = h; return false; } return true;
        }, IntPtr.Zero);
        return hit;
    }
    public static IntPtr FindChild(IntPtr parent, string want) {
        IntPtr hit = IntPtr.Zero;
        EnumChildWindows(parent, delegate(IntPtr h, IntPtr p) {
            var c = new StringBuilder(256); GetClassName(h, c, 256);
            if (c.ToString() == want) { hit = h; return false; } return true;
        }, IntPtr.Zero);
        return hit;
    }
    public static object Om(IntPtr h) {
        Guid i = new Guid("00020400-0000-0000-C000-000000000046");
        object o; int hr = AccessibleObjectFromWindow(h, 0xFFFFFFF0, ref i, out o);
        if (hr != 0) throw new COMException("AOFW", hr);
        return o;
    }

    // Capture the DOCUMENT child window only -- that is what a viewer would show,
    // and it excludes ribbon chrome from the colour measurement.
    public static double[] Shot(IntPtr h, string path) {
        RECT r; GetWindowRect(h, out r);
        int w = r.R-r.L, ht = r.B-r.T;
        if (w <= 0 || ht <= 0) return new double[]{ -1, -1 };
        using (Bitmap b = new Bitmap(w, ht, PixelFormat.Format32bppArgb))
        using (Graphics g = Graphics.FromImage(b)) {
            IntPtr dc = g.GetHdc();
            try { PrintWindow(h, dc, 2); } finally { g.ReleaseHdc(dc); }
            long sum = 0; int n = 0; var seen = new System.Collections.Generic.HashSet<int>();
            for (int y = 0; y < ht; y += 7) for (int x = 0; x < w; x += 7) {
                Color c = b.GetPixel(x, y); sum += (c.R+c.G+c.B)/3; n++; seen.Add(c.ToArgb());
            }
            if (path != null) b.Save(path, ImageFormat.Jpeg);
            return new double[]{ n > 0 ? (double)sum/n : -1, seen.Count };
        }
    }
    public static double Bench(IntPtr h, int frames) {
        RECT r; GetWindowRect(h, out r);
        int w = r.R-r.L, ht = r.B-r.T;
        var sw = System.Diagnostics.Stopwatch.StartNew();
        for (int i = 0; i < frames; i++) {
            using (Bitmap b = new Bitmap(w, ht, PixelFormat.Format32bppArgb))
            using (Graphics g = Graphics.FromImage(b)) {
                IntPtr dc = g.GetHdc();
                try { PrintWindow(h, dc, 2); } finally { g.ReleaseHdc(dc); }
                using (var ms = new System.IO.MemoryStream()) b.Save(ms, ImageFormat.Jpeg);
            }
        }
        sw.Stop();
        return sw.Elapsed.TotalMilliseconds / frames;
    }
}
"@ -ReferencedAssemblies System.Drawing, System.Drawing.Primitives

function Rep($l, $v) { Write-Output ("{0,-30} {1}" -f $l, $v) }

$themeKey = 'HKCU:\Software\Microsoft\Office\16.0\Common'
$origTheme = (Get-ItemProperty -Path $themeKey -Name 'UI Theme' -ErrorAction SilentlyContinue).'UI Theme'
Rep "user UI Theme (before)" $origTheme

$deskName = "CopilotWordBench"
$desk = [P2]::CreateDesktop($deskName, [IntPtr]::Zero, [IntPtr]::Zero, 0, 0x10000000, [IntPtr]::Zero)
$exe  = "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
$doc  = "$env:TEMP\desktop-probe.docx"
$ownPid = 0

try {
    # Word reads the Office theme at process start. Flip it to White (5) just for
    # the launch, then put the user's value straight back.
    Set-ItemProperty -Path $themeKey -Name 'UI Theme' -Value 5 -Type DWord
    Rep "UI Theme forced to" 5

    $si = New-Object P2+STARTUPINFO
    $si.cb = [Runtime.InteropServices.Marshal]::SizeOf($si)
    $si.desktop = $deskName
    $pi = New-Object P2+PROCESS_INFORMATION
    # Launch with NO document -- we want to test Documents.Open across desktops.
    $cmd = '"' + $exe + '" /w /q'
    $ok = [P2]::CreateProcess([IntPtr]::Zero, $cmd, [IntPtr]::Zero, [IntPtr]::Zero, $false, 0, [IntPtr]::Zero, [IntPtr]::Zero, [ref]$si, [ref]$pi)
    $ownPid = $pi.pid
    Rep "launched (no doc)" "$ok pid=$ownPid"

    Start-Sleep -Seconds 3
    Set-ItemProperty -Path $themeKey -Name 'UI Theme' -Value $origTheme -Type DWord
    Rep "UI Theme restored to" $origTheme

    Start-Sleep -Seconds 8
    $frame = [P2]::Find($desk, [uint32]$ownPid, "OpusApp")
    Rep "frame hwnd" $frame
    $wwg = [P2]::FindChild($frame, "_WwG")
    Rep "doc surface hwnd" $wwg

    $win = [P2]::Om($wwg)
    $app = $win.Application
    Rep "COM bound" $app.Version

    # --- Q3: does Documents.Open work across the desktop boundary? ---
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $app.AutomationSecurity = 3
    $opened = $app.Documents.Open($doc, $false, $true, $false)
    $sw.Stop()
    Rep "Documents.Open cross-desktop" ("OK in {0} ms -> {1}" -f [int]$sw.ElapsedMilliseconds, $opened.Name)
    Rep "page count" $opened.Content.Information(4)

    $dwin = $app.ActiveWindow
    $dwin.View.Type = 3
    [P2]::SetWindowPos($frame, [IntPtr]::Zero, 0, 0, 900, 1200, 0x0014) | Out-Null
    Start-Sleep -Milliseconds 900
    $wwg2 = [P2]::FindChild($frame, "_WwG")

    # --- Q1: page colour with the light theme forced at launch ---
    $s = [P2]::Shot($wwg2, "$env:TEMP\bench-page.jpg")
    Rep "PAGE brightness / colours" ("{0:F1} / {1}" -f $s[0], $s[1])

    # --- Q2: sustained capture rate on a non-input desktop ---
    Rep "capture mean ms (30 frames)" ("{0:F1}" -f [P2]::Bench($wwg2, 30))

    # --- Q4: scroll -> frame latency through COM ---
    $lat = @()
    for ($i = 0; $i -lt 5; $i++) {
        $t = [Diagnostics.Stopwatch]::StartNew()
        $dwin.SmallScroll(3, 0, 0, 0)
        [P2]::Shot($wwg2, $null) | Out-Null
        $t.Stop(); $lat += $t.Elapsed.TotalMilliseconds
    }
    Rep "scroll->frame mean ms" ("{0:F1}" -f (($lat | Measure-Object -Average).Average))

    $opened.Close(0)
}
catch { R "ERROR" $_.Exception.Message }
finally {
    if ($origTheme -ne $null) { Set-ItemProperty -Path $themeKey -Name 'UI Theme' -Value $origTheme -Type DWord }
    if ($ownPid -and (Get-Process -Id $ownPid -ErrorAction SilentlyContinue)) { Stop-Process -Id $ownPid -Force }
    Start-Sleep -Milliseconds 600
    [P2]::CloseDesktop($desk) | Out-Null
    Rep "cleanup" "done; theme=$((Get-ItemProperty -Path $themeKey -Name 'UI Theme').'UI Theme')"
}


