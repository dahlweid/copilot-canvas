# Probe: does Window.RangeFromPoint / GetPoint work for a Word window living on a
# NON-INPUT desktop object? These two calls are the linchpin of the read-write
# design (click -> caret, and caret -> overlay), and the docs say RangeFromPoint
# needs screen coordinates of a currently-rendered region. Untested on a desktop
# that is never switched to.
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot '_common.ps1')

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class P3
{
    [DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
    public static extern IntPtr CreateDesktop(string n, IntPtr d, IntPtr m, int f, uint a, IntPtr s);
    [DllImport("user32.dll")] public static extern bool CloseDesktop(IntPtr h);
    [DllImport("user32.dll")] public static extern bool EnumDesktopWindows(IntPtr d, EnumProc cb, IntPtr p);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int m);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
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
}
"@

function Rep($l, $v) { Write-Output ("{0,-34} {1}" -f $l, $v) }

$deskName = "CopilotWordHit"
$desk = [P3]::CreateDesktop($deskName, [IntPtr]::Zero, [IntPtr]::Zero, 0, 0x10000000, [IntPtr]::Zero)
$exe = "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
$doc = "$env:TEMP\desktop-probe.docx"
$ownPid = 0
try {
    $si = New-Object P3+STARTUPINFO
    $si.cb = [Runtime.InteropServices.Marshal]::SizeOf($si)
    $si.desktop = $deskName
    $pi = New-Object P3+PROCESS_INFORMATION
    $cmd = '"' + $exe + '" /w /q "' + $doc + '"'
    [P3]::CreateProcess([IntPtr]::Zero, $cmd, [IntPtr]::Zero, [IntPtr]::Zero, $false, 0, [IntPtr]::Zero, [IntPtr]::Zero, [ref]$si, [ref]$pi) | Out-Null
    $ownPid = $pi.pid
    # Recorded immediately: Stop-VerifiedWord declines a pid whose StartTime was
    # never captured, and a silent decline in teardown is a leak.
    $ownStart = Get-WordStartTime $ownPid
    Rep "launched pid" $ownPid
    Start-Sleep -Seconds 12

    $frame = [P3]::Find($desk, [uint32]$ownPid, "OpusApp")
    # Park at a normal on-desktop position; this desktop has its own coordinate
    # space and nothing else is on it.
    [P3]::SetWindowPos($frame, [IntPtr]::Zero, 0, 0, 1000, 1200, 0x0014) | Out-Null
    Start-Sleep -Milliseconds 900
    $wwg = [P3]::FindChild($frame, "_WwG")
    $win = [P3]::Om($wwg)
    $app = $win.Application
    $dwin = $app.ActiveWindow
    $dwin.View.Type = 3
    Start-Sleep -Milliseconds 500

    $r = New-Object P3+RECT
    [P3]::GetWindowRect($wwg, [ref]$r) | Out-Null
    Rep "_WwG screen rect" ("L={0} T={1} R={2} B={3}" -f $r.L, $r.T, $r.R, $r.B)

    # --- GetPoint: caret/range -> screen rectangle (drives the overlay) ---
    try {
        $x = 0; $y = 0; $w = 0; $h = 0
        $target = $app.ActiveDocument.Paragraphs.Item(3).Range
        $dwin.ScrollIntoView($target)
        Start-Sleep -Milliseconds 300
        $dwin.GetPoint([ref]$x, [ref]$y, [ref]$w, [ref]$h, $target)
        Rep "GetPoint(para 3)" ("x=$x y=$y w=$w h=$h")
        $hit = ($w -gt 0 -and $h -gt 0)
        Rep "GetPoint usable?" $hit
    } catch { Rep "GetPoint FAILED" $_.Exception.Message }

    # --- RangeFromPoint: screen point -> document range (drives click-to-caret) ---
    foreach ($probe in @(@(0.5, 0.35), @(0.5, 0.5), @(0.4, 0.6))) {
        $px = [int]($r.L + ($r.R - $r.L) * $probe[0])
        $py = [int]($r.T + ($r.B - $r.T) * $probe[1])
        try {
            $obj = $dwin.RangeFromPoint($px, $py)
            if ($null -eq $obj) {
                Rep "RangeFromPoint($px,$py)" "returned Nothing"
            } else {
                $txt = ""
                try { $txt = ($obj.Text -replace '\s+', ' ') } catch { $txt = "<no .Text>" }
                if ($txt.Length -gt 46) { $txt = $txt.Substring(0, 46) + "..." }
                Rep "RangeFromPoint($px,$py)" ("OK start=$($obj.Start) -> '$txt'")
            }
        } catch { Rep "RangeFromPoint($px,$py)" ("FAILED: " + $_.Exception.Message) }
    }

    # --- Round trip: can we actually place a caret and type? ---
    try {
        $px = [int]($r.L + ($r.R - $r.L) * 0.5)
        $py = [int]($r.T + ($r.B - $r.T) * 0.5)
        $rng = $dwin.RangeFromPoint($px, $py)
        $rng.Select()
        $app.Selection.TypeText("[EDIT]")
        Rep "click->select->TypeText" ("OK, selection now at " + $app.Selection.Start)
        $app.ActiveDocument.Undo() | Out-Null
        Rep "Undo()" "OK"
    } catch { Rep "edit round trip FAILED" $_.Exception.Message }

    $app.ActiveDocument.Close(0)
}
catch { Rep "ERROR" $_.Exception.Message }
finally {
    # $ownPid came back from CreateProcess -- a kernel fact about a process made
    # for us, not a census difference (#136) -- so the kill is kept, and routed
    # through Stop-VerifiedWord, which pins the handle and re-verifies name and
    # StartTime before terminating.
    Rep "teardown pid $ownPid" (Stop-VerifiedWord $ownPid $ownStart)
    Start-Sleep -Milliseconds 600
    Rep "cleanup" "done"
}
