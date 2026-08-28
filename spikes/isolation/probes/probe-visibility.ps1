# Probe: verify the user-visible consequence of the hidden-desktop approach.
# A window on a non-input desktop object SHOULD be invisible to the shell, but
# "should" has been wrong twice in this investigation, so measure it:
#   - Is the Word window enumerable on the DEFAULT desktop at all?
#   - Does the taskbar (via the shell's window list) know about it?
#   - Does it appear in the Alt+Tab candidate set?
$ErrorActionPreference = 'Stop'

Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public static class P4
{
    [DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
    public static extern IntPtr CreateDesktop(string n, IntPtr d, IntPtr m, int f, uint a, IntPtr s);
    [DllImport("user32.dll")] public static extern bool CloseDesktop(IntPtr h);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr p);
    [DllImport("user32.dll")] public static extern bool EnumDesktopWindows(IntPtr d, EnumProc cb, IntPtr p);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int m);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int m);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern IntPtr GetWindow(IntPtr h, uint c);
    [DllImport("user32.dll")] public static extern int GetWindowLong(IntPtr h, int i);
    [DllImport("kernel32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
    public static extern bool CreateProcess(IntPtr a, string c, IntPtr pa, IntPtr ta, bool inh, uint f,
        IntPtr env, IntPtr dir, ref STARTUPINFO si, out PROCESS_INFORMATION pi);
    public delegate bool EnumProc(IntPtr h, IntPtr p);
    [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)]
    public struct STARTUPINFO { public int cb; public string r1, desktop, title;
        public int x,y,xs,ys,xc,yc,fill,flags; public short show, cb2; public IntPtr r2, si, so, se; }
    [StructLayout(LayoutKind.Sequential)]
    public struct PROCESS_INFORMATION { public IntPtr hP, hT; public int pid, tid; }

    // Count windows belonging to a pid that EnumWindows (current desktop) sees.
    public static List<string> OnCurrentDesktop(uint pid) {
        var found = new List<string>();
        EnumWindows(delegate(IntPtr h, IntPtr p) {
            uint g; GetWindowThreadProcessId(h, out g); if (g != pid) return true;
            var c = new StringBuilder(256); GetClassName(h, c, 256);
            var t = new StringBuilder(256); GetWindowText(h, t, 256);
            found.Add(string.Format("{0} visible={1} title='{2}'", c, IsWindowVisible(h), t));
            return true;
        }, IntPtr.Zero);
        return found;
    }

    // Approximate the Alt+Tab rule: visible, unowned, and not WS_EX_TOOLWINDOW.
    public static List<string> AltTabCandidates(uint pid) {
        var found = new List<string>();
        EnumWindows(delegate(IntPtr h, IntPtr p) {
            uint g; GetWindowThreadProcessId(h, out g); if (g != pid) return true;
            if (!IsWindowVisible(h)) return true;
            if (GetWindow(h, 4) != IntPtr.Zero) return true;      // GW_OWNER
            int ex = GetWindowLong(h, -20);                        // GWL_EXSTYLE
            if ((ex & 0x00000080) != 0) return true;               // WS_EX_TOOLWINDOW
            var t = new StringBuilder(256); GetWindowText(h, t, 256);
            found.Add(t.ToString());
            return true;
        }, IntPtr.Zero);
        return found;
    }
    public static int CountOnDesktop(IntPtr desk, uint pid) {
        int n = 0;
        EnumDesktopWindows(desk, delegate(IntPtr h, IntPtr p) {
            uint g; GetWindowThreadProcessId(h, out g); if (g == pid) n++;
            return true;
        }, IntPtr.Zero);
        return n;
    }
}
"@

function Rep($l, $v) { Write-Output ("{0,-40} {1}" -f $l, $v) }

$deskName = "CopilotWordVis"
$desk = [P4]::CreateDesktop($deskName, [IntPtr]::Zero, [IntPtr]::Zero, 0, 0x10000000, [IntPtr]::Zero)
$exe = "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
$doc = "$env:TEMP\desktop-probe.docx"
$ownPid = 0
try {
    $si = New-Object P4+STARTUPINFO
    $si.cb = [Runtime.InteropServices.Marshal]::SizeOf($si)
    $si.desktop = $deskName
    $pi = New-Object P4+PROCESS_INFORMATION
    $cmd = '"' + $exe + '" /w /q "' + $doc + '"'
    [P4]::CreateProcess([IntPtr]::Zero, $cmd, [IntPtr]::Zero, [IntPtr]::Zero, $false, 0, [IntPtr]::Zero, [IntPtr]::Zero, [ref]$si, [ref]$pi) | Out-Null
    $ownPid = $pi.pid
    Rep "Word pid on hidden desktop" $ownPid
    Start-Sleep -Seconds 12

    Rep "windows on the HIDDEN desktop" ([P4]::CountOnDesktop($desk, [uint32]$ownPid))

    $onDefault = [P4]::OnCurrentDesktop([uint32]$ownPid)
    Rep "windows visible to EnumWindows" $onDefault.Count
    foreach ($w in $onDefault) { Rep "   " $w }

    $alt = [P4]::AltTabCandidates([uint32]$ownPid)
    Rep "Alt+Tab candidates for that pid" $alt.Count
    foreach ($a in $alt) { Rep "   alt-tab" $a }

    # What the shell itself reports -- this is the taskbar's own view.
    $shellWindows = (New-Object -ComObject Shell.Application).Windows()
    Rep "Shell.Application window count" $shellWindows.Count

    # Sanity control: the process really is alive and holding the document.
    $proc = Get-Process -Id $ownPid -ErrorAction SilentlyContinue
    Rep "process alive / title" ("{0} / '{1}'" -f ($null -ne $proc), $proc.MainWindowTitle)
}
catch { Rep "ERROR" $_.Exception.Message }
finally {
    if ($ownPid -and (Get-Process -Id $ownPid -ErrorAction SilentlyContinue)) { Stop-Process -Id $ownPid -Force }
    Start-Sleep -Milliseconds 600
    [P4]::CloseDesktop($desk) | Out-Null
    Rep "cleanup" "done"
}
