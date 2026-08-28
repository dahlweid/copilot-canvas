# Shared machinery for launching a PowerPoint that is genuinely OURS.
#
# probe-single-instance.ps1 showed that New-Object -ComObject PowerPoint.Application
# attaches to whatever PowerPoint is already running, and probe-second-process.ps1
# showed that launching POWERPNT.EXE on the current desktop just hands off to that
# same instance and exits. The only route to a private instance is the one
# spikes/isolation already uses for Word: create a separate window station
# desktop, launch the executable into it, and bind to the object model through
# AccessibleObjectFromWindow/OBJID_NATIVEOM.
#
# For Word the bind target is the document surface, class _WwG.
# For PowerPoint it is class mdiClass -- established empirically by
# probe-second-process.ps1, which offered all 23 descendant windows to the binder
# and found that only mdiClass hands back an object that can name the open deck.
# MsoCommandBar and RICHEDIT60W also answer, but their .Application is a stub
# with an empty Version and no Presentations, so "it answered" is not the test.
#
# SAFETY: everything here is keyed to a pid returned by CreateProcess. Nothing in
# this file enumerates processes by name, so it cannot touch another session's
# PowerPoint.

Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

public static class PptIso
{
    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern IntPtr CreateDesktop(string name, IntPtr dev, IntPtr mode, int flags, uint access, IntPtr sa);
    [DllImport("user32.dll")] public static extern bool CloseDesktop(IntPtr h);
    [DllImport("user32.dll")] public static extern bool EnumDesktopWindows(IntPtr desk, EnumProc cb, IntPtr p);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent, EnumProc cb, IntPtr p);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int max);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);

    [DllImport("oleacc.dll")]
    public static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint objId, ref Guid iid,
        [MarshalAs(UnmanagedType.IDispatch)] out object obj);

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern bool CreateProcess(IntPtr app, string cmd, IntPtr pa, IntPtr ta, bool inherit,
        uint flags, IntPtr env, IntPtr dir, ref STARTUPINFO si, out PROCESS_INFORMATION pi);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetWindowText(IntPtr h, StringBuilder s, int max);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern bool PostMessage(IntPtr h, uint msg, IntPtr wp, IntPtr lp);

    public delegate bool EnumProc(IntPtr hwnd, IntPtr param);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct STARTUPINFO
    {
        public int cb; public string reserved; public string desktop; public string title;
        public int x, y, xSize, ySize, xCount, yCount, fill, flags;
        public short showWindow, cbReserved2; public IntPtr reserved2, hStdIn, hStdOut, hStdErr;
    }
    [StructLayout(LayoutKind.Sequential)]
    public struct PROCESS_INFORMATION { public IntPtr hProcess, hThread; public int pid, tid; }

    public static string ClassOf(IntPtr h)
    {
        var cls = new StringBuilder(256); GetClassName(h, cls, 256);
        return cls.ToString();
    }

    public static string Classes(IntPtr desk, uint pid)
    {
        var sb = new StringBuilder();
        EnumDesktopWindows(desk, delegate(IntPtr h, IntPtr p) {
            uint got; GetWindowThreadProcessId(h, out got);
            if (got == pid) sb.Append(ClassOf(h)).Append(" ");
            return true;
        }, IntPtr.Zero);
        return sb.ToString().Trim();
    }

    public static IntPtr FindClass(IntPtr desk, uint pid, string wanted)
    {
        IntPtr hit = IntPtr.Zero;
        EnumDesktopWindows(desk, delegate(IntPtr h, IntPtr p) {
            uint got; GetWindowThreadProcessId(h, out got);
            if (got != pid) return true;
            if (ClassOf(h) == wanted) { hit = h; return false; }
            return true;
        }, IntPtr.Zero);
        return hit;
    }

    public static IntPtr[] Descendants(IntPtr parent)
    {
        var list = new List<IntPtr>();
        EnumChildWindows(parent, delegate(IntPtr h, IntPtr p) { list.Add(h); return true; }, IntPtr.Zero);
        return list.ToArray();
    }

    public static object NativeOm(IntPtr hwnd)
    {
        Guid iid = new Guid("00020400-0000-0000-C000-000000000046"); // IID_IDispatch
        object obj;
        int hr = AccessibleObjectFromWindow(hwnd, 0xFFFFFFF0, ref iid, out obj); // OBJID_NATIVEOM
        if (hr != 0) throw new COMException("AccessibleObjectFromWindow failed", hr);
        return obj;
    }

    public static string TextOf(IntPtr h)
    {
        var sb = new StringBuilder(512); GetWindowText(h, sb, 512);
        return sb.ToString();
    }

    // A modal #32770 is PowerPoint refusing to proceed. Reading it is the whole
    // point -- the caption and static text say WHY, which is the answer to the
    // question rather than an obstacle to it.
    public static string DialogText(IntPtr desk, uint pid)
    {
        var sb = new StringBuilder();
        EnumDesktopWindows(desk, delegate(IntPtr h, IntPtr p) {
            uint got; GetWindowThreadProcessId(h, out got);
            if (got != pid || ClassOf(h) != "#32770") return true;
            sb.Append("[").Append(TextOf(h)).Append("] ");
            EnumChildWindows(h, delegate(IntPtr c, IntPtr q) {
                string t = TextOf(c);
                if (t.Length > 0) sb.Append(t.Replace("\r", " ").Replace("\n", " ")).Append(" | ");
                return true;
            }, IntPtr.Zero);
            return true;
        }, IntPtr.Zero);
        return sb.ToString().Trim();
    }

    public static int CloseDialogs(IntPtr desk, uint pid)
    {
        var hits = new List<IntPtr>();
        EnumDesktopWindows(desk, delegate(IntPtr h, IntPtr p) {
            uint got; GetWindowThreadProcessId(h, out got);
            if (got == pid && ClassOf(h) == "#32770") hits.Add(h);
            return true;
        }, IntPtr.Zero);
        foreach (var h in hits)
        {
            // WM_CLOSE is not enough for the safe-mode prompt: it has Yes/No and
            // closing it is not an answer. IDNO via WM_COMMAND is language-neutral,
            // which matters because this Office is German ("&Nein", not "&No").
            PostMessage(h, 0x0111, new IntPtr(7), IntPtr.Zero);  // WM_COMMAND, IDNO
            PostMessage(h, 0x0111, new IntPtr(2), IntPtr.Zero);  // WM_COMMAND, IDCANCEL
            PostMessage(h, 0x0010, IntPtr.Zero, IntPtr.Zero);    // WM_CLOSE
        }
        return hits.Count;
    }
}
"@ -ErrorAction SilentlyContinue

function Find-PowerPointExe {
    foreach ($p in @(
            "$env:ProgramFiles\Microsoft Office\root\Office16\POWERPNT.EXE",
            "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\POWERPNT.EXE")) {
        if (Test-Path $p) { return $p }
    }
    $c = Get-Command POWERPNT.EXE -ErrorAction SilentlyContinue
    if ($c) { return $c.Source }
    throw 'POWERPNT.EXE not found'
}

# Launch a private PowerPoint on its own desktop with $File open, and bind to it.
# Returns $null on failure. Always pair with Stop-IsolatedPowerPoint.
function Start-IsolatedPowerPoint {
    param(
        [Parameter(Mandatory)][string]$File,
        [string]$DesktopName = ('CopilotPpt' + [guid]::NewGuid().ToString('N').Substring(0, 8)),
        [int]$TimeoutSeconds = 40,
        # PowerPoint needs to finish starting before it is poked. Hammering
        # AccessibleObjectFromWindow from the first half-second produced timeouts
        # in BOTH arms of probe-cross-instance-lock.ps1, including the control --
        # which is how the flaw was caught rather than published as a finding.
        [int]$SettleSeconds = 12
    )
    $exe = Find-PowerPointExe
    $desk = [PptIso]::CreateDesktop($DesktopName, [IntPtr]::Zero, [IntPtr]::Zero, 0, 0x10000000, [IntPtr]::Zero)
    if ($desk -eq [IntPtr]::Zero) { throw 'CreateDesktop failed' }

    $si = New-Object PptIso+STARTUPINFO
    $si.cb = [Runtime.InteropServices.Marshal]::SizeOf($si)
    $si.desktop = $DesktopName
    $pi = New-Object PptIso+PROCESS_INFORMATION
    $cmd = '"' + $exe + '" "' + $File + '"'
    $ok = [PptIso]::CreateProcess([IntPtr]::Zero, $cmd, [IntPtr]::Zero, [IntPtr]::Zero,
        $false, 0, [IntPtr]::Zero, [IntPtr]::Zero, [ref]$si, [ref]$pi)

    $ctx = [pscustomobject]@{
        Pid = $pi.Pid; Desk = $desk; DesktopName = $DesktopName
        App = $null; Pres = $null; BoundMs = -1; Ok = $ok; Diag = ''; Dialogs = @()
    }
    if (-not $ok) { return $ctx }

    $sw = [Diagnostics.Stopwatch]::StartNew()
    Start-Sleep -Seconds $SettleSeconds
    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        if (-not (Get-Process -Id $ctx.Pid -ErrorAction SilentlyContinue)) { $ctx.Diag = 'process exited'; break }

        # Record any modal dialog before dismissing it: its text is evidence.
        $dlg = [PptIso]::DialogText($desk, [uint32]$ctx.Pid)
        if ($dlg -and ($ctx.Dialogs -notcontains $dlg)) { $ctx.Dialogs += $dlg }
        if ($dlg) { [PptIso]::CloseDialogs($desk, [uint32]$ctx.Pid) | Out-Null; Start-Sleep -Milliseconds 700 }

        $frame = [PptIso]::FindClass($desk, [uint32]$ctx.Pid, 'PPTFrameClass')
        if ($frame -eq [IntPtr]::Zero) {
            $ctx.Diag = 'no PPTFrameClass; top-level: ' + [PptIso]::Classes($desk, [uint32]$ctx.Pid)
            Start-Sleep -Milliseconds 1000
            continue
        }
        $kids = @([PptIso]::Descendants($frame))
        $ctx.Diag = 'frame ok; mdiClass present: ' +
            [bool](@($kids | Where-Object { [PptIso]::ClassOf($_) -eq 'mdiClass' }).Count)
        foreach ($h in $kids) {
            if ([PptIso]::ClassOf($h) -ne 'mdiClass') { continue }
            try {
                $om = [PptIso]::NativeOm($h)
                $app = $om.Application
                if ($app.Presentations.Count -ge 1) {
                    $ctx.App = $app
                    $ctx.Pres = $app.Presentations.Item(1)
                    $ctx.BoundMs = $sw.ElapsedMilliseconds
                    return $ctx
                }
                $ctx.Diag = 'bound mdiClass but Presentations.Count=0'
            }
            catch { $ctx.Diag = 'mdiClass NativeOm failed: ' + $_.Exception.Message.Split([char]10)[0] }
        }
        Start-Sleep -Milliseconds 1000
    }
    return $ctx
}

function Stop-IsolatedPowerPoint {
    param($Ctx)
    if (-not $Ctx) { return }
    # Prefer a graceful Quit. Killing POWERPNT.EXE makes the NEXT launch show a
    # safe-mode prompt ("PowerPoint konnte beim letzten Mal nicht gestartet
    # werden"), which broke the control arm of probe-cross-instance-lock.ps1
    # until it was diagnosed. Quit is still not trusted to terminate -- 15/15
    # cycles in probe-stability.ps1 left the process running -- so the kill
    # remains, just as a fallback rather than the first move.
    if ($Ctx.App) {
        try {
            foreach ($p in @($Ctx.App.Presentations)) { try { $p.Saved = -1 } catch { } }
            $Ctx.App.Quit()
        }
        catch { }
        for ($i = 0; $i -lt 16; $i++) {
            if (-not (Get-Process -Id $Ctx.Pid -ErrorAction SilentlyContinue)) { break }
            Start-Sleep -Milliseconds 500
        }
    }
    if ($Ctx.Pid -and (Get-Process -Id $Ctx.Pid -ErrorAction SilentlyContinue)) {
        Stop-Process -Id $Ctx.Pid -Force -ErrorAction SilentlyContinue
        Start-Sleep -Milliseconds 900
    }
    if ($Ctx.Desk -and $Ctx.Desk -ne [IntPtr]::Zero) { [PptIso]::CloseDesktop($Ctx.Desk) | Out-Null }
}
