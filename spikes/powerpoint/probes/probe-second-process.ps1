# Probe: can a SECOND PowerPoint process exist at all?
#
# probe-single-instance.ps1 proved that New-Object -ComObject PowerPoint.Application
# attaches to the running instance instead of starting a new one. That is a
# statement about the CLSID registration, not necessarily about the executable.
# Before concluding that PowerPoint cannot be isolated, the stronger claim has to
# be tested: launch POWERPNT.EXE directly and see whether the process survives.
#
#   P1  launch POWERPNT.EXE on THIS desktop while an instance is already running
#   P2  launch it on a SEPARATE window station desktop -- this is exactly the
#       technique spikes/isolation/probes/probe-bind.ps1 uses to give Word a
#       private instance, so if anything works it is this
#   P3  if a second process survives, try to bind to it specifically via
#       AccessibleObjectFromWindow/OBJID_NATIVEOM, and check whether the object
#       it hands back belongs to that pid or to the original instance
#
# P2 is the control that matters: the identical technique demonstrably produces a
# private Word. If it does not produce a private PowerPoint, the difference is
# the application, not the method.
#
# SAFETY: every process this probe touches is one it launched itself, by pid
# captured from CreateProcess. It never enumerates-and-kills by name.

$ErrorActionPreference = 'Continue'
. (Join-Path $PSScriptRoot '_common.ps1')

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Text;

public static class Pw
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

    public static string Classes(IntPtr desk, uint pid)
    {
        var sb = new StringBuilder();
        EnumDesktopWindows(desk, delegate(IntPtr h, IntPtr p) {
            uint got; GetWindowThreadProcessId(h, out got);
            if (got != pid) return true;
            var cls = new StringBuilder(256); GetClassName(h, cls, 256);
            sb.Append(cls.ToString()).Append(" ");
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
            var cls = new StringBuilder(256); GetClassName(h, cls, 256);
            if (cls.ToString() == wanted) { hit = h; return false; }
            return true;
        }, IntPtr.Zero);
        return hit;
    }

    public static string ChildClasses(IntPtr parent)
    {
        var sb = new StringBuilder();
        EnumChildWindows(parent, delegate(IntPtr h, IntPtr p) {
            var cls = new StringBuilder(256); GetClassName(h, cls, 256);
            sb.Append(cls.ToString()).Append(" ");
            return true;
        }, IntPtr.Zero);
        return sb.ToString().Trim();
    }

    // Word binds on the document surface (_WwG), not the frame, so the whole
    // descendant tree has to be offered to AccessibleObjectFromWindow.
    public static IntPtr[] Descendants(IntPtr parent)
    {
        var list = new System.Collections.Generic.List<IntPtr>();
        EnumChildWindows(parent, delegate(IntPtr h, IntPtr p) { list.Add(h); return true; }, IntPtr.Zero);
        return list.ToArray();
    }

    public static string ClassOf(IntPtr h)
    {
        var cls = new StringBuilder(256); GetClassName(h, cls, 256);
        return cls.ToString();
    }

    public static object NativeOm(IntPtr hwnd)
    {
        Guid iid = new Guid("00020400-0000-0000-C000-000000000046");
        object obj;
        int hr = AccessibleObjectFromWindow(hwnd, 0xFFFFFFF0, ref iid, out obj);
        if (hr != 0) throw new COMException("AccessibleObjectFromWindow failed", hr);
        return obj;
    }
}
"@

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

function Start-Ppt([string]$Exe, [string]$Desktop, [string]$File) {
    $si = New-Object Pw+STARTUPINFO
    $si.cb = [Runtime.InteropServices.Marshal]::SizeOf($si)
    if ($Desktop) { $si.desktop = $Desktop }
    $pi = New-Object Pw+PROCESS_INFORMATION
    $cmd = '"' + $Exe + '"'
    if ($File) { $cmd += ' "' + $File + '"' }
    $ok = [Pw]::CreateProcess([IntPtr]::Zero, $cmd, [IntPtr]::Zero, [IntPtr]::Zero,
        $false, 0, [IntPtr]::Zero, [IntPtr]::Zero, [ref]$si, [ref]$pi)
    # StartTime is captured here, at launch, because Stop-VerifiedPpt refuses a
    # kill it cannot match on both name and start. Without it every sweep below
    # would decline and the teardown would quietly become a leak. It matters more
    # here than anywhere: :178 reports "process exited - handed off" as a VERDICT
    # of this very experiment, so a pid that has been recycled is the expected
    # case, not the exotic one.
    $start = $null
    if ($ok) { try { $start = (Get-Process -Id $pi.pid -ErrorAction SilentlyContinue).StartTime } catch { } }
    [pscustomobject]@{ Ok = $ok; Pid = $pi.pid; StartTime = $start }
}

# P2/P3 must launch WITH a deck open: Word binds on the document surface, so an
# empty PowerPoint with no document window would fail the test for the wrong
# reason.
$fixture = Join-Path (Split-Path $PSScriptRoot -Parent) '.fixtures\deck.pptx'
$work = Join-Path $env:TEMP ("pptproc-" + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Force -Path $work | Out-Null
$deck = Join-Path $work 'deck.pptx'
if (Test-Path $fixture) { Copy-Item $fixture $deck } else { $deck = $null }

$exe = Find-PowerPointExe
Rep "POWERPNT.EXE" $exe
$before = @(Get-PptPids)
Rep "POWERPNT pids before" ($(if ($before) { $before -join ',' } else { '(none)' }))

$anchor = $null      # a COM instance, so there is definitely one running
$launched = @()      # launcher records (pid + StartTime) -- the only kill targets
$desk = [IntPtr]::Zero

try {
    $anchor = New-PowerPointInstance
    $anchorPid = @(Get-PptPids | Where-Object { $before -notcontains $_ })
    Rep "anchor instance pid" ($(if ($anchorPid) { $anchorPid -join ',' } else { '(attached to existing)' }))

    Say "== P1: launch POWERPNT.EXE on THIS desktop, with an instance already running =="
    $p1 = Start-Ppt $exe $null $null
    Rep "  CreateProcess ok / pid" ("$($p1.Ok) / $($p1.Pid)")
    if ($p1.Pid) { $launched += $p1 }
    Start-Sleep -Seconds 8
    $alive1 = [bool](Get-Process -Id $p1.Pid -ErrorAction SilentlyContinue)
    Rep "  pid $($p1.Pid) still alive after 8s" $alive1
    Rep "  VERDICT" $(if ($alive1) { 'a second PROCESS exists' } else { 'process exited - handed off to the running instance' })
    if ($alive1) {
        $launched = @($launched | Where-Object { $_.Pid -ne $p1.Pid })
        $r1 = Stop-VerifiedPpt $p1.Pid $p1.StartTime
        Say ("      sweep of $($p1.Pid): $r1")
    }

    Say "== P2: launch it on a SEPARATE desktop (the technique that isolates Word) =="
    $deskName = 'CopilotPptBind'
    $desk = [Pw]::CreateDesktop($deskName, [IntPtr]::Zero, [IntPtr]::Zero, 0, 0x10000000, [IntPtr]::Zero)
    Rep "  desktop handle" $desk
    $p2 = Start-Ppt $exe $deskName $deck
    Rep "  launched with a deck" $(if ($deck) { 'yes' } else { 'NO - fixture missing, P3 is invalid' })
    Rep "  CreateProcess ok / pid" ("$($p2.Ok) / $($p2.Pid)")
    if ($p2.Pid) { $launched += $p2 }
    Start-Sleep -Seconds 15
    $alive2 = [bool](Get-Process -Id $p2.Pid -ErrorAction SilentlyContinue)
    Rep "  pid $($p2.Pid) still alive after 15s" $alive2
    Rep "  VERDICT" $(if ($alive2) { 'a second PROCESS exists on its own desktop' } else { 'process exited even on its own desktop' })

    if ($alive2) {
        Say "== P3: can we bind to THAT process specifically? =="
        $classes = [Pw]::Classes($desk, [uint32]$p2.Pid)
        Rep "  top-level classes for that pid" $(if ($classes) { $classes } else { '(none)' })
        $frame = [Pw]::FindClass($desk, [uint32]$p2.Pid, 'PPTFrameClass')
        Rep "  PPTFrameClass hwnd" $frame
        if ($frame -ne [IntPtr]::Zero) {
            # Try the frame and every descendant. Word succeeds on _WwG, a child,
            # so testing only the frame would understate what is possible.
            $targets = @($frame) + @([Pw]::Descendants($frame))
            Rep "  windows offered to NativeOm" $targets.Count
            # Do NOT stop at the first success: a MsoCommandBar answers
            # AccessibleObjectFromWindow but is not an entry point into the
            # presentation object model. Only a bind that can name the open deck
            # counts as isolation.
            $hits = @()
            foreach ($h in $targets) {
                $cls = [Pw]::ClassOf($h)
                try { $om = [Pw]::NativeOm($h) } catch { continue }
                if (-not $om) { continue }
                $ver = ''; $cnt = -1; $deckName = ''
                try { $ver = $om.Application.Version } catch { }
                try { $cnt = $om.Application.Presentations.Count } catch { }
                try { $deckName = $om.Application.Presentations.Item(1).Name } catch { }
                $hits += [pscustomobject]@{ Class = $cls; Type = $om.GetType().FullName; Ver = $ver; Count = $cnt; Deck = $deckName }
                Say ("      bound '{0}' ver='{1}' Presentations={2} deck='{3}'" -f $cls, $ver, $cnt, $deckName)
            }
            Rep "  windows that answered NativeOm" $hits.Count
            $real = @($hits | Where-Object { $_.Deck })
            if ($real.Count) {
                Rep "  BOUND to the isolated deck via" ("'" + $real[0].Class + "' -> " + $real[0].Deck)
                $anchorCount = try { $anchor.App.Presentations.Count } catch { '?' }
                Rep "  anchor .Presentations.Count" $anchorCount
                Rep "  ISOLATION" $(if ($anchorCount -eq 0) {
                        'YES - the bound object sees a deck the anchor instance does not' }
                    else { 'INCONCLUSIVE - anchor has decks open too' })
            }
            else {
                Rep "  binds that could name the deck" '0'
                Rep "  ISOLATION" 'NO - process exists, but no window exposes the presentation object model'
            }
        }
        else { Rep "  binding" 'no frame window to bind to' }
    }
}
catch { Rep "ERROR" $_.Exception.Message.Split([char]10)[0] }
finally {
    foreach ($rec in ($launched | Sort-Object -Property Pid -Unique)) {
        $r = Stop-VerifiedPpt $rec.Pid $rec.StartTime
        if ($r -ne 'gone') { Say ("      sweep of launched pid {0}: {1}" -f $rec.Pid, $r) }
    }
    Start-Sleep -Milliseconds 900
    Close-PowerPointInstance $anchor
    if ($desk -ne [IntPtr]::Zero) { [Pw]::CloseDesktop($desk) | Out-Null }
    Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 900
    $after = @(Get-PptPids)
    Rep "POWERPNT pids after" ($(if ($after) { $after -join ',' } else { '(none)' }))
    $leaked = @($after | Where-Object { $before -notcontains $_ })
    Rep "leaked (should be none)" ($(if ($leaked) { $leaked -join ',' } else { '(none)' }))
}
