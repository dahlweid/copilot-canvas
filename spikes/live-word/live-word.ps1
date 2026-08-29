# Live Word pixel-streaming spike -- capture host.
#
# Drives a *visible* Word window parked off-screen and captures it with
# PrintWindow(PW_RENDERFULLCONTENT), which asks a window to paint itself even
# when it is not actually on screen. This is the only honest way to show "live
# Word" inside a canvas, since a native window cannot live inside an iframe.
#
# Speaks the same newline-delimited JSON protocol as the main host.
# Commands: start, resize, crop, capture, scroll, goto, info, stop, quit.

param([string]$PidDir)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)

$WD_PRINT_VIEW = 3
$WD_ALERTS_NONE = 0
$WD_GOTO_PAGE = 1
$WD_GOTO_ABSOLUTE = 1

Add-Type -ReferencedAssemblies System.Drawing -TypeDefinition @"
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;

public static class WordCapture
{
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left, Top, Right, Bottom; }

    [DllImport("user32.dll")] static extern bool PrintWindow(IntPtr hwnd, IntPtr hdc, uint flags);
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr hwnd, out RECT rect);
    [DllImport("user32.dll")] static extern bool SetWindowPos(IntPtr hwnd, IntPtr after, int x, int y, int cx, int cy, uint flags);
    [DllImport("user32.dll")] static extern bool IsWindow(IntPtr hwnd);
    [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr hwnd, int cmd);
    [DllImport("user32.dll")] static extern IntPtr GetDC(IntPtr hwnd);
    [DllImport("user32.dll")] static extern int ReleaseDC(IntPtr hwnd, IntPtr hdc);
    [DllImport("gdi32.dll")] static extern int GetDeviceCaps(IntPtr hdc, int index);
    [DllImport("user32.dll")] static extern uint GetDpiForWindow(IntPtr hwnd);

    const uint PW_RENDERFULLCONTENT = 2;
    const uint SWP_NOACTIVATE = 0x0010;
    const uint SWP_NOZORDER = 0x0004;
    const int SW_SHOWNOACTIVATE = 4;

    static ImageCodecInfo jpegCodec;

    static ImageCodecInfo Jpeg()
    {
        if (jpegCodec == null)
            foreach (var c in ImageCodecInfo.GetImageEncoders())
                if (c.FormatID == ImageFormat.Jpeg.Guid) jpegCodec = c;
        return jpegCodec;
    }

    /// Parks the window off the visible desktop at a fixed size, without
    /// activating it -- the user must never see it steal focus.
    public static void Place(IntPtr hwnd, int x, int y, int width, int height)
    {
        ShowWindow(hwnd, SW_SHOWNOACTIVATE);
        SetWindowPos(hwnd, IntPtr.Zero, x, y, width, height, SWP_NOACTIVATE | SWP_NOZORDER);
    }

    public static int[] Size(IntPtr hwnd)
    {
        RECT r;
        GetWindowRect(hwnd, out r);
        return new int[] { r.Right - r.Left, r.Bottom - r.Top };
    }

    public static bool Alive(IntPtr hwnd) { return IsWindow(hwnd); }

    // Scale factor Word is painting at (1.5 at 150% display scaling). It has to
    // be read from Word's own window: this host process is DPI-unaware, so
    // GetDeviceCaps reports a virtualized 96 DPI and would always say 1.0.
    // Word fits the page to the window rect in these units but paints at this
    // scale, so a capture of the physical window is clipped by exactly this
    // factor unless the zoom is divided by it.
    public static double Scale(IntPtr hwnd)
    {
        try
        {
            uint dpi = GetDpiForWindow(hwnd);
            if (dpi >= 48) return dpi / 96.0;
        }
        catch { /* pre-1607 Windows */ }
        IntPtr dc = GetDC(IntPtr.Zero);
        try { return GetDeviceCaps(dc, 88) / 96.0; }   // 88 = LOGPIXELSX
        finally { ReleaseDC(IntPtr.Zero, dc); }
    }

    /// Captures the window as JPEG. Returns null when the window refused to
    /// paint, which PrintWindow reports by return value rather than throwing.
    public static byte[] Capture(IntPtr hwnd, int quality, int cropTop, int cropBottom)
    {
        RECT r;
        if (!GetWindowRect(hwnd, out r)) return null;
        int w = r.Right - r.Left, h = r.Bottom - r.Top;
        if (w <= 0 || h <= 0) return null;

        using (var bmp = new Bitmap(w, h, PixelFormat.Format32bppArgb))
        {
            using (var g = Graphics.FromImage(bmp))
            {
                IntPtr hdc = g.GetHdc();
                try { if (!PrintWindow(hwnd, hdc, PW_RENDERFULLCONTENT)) return null; }
                finally { g.ReleaseHdc(hdc); }
            }

            // Crop away the ribbon and status bar so the canvas shows the page
            // rather than Word's chrome.
            int top = Math.Max(0, cropTop);
            int bottom = Math.Max(0, cropBottom);
            Bitmap frame = bmp;
            Bitmap cropped = null;
            if (top + bottom < h && (top > 0 || bottom > 0))
            {
                cropped = bmp.Clone(new Rectangle(0, top, w, h - top - bottom), bmp.PixelFormat);
                frame = cropped;
            }

            try
            {
                var enc = new EncoderParameters(1);
                enc.Param[0] = new EncoderParameter(Encoder.Quality, (long)quality);
                using (var ms = new MemoryStream())
                {
                    frame.Save(ms, Jpeg(), enc);
                    return ms.ToArray();
                }
            }
            finally { if (cropped != null) cropped.Dispose(); }
        }
    }
}
"@

# --- state -------------------------------------------------------------------

$script:app = $null
$script:doc = $null
$script:win = $null
$script:hwnd = [IntPtr]::Zero
$script:ownedPid = 0
$script:cropTop = 0
$script:cropBottom = 0

function Get-WordPids {
    @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
}

# wdPageFitBestFit -- fits the page width to the window, which is what a panel
# of arbitrary width needs. Word does not re-fit on its own after SetWindowPos.
$WD_PAGE_FIT_BEST = 2

function Set-PageFit {
    if (-not $script:win) { return }
    try {
        $script:win.View.Zoom.PageFit = $WD_PAGE_FIT_BEST
        Start-Sleep -Milliseconds 80
        # PageFit sizes the page to the window rect, but Word then paints at the
        # desktop scale factor, so the captured bitmap is clipped by that factor.
        # Dividing the fitted zoom out again makes the whole page land inside the
        # frame we actually capture.
        $scale = [WordCapture]::Scale($script:hwnd)
        if ($scale -gt 1.01) {
            $fitted = [int]$script:win.View.Zoom.Percentage
            # 0.92: PageFit measures against a client width that excludes chrome
            # we do capture (the page is not centred in the bitmap we get), so
            # the divided zoom still overshoots by a few percent. Empirical.
            $script:win.View.Zoom.Percentage =
                [Math]::Max(10, [int][Math]::Round(($fitted / $scale) * 0.92))
        }
    }
    catch { }
}

function Write-PidFile {
    if (-not $PidDir -or $script:ownedPid -le 0) { return }
    if (-not (Test-Path $PidDir)) { [void](New-Item -ItemType Directory -Path $PidDir -Force) }
    Set-Content -Path (Join-Path $PidDir "live-$PID.pid") -Value $script:ownedPid -Encoding ascii
}

function Remove-PidFile {
    if (-not $PidDir) { return }
    Remove-Item -Path (Join-Path $PidDir "live-$PID.pid") -Force -ErrorAction SilentlyContinue
}

function Stop-LiveWord {
    if ($script:doc) {
        try { $script:doc.Close(0) } catch { }
        $script:doc = $null
    }
    if ($script:app) {
        # Quit(), never Quit(<arg>): under Windows PowerShell 5.1 -- the runtime
        # this host runs under -- the argument form throws and the Word survives,
        # and process exit does not reap it either (probe-quit0-leak.ps1).
        # The report goes to STDERR on purpose: stdout here is the newline-
        # delimited JSON protocol, and run-spike.mjs forwards stderr with a [ps]
        # prefix. A "reporting" catch writing to a channel nobody reads is still
        # a swallow.
        try { $script:app.Quit() }
        catch { [Console]::Error.WriteLine("Quit() FAILED (Word may leak): " + $_.Exception.Message.Split([char]10)[0]) }
        $script:app = $null
    }
    $script:win = $null
    $script:hwnd = [IntPtr]::Zero
    if ($script:ownedPid -gt 0) {
        # Application.Quit() returns seconds before its process exits (2.7-6.1 s
        # idle, longer under load -- ADR 0005), so the fixed 200 ms wait that used
        # to sit here made this kill the actual reaper on every run: Word was
        # SIGKILLed while still shutting down normally, and whether Quit() worked
        # was unobservable. Poll to a generous deadline instead; on success that
        # costs only the real exit time. The ProcessName check keeps a recycled pid
        # from being killed as if it were the Word.
        $deadline = (Get-Date).AddSeconds(30)
        while ((Get-Date) -lt $deadline -and (Get-Process -Id $script:ownedPid -ErrorAction SilentlyContinue).ProcessName -eq 'WINWORD') {
            Start-Sleep -Milliseconds 250
        }
        if ((Get-Process -Id $script:ownedPid -ErrorAction SilentlyContinue).ProcessName -eq 'WINWORD') {
            # Swallow the exit race, then observe the outcome. A bare catch here
            # would turn a Word that survived the kill into silence, which is the
            # failure this whole file is being corrected for.
            try { (Get-Process -Id $script:ownedPid -ErrorAction SilentlyContinue).Kill() } catch { }
            $kd = (Get-Date).AddSeconds(15)
            while ((Get-Date) -lt $kd -and (Get-Process -Id $script:ownedPid -ErrorAction SilentlyContinue).ProcessName -eq 'WINWORD') {
                Start-Sleep -Milliseconds 250
            }
            if ((Get-Process -Id $script:ownedPid -ErrorAction SilentlyContinue).ProcessName -eq 'WINWORD') {
                [Console]::Error.WriteLine("STILL ALIVE after 30 s and after kill (leaked): pid $($script:ownedPid)")
            }
            else { [Console]::Error.WriteLine("STILL ALIVE after 30 s, killed: pid $($script:ownedPid)") }
        }
    }
    Remove-PidFile
    $script:ownedPid = 0
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

function Cmd-Start($a) {
    if ($script:app) { Stop-LiveWord }

    $path = [string]$a.path
    if (-not (Test-Path $path)) { throw "No such file: $path" }
    $width = 900; if ($null -ne $a.width) { $width = [int]$a.width }
    $height = 1200; if ($null -ne $a.height) { $height = [int]$a.height }

    $before = Get-WordPids
    $script:app = New-Object -ComObject Word.Application
    for ($i = 0; $i -lt 15; $i++) {
        $new = @((Get-WordPids) | Where-Object { $before -notcontains $_ })
        if ($new.Count -gt 0) { $script:ownedPid = [int]$new[0]; break }
        Start-Sleep -Milliseconds 100
    }
    Write-PidFile

    $script:app.DisplayAlerts = $WD_ALERTS_NONE
    $script:app.AutomationSecurity = 3
    # Visible is unavoidable: PrintWindow needs a real window to ask for paint.
    # It is parked far off the desktop immediately below.
    $script:app.Visible = $true

    $script:doc = $script:app.Documents.Open($path, $false, $true, $false, "wc-no-open", "wc-no-open")
    $script:win = $script:app.ActiveWindow
    $script:win.View.Type = $WD_PRINT_VIEW
    try { $script:win.DisplayRulers = $false } catch { }
    try { $script:win.DisplayVerticalRuler = $false } catch { }
    try { $script:win.View.ShowAll = $false } catch { }
    $script:hwnd = [IntPtr][int64]$script:win.Hwnd
    # Park it beyond the right edge of every plausible desktop.
    [WordCapture]::Place($script:hwnd, 30000, 0, $width, $height)
    Start-Sleep -Milliseconds 400
    # Zoom must be re-fitted *after* the resize, or the page stays at whatever
    # width the window happened to have when the document opened.
    Set-PageFit

    $size = [WordCapture]::Size($script:hwnd)
    @{
        pid       = $script:ownedPid
        hwnd      = [int64]$script:hwnd
        width     = $size[0]
        height    = $size[1]
        pageCount = [int]$script:doc.Content.Information(4)
        zoom      = [int]$script:win.View.Zoom.Percentage
        dpiScale  = [Math]::Round([WordCapture]::Scale($script:hwnd), 2)
    }
}

function Cmd-Zoom($a) {
    if (-not $script:win) { throw "No live document." }
    if ($null -ne $a.pageFit) { $script:win.View.Zoom.PageFit = [int]$a.pageFit }
    if ($null -ne $a.percent) { $script:win.View.Zoom.Percentage = [int]$a.percent }
    Start-Sleep -Milliseconds 120
    @{ percent = [int]$script:win.View.Zoom.Percentage; pageFit = [int]$script:win.View.Zoom.PageFit }
}

function Cmd-Resize($a) {
    if (-not $script:app) { throw "No live document." }
    [WordCapture]::Place($script:hwnd, 30000, 0, [int]$a.width, [int]$a.height)
    Start-Sleep -Milliseconds 150
    Set-PageFit
    $size = [WordCapture]::Size($script:hwnd)
    @{ width = $size[0]; height = $size[1]; zoom = [int]$script:win.View.Zoom.Percentage }
}

function Cmd-Crop($a) {
    $script:cropTop = [int]$a.top
    $script:cropBottom = [int]$a.bottom
    @{ top = $script:cropTop; bottom = $script:cropBottom }
}

function Cmd-Capture($a) {
    if (-not $script:app) { throw "No live document." }
    if (-not [WordCapture]::Alive($script:hwnd)) { throw "The Word window is gone." }
    $quality = 70; if ($null -ne $a.quality) { $quality = [int]$a.quality }

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $bytes = [WordCapture]::Capture($script:hwnd, $quality, $script:cropTop, $script:cropBottom)
    $sw.Stop()
    if ($null -eq $bytes) { throw "PrintWindow refused to paint the window." }

    $result = @{ bytes = $bytes.Length; ms = [int]$sw.ElapsedMilliseconds }
    if ($a.out) {
        [System.IO.File]::WriteAllBytes([string]$a.out, $bytes)
        $result.out = [string]$a.out
    }
    else {
        $result.jpegBase64 = [Convert]::ToBase64String($bytes)
    }
    $result
}

function Cmd-Scroll($a) {
    if (-not $script:win) { throw "No live document." }
    $down = 0; if ($null -ne $a.down) { $down = [int]$a.down }
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    if ($down -ge 0) { $script:win.SmallScroll($down, 0, 0, 0) }
    else { $script:win.SmallScroll(0, [Math]::Abs($down), 0, 0) }
    $sw.Stop()
    @{ ms = [int]$sw.ElapsedMilliseconds }
}

function Cmd-Goto($a) {
    if (-not $script:win) { throw "No live document." }
    $page = [int]$a.page
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $rng = $script:doc.GoTo($WD_GOTO_PAGE, $WD_GOTO_ABSOLUTE, $page, "")
    $script:win.ScrollIntoView($rng, $true)
    $sw.Stop()
    @{ ms = [int]$sw.ElapsedMilliseconds; page = $page }
}

function Cmd-Info($a) {
    if (-not $script:app) { throw "No live document." }
    $size = [WordCapture]::Size($script:hwnd)
    @{
        width     = $size[0]
        height    = $size[1]
        pageCount = [int]$script:doc.Content.Information(4)
        zoom      = [int]$script:win.View.Zoom.Percentage
        alive     = [WordCapture]::Alive($script:hwnd)
    }
}

function Cmd-Stop($a) {
    Stop-LiveWord
    @{ stopped = $true }
}

# --- dispatch ----------------------------------------------------------------

$writer = [Console]::Out

function Send($obj) {
    $writer.WriteLine(($obj | ConvertTo-Json -Depth 8 -Compress))
    $writer.Flush()
}

try {
    while ($true) {
        $line = [Console]::In.ReadLine()
        if ($null -eq $line) { break }
        if ([string]::IsNullOrWhiteSpace($line)) { continue }

        $req = $null
        try { $req = $line | ConvertFrom-Json } catch {
            Send @{ id = 0; ok = $false; error = "Malformed request: $($_.Exception.Message)" }
            continue
        }

        $id = 0
        if ($null -ne $req.id) { $id = [int]$req.id }
        try {
            $result = switch ([string]$req.cmd) {
                'start' { Cmd-Start $req.args }
                'resize' { Cmd-Resize $req.args }
                'zoom' { Cmd-Zoom $req.args }
                'crop' { Cmd-Crop $req.args }
                'capture' { Cmd-Capture $req.args }
                'scroll' { Cmd-Scroll $req.args }
                'goto' { Cmd-Goto $req.args }
                'info' { Cmd-Info $req.args }
                'stop' { Cmd-Stop $req.args }
                'quit' { Send @{ id = $id; ok = $true; result = @{ bye = $true } }; Stop-LiveWord; exit 0 }
                default { throw "Unknown command: $($req.cmd)" }
            }
            Send @{ id = $id; ok = $true; result = $result }
        }
        catch {
            Send @{ id = $id; ok = $false; error = $_.Exception.Message }
        }
    }
}
finally {
    Stop-LiveWord
}
