# What lock does Word actually hold, and which readers survive it?
#
# CONTEXT.md and .github/copilot-instructions.md both assert "Word takes
# FileShare::Read". The conclusion drawn from it -- a document open in Word can
# still be copied -- is correct and heavily relied on. This probe tests whether
# the stated mechanism is the one that produces it.
#
# The discriminator is a reader that asks for FileShare::Read itself. A caller's
# FileShare value is what it grants to OTHERS, so the two candidate mechanisms
# make opposite predictions for it, while agreeing on Copy-Item and Node.
#
# Run:  powershell -File probe-fileshare-algebra.ps1

$ErrorActionPreference = 'Stop'
$scratch = Join-Path $env:TEMP ("fileshare-probe-" + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $scratch | Out-Null

function Try-Open {
    param([string]$Path, [string]$Access, [string]$Share)
    try {
        $h = [System.IO.File]::Open(
            $Path,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::$Access,
            [System.IO.FileShare]::$Share)
        $h.Dispose()
        return 'ok'
    } catch [System.IO.IOException] {
        return 'sharing violation'
    } catch {
        return $_.Exception.GetType().Name
    }
}

$readers = @(
    @{ Label = 'read, grants ReadWrite  (Copy-Item / Node) '; Access = 'Read'; Share = 'ReadWrite' },
    @{ Label = 'read, grants Read       (DISCRIMINATOR)    '; Access = 'Read'; Share = 'Read' },
    @{ Label = 'read, grants None                          '; Access = 'Read'; Share = 'None' },
    @{ Label = 'ReadWrite, grants None (Test-FileWritable) '; Access = 'ReadWrite'; Share = 'None' }
)

Write-Output ''
Write-Output '=== PART A: sharing algebra against a synthetic holder ==='
Write-Output ''

$holders = @(
    @{ Label = 'holder: READ handle granting Read   (what our docs claim Word does)'; Access = 'Read'; Share = 'Read' },
    @{ Label = 'holder: WRITE handle granting ReadWrite (what L2 claims Word does)'; Access = 'Write'; Share = 'ReadWrite' }
)

foreach ($holder in $holders) {
    $target = Join-Path $scratch ("holder-" + [guid]::NewGuid().ToString('N').Substring(0, 6) + '.bin')
    Set-Content -LiteralPath $target -Value 'payload' -Encoding ascii

    $handle = [System.IO.File]::Open(
        $target,
        [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::$($holder.Access),
        [System.IO.FileShare]::$($holder.Share))

    Write-Output $holder.Label
    foreach ($r in $readers) {
        $result = Try-Open -Path $target -Access $r.Access -Share $r.Share
        Write-Output ("    {0} -> {1}" -f $r.Label, $result)
    }

    $copyDest = Join-Path $scratch ('copy-' + [guid]::NewGuid().ToString('N').Substring(0, 6) + '.bin')
    try {
        Copy-Item -LiteralPath $target -Destination $copyDest -ErrorAction Stop
        Write-Output '    Copy-Item                                  -> ok'
    } catch {
        Write-Output ("    Copy-Item                                  -> {0}" -f $_.Exception.GetType().Name)
    }

    $handle.Dispose()
    Write-Output ''
}

Write-Output '=== PART B: the same readers against a document held by real Word ==='
Write-Output ''

$doc = Join-Path $scratch 'held.docx'
$word = $null
$startedPids = @()
$before = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    $after = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $startedPids = @($after | Where-Object { $before -notcontains $_ })

    $new = $word.Documents.Add()
    $new.Content.Text = 'probe payload'
    [object]$savePath = [string]$doc
    [object]$saveFormat = 16
    $new.SaveAs([ref]$savePath, [ref]$saveFormat)
    [object]$noSave = 0
    $new.Close([ref]$noSave)

    # Reopen so Word holds the file the way it holds a user's open document.
    [object]$openPath = [string]$doc
    $opened = $word.Documents.Open([ref]$openPath)

    Write-Output ("document open in Word: {0}" -f $doc)
    foreach ($r in $readers) {
        $result = Try-Open -Path $doc -Access $r.Access -Share $r.Share
        Write-Output ("    {0} -> {1}" -f $r.Label, $result)
    }

    $copyDest = Join-Path $scratch 'word-copy.docx'
    try {
        Copy-Item -LiteralPath $doc -Destination $copyDest -ErrorAction Stop
        Write-Output '    Copy-Item                                  -> ok'
    } catch {
        Write-Output ("    Copy-Item                                  -> {0}" -f $_.Exception.GetType().Name)
    }

    $opened.Close([ref]$noSave)
} catch {
    Write-Output ("PART B FAILED: {0}" -f $_.Exception.Message)
} finally {
    if ($word) {
        try { $word.Quit() } catch { }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
    # Only ever reap a Word this probe started.
    foreach ($p in $startedPids) {
        $proc = Get-Process -Id $p -ErrorAction SilentlyContinue
        if ($proc) {
            $proc.WaitForExit(20000) | Out-Null
        }
    }
    Write-Output ''
    Write-Output ("word processes started by this probe: {0}" -f ($startedPids -join ', '))
    Remove-Item -Recurse -Force $scratch -ErrorAction SilentlyContinue
}
