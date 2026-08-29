# Does Word grant FileShare::Read or FileShare::ReadWrite on an open document?
#
# This question has now been answered three different ways in this repository
# (ADR 0005, then PR #22, then PR #24) and I was asked to verify the current
# answer rather than inherit it. This probe is written from scratch rather than
# re-running the existing instrument, so a shared mistake in that instrument
# cannot survive into this result.
#
# The trap this is built to avoid: Windows checks TWO things on every open --
#
#   (a) the access YOU request, against the share mode each existing handle GRANTS
#   (b) the access each existing HANDLE holds, against the share mode YOU grant
#
# A reader therefore only ever probes whichever of the holder's two properties
# its own request puts on the other side of a comparison, and is blind to the
# other. Every reader previously run here asked for READ access, so every one of
# them measured Word's ACCESS via rule (b) and not one could observe its SHARE
# mode -- which is exactly why the share half was asserted in both directions for
# months with nothing ever going red.
#
# The discriminator is a reader asking for WRITE access. That request is checked
# against the holder's share mode under rule (a), so it is the only shape that
# sees the share half at all.
#
# To keep this honest the probe runs the same five readers against two SYNTHETIC
# holders whose properties are known because this script opened them, and which
# differ in exactly one property:
#
#   synthetic-old : write access, grants ReadWrite   <- what PR #22 claimed Word is
#   synthetic-new : write access, grants Read        <- what PR #24 claims Word is
#
# If those two columns come back identical, the probe has separated nothing and
# its verdict on Word is worthless no matter how confident it looks. That check
# is enforced below rather than left to the reader of the output.

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# .NET's current directory is not PowerShell's, so every path here is absolute.
$scratch = Join-Path ([IO.Path]::GetTempPath()) ("fileshare-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $scratch | Out-Null

function Test-Reader {
    param([string]$Path, [string]$Access, [string]$Share)
    try {
        $fs = New-Object System.IO.FileStream(
            $Path,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::$Access,
            [System.IO.FileShare]::$Share)
        $fs.Dispose()
        return 'ok'
    } catch [System.IO.IOException] {
        # Discriminate on the exception TYPE, never the message -- Word and
        # Windows are both German on this machine. A sharing violation is
        # HResult 0x80070020; a lock violation 0x80070021.
        $code = $_.Exception.HResult -band 0xFFFF
        if ($code -eq 32 -or $code -eq 33) { return 'sharing violation' }
        return "IOException($code)"
    } catch [System.UnauthorizedAccessException] {
        return 'access denied'
    }
}

$readers = @(
    @{ label = 'read  / grants ReadWrite  (Node, Copy-Item)'; access = 'Read';      share = 'ReadWrite' }
    @{ label = 'read  / grants Read';                         access = 'Read';      share = 'Read' }
    @{ label = 'read  / grants None';                         access = 'Read';      share = 'None' }
    @{ label = 'WRITE / grants ReadWrite  <- discriminator';  access = 'Write';     share = 'ReadWrite' }
    @{ label = 'WRITE / grants None';                         access = 'Write';     share = 'None' }
)

$results = [ordered]@{}

# --- synthetic holders, whose properties are known by construction ------------

$synthetics = @(
    @{ name = 'synthetic-old (write, grants ReadWrite)'; access = 'Write'; share = 'ReadWrite' }
    @{ name = 'synthetic-new (write, grants Read)';      access = 'Write'; share = 'Read' }
)

foreach ($h in $synthetics) {
    $p = Join-Path $scratch ("holder-" + $h.share + ".bin")
    [IO.File]::WriteAllText($p, 'x')
    $held = New-Object System.IO.FileStream(
        $p,
        [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::$($h.access),
        [System.IO.FileShare]::$($h.share))
    try {
        $col = @()
        foreach ($r in $readers) { $col += Test-Reader -Path $p -Access $r.access -Share $r.share }
        $results[$h.name] = $col
    } finally {
        $held.Dispose()
    }
}

# --- real Word ----------------------------------------------------------------
#
# PID differencing, because three other sessions drive Word on this machine and
# killing one we did not start destroys someone's unsaved work.

$before = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$word = $null
$ownedPid = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    $after = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $new = @($after | Where-Object { $before -notcontains $_ })
    if ($new.Count -eq 1) { $ownedPid = $new[0] }
    Write-Host "started WINWORD pid=$ownedPid (there were $($before.Count) others)"

    $docPath = Join-Path $scratch 'held.docx'
    $doc = $word.Documents.Add()
    $doc.Content.Text = 'held open by Word'
    # 16 = wdFormatDocumentDefault (.docx). Numeric, never a named constant.
    $doc.SaveAs2($docPath, 16)

    if (-not (Test-Path -LiteralPath $docPath)) { throw "Word did not write $docPath" }

    $col = @()
    foreach ($r in $readers) { $col += Test-Reader -Path $docPath -Access $r.access -Share $r.share }
    $results['REAL WORD (document open)'] = $col

    $doc.Close(0)
    [Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
    $doc = $null

    # And once Word has let go, as a positive control that the file itself is not
    # the thing refusing. If this column is not all-ok, every other column is
    # measuring a property of the file rather than of the holder.
    $col = @()
    foreach ($r in $readers) { $col += Test-Reader -Path $docPath -Access $r.access -Share $r.share }
    $results['same file, Word closed it (control)'] = $col
} finally {
    if ($null -ne $word) {
        try { $word.Quit(0) } catch { }
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch { }
        $word = $null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()

        # Quit() returns in ~120ms and the process outlives it by seconds, longer
        # under load. Poll to a deadline; a flat sleep is a guess that goes red on
        # a busy machine and green on a quiet one.
        if ($null -ne $ownedPid) {
            $deadline = (Get-Date).AddSeconds(90)
            while ((Get-Date) -lt $deadline) {
                if ($null -eq (Get-Process -Id $ownedPid -ErrorAction SilentlyContinue)) { break }
                Start-Sleep -Milliseconds 250
            }
            $still = Get-Process -Id $ownedPid -ErrorAction SilentlyContinue
            if ($null -ne $still) { Write-Host "WARNING: pid $ownedPid still alive after 90s" }
            else { Write-Host "pid $ownedPid released" }
        }
    }
    Remove-Item -LiteralPath $scratch -Recurse -Force -ErrorAction SilentlyContinue
}

# --- report -------------------------------------------------------------------

Write-Host ''
$w = ($readers | ForEach-Object { $_.label.Length } | Measure-Object -Maximum).Maximum
foreach ($name in $results.Keys) {
    Write-Host $name
    for ($i = 0; $i -lt $readers.Count; $i++) {
        Write-Host ("  {0}  {1}" -f $readers[$i].label.PadRight($w), $results[$name][$i])
    }
    Write-Host ''
}

# --- the check that makes the verdict mean anything ---------------------------

$old = $results['synthetic-old (write, grants ReadWrite)']
$new = $results['synthetic-new (write, grants Read)']
$word_ = $results['REAL WORD (document open)']

$diff = @()
for ($i = 0; $i -lt $readers.Count; $i++) {
    if ($old[$i] -ne $new[$i]) { $diff += $readers[$i].label }
}

if ($diff.Count -eq 0) {
    Write-Host 'INCONCLUSIVE: the two synthetic holders behave identically under every'
    Write-Host 'reader here, so nothing in this probe can tell the two models apart and'
    Write-Host 'its verdict on real Word would be an artifact. Do not report a result.'
    exit 2
}

Write-Host ("discriminating reader(s): " + ($diff -join '; '))

$matchesOld = -not (Compare-Object $old $word_ -SyncWindow 0)
$matchesNew = -not (Compare-Object $new $word_ -SyncWindow 0)

if ($matchesNew -and -not $matchesOld) {
    Write-Host 'VERDICT: real Word matches synthetic-new -- write access, grants FileShare::Read.'
} elseif ($matchesOld -and -not $matchesNew) {
    Write-Host 'VERDICT: real Word matches synthetic-old -- write access, grants FileShare::ReadWrite.'
} else {
    Write-Host 'VERDICT: real Word matches NEITHER synthetic holder. The model is wrong again.'
}
