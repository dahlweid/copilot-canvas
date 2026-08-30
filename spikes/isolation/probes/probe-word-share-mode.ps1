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

param(
    # Bounded, because the first version of this probe called COM inline and hung
    # indefinitely with 16 concurrent WINWORD.EXE on the machine. Generous rather
    # than tight: Word's startup tail is load-dependent and is not bounded by
    # timings taken on an idle machine.
    [int]$StartupTimeoutSeconds = 120,
    [int]$ShutdownTimeoutSeconds = 120
)

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
    } catch {
        # The innermost exception, not $_.Exception. This call is `New-Object`
        # under an untyped catch, and in that cell PowerShell wraps what was
        # thrown in a MethodInvocationException whose own HResult is the generic
        # 0x80131501 -- so reading .HResult here without walking would report
        # every sharing violation as 5377, the low word of the wrapper.
        #
        # This comment used to add "and it does so for New-Object, [Type]::new()
        # and [IO.File]::Open alike -- there is no construction style that avoids
        # it". That was measured under an untyped catch and then stated as though
        # it held everywhere, which it does not: probe-fileshare-algebra.ps1
        # reads .HResult directly, from a static call under a TYPED catch, and is
        # correct to. Wrapping depends on the call kind AND the typed-ness of the
        # catch, not on either alone. The 2x2 is in
        # spikes/isolation/probes/probe-exception-wrapping.ps1.
        #
        # What does hold everywhere, and is the reason this walk stays: `catch
        # [System.IO.IOException]` matches in every cell, because PowerShell
        # tests the INNER type. So a typed catch firing is never evidence that
        # $_.Exception is the type it named -- which is what made the original
        # bug so easy to miss. The catch fired, the classification looked like it
        # was working, and the label was wrong.
        $root = $_.Exception
        while ($null -ne $root.InnerException) { $root = $root.InnerException }

        if ($root -is [System.UnauthorizedAccessException]) { return 'access denied' }
        if ($root -is [System.IO.IOException]) {
            # 32 = ERROR_SHARING_VIOLATION, 33 = ERROR_LOCK_VIOLATION. Discriminate
            # on the code, never the message: Windows is German on this machine.
            $code = $root.HResult -band 0xFFFF
            if ($code -eq 32 -or $code -eq 33) { return 'sharing violation' }
            return "IOException($code)"
        }
        return $root.GetType().Name
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
# Word runs in a child powershell.exe rather than in this process, for two
# reasons. First, this repo's standing rule: any COM call that can hang must be
# timeout-bounded, and Documents.Add / SaveAs2 have both wedged here under load
# -- the first version of this probe called them inline and hung indefinitely
# with 16 concurrent WINWORD.EXE on the machine. A child process can be waited on
# with a deadline; an in-process COM call cannot. Second, the readers must run
# *while* the document is held, which needs the holder to be somewhere else.
#
# Every path is passed as a discrete -Parameter value. Nothing is interpolated
# into a command line: a path through a PowerShell single-quoted literal breaks
# on an apostrophe, and the same path through cmd.exe is corrupted by &, ^ or a
# matched %VAR% pair.

$holder = Join-Path $PSScriptRoot 'probe-word-share-mode-holder.ps1'
if (-not (Test-Path -LiteralPath $holder)) { throw "holder script missing: $holder" }

$docPath   = Join-Path $scratch 'held.docx'
$readyFile = Join-Path $scratch 'ready'
$goFile    = Join-Path $scratch 'go'
$pidFile   = Join-Path $scratch 'pid'

$proc = Start-Process -FilePath 'powershell.exe' -PassThru -WindowStyle Hidden -ArgumentList @(
    '-NoProfile', '-ExecutionPolicy', 'Bypass',
    '-File', $holder,
    '-DocPath', $docPath,
    '-ReadyFile', $readyFile,
    '-GoFile', $goFile,
    '-PidFile', $pidFile
)

try {
    # Bounded wait for the holder to have the document open.
    $ready = $false
    $deadline = (Get-Date).AddSeconds($StartupTimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        if (Test-Path -LiteralPath $readyFile) { $ready = $true; break }
        if ($proc.HasExited) { break }
        Start-Sleep -Milliseconds 200
    }

    if (-not $ready) {
        Write-Host ''
        Write-Host ("INCONCLUSIVE: the holder did not open a document within {0}s." -f $StartupTimeoutSeconds)
        Write-Host 'This machine routinely carries a dozen WINWORD.EXE from other sessions and'
        Write-Host 'Word contends badly under that load. No verdict is reported, because a probe'
        Write-Host 'that could not create its own precondition has measured nothing.'
        exit 3
    }

    $ownedPid = (Get-Content -LiteralPath $pidFile -Raw -ErrorAction SilentlyContinue)
    Write-Host ("holder ready; its WINWORD pid is " + $(if ([string]::IsNullOrWhiteSpace($ownedPid)) { '<ambiguous, not claimed>' } else { $ownedPid.Trim() }))

    if (-not (Test-Path -LiteralPath $docPath)) { throw "holder reported ready but $docPath is absent" }

    $col = @()
    foreach ($r in $readers) { $col += Test-Reader -Path $docPath -Access $r.access -Share $r.share }
    $results['REAL WORD (document open)'] = $col
} finally {
    [IO.File]::WriteAllText($goFile, 'go')
    if (-not $proc.WaitForExit($ShutdownTimeoutSeconds * 1000)) {
        Write-Host "WARNING: holder did not exit within ${ShutdownTimeoutSeconds}s; not killing anything."
    }
    # The holder runs with its console hidden, so this note file is the only
    # channel by which its Quit outcome can reach anyone. Printing it here is
    # what stops the holder's reporting catch from being a swallow.
    $note = "$pidFile.note"
    if (Test-Path -LiteralPath $note) {
        Write-Host ("holder shutdown: " + (Get-Content -LiteralPath $note -Raw).Trim())
    } else {
        Write-Host "holder shutdown: no note written -- it did not reach its cleanup block."
    }
}

# And once Word has let go, as a positive control that the file itself is not the
# thing refusing. If this column is not all-ok, every other column is measuring a
# property of the file rather than of the holder.
if (Test-Path -LiteralPath $docPath) {
    $col = @()
    foreach ($r in $readers) { $col += Test-Reader -Path $docPath -Access $r.access -Share $r.share }
    $results['same file, Word closed it (control)'] = $col
}

Remove-Item -LiteralPath $scratch -Recurse -Force -ErrorAction SilentlyContinue

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
