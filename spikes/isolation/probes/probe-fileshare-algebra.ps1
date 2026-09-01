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

. (Join-Path $PSScriptRoot '_common.ps1')

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
        # Do not label every IOException a sharing violation. This probe is
        # cited as evidence, so a mislabelled row would not fail -- it would
        # fabricate a measurement, which is worse. Only ERROR_SHARING_VIOLATION
        # (32) and ERROR_LOCK_VIOLATION (33) are the thing under test; anything
        # else surfaces its HRESULT so it cannot be mistaken for a result.
        $hr = $_.Exception.HResult
        switch ($hr) {
            0x80070020 { return 'sharing violation' }
            0x80070021 { return 'lock violation' }
            default    { return ('IOException 0x{0:X8}' -f $hr) }
        }
    } catch {
        return $_.Exception.GetType().Name
    }
}

$readers = @(
    @{ Label = 'read, grants ReadWrite  (Copy-Item / Node) '; Access = 'Read'; Share = 'ReadWrite' },
    @{ Label = 'read, grants Read       (ACCESS discrim.)  '; Access = 'Read'; Share = 'Read' },
    @{ Label = 'read, grants None                          '; Access = 'Read'; Share = 'None' },
    @{ Label = 'ReadWrite, grants None (Test-FileWritable) '; Access = 'ReadWrite'; Share = 'None' },
    @{ Label = 'write, grants ReadWrite (SHARE discrim.)   '; Access = 'Write'; Share = 'ReadWrite' }
)

Write-Output ''
Write-Output '=== PART A: sharing algebra against a synthetic holder ==='
Write-Output ''

# The three models this claim has passed through -- the third of which is also
# the configuration read-smoke.mjs now uses as its stand-in for Word -- plus the
# row that proves the model's own limit: a holder differing from row 3 in access
# alone, share mode fixed.
#
# That fourth row is not bookkeeping, but it proves something narrower than it
# used to claim. It shows that a WRITE holder and a READWRITE holder granting
# the same share mode are indistinguishable -- so a reader cannot tell whether
# a holder that writes *also* reads. It does not show that access is invisible.
# A read-requesting reader that grants Read sees the holder's access through
# rule (b), and that is how "Word's access includes write" is measured rather
# than inferred. What stays unobservable is only the read half of it. Rows 3
# and 4 coming back byte-identical is the demonstration, asserted below rather
# than left for a reader to notice. If those rows ever diverge, Write and
# ReadWrite have become distinguishable and the shorthand needs revisiting --
# so the probe says so instead of printing a table nobody diffs.
#
# This paragraph previously read "a holder's ACCESS mode cannot be observed
# from outside at all ... only 'grants Read' is measured." Both halves were
# backwards, and they were backwards in the direction that made the then-current
# model unfalsifiable: if access is unobservable, no measurement can contradict
# a claim about it.
#
# Two of these readers discriminate, and they discriminate different things.
# Windows checks the access you request against the holder's SHARE mode, and the
# holder's ACCESS against the share mode you offer. So a read-requesting reader
# only ever exercises the second check and measures the holder's access; it is
# blind to the share mode. A write-requesting reader that grants ReadWrite
# inverts that and is the only shape here that sees the share half.
#
# Until the last row of $readers existed, every reader in this file asked for
# read access, so the share half could be -- and was -- asserted in either
# direction with nothing going red.
$holders = @(
    @{ Key = 'r-read';  Label = 'holder: READ  access, grants Read      (the original claim)      '; Access = 'Read';      Share = 'Read' },
    @{ Key = 'w-rw';    Label = 'holder: WRITE access, grants ReadWrite (the first correction)    '; Access = 'Write';     Share = 'ReadWrite' },
    @{ Key = 'w-read';  Label = 'holder: WRITE access, grants Read      (WHAT WORD ACTUALLY DOES) '; Access = 'Write';     Share = 'Read' },
    @{ Key = 'rw-read'; Label = 'holder: READWRITE access, grants Read  (WRITE VS READWRITE ALIKE)'; Access = 'ReadWrite'; Share = 'Read' }
)

# Per-holder result vectors, so the identity above can be asserted rather than
# eyeballed.
$vectors = @{}
$script:identityBroken = $false

foreach ($holder in $holders) {
    $target = Join-Path $scratch ("holder-" + [guid]::NewGuid().ToString('N').Substring(0, 6) + '.bin')
    Set-Content -LiteralPath $target -Value 'payload' -Encoding ascii

    $handle = [System.IO.File]::Open(
        $target,
        [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::$($holder.Access),
        [System.IO.FileShare]::$($holder.Share))

    $vector = @()
    Write-Output $holder.Label
    foreach ($r in $readers) {
        $result = Try-Open -Path $target -Access $r.Access -Share $r.Share
        $vector += $result
        Write-Output ("    {0} -> {1}" -f $r.Label, $result)
    }

    $copyDest = Join-Path $scratch ('copy-' + [guid]::NewGuid().ToString('N').Substring(0, 6) + '.bin')
    try {
        Copy-Item -LiteralPath $target -Destination $copyDest -ErrorAction Stop
        $vector += 'ok'
        Write-Output '    Copy-Item                                  -> ok'
    } catch {
        $vector += $_.Exception.GetType().Name
        Write-Output ("    Copy-Item                                  -> {0}" -f $_.Exception.GetType().Name)
    }

    $vectors[$holder.Key] = $vector
    $handle.Dispose()
    Write-Output ''
}

# The assertion the fourth row exists for. Two holders differing only in access
# mode must be indistinguishable to every reader; if they are not, the shorthand
# "Word holds a write handle" has become more precise than the evidence, since
# it would then be claiming the read half too.
#
# Columns are named, not indexed. An index is the failure this file is about in
# miniature: "cell 4" makes the reader reconstruct which column that was from
# another part of the script, and a reconstruction is exactly what an assertion
# is supposed to remove.
$columnLabels = @($readers | ForEach-Object { $_.Label.Trim() }) + @('Copy-Item')
$wRead  = $vectors['w-read']
$rwRead = $vectors['rw-read']
$differences = @()
for ($i = 0; $i -lt $wRead.Count; $i++) {
    if ($wRead[$i] -ne $rwRead[$i]) {
        $differences += ("{0}: WRITE holder -> {1} ; READWRITE holder -> {2}" -f $columnLabels[$i], $wRead[$i], $rwRead[$i])
    }
}
if ($differences.Count -eq 0) {
    Write-Output ("ACCESS-MODE IDENTITY HOLDS: WRITE and READWRITE holders granting Read are identical across all {0} readers plus Copy-Item." -f $readers.Count)
    # Single-quoted so the strings stay free of interpolation by construction;
    # they contain `*` and `/` that would otherwise invite it.
    Write-Output '  => a holder that writes cannot be told apart from one that reads and writes.'
    Write-Output '     Its access is still observable as *including write* -- that is what the'
    Write-Output '     read/grants-Read column measures. Only the read half is invisible.'
} else {
    # Recorded rather than thrown, so PART B still runs and reports against real
    # Word -- but the run must not exit 0, or an assertion nobody reads is worth
    # no more than the table nobody diffed.
    $script:identityBroken = $true
    Write-Output 'ACCESS-MODE IDENTITY BROKEN -- the inference this repo relies on no longer holds:'
    foreach ($d in $differences) { Write-Output ("  {0}" -f $d) }
}
Write-Output ''

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
    # $startedPids is a census DIFFERENCE, not an ownership fact. Differencing is
    # measured unsound here (#136): probe-init-attribution.ps1 got 2 new pids for
    # 1 instance created, and a census control saw 2 strangers' WINWORDs appear
    # in a 40 s window with nothing launched. The old comment here said "only
    # ever reap a Word this probe started", which is the claim the measurement
    # falsifies.
    #
    # The wait stays and is still worth its 20 s: Quit() returns in 3-28 ms
    # (measured, spikes/isolation/probes/probe-quit-exit-gap.ps1) while the
    # process lives on for seconds, so a survivor observed without the wait is a
    # stopwatch artefact. What follows the wait is now a report. The earlier run
    # that left a WINWORD alive past the full 20 s is still visible -- it is just
    # reported instead of resolved by killing a process we cannot identify.
    foreach ($p in $startedPids) {
        $proc = Get-Process -Id $p -ErrorAction SilentlyContinue
        if ($proc) {
            $proc.WaitForExit(20000) | Out-Null
        }
    }
    $survivors = @($startedPids | Where-Object { Get-Process -Id $_ -ErrorAction SilentlyContinue })
    Write-CensusSurvivors $survivors
    Write-Output ''
    Write-Output ("WINWORD that appeared during this probe: {0}" -f ($startedPids -join ', '))
    Remove-Item -Recurse -Force $scratch -ErrorAction SilentlyContinue
}

# Placed after the finally so PART B's Word is always reaped first: an exit that
# skips cleanup would trade a reported failure for a leaked WINWORD.
if ($script:identityBroken) {
    Write-Output ''
    Write-Output 'FAILED: access-mode identity broken (see PART A).'
    exit 1
}
