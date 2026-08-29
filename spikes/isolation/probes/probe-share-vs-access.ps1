# Does a holder's ACCESS mode or its SHARE mode cause a sharing violation?
#
# Context. ADR 0006 records that of "Word holds a write handle granting
# ReadWrite", the *share* half is measured and the *access* half is not
# observable from outside at all. This probe was written to check that, and
# measures the opposite on both counts:
#
#   * the discriminating reader already in the repo measures ACCESS, not share;
#   * nothing in the repo measures SHARE, because every reader it uses asks for
#     read access;
#   * and when a reader that *can* measure the share half is finally pointed at
#     real Word, Word grants `Read` -- not `ReadWrite`.
#
# The mechanism. Windows checks two things on every open:
#
#   (a) the ACCESS you request, against the SHARE mode of each existing handle
#   (b) the ACCESS of each existing handle, against the SHARE mode you offer
#
# A reader therefore probes whichever of the holder's two properties its own
# request puts on the other side of the comparison, and is blind to the other:
#
#   reader access=Read,  share=Read      -> rule (b) -> measures holder ACCESS
#   reader access=Write, share=ReadWrite -> rule (a) -> measures holder SHARE
#
# The three synthetic holders below differ in exactly one property per adjacent
# pair, which is what makes them discriminating. A probe whose cases all agree
# has measured nothing -- the lesson from `probe-errno-mapping.mjs`, which
# modelled Word with a holder granting Read and could not have detected its own
# mispremise. This probe is that lesson applied to the correction itself.
#
# Run:  powershell -NoProfile -ExecutionPolicy Bypass -File probe-share-vs-access.ps1 [-Doc <path.docx>]
#
# Omit -Doc to run only the synthetic controls, which need no Word.

param([string]$Doc)

$ErrorActionPreference = "Stop"

function Try-Open {
    param([string]$Path, [System.IO.FileAccess]$Access, [System.IO.FileShare]$Share)
    try {
        $s = [IO.File]::Open($Path, [IO.FileMode]::Open, $Access, $Share)
        $s.Close()
        return "ok"
    } catch [System.IO.IOException] {
        # Specifically a sharing violation. A read-only attribute or a denying
        # ACL surfaces as UnauthorizedAccessException and lands in the catch
        # below, so the two are never confused for one another.
        return "violation"
    } catch {
        return "other: $($_.Exception.GetType().Name)"
    }
}

function Probe([string]$Path) {
    [pscustomobject]@{
        "read/Read (ACCESS)"      = Try-Open -Path $Path -Access Read  -Share Read
        "write/ReadWrite (SHARE)" = Try-Open -Path $Path -Access Write -Share ReadWrite
        "read/ReadWrite (copy)"   = Try-Open -Path $Path -Access Read  -Share ReadWrite
    }
}

$rows = @()

# --- Synthetic controls, where both properties are known by construction ------
$ctl = Join-Path $env:TEMP "probe-share-vs-access.bin"
Set-Content -LiteralPath $ctl -Value "control"

# The free-file row matters: it proves the readers themselves can succeed on
# this file, so every violation below is contention and not an attribute.
$rows += (Probe $ctl | Add-Member holder "no holder (file free)" -PassThru)

$holders = @(
    @{ label = "access=Read      grants Read     "; access = [IO.FileAccess]::Read;      share = [IO.FileShare]::Read }
    @{ label = "access=Read      grants ReadWrite"; access = [IO.FileAccess]::Read;      share = [IO.FileShare]::ReadWrite }
    @{ label = "access=ReadWrite grants ReadWrite"; access = [IO.FileAccess]::ReadWrite; share = [IO.FileShare]::ReadWrite }
)
foreach ($h in $holders) {
    $handle = [IO.File]::Open($ctl, [IO.FileMode]::Open, $h.access, $h.share)
    $rows += (Probe $ctl | Add-Member holder $h.label -PassThru)
    $handle.Close()
}
Remove-Item -LiteralPath $ctl -Force

# --- Real Word ----------------------------------------------------------------
if ($Doc) {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $opened = $null
    try {
        $opened = $word.Documents.Open($Doc)
        $rows += (Probe $Doc | Add-Member holder "REAL WORD" -PassThru)
    } finally {
        if ($null -ne $opened) { $opened.Close(0) }
        $word.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
}

$rows |
    Select-Object holder, "read/Read (ACCESS)", "write/ReadWrite (SHARE)", "read/ReadWrite (copy)" |
    Format-Table -AutoSize |
    Out-String -Width 220

# Measured 2026-08-29, Word 16 (de-DE), twice:
#
#   holder                            read/Read   write/ReadWrite   read/ReadWrite
#   no holder (file free)             ok          ok                ok
#   access=Read      grants Read      ok          violation         ok
#   access=Read      grants ReadWrite ok          ok                ok
#   access=ReadWrite grants ReadWrite violation   ok                ok
#   REAL WORD                         violation   violation         ok
#
# Word matches no synthetic row: it is the missing fourth combination, write
# access granting `Read`. The first column separates it from a read-access
# holder; the second separates it from a ReadWrite-granting one. Neither column
# alone identifies it, which is exactly why one column alone kept producing
# confident and half-wrong answers.
