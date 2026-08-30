# probe-strict-lock.ps1 -- holds a file the way Word does *not*, for N seconds.
#
# The positive control for probe-close-during-retry.mjs arm 2. An arm that
# counts copy failures and reports zero is worthless unless the counter is shown
# capable of counting one -- "all arms agree" is this repo's oldest documented
# defect, and a dead instrument reads exactly like a clean result.
#
# `FileShare::None` is the share mode `word-host.ps1` names in its own measured
# table as the one that produces `System.IO.IOException` from `Copy-Item`. It is
# deliberately *not* how Word holds a document: Word takes write access and
# grants `FileShare::Read`, which a reader granting ReadWrite passes.
#
# Prints "locked" once the handle is held, so the caller can start the clock on
# a fact rather than on a sleep.
param(
    [Parameter(Mandatory = $true)][string]$Path,
    [int]$Seconds = 3
)

$ErrorActionPreference = 'Stop'

$stream = [System.IO.File]::Open(
    $Path,
    [System.IO.FileMode]::Open,
    [System.IO.FileAccess]::Read,
    [System.IO.FileShare]::None)
try {
    Write-Output 'locked'
    Start-Sleep -Seconds $Seconds
} finally {
    $stream.Dispose()
}
