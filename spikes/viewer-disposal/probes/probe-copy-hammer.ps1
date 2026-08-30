# probe-copy-hammer.ps1 -- repeatedly copies a file and reports what refused.
#
# The instrument for arm 2 of probe-close-during-retry.mjs. It must be
# `Copy-Item` and nothing else: `Copy-Item` is *literally the call* that raises
# "Another process is holding <name> open more strictly than Word does" in
# `word-host.ps1` (Open-Doc's copy of the original). A reader written in Node
# would grant FILE_SHARE_READ|WRITE|DELETE and would therefore be a strictly
# weaker instrument than the one whose failure we are trying to reproduce.
#
# Reports the exception *type*, never the message -- messages are localized on
# this machine, and matching them is the trap this project has paid for twice.
# The type is also what `word-host.ps1` itself branches on, so a failure counted
# here maps onto the code that would be raised there.
#
# Emits one JSON object on stdout so the caller parses rather than scrapes.
param(
    [Parameter(Mandatory = $true)][string]$Src,
    [Parameter(Mandatory = $true)][string]$Dst,
    [int]$Seconds = 10,
    [int]$IntervalMs = 50
)

$ErrorActionPreference = 'Stop'

$attempts = 0
$failures = @()
$deadline = (Get-Date).AddSeconds($Seconds)

while ((Get-Date) -lt $deadline) {
    $attempts++
    try {
        Copy-Item -LiteralPath $Src -Destination $Dst -Force -ErrorAction Stop
    } catch {
        $failures += [pscustomobject]@{
            at   = (Get-Date).ToString('o')
            type = $_.Exception.GetType().FullName
        }
    }
    Start-Sleep -Milliseconds $IntervalMs
}

[pscustomobject]@{
    attempts = $attempts
    failures = @($failures)
} | ConvertTo-Json -Depth 4 -Compress
