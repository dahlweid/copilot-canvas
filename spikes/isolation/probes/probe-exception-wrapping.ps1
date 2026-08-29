# When does PowerShell 5.1 wrap a .NET exception, and when does it not?
#
# Two probes in this repo read exceptions from a failed FileStream open and
# reached opposite conclusions about what lands in $_.Exception:
#
#   probe-word-share-mode.ps1   walks InnerException before classifying, and its
#                               comment claimed PowerShell wraps for New-Object,
#                               [Type]::new() and a static call "alike -- there
#                               is no construction style that avoids it".
#   probe-fileshare-algebra.ps1 reads $_.Exception.HResult directly, and prints
#                               correct 'sharing violation' labels, which is only
#                               possible if the exception arrived UNWRAPPED.
#
# Both are right, and neither claim generalises. The discriminator is not the
# construction style and not the scope of the catch -- both were measured flat.
# It is the INTERACTION of the call kind with whether the catch is TYPED:
#
#                          catch [System.IO.IOException]   catch { }
#   [IO.File]::Open  (method)   UNWRAPPED IOException      MethodInvocationException
#   New-Object       (cmdlet)   MethodInvocationException  MethodInvocationException
#
# A typed catch matches in every cell, because PowerShell tests the INNER type.
# So the catch firing is never evidence that $_.Exception is the type named.
#
# The dangerous cell is New-Object + typed catch: the catch matches, so the code
# looks like it is working, while $_.Exception is the wrapper and .HResult is the
# generic 0x80131501 (low word 5377). Any switch on .HResult there silently takes
# its default branch forever.
#
# Consequence for instrumentation: neither "read .HResult directly" nor "always
# walk InnerException" is correct in general. Walking is correct everywhere --
# GetBaseException() on an unwrapped exception returns itself -- so it is the one
# shape that is safe without knowing which cell you are in.
#
# Exits 0 if the matrix above reproduces, 1 if this machine disagrees with it.
#
# Run:  powershell -File probe-exception-wrapping.ps1

$ErrorActionPreference = 'Stop'

$path = [IO.Path]::Combine($env:TEMP, "exception-wrapping-$PID.bin")
[IO.File]::WriteAllText($path, 'x')

$results = @{}

function Record($cell, $ex, $matchedTyped) {
    $script:results[$cell] = [pscustomobject]@{
        Cell         = $cell
        Type         = $ex.GetType().Name
        HResult      = '0x{0:X8}' -f $ex.HResult
        Inner        = if ($null -ne $ex.InnerException) { $ex.InnerException.GetType().Name } else { '(none)' }
        Base         = $ex.GetBaseException().GetType().Name
        BaseHResult  = '0x{0:X8}' -f $ex.GetBaseException().HResult
        TypedMatched = $matchedTyped
    }
}

$holder = $null
try {
    # A holder that grants nothing, so every reader below fails the same way.
    $holder = New-Object System.IO.FileStream(
        $path,
        [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::ReadWrite,
        [System.IO.FileShare]::None)

    try {
        $fs = [System.IO.File]::Open($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
        $fs.Dispose()
    } catch [System.IO.IOException] { Record 'method + typed'    $_.Exception $true }
      catch                         { Record 'method + typed'    $_.Exception $false }

    try {
        $fs = [System.IO.File]::Open($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
        $fs.Dispose()
    } catch { Record 'method + untyped' $_.Exception 'n/a' }

    try {
        $fs = New-Object System.IO.FileStream($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
        $fs.Dispose()
    } catch [System.IO.IOException] { Record 'cmdlet + typed'    $_.Exception $true }
      catch                         { Record 'cmdlet + typed'    $_.Exception $false }

    try {
        $fs = New-Object System.IO.FileStream($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
        $fs.Dispose()
    } catch { Record 'cmdlet + untyped' $_.Exception 'n/a' }
} finally {
    if ($null -ne $holder) { $holder.Dispose() }
    Remove-Item $path -Force -ErrorAction SilentlyContinue
}

Write-Output ''
Write-Output ("PSVersion: {0}" -f $PSVersionTable.PSVersion.ToString())
Write-Output ''
Write-Output ("{0,-18} {1,-28} {2,-12} {3,-26} {4,-14} {5}" -f 'cell', '$_.Exception', 'HResult', 'InnerException', 'GetBaseException', 'typed catch matched')
Write-Output ("-" * 130)
foreach ($cell in 'method + typed', 'method + untyped', 'cmdlet + typed', 'cmdlet + untyped') {
    $r = $results[$cell]
    Write-Output ("{0,-18} {1,-28} {2,-12} {3,-26} {4,-14} {5}" -f $r.Cell, $r.Type, $r.HResult, $r.Inner, $r.Base, $r.TypedMatched)
}

# The recorded finding, asserted rather than described, so this probe fails on a
# machine where it does not hold instead of printing a table nobody re-reads.
$expected = @{
    'method + typed'   = 'IOException'
    'method + untyped' = 'MethodInvocationException'
    'cmdlet + typed'   = 'MethodInvocationException'
    'cmdlet + untyped' = 'MethodInvocationException'
}

$disagreements = @()
foreach ($cell in $expected.Keys) {
    if ($results[$cell].Type -ne $expected[$cell]) {
        $disagreements += ("{0}: expected {1}, got {2}" -f $cell, $expected[$cell], $results[$cell].Type)
    }
}

# GetBaseException is the shape that is correct in every cell; if that ever stops
# being true the "just always walk" advice above is wrong and must be retracted.
foreach ($cell in $expected.Keys) {
    if ($results[$cell].BaseHResult -ne '0x80070020') {
        $disagreements += ("{0}: GetBaseException().HResult was {1}, not the sharing violation 0x80070020" -f $cell, $results[$cell].BaseHResult)
    }
}

Write-Output ''
if ($disagreements.Count -gt 0) {
    Write-Output 'DISAGREES with the recorded matrix:'
    $disagreements | ForEach-Object { Write-Output ("  - " + $_) }
    exit 1
}

Write-Output 'Matrix reproduces. Wrapping depends on call kind AND typed-ness, not on either alone.'
Write-Output 'GetBaseException() reached the sharing violation in all four cells.'
exit 0
