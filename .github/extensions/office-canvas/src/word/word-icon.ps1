# Extracts the Word application icon from the Word installed on this machine and
# writes it to stdout as base64 PNG.
#
# Takes no parameters, and that is the point rather than an omission. A path
# handed to PowerShell as an interpolated argument breaks on an apostrophe
# (C:\Users\O'Brien\...) and a path handed through cmd breaks on & or a matched
# %VAR% pair; the repo's rule is to remove the parser rather than escape it. So
# this script finds WINWORD.EXE itself and no caller-supplied string is ever
# parsed as code.
#
# Redistributes nothing: the bytes come from the user's own installation and
# never enter the repository. That constraint is the whole reason this exists --
# Microsoft's marks are absent from the open icon sets for trademark reasons.
# Measured in spikes/word-icon/probes/probe-icon-sources.mjs.
#
# Does not launch Word: ExtractAssociatedIcon reads the executable's resources.
# Measured by the same probe, which counts WINWORD.EXE either side (14 -> 14).

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Where the installers put it, newest layout first. The registry App Paths entry
# is consulted first because it is what the shell itself resolves `winword` by,
# so it survives a non-default install directory that a hardcoded list cannot.
function Find-Winword {
    foreach ($hive in 'HKLM:\SOFTWARE', 'HKLM:\SOFTWARE\WOW6432Node') {
        $key = Join-Path $hive 'Microsoft\Windows\CurrentVersion\App Paths\winword.exe'
        try {
            $registered = (Get-ItemProperty -LiteralPath $key -ErrorAction Stop).'(default)'
        } catch {
            continue
        }
        if ($registered -and (Test-Path -LiteralPath $registered -PathType Leaf)) { return $registered }
    }

    foreach ($root in $env:ProgramFiles, ${env:ProgramFiles(x86)}) {
        if (-not $root) { continue }
        foreach ($leaf in 'Microsoft Office\Root\Office16\WINWORD.EXE',
                          'Microsoft Office\Office16\WINWORD.EXE',
                          'Microsoft Office\Office15\WINWORD.EXE') {
            $candidate = Join-Path $root $leaf
            if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }
        }
    }

    return $null
}

$exe = Find-Winword
if (-not $exe) {
    # A machine with no Word is an ordinary state, not a failure: the viewer
    # renders a PDF and only the decoration is missing. Say so on stderr and
    # exit non-zero so the caller can tell "no Word" from "empty output".
    [Console]::Error.WriteLine('winword-not-found')
    exit 2
}

Add-Type -AssemblyName System.Drawing

$icon = [System.Drawing.Icon]::ExtractAssociatedIcon($exe)
try {
    $bitmap = $icon.ToBitmap()
    try {
        $stream = New-Object System.IO.MemoryStream
        try {
            $bitmap.Save($stream, [System.Drawing.Imaging.ImageFormat]::Png)
            [Console]::Out.Write([Convert]::ToBase64String($stream.ToArray()))
        } finally {
            $stream.Dispose()
        }
    } finally {
        $bitmap.Dispose()
    }
} finally {
    $icon.Dispose()
}
