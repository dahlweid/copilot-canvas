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

# Pin the outbound console encoding before anything is written.
#
# The payload below is base64, which is ASCII and therefore identical under
# every codepage in play -- so this is not fixing a corrupted icon. It is fixing
# the *failure* path, which is where the non-ASCII actually lives: an
# unhandled PowerShell error on this machine is reported in German, and
# `$ErrorActionPreference = 'Stop'` above means a failure surfaces exactly
# that way.
#
# Measured, spawning this script's shape (`-NoProfile -NonInteractive`, stderr
# piped, `setEncoding("utf8")` on the Node side): with the encoding unpinned,
# "Gr[U+00FC][U+00DF]e [U+00C4][U+00D6][U+00DC]" on stderr arrives as
# U+FFFD U+FFFD for every non-ASCII character; with it pinned, it arrives
# intact. Note that this direction fails *lossily* -- U+FFFD cannot be decoded
# back -- unlike the inbound direction, whose OEM mojibake is at least
# reversible. A diagnostic is the one thing that must survive the case where
# everything else did not.
#
# This is also what makes the rule in `console-encoding.test.mjs` true of the
# whole shipped surface rather than of stdin-readers only: a script that pins
# nothing because its current payload happens to be ASCII is one edit away from
# a lossy channel, and nothing about the ASCII would have to change for that
# edit to be wrong.
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)

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
