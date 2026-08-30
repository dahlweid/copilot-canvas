# Reads ONE line from stdin the same way word-host.ps1 does and reports the
# codepoints it actually received. $env:PROBE_ARM selects the remedy under test.
# Nothing non-ASCII appears in this file; the expectation is passed in as
# codepoints so this probe cannot be fooled by its own file encoding.
#
# This is the one .ps1 in the tree exempt from the "set both console encodings"
# rule that console-encoding.test.mjs enforces, because varying that assignment
# across arms is the whole measurement.
$ErrorActionPreference = 'Stop'

$arm = $env:PROBE_ARM
$applied = 'none'
$setError = ''

try {
    switch ($arm) {
        'utf8-nobom' {
            [Console]::InputEncoding = New-Object System.Text.UTF8Encoding($false)
            $applied = 'UTF8Encoding($false)'
        }
        'utf8-static' {
            [Console]::InputEncoding = [System.Text.Encoding]::UTF8
            $applied = '[Text.Encoding]::UTF8'
        }
        'default' {
            # NOT a remedy arm. `[Text.Encoding]::Default` names a different
            # encoding on each runtime -- the ANSI codepage under Windows
            # PowerShell 5.1, UTF-8 under PowerShell 7 on .NET Core -- so code
            # reaching for it as a "safe" value behaves differently depending on
            # which host launched it, and neither behaviour raises an error.
            # This arm exists to record which one you get, per runtime, rather
            # than to assert what it ought to be.
            [Console]::InputEncoding = [System.Text.Encoding]::Default
            $applied = '[Text.Encoding]::Default'
        }
        default { $applied = 'none (control)' }
    }
}
catch {
    # The one real risk of pinning InputEncoding is that the setter throws when
    # stdin is a redirected pipe rather than a console -- which is exactly how
    # word-host.ps1 and live-word.ps1 are spawned. Captured as a field rather
    # than allowed to kill the run, so the answer is reported either way. The
    # hosts themselves assign unguarded, on the strength of this arm.
    $setError = $_.Exception.GetType().Name + ': ' + $_.Exception.Message
}

$line = [Console]::In.ReadLine()
$req = $line | ConvertFrom-Json
$text = $req.text

$cps = ($text.ToCharArray() | ForEach-Object { "U+{0:X4}" -f [int]$_ }) -join ' '

# The runtime and the two encodings are reported as measurements, not inferred
# by the caller from the executable name: `powershell.exe` is 5.1 and `pwsh` is
# 7.x today, but the claim being recorded is about the runtime, so the runtime is
# what says so.
[pscustomobject]@{
    arm             = if ($arm) { $arm } else { 'control' }
    applied         = $applied
    setError        = $setError
    psVersion       = $PSVersionTable.PSVersion.ToString()
    psEdition       = $PSVersionTable.PSEdition
    defaultWebName  = [System.Text.Encoding]::Default.WebName
    defaultCodePage = [System.Text.Encoding]::Default.CodePage
    inputWebName    = [Console]::InputEncoding.WebName
    inputCodePage   = [Console]::InputEncoding.CodePage
    codepoints      = $cps
} | ConvertTo-Json -Compress
