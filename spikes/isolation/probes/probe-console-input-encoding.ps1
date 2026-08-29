# Reads ONE line from stdin the same way word-host.ps1 does and reports the
# codepoints it actually received. $env:PROBE_ARM selects the remedy under test.
# Nothing non-ASCII appears in this file; the expectation is passed in as
# codepoints so this probe cannot be fooled by its own file encoding.
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
        default { $applied = 'none (control)' }
    }
}
catch {
    $setError = $_.Exception.GetType().Name + ': ' + $_.Exception.Message
}

$line = [Console]::In.ReadLine()
$req = $line | ConvertFrom-Json
$text = $req.text

$cps = ($text.ToCharArray() | ForEach-Object { "U+{0:X4}" -f [int]$_ }) -join ' '

[pscustomobject]@{
    arm      = if ($arm) { $arm } else { 'control' }
    applied  = $applied
    setError = $setError
    codepoints = $cps
} | ConvertTo-Json -Compress
