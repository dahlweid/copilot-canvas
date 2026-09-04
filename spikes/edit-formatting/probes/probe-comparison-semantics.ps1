# probe-comparison-semantics.ps1 -- why Set-ParagraphText's diff compares
# ORDINALLY, not with -eq or -ceq. Office-free: this measures PowerShell's own
# string/char comparison operators, so it starts no Word and needs no Office.
#
# Source is deliberately all-ASCII: PowerShell 5.1 reads a BOM-less .ps1 as the
# system ANSI code page, so a literal eszett or a-umlaut in an executable string
# is mis-decoded and breaks the parse. Every special character below is built
# with [char]0x.. so the file runs identically whatever the runner's code page.
#
# The fix in word-host.ps1 diffs the old visible text against the new to find the
# changed middle span, and returns early when they are equal. Both of PowerShell's
# built-in equality operators are wrong for that job in the German locale this
# repo runs under, and each wrong in a way that silently drops a real edit -- the
# exact #170 failure, reintroduced:
#
#   -eq   is case-INSENSITIVE. "github" -eq "GitHub" is True, so a case-only edit
#         compares equal and the fix writes nothing.
#   -ceq  is case-sensitive but still CULTURE-sensitive. Under de-DE the eszett
#         (U+00DF) folds to "ss", so "strasse" -ceq "stra<eszett>e" is True, and a
#         real German edit strasse<->strasse-with-eszett compares equal and
#         vanishes too.
#
# StringComparison::Ordinal folds neither. This probe prints both traps and shows
# ordinal getting them right, then sweeps the Latin range to establish that a
# char-level -ceq happens to coincide with ordinal there (the scans do not depend
# on that -- they compare [int] code units -- but it is worth recording).
#
# Forced to de-DE so the result does not depend on the runner's regional settings.
$ErrorActionPreference = 'Stop'
$deDE = [System.Globalization.CultureInfo]::GetCultureInfo('de-DE')
[System.Threading.Thread]::CurrentThread.CurrentCulture = $deDE
[System.Threading.Thread]::CurrentThread.CurrentUICulture = $deDE

Write-Host "PSVersion: $($PSVersionTable.PSVersion)   Culture: $([System.Threading.Thread]::CurrentThread.CurrentCulture.Name)"
Write-Host ''

# Distinct variable names throughout: PowerShell variable names are themselves
# case-insensitive, so $a and $A are ONE variable -- reusing them to hold 'a' and
# 'A' silently overwrites, which is its own instance of the bug under study.
$eszett  = [string][char]0x00DF   # the eszett, folds to "ss" under de-DE
$strasse = 'strasse'
$strasze = 'stra' + $eszett + 'e' # the same word spelled with the eszett

Write-Host '=== string-level: the early-return  if ($old -<op> $text) { return } ==='
Write-Host ("  -eq   github vs GitHub            : {0}   (case-insensitive; drops a case-only edit)" -f ('github' -eq 'GitHub'))
Write-Host ("  -ceq  strasse vs eszett-spelling  : {0}   (culture fold; drops a German edit)" -f ($strasse -ceq $strasze))
Write-Host ("  Ordinal strasse vs eszett         : {0}   (correct: they differ, so the edit proceeds)" -f ([string]::Equals($strasse, $strasze, [System.StringComparison]::Ordinal)))
Write-Host ("  Ordinal github vs GitHub          : {0}   (correct: they differ, so the edit proceeds)" -f ([string]::Equals('github', 'GitHub', [System.StringComparison]::Ordinal)))
Write-Host ("  Ordinal identical text            : {0}   (correct: a true no-op still returns early)" -f ([string]::Equals($strasse, 'strasse', [System.StringComparison]::Ordinal)))
Write-Host ''

Write-Host '=== char-level: the prefix/suffix scans compare [int] code units ==='
$lowerA = [char]0x0061; $upperA = [char]0x0041; $lowerAuml = [char]0x00E4; $upperAuml = [char]0x00C4
Write-Host ("  [int]lower-a  -eq [int]upper-A     : {0}   (expects False -- distinct code units)" -f ([int]$lowerA -eq [int]$upperA))
Write-Host ("  [int]lower-a  -eq [int]lower-a     : {0}   (expects True)" -f ([int]$lowerA -eq [int][char]0x0061))
Write-Host ("  [int]a-umlaut -eq [int]a-umlaut    : {0}   (expects True)" -f ([int]$lowerAuml -eq [int][char]0x00E4))
Write-Host ("  [int]a-umlaut -eq [int]A-umlaut    : {0}   (expects False)" -f ([int]$lowerAuml -eq [int]$upperAuml))
Write-Host ''

Write-Host '=== sweep: does char -ceq ever diverge from code-unit equality over U+0000..U+024F? ==='
$divergences = 0
for ($i = 0; $i -le 0x24F; $i++) {
    for ($j = 0; $j -le 0x24F; $j++) {
        $ceq = ([char]$i -ceq [char]$j)
        $ordinal = ($i -eq $j)
        if ($ceq -ne $ordinal) {
            $divergences++
            if ($divergences -le 10) { Write-Host ("  DIVERGE U+{0:X4} vs U+{1:X4}: -ceq={2} ordinal={3}" -f $i, $j, $ceq, $ordinal) }
        }
    }
}
Write-Host ("  divergences: {0}   (0 means char -ceq coincides with ordinal over this range)" -f $divergences)
