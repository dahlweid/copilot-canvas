# Probe: can we force print colours (white page) on a Word instance that runs
# under a user profile with Office dark mode enabled?
$ErrorActionPreference = 'Continue'

$word = New-Object -ComObject Word.Application
Write-Output "created via COM"

$candidates = @(
    'DarkModeDocumentColor',
    'DarkMode',
    'DisableDarkMode',
    'UseDarkModeDocumentColor'
)
foreach ($name in $candidates) {
    try {
        $v = $word.Options.PSObject.Properties[$name]
        $val = $word.Options.$name
        Write-Output ("option {0,-28} = {1}" -f $name, $val)
    } catch {
        Write-Output ("option {0,-28} : NOT PRESENT" -f $name)
    }
}

# Whatever the theme, a document Window exposes per-window colour behaviour.
try {
    $doc = $word.Documents.Add()
    $win = $word.ActiveWindow
    foreach ($name in @('DisplayRulers','View')) { }
    Write-Output ("ActiveWindow.View.Type      = " + $win.View.Type)
    try { Write-Output ("View.DisplayBackgrounds     = " + $win.View.DisplayBackgrounds) } catch { Write-Output "View.DisplayBackgrounds     : NOT PRESENT" }
    try { Write-Output ("View.DisplayPageBoundaries  = " + $win.View.DisplayPageBoundaries) } catch { Write-Output "View.DisplayPageBoundaries  : NOT PRESENT" }
    $doc.Close(0)
} catch { Write-Output ("doc probe failed: " + $_.Exception.Message) }

# The theme itself is a per-user registry value Word reads at startup.
$uiTheme = Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Office\16.0\Common' -Name 'UI Theme' -ErrorAction SilentlyContinue
Write-Output ("HKCU UI Theme               = " + $(if ($uiTheme) { $uiTheme.'UI Theme' } else { '<unset>' }))

$pid0 = (Get-Process WINWORD -ErrorAction SilentlyContinue | Select-Object -First 1).Id
$word.Quit(0)
[Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
Start-Sleep -Seconds 2
Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { Write-Output "still running: $($_.Id)" }
