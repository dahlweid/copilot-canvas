# probe-saveas-apartment.ps1
#
# Why this probe exists
# ---------------------
# probe-autocorrect.ps1 ran its arms inside `Start-Job` and every arm that
# reached `Document.SaveAs2` hung -- on both the `Range.Text` path and the
# `Selection.TypeText` path, at exactly the same call, with `DisplayAlerts = 0`
# and the target directory already created. `Documents.Add`, `Range.Text` and
# `Selection.TypeText` all returned in milliseconds in the same job.
#
# `word-host.ps1` never calls `SaveAs2` -- it only ever calls `Save()` on a
# document it opened -- so nothing in the shipping code proves `SaveAs2` works.
# `create_document` has to call it, so the hang has to be explained before any
# authoring code is written on top of it.
#
# The hypothesis: apartment state. `powershell.exe` 5.1 hosts the console
# runspace in an STA, but `Start-Job` children run MTA. A COM call that pumps a
# modal message loop deadlocks when it is marshalled out of an MTA, and `SaveAs2`
# is a far heavier call than `Save()` -- it negotiates a converter, which is
# exactly the kind of call that pumps.
#
# If that is the cause, the hang is an artefact of how the earlier probe was
# run and not a constraint on the product, because the real host is STA. That
# distinction decides whether `create_document` can save at all, so it is worth
# a probe of its own rather than a guess.
#
# Method
# ------
# One worker script, run as a separate `powershell.exe` under `-STA` and again
# under `-MTA`, with everything else held identical. The worker is launched with
# discrete argv elements via `Start-Process -ArgumentList` -- never a command
# line built from a path -- so the document path cannot be re-parsed by
# `cmd.exe` or by PowerShell's own quoting. Each run is watched to a deadline
# and the worker's step trace is printed whether it finished or wedged, so a
# hang is localized rather than reported as silence.
#
# Cleanup kills only PIDs that were absent before the probe started. Several
# WINWORD.EXE belonging to other sessions are routinely alive on this machine
# and killing one destroys someone's unsaved work.

$ErrorActionPreference = 'Stop'

function Get-WordPids {
    @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }) | Sort-Object
}

$pidsBefore = Get-WordPids
Write-Host "WINWORD before: $($pidsBefore -join ', ')"

$root = Join-Path $env:TEMP ("apt-" + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $root | Out-Null

# The worker. Written to disk once and run twice; the only difference between
# the runs is the apartment switch on powershell.exe.
$worker = Join-Path $root 'worker.ps1'
@'
param(
    [Parameter(Mandatory = $true)][string] $Out,
    [Parameter(Mandatory = $true)][string] $Trace,
    [Parameter(Mandatory = $true)][string] $DocPath
)

# Note the parameter is $DocPath, not $Doc. PowerShell variable names are
# case-insensitive and a param() type constraint follows the variable for the
# rest of the scope, so a later `$doc = $w.Documents.Add()` would be *coerced to
# a string* by the [string] on $Doc rather than rejected. The COM object is then
# gone, `$doc.Content` reads as $null, and the failure surfaces two lines later
# as "the property Text was not found on this object" -- pointing at a line that
# is correct. The first version of this probe did exactly that and cost a run.

$ErrorActionPreference = 'Stop'
$sw = [Diagnostics.Stopwatch]::StartNew()
function Step($m) {
    Add-Content -LiteralPath $Trace -Value ("{0,7}ms  {1}" -f $sw.ElapsedMilliseconds, $m)
}

$result = [ordered]@{}
$result.apartment = [Threading.Thread]::CurrentThread.GetApartmentState().ToString()
Step "worker started, apartment $($result.apartment)"

$w = $null
try {
    $w = New-Object -ComObject Word.Application
    $w.Visible = $false
    $w.DisplayAlerts = 0
    Step 'Word ready'

    $document = $w.Documents.Add()
    Step "Documents.Add returned, type $($document.GetType().FullName)"

    $r = $document.Paragraphs.Item(1).Range
    $visible = ([string]$r.Text).TrimEnd("`r", [char]7, "`n")
    $drop = ($r.End - $r.Start) - $visible.Length
    if ($drop -gt 0) { $r.MoveEnd(1, -$drop) | Out-Null }
    $r.Text = 'probe body'
    Step 'wrote body'

    # wdFormatDocumentDefault = 16 (.docx)
    Step 'calling SaveAs2(path, 16)'
    $document.SaveAs2($DocPath, 16)
    Step 'SaveAs2 returned'
    $result.saveAs2 = 'ok'

    $document.Close(0)
    Step 'Close returned'
    $result.exists = (Test-Path -LiteralPath $DocPath)
} catch {
    Step "ERROR $($_.Exception.GetType().FullName): $($_.Exception.Message)"
    $result.saveAs2 = "error: $($_.Exception.Message)"
} finally {
    if ($null -ne $w) {
        try { $w.Quit(0) } catch { }
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($w) | Out-Null } catch { }
    }
    Step 'quit'
}

$result | ConvertTo-Json -Compress | Set-Content -LiteralPath $Out -Encoding UTF8
Step 'result written'
'@ | Set-Content -LiteralPath $worker -Encoding UTF8

function Invoke-Worker([string] $apartmentSwitch) {
    $tag = $apartmentSwitch.TrimStart('-')
    $out = Join-Path $root "$tag.json"
    $trace = Join-Path $root "$tag.trace"
    $doc = Join-Path $root "$tag.docx"
    New-Item -ItemType File -Path $trace | Out-Null

    Write-Host ""
    Write-Host "=== powershell.exe $apartmentSwitch ==="

    # Discrete argv elements. The document path is a -Doc parameter value, never
    # interpolated into a command string, so an apostrophe or an '&' in it
    # cannot reach a parser.
    $p = Start-Process -FilePath 'powershell.exe' -PassThru -WindowStyle Hidden -ArgumentList @(
        $apartmentSwitch
        '-NoProfile'
        '-NonInteractive'
        '-ExecutionPolicy', 'Bypass'
        '-File', $worker
        '-Out', $out
        '-Trace', $trace
        '-DocPath', $doc
    )

    # Poll to a deadline. Word's own teardown is load-dependent, so a flat sleep
    # would decide the outcome on machine quietness rather than on behaviour.
    $deadline = (Get-Date).AddSeconds(90)
    while (-not $p.HasExited -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 250 }

    if (-not $p.HasExited) {
        Write-Host "  HUNG (no exit within 90s)"
        try { Stop-Process -Id $p.Id -Force } catch { }
    } elseif (Test-Path -LiteralPath $out) {
        Write-Host "  $(Get-Content -LiteralPath $out -Raw)"
    } else {
        Write-Host "  exited $($p.ExitCode) with no result file"
    }

    Write-Host "  --- step trace ---"
    Get-Content -LiteralPath $trace | ForEach-Object { Write-Host "  $_" }
}

Invoke-Worker '-STA'
Invoke-Worker '-MTA'

$leaked = @(Get-WordPids | Where-Object { $pidsBefore -notcontains $_ })
if ($leaked.Count -gt 0) {
    Write-Host ""
    Write-Host "cleaning up WINWORD started by this probe: $($leaked -join ', ')"
    foreach ($p in $leaked) { try { Stop-Process -Id $p -Force -ErrorAction SilentlyContinue } catch { } }
} else {
    Write-Host ""
    Write-Host "no WINWORD left behind"
}

Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
