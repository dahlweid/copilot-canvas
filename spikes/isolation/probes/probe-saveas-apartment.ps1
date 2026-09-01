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
# `word-host.ps1`'s `create_document` path now calls
# `$doc.SaveAs2($path, $WD_FORMAT_XML_DOCUMENT)` (in `Cmd-Create`); the edit
# path calls `Save()` on a document it opened. So the shipping code does exercise
# `SaveAs2`, and the hang seen under `Start-Job` had to be explained before that
# authoring code could be trusted.
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

. (Join-Path $PSScriptRoot '_common.ps1')

# Set when a worker powershell.exe misses its deadline and is force-killed, which
# orphans the Word that worker was driving. The census report at the end is the
# only place that leak is now named, so it must know.
$script:workerKilled = $false

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
        # Quit(), never Quit(<arg>). Under Windows PowerShell 5.1 the argument
        # form does not bind: it throws and the Word survives, and process exit
        # does not reap it (probe-quit0-leak.ps1). The no-argument
        # form takes the same default. Reporting via Step rather than swallowing,
        # because the empty catch is what hid the leak.
        try { $w.Quit() } catch { Step "Quit() FAILED (Word may leak) -- $($_.Exception.Message.Split([char]10)[0])" }
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
        # Sound: $p.Id came back from Start-Process. But killing a COM client
        # orphans the Word it was driving -- recorded so the census report at the
        # end can name a cause it actually knows (#136).
        try { Stop-Process -Id $p.Id -Force } catch { }
        $script:workerKilled = $true
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
# This used to force-kill $leaked. It is a census DIFFERENCE, measured unsound
# here (#136): 2 new pids for 1 instance in probe-init-attribution.ps1, and 2
# strangers' WINWORDs in a 40 s window with nothing launched. Removing the sweep
# exposes the leak it masked -- a force-killed worker orphans its Word -- so that
# is named rather than left silent.
if ($leaked.Count -gt 0) {
    Write-Host ""
    Write-CensusSurvivors $leaked -WorkerKilled:$script:workerKilled
} else {
    Write-Host ""
    Write-Host "no WINWORD left behind"
}

Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
