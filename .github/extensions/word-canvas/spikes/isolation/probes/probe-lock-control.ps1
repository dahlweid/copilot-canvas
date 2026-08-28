# Control for T3: is the hang caused by the LOCK, or merely by running a
# second Word instance at all?
#
# Arm A: instance 1 holds original.docx; instance 2 opens a DIFFERENT file.
# Arm B: instance 1 holds original.docx; instance 2 opens THE SAME file.
#
# If A completes and B hangs, the hang is attributable to the lock.
# Each arm runs in its own job with a hard timeout so nothing wedges.

$ErrorActionPreference = 'Stop'
$src = $args[0]

function Run-Arm {
    param([string]$Arm, [string]$Src)

    $job = Start-Job -ArgumentList $Arm, $Src -ScriptBlock {
        param($Arm, $Src)
        $root = Join-Path $env:TEMP ("ctrl-" + [guid]::NewGuid().ToString('N').Substring(0,8))
        New-Item -ItemType Directory -Path $root | Out-Null
        $a = Join-Path $root 'original.docx'
        $b = Join-Path $root 'other.docx'
        Copy-Item $Src $a
        Copy-Item $Src $b
        $target = if ($Arm -eq 'A') { $b } else { $a }

        $w1 = New-Object -ComObject Word.Application
        $w1.Visible = $false; $w1.DisplayAlerts = 0
        $d1 = $w1.Documents.Open($a, $false, $false)

        $w2 = New-Object -ComObject Word.Application
        $w2.Visible = $false; $w2.DisplayAlerts = 0
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $result = try {
            $d2 = $w2.Documents.Open($target, $false, $false)
            $ro = $d2.ReadOnly
            $d2.Close(0)
            "opened ReadOnly=$ro in $($sw.ElapsedMilliseconds) ms"
        } catch {
            "threw after $($sw.ElapsedMilliseconds) ms -> $($_.Exception.Message.Split([char]10)[0])"
        }

        try { $d1.Close(0) } catch {}
        try { $w2.Quit() } catch {}
        try { $w1.Quit() } catch {}
        Remove-Item $root -Recurse -Force -ErrorAction SilentlyContinue
        $result
    }

    $done = Wait-Job $job -Timeout 45
    if ($done) {
        "Arm ${Arm}: " + (Receive-Job $job)
    } else {
        "Arm ${Arm}: HUNG (no result within 45 s)"
        Stop-Job $job
    }
    Remove-Job $job -Force -ErrorAction SilentlyContinue
}

Run-Arm -Arm 'A' -Src $src
Run-Arm -Arm 'B' -Src $src

# sweep any Word left wedged by a hung arm
Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Get-ChildItem $env:TEMP -Directory -Filter 'ctrl-*' -ErrorAction SilentlyContinue |
    Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
"swept"
