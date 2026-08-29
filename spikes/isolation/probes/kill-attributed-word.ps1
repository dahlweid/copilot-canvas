# Kills only the WINWORD.EXE processes named on the command line, and only after
# confirming each one's start time falls inside the window given.
#
# Three other sessions drive Word on this machine. Never kill a Word this script
# cannot attribute to the caller: a COM-attached PowerPoint or Word instance *is*
# the user's, and quitting it destroys their unsaved work.

param(
    [Parameter(Mandatory = $true)][int[]]$Pid_,
    [Parameter(Mandatory = $true)][datetime]$StartedAfter
)

foreach ($id in $Pid_) {
    $p = Get-Process -Id $id -ErrorAction SilentlyContinue
    if ($null -eq $p) { Write-Host "$id : already gone"; continue }
    if ($p.ProcessName -ne 'WINWORD') { Write-Host "$id : not WINWORD ($($p.ProcessName)), refusing"; continue }
    if ($p.StartTime -lt $StartedAfter) {
        Write-Host "$id : started $($p.StartTime), before the window -- refusing"
        continue
    }
    Stop-Process -Id $id -Force
    Write-Host "$id : killed (started $($p.StartTime))"
}
