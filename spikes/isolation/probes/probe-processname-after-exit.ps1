# What a `System.Diagnostics.Process` reports about a process that has EXITED,
# and what `Kill()` does to it -- measured under Windows PowerShell 5.1, which
# is what `word-host.ps1` runs under.
#
# Three files kill a Word through a helper that reads `ProcessName` to check the
# pid still holds a WINWORD. Two of them GUARD that read with a try/catch and
# say the read can throw; the shipped host does not guard it. Nothing in the
# tree had ever measured which is right, so the divergence was being argued
# from .NET semantics. This measures it instead.
#
# The headline result is that BOTH sides were reasoning about the wrong layer.
# The .NET getter does throw (arm A2). PowerShell's property adapter converts
# that throw into `$null` before any caller sees it -- and `$Error` does not
# grow, under `$ErrorActionPreference = 'Stop'`, which is what makes the
# swallow invisible to the one setting an author would expect to catch it.
# So the guard cannot fire and the unguarded read cannot throw. The real
# consequence is elsewhere: arm C, where a process that exits mid-sequence
# passes every identity check and then fails the terminate.
#
# Office-free BY CONSTRUCTION, and that is the point: the question is about
# `System.Diagnostics.Process`, not about Word. The subject is a short-lived
# `cmd.exe` this probe starts itself, so it needs no licence, no COM, and
# cannot touch anybody's document.
#
# Nothing alive is terminated. Every `Kill()` below is aimed at a subject this
# probe started AND has already waited for the exit of -- that is the whole
# point of arms C and D, which are about what a terminate does to a corpse.
#
#   powershell.exe -NoProfile -ExecutionPolicy Bypass -File spikes/isolation/probes/probe-processname-after-exit.ps1
#
# Arms:
#
#   A1  name read after exit, via the helper's acquisition route
#   A2  the same read forced through reflection, bypassing the adapter
#   A3  the decision the shipped helper's name check then takes
#   B   the full guard chain run against a corpse: does it reach the kill?
#   C   Kill() on an exited subject, handle pinned -- and which catch fires
#   D   Kill() on an exited subject, NOT pinned -- the file's own discriminator
#   E   a process that exists but is protected, as a separate failure shape
#   F   .Handle itself, which turns out to be swallowed the same way
#
# Exception TYPES are reported, never just messages. Messages here are
# localized -- this machine's Windows speaks German -- so any code reacting to
# one of these must discriminate on the type, and a probe printing only the
# message would invite exactly that mistake.

$ErrorActionPreference = 'Stop'

function Start-Subject {
    # `cmd /c exit` is the cheapest process on the machine that is neither ours
    # nor Office. -PassThru yields the pid the KERNEL minted for it.
    return Start-Process -FilePath $env:ComSpec -ArgumentList '/c', 'exit' -PassThru -WindowStyle Hidden
}

function Get-Corpse {
    # Reproduces how the helper acquires its object: `Get-Process -Id` while the
    # process is still alive, `.Handle` pinned, and only then the exit. The
    # acquisition route matters -- see arm A1.
    $s = Start-Subject
    $g = Get-Process -Id $s.Id -ErrorAction SilentlyContinue
    if ($Pin) { $null = $g.Handle }
    $recorded = $g.StartTime
    $s.WaitForExit()
    Start-Sleep -Milliseconds 300
    return [pscustomobject]@{ Proc = $g; Recorded = $recorded }
}

function Format-Root($errorRecord) {
    $root = $errorRecord.Exception
    while ($null -ne $root.InnerException) { $root = $root.InnerException }
    return $root
}

Write-Output "PowerShell $($PSVersionTable.PSVersion)"
Write-Output ''

# --- A: can the name read throw into a PowerShell caller? ---------------------
$Pin = $true
$a = Get-Corpse
$p = $a.Proc

# Read it as the helper does, unguarded. Note that it is NOT $null yet: a
# `Get-Process -Id` object carries the name materialized by the enumeration
# that produced it, and that copy survives the process. `Refresh()` discards it
# and forces the read to go to the (now absent) process -- which is the state
# the guard was written for, and nothing in the shipped host calls `Refresh()`.
$cached = $p.ProcessName
$p.Refresh()
$before = $Error.Count
$after = $p.ProcessName
$grew = $Error.Count - $before
Write-Output ("A1 name, cached by Get-Process       -> '{0}'" -f $cached)
Write-Output ("A1 name, after Refresh(), unguarded  -> {0}   threw=False   `$Error grew by {1}" -f $(if ($null -eq $after) { '$null' } else { "'$after'" }), $grew)

try {
    $null = [System.Diagnostics.Process].GetProperty('ProcessName').GetValue($p, $null)
    Write-Output 'A2 same read, via reflection         -> returned (did not throw)'
} catch {
    $root = Format-Root $_
    Write-Output ("A2 same read, via reflection         -> THREW {0}" -f $root.GetType().FullName)
}

# So the guarded and unguarded forms are indistinguishable from PowerShell, and
# the shipped helper's name check takes this branch on an unreadable name:
Write-Output ("A3 (`$name -ne 'WINWORD') where name is `$null -> {0}   [helper returns 'gone': the SAFE branch]" -f ($after -ne 'WINWORD'))
Write-Output ''

# --- B: does a corpse pass the whole identity chain? --------------------------
$b = Get-Corpse
$q = $b.Proc
$stillThere = Get-Process -Id $q.Id -ErrorAction SilentlyContinue
Write-Output ("B0 Get-Process -Id, re-taken         -> {0}" -f $(if ($null -eq $stillThere) { "`$null   [a FRESH acquisition returns 'gone' at the helper's Get-Process]" } else { 'an object' }))
# But the helper does not re-acquire. Within one call the object is taken once
# and the process may exit at any point after. On THAT object:
$name = $q.ProcessName
$actual = try { $q.StartTime } catch { $null }
Write-Output ("B1 name on the held object           -> '{0}'   passes the WINWORD check: {1}" -f $name, ($name -eq 'cmd'))
Write-Output ("B2 StartTime through the pinned handle -> {0}" -f $actual)
Write-Output ("B3 matches the start time recorded before the exit: {0}   [every guard passes; control reaches the kill]" -f ($actual -eq $b.Recorded))
Write-Output ''

# --- C: the terminate that the passing chain then reaches ---------------------
# Three DIFFERENT propositions are separated here, because conflating them is
# how this arm was got wrong once already. "the outer exception is an
# InvalidOperationException" and "a typed catch for InvalidOperationException
# fires" are not the same claim, and here they have different answers.
try {
    $q.Kill()
    Write-Output 'C1 Kill(), pinned, subject exited    -> returned (did not throw)'
} catch {
    $root = Format-Root $_
    Write-Output ("C1 Kill(), pinned, subject exited    -> THREW" )
    Write-Output ("C2 outer exception type              -> {0}" -f $_.Exception.GetType().FullName)
    Write-Output ("C3 root (innermost) type             -> {0}" -f $root.GetType().FullName)
    Write-Output ("C4 outer -is [InvalidOperationException] -> {0}" -f ($_.Exception -is [System.InvalidOperationException]))
    $sameText = ($_.Exception.Message -eq $root.Message)
    Write-Output ("C6 outer message identical to root's -> {0}   (outer {1} chars, root {2}; outer is ONE line and CONTAINS the root: {3})" -f $sameText, $_.Exception.Message.Length, $root.Message.Length, (($_.Exception.Message -split "`n").Count -eq 1 -and $_.Exception.Message.Contains($root.Message)))
    Write-Output ("C6 so quoting the outer loses no cause -- it prefixes PowerShell's own wrapper prose. Reporting the root is a tidier message, not a bug fix.")
}
# C5 is the one that decides how the shipped host should be written: whether a
# type-filtered catch clause actually fires. It is measured by running one,
# not inferred from C2 or C4.
$q2 = (Get-Corpse).Proc
$clause = 'none: Kill() returned'
try { $q2.Kill() }
catch [System.InvalidOperationException] { $clause = 'catch [System.InvalidOperationException]' }
catch { $clause = 'the bare catch: ' + $_.Exception.GetType().FullName }
Write-Output ("C5 which catch clause fired          -> {0}" -f $clause)
Write-Output ''

# --- D: the same terminate without the pin ------------------------------------
# `word-host.ps1` documents these two as different failures and uses the
# difference to argue the pin is load-bearing. Reproduced here independently,
# because that comment carried no probe of its own.
$Pin = $false
$d = Get-Corpse
try {
    $d.Proc.Kill()
    Write-Output 'D  Kill(), NOT pinned, exited       -> returned'
} catch {
    $root = Format-Root $_
    Write-Output ("D  Kill(), NOT pinned, exited       -> THREW; root {0}" -f $root.GetType().FullName)
}
Write-Output ''

# --- E: a live process whose reads are refused rather than absent -------------
# Separated from the arms above because "the process is gone" and "the process
# is there and will not answer" are different states, and only the first is
# what the helper's name check is about. pid 4 is the kernel's System process; it
# is only read here, never touched.
$sys = Get-Process -Id 4 -ErrorAction SilentlyContinue
if ($null -eq $sys) {
    Write-Output 'E  pid 4 not present; arm skipped'
} else {
    $before = $Error.Count
    $sysStart = $sys.StartTime
    Write-Output ("E  protected process: name '{0}', StartTime {1}, `$Error grew by {2}" -f $sys.ProcessName, $(if ($null -eq $sysStart) { '$null' } else { $sysStart }), ($Error.Count - $before))
    Write-Output '   [the same silent conversion: a guarded and an unguarded StartTime read agree here too]'
    # E2 is the OTHER cause of a $null pin, and it is measured here rather than
    # asserted: arm F2 below shows what an EXITED process answers, which says
    # nothing about one that is present and refuses the open. pid 4 is opened
    # for a handle and nothing else; no attempt is made to touch it.
    $before = $Error.Count
    $threwE = $false
    $sysHandle = $null
    try { $sysHandle = $sys.Handle } catch { $threwE = $true }
    Write-Output ("E2 .Handle on that same process     -> {0}  threw={1}  `$Error grew by {2}" -f `
        $(if ($null -eq $sysHandle) { '$null' } else { 'a handle' }), $threwE, ($Error.Count - $before))
}
Write-Output ''

# --- F: the pin itself ---------------------------------------------------------
# Arms A-E were all about reads taken AFTER the pin. This arm asks whether the
# pin happened at all. `word-host.ps1` takes it with `$null = $p.Handle` inside a
# try/catch and treats reaching the next line as proof that the pid is now held.
# Every later guard, and the safety of the terminate itself, rests on that.
#
# F1 is the control: an ordinary live process, so a $null in F2 cannot be read as
# ".Handle just always answers nothing here".
$before = $Error.Count
$live = Get-Process -Id $PID
$liveHandle = $live.Handle
Write-Output ("F1 .Handle on a live process        -> {0}  `$Error grew by {1}" -f `
    $(if ($null -eq $liveHandle) { '$null' } else { 'a handle' }), ($Error.Count - $before))

# F2 is the case the helper is written for: the process exited before the pin was
# taken. This is not a rare interleaving -- it is the ordinary end of a Word that
# quit on its own while the sweep was walking the pid list.
$Pin = $false
$f = Get-Corpse
$before = $Error.Count
$threw = $false
$handle = $null
try { $handle = $f.Proc.Handle } catch { $threw = $true }
Write-Output ("F2 .Handle after exit               -> {0}  threw={1}  `$Error grew by {2}" -f `
    $(if ($null -eq $handle) { '$null' } else { 'a handle' }), $threw, ($Error.Count - $before))

# F3 asks the only question that matters for the shipped code: with the read
# written exactly as the helper writes it, does control enter the catch?
$g = Get-Corpse
$reached = 'the catch'
try { $null = $g.Proc.Handle; $reached = 'the line after it' } catch { }
Write-Output ("F3 `$null = `$p.Handle, then?         -> control reached {0}" -f $reached)

# F4 walks the rest of the guard chain on that unpinned object, because a pin
# that did not happen is only a defect if the guards below fail to notice.
$name = $g.Proc.ProcessName
$actual = try { $g.Proc.StartTime } catch { $null }
Write-Output ("F4 unpinned, the guards then see    -> name '{0}', StartTime {1}, matches recorded: {2}" -f `
    $(if ($null -eq $name) { '$null' } else { $name }), `
    $(if ($null -eq $actual) { '$null' } else { $actual }), `
    ($actual -eq $g.Recorded))
Write-Output '   [so the identity checks pass on an object that was never pinned, and the terminate is reached]'
