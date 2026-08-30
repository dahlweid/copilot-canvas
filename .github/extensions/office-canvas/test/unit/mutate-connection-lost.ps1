# Mutation check for the connection-lost panel test.
#
# Not part of the suite: a test that cannot fail is invisible to a green run, so
# each defect below is reintroduced, the tests are run, and the run must go red.
# Anything reported as SURVIVED is a test that proves nothing.
#
# ## Why this file exists at all
#
# The suite here is **only** `test/unit/connection-lost.test.mjs`, deliberately.
# Most mutants below also break `connection-status.test.mjs`, and running both
# would credit the new file with kills the old one made. Narrowing the suite to
# one file makes every KILLED line attributable to the test this PR adds.
#
# ## The mutant this whole file exists for
#
# `the monitor is called, and does nothing` keeps `monitorConnection(source,`
# textually intact while making it inert. That matters because the only thing
# guarding `app.js`'s wiring before this PR was a **regex** in
# `ui-contract.test.mjs` -- `/monitorConnection\s*\(\s*source\s*,/` -- which this
# mutant still matches. So it is a defect the old guard passes and the new test
# fails, which is the definition of the gap being closed rather than restated.
# `ui-contract.test.mjs` says as much itself: it "cannot fail on a monitor that
# is imported, called, and then never reached at runtime".
#
# ## This file must keep its UTF-8 BOM
#
# It is run with `powershell -File`, which is Windows PowerShell 5.1, and 5.1
# decodes a BOM-less script as ANSI. No anchor below is non-ASCII today, but the
# strings this area of the code deals in are (`LOST` carries an em dash), so the
# guard is kept rather than left to be rediscovered by whoever adds the first
# such anchor and gets a silent no-op. `mutate-webview.ps1` documents the family.
#
# Run from the extension root:
#   powershell -File test/unit/mutate-connection-lost.ps1

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

# An em dash is one character decoded as UTF-8 and three decoded as ANSI. This
# has to come before any mutant runs: a mangled anchor still produces a green
# `KILLED` line for every mutant whose anchor happens to be ASCII, and a partial
# run is the thing hardest to notice.
if ('X—X'.Length -ne 3) {
    Write-Host "This script was decoded as ANSI, not UTF-8 -- its BOM is missing." -ForegroundColor Red
    Write-Host "Every non-ASCII anchor would silently fail to match. Restore the BOM." -ForegroundColor Red
    exit 1
}

$suite = @(
    'test/unit/connection-lost.test.mjs'
)

# Anchors are single lines on purpose. A multi-line anchor carries the line
# ending of *this* file, so it matches a CRLF checkout and silently misses an LF
# one -- the MISSING category would catch it, but only on the machine that has
# the other ending. Every mutant here is one line in, one line out.
$mutants = @(
    # --- the wiring, which is what this file is for --------------------------

    # The gap this PR closes. Textually still a call, so the regex guard in
    # `ui-contract.test.mjs` is satisfied; behaviourally the panel is deaf.
    @{ name = 'the monitor is called, and does nothing'
       file = 'src/ui/app.js'
       from = '    monitorConnection(source, { setStatus });'
       to   = '    monitorConnection(source, { setStatus: () => {} });' }

    # The other half of the wiring: a monitor watching a source the panel is not
    # actually using would report on nothing the reader can see.
    @{ name = 'the monitor watches a source the panel never opened'
       file = 'src/ui/app.js'
       from = '    monitorConnection(source, { setStatus });'
       to   = '    monitorConnection({ readyState: 0, close() {} }, { setStatus });' }

    # --- what the panel renders ----------------------------------------------

    @{ name = 'the terminal message is written into a still-hidden element'
       file = 'src/ui/app.js'
       from = '    el.status.hidden = false;'
       to   = '    el.status.hidden = true;' }

    @{ name = 'the terminal message is not marked as an error'
       file = 'src/ui/app.js'
       from = '    el.status.classList.toggle("error", error);'
       to   = '    el.status.classList.toggle("error", false);' }

    # --- the give-up itself ---------------------------------------------------

    @{ name = 'the panel never gives up, and claims a recovery forever (#66)'
       file = 'src/ui/connection-status.mjs'
       from = '            if (current()) giveUp();'
       to   = '            if (false) giveUp();' }

    @{ name = 'the panel says it is still reconnecting instead of that it stopped'
       file = 'src/ui/connection-status.mjs'
       from = '        say(LOST, { error: true });'
       to   = '        say(RECONNECTING, { error: true });' }

    # "cannot resume" is only true because the source is shut first.
    @{ name = 'the panel claims it cannot resume while holding an open source'
       file = 'src/ui/connection-status.mjs'
       from = '            source.close();'
       to   = '            void source;' }

    @{ name = 'the grace period is skipped and the drop is announced at once'
       file = 'src/ui/connection-status.mjs'
       from = '            if (current()) say(RECONNECTING, { busy: true });'
       to   = '            say(RECONNECTING, { busy: true });' }

    # --- recovery, the negative case -----------------------------------------

    @{ name = 'a reconnect leaves the stale message on screen'
       file = 'src/ui/connection-status.mjs'
       from = '        clear();'
       to   = '        void 0;' }

    # --- the harness itself ---------------------------------------------------
    #
    # A stub that cannot observe the thing under test would make the assertions
    # above decorative. These are not product defects; they ask whether this
    # harness is capable of seeing a close and a status at all.

    @{ name = 'HARNESS: the stub source does not record being closed'
       file = 'test/unit/ui-harness.mjs'
       from = '        this.closed += 1;'
       to   = '        void 0;' }

    @{ name = 'HARNESS: the stub element reports no text'
       file = 'test/unit/ui-dom.mjs'
       from = '        return this.children.map((child) => child.textContent).join("");'
       to   = '        return "";' }
)

$survived = @()
$missing = @()

# A mutant is only informative against a suite that is otherwise green: a red
# baseline kills every mutant and the run reads as a clean sweep.
Push-Location $root
try {
    & node --test @suite *> $null
    if ($LASTEXITCODE -ne 0) {
        Write-Host 'Baseline is red. Fix the suite before mutating; a red baseline kills every mutant.' -ForegroundColor Red
        exit 1
    }

    foreach ($m in $mutants) {
        $file = Join-Path $root $m.file
        $original = [IO.File]::ReadAllText($file)
        if (-not $original.Contains($m.from)) {
            # A stale anchor is its own failure, not a survivor. Reported as one
            # it reads as "the tests do not defend this", when in fact nothing
            # was mutated at all.
            Write-Host ("MISSING  {0} -- anchor not found in {1}" -f $m.name, $m.file) -ForegroundColor Yellow
            $missing += $m.name
            continue
        }

        [IO.File]::WriteAllText($file, $original.Replace($m.from, $m.to))
        try {
            & node --test @suite *> $null
            $red = $LASTEXITCODE -ne 0
        } finally {
            [IO.File]::WriteAllText($file, $original)
        }

        if ($red) {
            Write-Host ("KILLED   {0}" -f $m.name) -ForegroundColor Green
        } else {
            Write-Host ("SURVIVED {0}" -f $m.name) -ForegroundColor Red
            $survived += $m.name
        }
    }
} finally {
    Pop-Location
}

Write-Host ''
Write-Host ("{0} of {1} mutants killed" -f ($mutants.Count - $survived.Count - $missing.Count), $mutants.Count)
if ($missing.Count -gt 0) {
    Write-Host 'Never applied (stale anchor -- these measured nothing):' -ForegroundColor Yellow
    $missing | ForEach-Object { Write-Host ("  - {0}" -f $_) -ForegroundColor Yellow }
}
if ($survived.Count -gt 0) {
    Write-Host 'Unkilled:' -ForegroundColor Red
    $survived | ForEach-Object { Write-Host ("  - {0}" -f $_) -ForegroundColor Red }
}
if ($survived.Count -gt 0 -or $missing.Count -gt 0) { exit 1 }
