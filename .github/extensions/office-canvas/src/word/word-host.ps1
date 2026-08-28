# word-host.ps1
#
# Long-lived PowerShell process that owns a hidden Microsoft Word instance and
# acts as the rendering/query engine for the office-canvas extension.
#
# Protocol: newline-delimited JSON over stdio.
#   in  -> {"id":1,"cmd":"open","args":{...}}
#   out -> {"id":1,"ok":true,"result":{...}}
#       -> {"id":1,"ok":false,"error":{"code":"...","message":"..."}}
#
# Hard rules (see plan):
#   * Numeric wd*/mso* constants only. Localized Word UIs do not have
#     "Heading 1" and string style names throw.
#   * Macros are force-disabled before any document is opened.
#   * Only an unblocked temp copy is ever opened, never the user's original.
#   * A Word instance we did not create is never quit and never hidden.
#   * Global Application.Options are never modified; they persist for the user.

param(
    # Directory used to record which WINWORD process this host owns, so an
    # instance orphaned by a crash can be reaped by the next host that starts.
    [string]$PidDir
)

$ErrorActionPreference = 'Stop'

[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)

# --- Word / Office constants -------------------------------------------------
$WD_ALERTS_NONE = 0
$MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3
$WD_DO_NOT_SAVE_CHANGES = 0
$WD_EXPORT_FORMAT_PDF = 17
$WD_EXPORT_ALL_DOCUMENT = 0
$WD_EXPORT_FROM_TO = 3
$WD_EXPORT_OPTIMIZE_FOR_PRINT = 0
$WD_EXPORT_DOCUMENT_CONTENT = 0
$WD_EXPORT_CREATE_HEADING_BOOKMARKS = 1
$WD_INFO_ACTIVE_END_PAGE_NUMBER = 3
$WD_INFO_NUMBER_OF_PAGES = 4
$WD_STATISTIC_WORDS = 0
$WD_STATISTIC_PAGES = 2
$WD_FIND_STOP = 0
$WD_GOTO_PAGE = 1
$WD_GOTO_ABSOLUTE = 1
$WD_OUTLINE_BODY_TEXT = 10

# wdStyleHeading1 = -2 ... wdStyleHeading9 = -10
function Get-HeadingStyleId([int]$level) { return -1 - $level }

# A password that will never match, so an encrypted document fails fast with an
# error instead of blocking automation on a modal password dialog.
$FAIL_FAST_PASSWORD = '#word-canvas-no-password#'

$script:App = $null
$script:OwnedPid = $null
$script:Docs = @{}
# Remembers how each document was opened so a dead COM connection can be
# recovered without the caller noticing.
$script:DocArgs = @{}

# --- plumbing ----------------------------------------------------------------

function Send-Message($obj) {
    $json = $obj | ConvertTo-Json -Compress -Depth 12
    [Console]::Out.WriteLine($json)
    [Console]::Out.Flush()
}

function Send-Ok($id, $result) {
    Send-Message @{ id = $id; ok = $true; result = $result }
}

function Send-Fail($id, $code, $message, $data = $null) {
    $err = @{ code = $code; message = $message }
    if ($null -ne $data) { $err.data = $data }
    Send-Message @{ id = $id; ok = $false; error = $err }
}

# A typed failure.
#
# Every failure used to reach the caller as `word_error`, which meant the only
# way to tell "the file is missing" from "the file is locked" from "Word itself
# died" was to pattern-match `$_.Exception.Message` -- and those messages are
# localized, so matching them on this German machine is matching a translation.
# `throw` records the thrown object on `$_.TargetObject`, so the dispatch loop
# can recover the code without parsing anything.
#
# `data` carries facts the caller needs *on the failure path*, which is where
# they matter most: `writable` is the obvious one, since "the original is locked"
# is precisely the case where a read fails and the caller wants to know why.
function New-HostError([string]$code, [string]$message, $data = $null) {
    return @{ __hostError = $true; code = $code; message = $message; data = $data }
}

function Get-WordPids {
    return @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
}

# --- owned-process registry --------------------------------------------------
#
# Each host records the WINWORD process it created in $PidDir\<hostPid>.pid.
# A host that dies without running its teardown leaves both the file and the
# Word process behind; the next host to start finds the file, sees that the
# recording host is gone, and ends the orphan. Without this, a crash leaks a
# hidden Word that nothing will ever clean up -- and worse, a later CreateObject
# can attach to that orphan, at which point ownership detection sees no new
# process and (correctly, but unhelpfully) refuses to quit it.

function Get-PidFilePath {
    if ([string]::IsNullOrWhiteSpace($PidDir)) { return $null }
    return (Join-Path $PidDir "$PID.pid")
}

function Test-IsWordPid([int]$candidate) {
    $p = Get-Process -Id $candidate -ErrorAction SilentlyContinue
    return ($null -ne $p -and $p.ProcessName -eq 'WINWORD')
}

function Clear-OrphanedWord {
    if ([string]::IsNullOrWhiteSpace($PidDir)) { return }
    try {
        if (-not (Test-Path -LiteralPath $PidDir)) {
            New-Item -ItemType Directory -Force -Path $PidDir | Out-Null
            return
        }
        foreach ($file in Get-ChildItem -LiteralPath $PidDir -Filter '*.pid' -ErrorAction SilentlyContinue) {
            $hostPid = 0
            if (-not [int]::TryParse([IO.Path]::GetFileNameWithoutExtension($file.Name), [ref]$hostPid)) { continue }
            if ($hostPid -eq $PID) { continue }
            if ($null -ne (Get-Process -Id $hostPid -ErrorAction SilentlyContinue)) { continue }  # still running
            $wordPid = 0
            if ([int]::TryParse((Get-Content -LiteralPath $file.FullName -Raw -ErrorAction SilentlyContinue).Trim(), [ref]$wordPid)) {
                if (Test-IsWordPid $wordPid) {
                    try { (Get-Process -Id $wordPid).Kill() } catch { }
                }
            }
            Remove-Item -LiteralPath $file.FullName -Force -ErrorAction SilentlyContinue
        }
    } catch { }
}

function Register-OwnedWord([int]$wordPid) {
    $file = Get-PidFilePath
    if ($null -eq $file) { return }
    try {
        if (-not (Test-Path -LiteralPath $PidDir)) { New-Item -ItemType Directory -Force -Path $PidDir | Out-Null }
        Set-Content -LiteralPath $file -Value "$wordPid" -Encoding ascii
    } catch { }
}

function Unregister-OwnedWord {
    $file = Get-PidFilePath
    if ($null -eq $file) { return }
    try { Remove-Item -LiteralPath $file -Force -ErrorAction SilentlyContinue } catch { }
}

# PowerShell cannot call the parameterized `Item` property on Office's
# DocumentProperties collection directly -- `$doc.BuiltInDocumentProperties.Item($name)`
# throws NullReferenceException. Late-bound InvokeMember is the working path.
function Get-DocProp($doc, [string]$name) {
    try {
        $props = $doc.BuiltInDocumentProperties
        $prop = [System.__ComObject].InvokeMember('Item', 'GetProperty', $null, $props, @($name))
        return [System.__ComObject].InvokeMember('Value', 'GetProperty', $null, $prop, @())
    } catch {
        return $null
    }
}

function Get-PageOf($range) {
    try { return [int]$range.Information($WD_INFO_ACTIVE_END_PAGE_NUMBER) } catch { return 0 }
}

# Whether the file could be opened for writing right now -- 4 ms when another
# process holds it, 9 ms when free, correct in both directions.
#
# ADR 0005: this is the *only* safe way to ask. Probing by asking Word to open
# the document is precisely the call that hangs indefinitely on a held file,
# with DisplayAlerts already off, leaving two processes to be killed by hand.
function Test-FileWritable([string]$path) {
    try {
        $stream = [IO.File]::Open($path, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        $stream.Close()
        $stream.Dispose()
        return $true
    } catch {
        return $false
    }
}

# ComputeStatistics(wdStatisticPages) returns 1 on a hidden, window-less Word
# instance because the document is never laid out for the screen. Information()
# reports the real paginated count, so it is the primary source here.
function Get-PageCount($doc) {
    try {
        $n = [int]$doc.Content.Information($WD_INFO_NUMBER_OF_PAGES)
        if ($n -gt 0) { return $n }
    } catch { }
    try { return [int]$doc.ComputeStatistics($WD_STATISTIC_PAGES) } catch { return 0 }
}

function Get-WordCount($doc) {
    try { return [int]$doc.ComputeStatistics($WD_STATISTIC_WORDS) } catch { return 0 }
}

function Test-AppAlive {
    if ($null -eq $script:App) { return $false }
    try {
        # PowerShell does not reliably throw when reading a property off a
        # disconnected COM proxy -- it can hand back $null instead. Probing the
        # *value* is therefore the only trustworthy liveness check.
        $version = $script:App.Version
        return -not [string]::IsNullOrEmpty([string]$version)
    } catch {
        return $false
    }
}

function Test-DocAlive($doc) {
    if ($null -eq $doc) { return $false }
    try {
        $name = $doc.Name
        return -not [string]::IsNullOrEmpty([string]$name)
    } catch {
        return $false
    }
}

function Resolve-Doc($docId) {
    # Re-open transparently if Word died underneath us (crash, stuck dialog that
    # was killed, user ending the process). The caller should not have to care.
    if ($script:Docs.ContainsKey($docId)) {
        $doc = $script:Docs[$docId]
        if (Test-DocAlive $doc) { return $doc }
        $script:Docs.Remove($docId) | Out-Null
    }
    if ($script:DocArgs.ContainsKey($docId)) {
        $saved = $script:DocArgs[$docId]
        Open-DocInternal $docId $saved.path $saved.workDir | Out-Null
        return $script:Docs[$docId]
    }
    throw (New-HostError 'no_such_document' "No open document with id '$docId'.")
}

# --- Word lifecycle ----------------------------------------------------------

function Initialize-Word {
    if (Test-AppAlive) { return }

    if ($null -ne $script:App) {
        # The instance died underneath us. Drop every stale handle before
        # creating a replacement, or later calls fail with RPC_E_DISCONNECTED.
        $script:App = $null
        $script:OwnedPid = $null
        $script:Docs = @{}
    }

    $before = Get-WordPids
    try {
        $script:App = New-Object -ComObject Word.Application
    } catch {
        throw (New-HostError 'word_unavailable' `
            "Microsoft Word could not be started. Word must be installed to use this canvas. ($($_.Exception.Message))")
    }

    # Ownership detection: if a brand new WINWORD process appeared we created it
    # and may hide and later quit it. If we merely attached to a Word the user
    # already had running, leave its visibility alone and never quit it out from
    # under them.
    #
    # The process does not always show up in the process list by the time
    # CreateObject returns -- especially right after a previous instance was
    # killed -- so poll briefly rather than taking a single snapshot. Getting
    # this wrong leaks a Word process, so the retry is worth the latency.
    $new = @()
    for ($attempt = 0; $attempt -lt 15; $attempt++) {
        $new = @(Get-WordPids | Where-Object { $before -notcontains $_ })
        if ($new.Count -ge 1) { break }
        Start-Sleep -Milliseconds 100
    }
    if ($new.Count -ge 1) {
        $script:OwnedPid = [int]$new[0]
        Register-OwnedWord $script:OwnedPid
    } else {
        $script:OwnedPid = $null
    }

    if ($null -ne $script:OwnedPid) {
        try { $script:App.Visible = $false } catch { }
    }
    try { $script:App.DisplayAlerts = $WD_ALERTS_NONE } catch { }
    # Force-disable macros before any document is opened. We render documents
    # the user may not have authored.
    try { $script:App.AutomationSecurity = $MSO_AUTOMATION_SECURITY_FORCE_DISABLE } catch { }
}

function Close-Doc($docId) {
    if (-not $script:Docs.ContainsKey($docId)) { return }
    $doc = $script:Docs[$docId]
    $script:Docs.Remove($docId)
    try { $doc.Close($WD_DO_NOT_SAVE_CHANGES) } catch { }
    try { [Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null } catch { }
}

function Stop-Word {
    foreach ($docId in @($script:Docs.Keys)) { Close-Doc $docId }
    $script:DocArgs = @{}
    if ($null -ne $script:App) {
        if ($null -ne $script:OwnedPid) {
            try { $script:App.Quit($WD_DO_NOT_SAVE_CHANGES) } catch { }
        }
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($script:App) | Out-Null } catch { }
        $script:App = $null
    }
    # Belt and braces: if the instance we own somehow survived Quit, end it.
    if ($null -ne $script:OwnedPid) {
        try {
            Start-Sleep -Milliseconds 300
            $p = Get-Process -Id $script:OwnedPid -ErrorAction SilentlyContinue
            if ($null -ne $p -and $p.ProcessName -eq 'WINWORD') { $p.Kill() }
        } catch { }
        $script:OwnedPid = $null
    }
    Unregister-OwnedWord
}

# --- commands ----------------------------------------------------------------

function Cmd-Ping($a) {
    Initialize-Word
    return @{
        ready       = $true
        wordVersion = [string]$script:App.Version
        ownedPid    = $script:OwnedPid
        owned       = ($null -ne $script:OwnedPid)
    }
}

function Open-DocInternal([string]$docId, [string]$path, [string]$workDir, [bool]$withStats = $true) {
    Initialize-Word

    if ([string]::IsNullOrWhiteSpace($path)) { throw (New-HostError 'invalid_request' "No document path supplied.") }
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw (New-HostError 'file_not_found' "File not found: $path")
    }

    if (-not (Test-Path -LiteralPath $workDir)) {
        New-Item -ItemType Directory -Force -Path $workDir | Out-Null
    }

    # Never open the user's original file: copy it, strip the mark-of-the-web
    # (Protected View refuses automation) and open the copy read-only. This is
    # what makes "read-only" structurally true rather than merely intended.
    $ext = [IO.Path]::GetExtension($path)
    $copy = Join-Path $workDir ("source" + $ext)
    try {
        Copy-Item -LiteralPath $path -Destination $copy -Force -ErrorAction Stop
    } catch {
        # The copy is the first thing that touches the original, so an exclusive
        # lock held by another process surfaces here rather than at Documents.Open
        # -- which is the good outcome, because that is the call that hangs.
        if (-not (Test-FileWritable $path)) {
            throw (New-HostError 'file_locked' `
                "Another process is holding $([IO.Path]::GetFileName($path)) open. Close it and try again." `
                    @{ writable = $false })
        }
        throw (New-HostError 'copy_failed' "Could not make a working copy of the document. ($($_.Exception.Message))")
    }
    try { Unblock-File -LiteralPath $copy -ErrorAction SilentlyContinue } catch { }

    Close-Doc $docId

    try {
        # FileName, ConfirmConversions, ReadOnly, AddToRecentFiles,
        # PasswordDocument, PasswordTemplate, Revert
        $doc = $script:App.Documents.Open($copy, $false, $true, $false, $FAIL_FAST_PASSWORD, $FAIL_FAST_PASSWORD, $false)
    } catch {
        throw (New-HostError 'document_unreadable' `
            "Word could not open the document. It may be password-protected or corrupt. ($($_.Exception.Message))")
    }

    $script:Docs[$docId] = $doc
    $script:DocArgs[$docId] = @{ path = $path; workDir = $workDir }

    $item = Get-Item -LiteralPath $path
    # Pagination and word count are a full pass over the document. The renderer
    # needs them; a structure read does not, and pays for them in latency.
    $pageCount = 0
    $wordCount = 0
    if ($withStats) {
        $pageCount = (Get-PageCount $doc)
        $wordCount = (Get-WordCount $doc)
    }
    return @{
        docId       = $docId
        path        = $path
        name        = $item.Name
        pageCount   = $pageCount
        wordCount   = $wordCount
        sizeBytes   = [int64]$item.Length
        modifiedIso = $item.LastWriteTimeUtc.ToString('o')
        title       = [string](Get-DocProp $doc 'Title')
        author      = [string](Get-DocProp $doc 'Author')
        workingCopy = $copy
    }
}

function Cmd-Open($a) {
    return Open-DocInternal ([string]$a.docId) ([string]$a.path) ([string]$a.workDir)
}

# One `Content.WordOpenXML` call, written to a file, document closed at once.
#
# Measured (PLAN.md §18.1) on a 219-paragraph document: walking Paragraphs and
# touching Range.Text / OutlineLevel per paragraph cost 3724 ms, because every
# property touch is a cross-process COM call; one WordOpenXML call cost 289 ms
# and returned strictly more -- text, style and full markup. A per-paragraph
# property walk is a defect, not a slow path.
#
# The markup goes to a file rather than back through the protocol: it is
# routinely 100 KB and can be megabytes, and JSON-escaping that onto a single
# stdio line is pure overhead when the reader is on the same machine.
function Cmd-Structure($a) {
    $docId = [string]$a.docId
    $path = [string]$a.path
    $out = [string]$a.out
    if ([string]::IsNullOrWhiteSpace($out)) {
        throw (New-HostError 'invalid_request' "No output path supplied for the structure read.")
    }
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw (New-HostError 'file_not_found' "File not found: $path")
    }

    # Reported, not enforced: a read is served from a copy and so never needs
    # the original. It tells the caller whether an *edit* would collide -- and it
    # is attached to failures too, because a failed read is exactly when the
    # caller most needs to know the original is held by someone else.
    $writable = Test-FileWritable $path

    try {
        $meta = Open-DocInternal $docId $path ([string]$a.workDir) $false
    } catch {
        $thrown = $_.TargetObject
        if ($thrown -is [hashtable] -and $thrown.__hostError) {
            if ($null -eq $thrown.data) { $thrown.data = @{} }
            $thrown.data.writable = $writable
        }
        throw
    }
    try {
        $xml = [string](Resolve-Doc $docId).Content.WordOpenXML
    } finally {
        # The lock window ends the moment the markup is in hand. Nothing below
        # this point needs Word.
        $script:DocArgs.Remove($docId) | Out-Null
        Close-Doc $docId
    }

    $outDir = Split-Path -Parent $out
    if (-not [string]::IsNullOrWhiteSpace($outDir) -and -not (Test-Path -LiteralPath $outDir)) {
        New-Item -ItemType Directory -Force -Path $outDir | Out-Null
    }
    # No BOM: the caller parses this as XML, and a BOM ahead of the declaration
    # is a parse error in stricter readers.
    try {
        [IO.File]::WriteAllText($out, $xml, (New-Object System.Text.UTF8Encoding($false)))
    } catch {
        # A short write here is what produces truncated markup downstream, so it
        # gets its own code rather than arriving as a parse failure later.
        throw (New-HostError 'write_failed' `
            "Could not write the document markup to $out. ($($_.Exception.Message))" @{ writable = $writable })
    }

    return @{
        out         = $out
        bytes       = [int64](Get-Item -LiteralPath $out).Length
        writable    = $writable
        name        = $meta.name
        sizeBytes   = $meta.sizeBytes
        modifiedIso = $meta.modifiedIso
        title       = $meta.title
        author      = $meta.author
    }
}

function Cmd-Export($a) {
    $doc = Resolve-Doc ([string]$a.docId)
    $out = [string]$a.out
    $outDir = Split-Path -Parent $out
    if (-not (Test-Path -LiteralPath $outDir)) {
        New-Item -ItemType Directory -Force -Path $outDir | Out-Null
    }

    $from = 0
    $to = 0
    $range = $WD_EXPORT_ALL_DOCUMENT
    if ($null -ne $a.from -and $null -ne $a.to) {
        $from = [int]$a.from
        $to = [int]$a.to
        if ($from -gt 0 -and $to -ge $from) { $range = $WD_EXPORT_FROM_TO }
    }

    # OutputFileName, ExportFormat, OpenAfterExport, OptimizeFor, Range, From,
    # To, Item, IncludeDocProps, KeepIRM, CreateBookmarks, DocStructureTags,
    # BitmapMissingFonts, UseISO19005_1
    $doc.ExportAsFixedFormat(
        $out,
        $WD_EXPORT_FORMAT_PDF,
        $false,
        $WD_EXPORT_OPTIMIZE_FOR_PRINT,
        $range,
        $from,
        $to,
        $WD_EXPORT_DOCUMENT_CONTENT,
        $true,
        $true,
        $WD_EXPORT_CREATE_HEADING_BOOKMARKS,
        $true,
        $true,
        $false)

    return @{
        out       = $out
        pageCount = (Get-PageCount $doc)
        sizeBytes = [int64](Get-Item -LiteralPath $out).Length
    }
}

function Cmd-Outline($a) {
    $doc = Resolve-Doc ([string]$a.docId)
    $limit = 2000
    if ($null -ne $a.limit) { $limit = [int]$a.limit }

    $found = New-Object System.Collections.ArrayList

    # Style-based Find costs one COM round trip per heading. Walking every
    # paragraph would cost one per paragraph, which is unusably slow on long
    # documents.
    foreach ($level in 1..9) {
        if ($found.Count -ge $limit) { break }
        $styleId = Get-HeadingStyleId $level
        try {
            $rng = $doc.Content
            $f = $rng.Find
            $f.ClearFormatting()
            $f.Text = ''
            $f.Forward = $true
            $f.Wrap = $WD_FIND_STOP
            $f.Format = $true
            $f.Style = $styleId
            $guard = 0
            while ($guard -lt $limit -and $f.Execute()) {
                $guard++
                $text = ([string]$rng.Text).Trim([char]13, [char]7, [char]11, [char]12, ' ')
                if (-not [string]::IsNullOrWhiteSpace($text)) {
                    [void]$found.Add(@{
                            level = $level
                            text  = $text
                            page  = (Get-PageOf $rng)
                            start = [int]$rng.Start
                        })
                }
                if ($found.Count -ge $limit) { break }
            }
        } catch {
            # A document that does not define this heading style simply has no
            # headings at this level.
        }
    }

    # Fallback for documents using custom heading styles: scan outline levels
    # directly, but only when the document is small enough for that to be cheap.
    if ($found.Count -eq 0) {
        try {
            $paraCount = [int]$doc.Paragraphs.Count
            if ($paraCount -le 2000) {
                for ($i = 1; $i -le $paraCount; $i++) {
                    $p = $doc.Paragraphs.Item($i)
                    $lvl = [int]$p.OutlineLevel
                    if ($lvl -lt $WD_OUTLINE_BODY_TEXT) {
                        $text = ([string]$p.Range.Text).Trim([char]13, [char]7, ' ')
                        if (-not [string]::IsNullOrWhiteSpace($text)) {
                            [void]$found.Add(@{
                                    level = $lvl
                                    text  = $text
                                    page  = (Get-PageOf $p.Range)
                                    start = [int]$p.Range.Start
                                })
                        }
                    }
                    if ($found.Count -ge $limit) { break }
                }
            }
        } catch { }
    }

    $sorted = @($found | Sort-Object { $_.start })
    return @{ headings = $sorted; count = $sorted.Count }
}

function Cmd-Search($a) {
    $doc = Resolve-Doc ([string]$a.docId)
    $query = [string]$a.query
    if ([string]::IsNullOrWhiteSpace($query)) { throw (New-HostError 'invalid_request' "Search query is empty.") }

    $limit = 200
    if ($null -ne $a.limit) { $limit = [int]$a.limit }
    $matchCase = $false
    if ($null -ne $a.matchCase) { $matchCase = [bool]$a.matchCase }
    $wholeWord = $false
    if ($null -ne $a.wholeWord) { $wholeWord = [bool]$a.wholeWord }

    $docEnd = [int]$doc.Content.End
    $hits = New-Object System.Collections.ArrayList

    $rng = $doc.Content
    $f = $rng.Find
    $f.ClearFormatting()
    $f.Text = $query
    $f.Forward = $true
    $f.Wrap = $WD_FIND_STOP
    $f.Format = $false
    $f.MatchCase = $matchCase
    $f.MatchWholeWord = $wholeWord

    $guard = 0
    while ($guard -lt $limit -and $f.Execute()) {
        $guard++
        $start = [int]$rng.Start
        $end = [int]$rng.End
        $ctxStart = [Math]::Max(0, $start - 60)
        $ctxEnd = [Math]::Min($docEnd - 1, $end + 60)
        $snippet = ''
        try {
            $snippet = ([string]$doc.Range($ctxStart, $ctxEnd).Text) -replace '[\r\n\a\v\f\t]+', ' '
            $snippet = $snippet.Trim()
        } catch { }
        [void]$hits.Add(@{
                page    = (Get-PageOf $rng)
                start   = $start
                end     = $end
                snippet = $snippet
            })
    }

    $result = @($hits)
    return @{ hits = $result; count = $result.Count; query = $query }
}

function Cmd-Text($a) {
    $doc = Resolve-Doc ([string]$a.docId)

    $fromPage = 0
    if ($null -ne $a.fromPage) { $fromPage = [int]$a.fromPage }
    $toPage = 0
    if ($null -ne $a.toPage) { $toPage = [int]$a.toPage }

    $totalPages = Get-PageCount $doc

    if ($fromPage -le 0) {
        return @{ text = [string]$doc.Content.Text; fromPage = 1; toPage = $totalPages }
    }

    if ($toPage -le 0 -or $toPage -gt $totalPages) { $toPage = $totalPages }
    if ($fromPage -gt $totalPages) { throw (New-HostError 'page_out_of_range' "Page $fromPage is beyond the end of the document ($totalPages pages).") }

    $startRange = $doc.GoTo($WD_GOTO_PAGE, $WD_GOTO_ABSOLUTE, $fromPage)
    $start = [int]$startRange.Start
    if ($toPage -ge $totalPages) {
        $end = [int]$doc.Content.End
    } else {
        $nextRange = $doc.GoTo($WD_GOTO_PAGE, $WD_GOTO_ABSOLUTE, $toPage + 1)
        $end = [int]$nextRange.Start
    }
    if ($end -lt $start) { $end = $start }

    return @{ text = [string]$doc.Range($start, $end).Text; fromPage = $fromPage; toPage = $toPage }
}

function Cmd-Info($a) {
    $doc = Resolve-Doc ([string]$a.docId)
    return @{
        name       = [string]$doc.Name
        pageCount  = (Get-PageCount $doc)
        wordCount  = (Get-WordCount $doc)
        paragraphs = [int]$doc.Paragraphs.Count
        title      = [string](Get-DocProp $doc 'Title')
        author     = [string](Get-DocProp $doc 'Author')
        subject    = [string](Get-DocProp $doc 'Subject')
        company    = [string](Get-DocProp $doc 'Company')
        created    = [string](Get-DocProp $doc 'Creation date')
        revision   = [string](Get-DocProp $doc 'Revision number')
    }
}

function Cmd-Close($a) {
    $docId = [string]$a.docId
    $script:DocArgs.Remove($docId) | Out-Null
    Close-Doc $docId
    return @{ closed = $true }
}

# --- dispatch loop -----------------------------------------------------------

# Clean up after any previous host that died without tearing its Word down.
Clear-OrphanedWord

$exitRequested = $false
try {
    while (-not $exitRequested) {
        $line = [Console]::In.ReadLine()
        if ($null -eq $line) { break }
        if ([string]::IsNullOrWhiteSpace($line)) { continue }

        $id = $null
        try {
            $req = $line | ConvertFrom-Json
            $id = $req.id
            $cmdArgs = $req.args
            switch ([string]$req.cmd) {
                'ping' { Send-Ok $id (Cmd-Ping $cmdArgs) }
                'open' { Send-Ok $id (Cmd-Open $cmdArgs) }
                'structure' { Send-Ok $id (Cmd-Structure $cmdArgs) }
                'export' { Send-Ok $id (Cmd-Export $cmdArgs) }
                'outline' { Send-Ok $id (Cmd-Outline $cmdArgs) }
                'search' { Send-Ok $id (Cmd-Search $cmdArgs) }
                'text' { Send-Ok $id (Cmd-Text $cmdArgs) }
                'info' { Send-Ok $id (Cmd-Info $cmdArgs) }
                'close' { Send-Ok $id (Cmd-Close $cmdArgs) }
                'quit' {
                    Stop-Word
                    Send-Ok $id @{ stopped = $true }
                    $exitRequested = $true
                }
                default { Send-Fail $id 'unknown_command' "Unknown command '$([string]$req.cmd)'." }
            }
        } catch {
            # A typed failure carries its own code; anything else is genuinely
            # unclassified and stays `word_error`.
            $thrown = $_.TargetObject
            if ($thrown -is [hashtable] -and $thrown.__hostError) {
                Send-Fail $id ([string]$thrown.code) ([string]$thrown.message) $thrown.data
            } else {
                Send-Fail $id 'word_error' $_.Exception.Message
            }
        }
    }
} finally {
    Stop-Word
}
