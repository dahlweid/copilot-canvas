# Mutation check for the create_document unit tests.
#
# Not part of the suite: a test that cannot fail is invisible to a green run, so
# each defect below is reintroduced, the suite is run, and the run must go red.
# Anything reported as SURVIVED is a test that proves nothing -- and anything
# reported as MISSING or AMBIGUOUS is a *mutant* that proves nothing, which is
# the worse of the two because the total keeps counting it. See the comment
# above the counters for why the three are reported apart.
#
# Run from the extension root:
#   powershell -File test/unit/mutate-create.ps1

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

# Built from a code point rather than typed. Windows PowerShell reads a BOM-less
# UTF-8 script as ANSI, so an en dash written literally here would not match the
# one in the source file -- and the mutant would report MISSING, which reads like
# a broken anchor rather than an unkilled defect.
$dash = [char]0x2013

$mutants = @(
    @{ name = 'unknown fields ignored instead of refused'
       file = 'src/word/create-intent.mjs'
       from  = '    if (extra.length > 0) {
        fail(
            "invalid_block",'
       to    = '    if (false) {
        fail(
            "invalid_block",' }

    @{ name = 'ragged tables padded instead of refused'
       file = 'src/word/create-intent.mjs'
       from  = '                if (row.length !== width) {'
       to    = '                if (false) {' }

    @{ name = 'spec-level extra fields dropped'
       file = 'src/word/create-intent.mjs'
       from  = '    const extra = Object.keys(rest);
    if (extra.length > 0) {
        fail("invalid_spec"'
       to    = '    const extra = Object.keys(rest);
    if (false) {
        fail("invalid_spec"' }

    @{ name = 'BLOCK_HELP heading range hardcoded'
       file = 'src/word/create-intent.mjs'
       from  = '${MIN_BLOCK_HEADING_LEVEL}' + $dash + '${MAX_HEADING_LEVEL}.' + [char]0x60 + ','
       to    = '1' + $dash + '9.' + [char]0x60 + ',' }

    @{ name = 'schema maxItems hardcoded'
       file = 'extension.mjs'
       from  = '                maxItems: MAX_BLOCKS,'
       to    = '                maxItems: 500,' }

    @{ name = 'schema block kinds spelled out'
       file = 'extension.mjs'
       from  = '            enum: BLOCK_KINDS,'
       to    = '            enum: ["heading", "paragraph", "list", "table"],' }

    @{ name = 'creatable extensions spelled out'
       file = 'extension.mjs'
       from  = 'description: `Absolute or workspace-relative path to create (${creatableList()}).'
       to    = 'description: `Absolute or workspace-relative path to create (.docx).' }

    # Round 3 of review found this message opened with the bare extension --
    # ".rtf cannot be created" -- which reads as though the extension were the
    # thing being created, and never named the file. An agent that authored
    # several documents in one turn cannot tell from that which call to change.
    @{ name = 'unsupported_type refusal does not name the file it refused'
       file = 'src/word/document-author.mjs'
       from  = 'Cannot create ${docPath}: create_document does not write ${ext} files. It writes ${creatableList()}.'
       to    = '${ext} cannot be created. create_document writes ${creatableList()}.' }

    # Replaces two mutants that each moved one detail from `data` to a top-level
    # property. Both are now unbuildable: `CreateError` mirrors `details` onto
    # `data`, so a top-level detail is no longer a dropped one, and one of those
    # mutants' `to` text is verbatim what the file now ships. That is the
    # constructor closing the class rather than the mutants going stale --
    # per-site defects cannot be injected once the site is not where the
    # guarantee lives. So the mutant moves to where the guarantee moved.
    #
    # Single-line anchor deliberately: the two it replaces were multi-line, and
    # a multi-line anchor here fails silently on the CRLF/LF split.
    @{ name = 'details never reach data, so every create failure arrives bare'
       file = 'src/word/document-author.mjs'
       from  = '        if (details && Object.keys(details).length > 0) this.data = { ...details };'
       to    = '' }

    @{ name = 'no check that the file actually exists after SaveAs2'
       file = 'src/word/document-author.mjs'
       from  = '        if (!revisionToken) {'
       to    = '        if (false) {' }

    @{ name = 'a read-back failure reported as a failed create'
       file = 'src/word/document-author.mjs'
       from  = '                "document_unreadable",
                `${path.basename(docPath)} was created and saved, but reading it back afterwards failed `'
       to    = '                "create_failed",
                `${path.basename(docPath)} could not be created. `' }

    @{ name = 'an existing file overwritten'
       file = 'src/word/document-author.mjs'
       from  = '        if (await isExistingFile(docPath)) failFromStatus({ status: "file_exists" }, docPath);'
       to    = '        if (false) failFromStatus({ status: "file_exists" }, docPath);' }

    # Re-anchored. The original anchored on an object literal at a call site;
    # the three literals were replaced by one `reportedAutoCorrect` helper, so
    # the anchor stopped matching and the runner reported MISSING -- the mutant
    # silently stopped being evidence while still being listed. Anchoring in the
    # helper is also stronger: it mutates all three call sites at once.
    @{ name = 'autocorrect outcome reported as always suppressed'
       file = 'src/word/document-author.mjs'
       from  = '        suppressed: Boolean(ac.suppressed),
        reason: ac.reason ?? null,'
       to    = '        suppressed: true,
        reason: null,' }

    # The restore half of the same defect. `suppressed` was pinned in both
    # directions by two fixtures and `restored` in neither, so this mutant
    # survived every unit test until an all-fields-populated fixture was added.
    # Suppression proven and restoration unproven is the same defect one step
    # later, and that is true of the mutants as well as of the code.
    @{ name = 'autocorrect restore outcome reported as never restored'
       file = 'src/word/document-author.mjs'
       from  = '        restored: Boolean(ac.restored),
        restoreReason: ac.restoreReason ?? null,'
       to    = '        restored: false,
        restoreReason: null,' }

    @{ name = 'lock release never flagged'
       file = 'src/word/document-author.mjs'
       from  = '        const lockReleased = result.released !== false;'
       to    = '        const lockReleased = true;' }

    @{ name = 'unsupported extensions accepted'
       file = 'src/word/document-author.mjs'
       from  = '        if (!CREATABLE.has(ext)) {'
       to    = '        if (false) {' }

    @{ name = 'spec validated after the host is called'
       file = 'src/word/document-author.mjs'
       from  = '        const spec = validateSpec(rawSpec);'
       to    = '        const spec = { blocks: rawSpec?.blocks ?? [] };' }

    # A predicted paragraph count used to be returned here and was wrong by
    # construction — it counted in Word's COM coordinate system while the map is
    # OOXML-derived. Reinstating it must go red, or the pin against it is decor.
    @{ name = 'a predicted paragraph count is put back'
       file = 'src/word/create-intent.mjs'
       from  = '/** A short, human-readable description, for logs and snapshot manifests. */'
       to    = @'
export function paragraphsIn(spec) { return spec.blocks.length; }

/** A short, human-readable description, for logs and snapshot manifests. */
'@ }

    # The message used to assert "no document was written" unconditionally, which
    # the host cannot promise -- its cleanup delete is best-effort inside a
    # swallowed catch. Going back to a fixed claim must go red.
    @{ name = 'failed-create message asserts a cleanup it never checked'
       file = 'src/word/document-author.mjs'
       from  = '                    (leftBehind'
       to    = '                    (false' }

    # --- the schema/runtime divergence found in review round 1 ---------------
    #
    # All five are one defect class: the tool schema is a second copy of a
    # contract create-intent.mjs already defines. The shipped version declared no
    # floors at all while the validator refused every empty collection, so a
    # caller could construct a request that was schema-valid and rejected every
    # time. A model cannot learn a rule the schema does not state.

    @{ name = 'schema floor dropped, so an empty list is schema-valid and always refused'
       file = 'extension.mjs'
       from  = '            minItems: MIN_LIST_ITEMS,
'
       to    = '' }

    @{ name = 'schema floor hardcoded rather than derived'
       file = 'extension.mjs'
       from  = '                minItems: MIN_BLOCKS,'
       to    = '                minItems: 1,' }

    @{ name = 'per-kind clause written by hand beside BLOCKS instead of derived'
       file = 'extension.mjs'
       from  = 'description: `${fieldUsage("ordered")} True numbers the items'
       to    = 'description: `list only: true numbers the items' }

    @{ name = 'fieldUsage answers for a field no kind takes'
       file = 'src/word/create-intent.mjs'
       from  = '        throw new Error(`no block kind takes'
       to    = '        return ""; throw new Error(`no block kind takes' }

    @{ name = 'description claims autocorrect is unconditionally off'
       file = 'extension.mjs'
       from  = '        "Text is written verbatim. Autocorrect is switched off first on a Word this tool started, so",'
       to    = '        "Text is written verbatim ' + $dash + ' Word' + "'" + 's autocorrect is switched off on the instance that authors it, so",' }

    # Round 1 flagged the missing floors; auditing the whole boundary rather than
    # the flagged fields found the text length undeclared as well, on all three
    # string-bearing fields, and the autoCorrect outcome invisible on the one
    # failure path where a document really was authored.

    @{ name = 'text length enforced but not declared to the model'
       file = 'extension.mjs'
       from  = '            maxLength: MAX_TEXT_LENGTH,
'
       to    = '' }

    # Anchored on the `cause:` line above the block, not on the block alone.
    # `create_unverified` carries an identical `autoCorrect` line, so the shorter
    # anchor matched twice and the runner reported AMBIGUOUS -- it would have
    # deleted the field from both failures while claiming to test one. That is
    # now truer than when it was written: the literal became a shared
    # `reportedAutoCorrect(result)` call appearing at three sites, so the one
    # distinguishing line is `cause: err.code ?? null` -- `create_unverified`
    # uses the shorthand `cause,`.
    @{ name = 'autoCorrect dropped from the one failure that still authored a document'
       file = 'src/word/document-author.mjs'
       from  = '                    cause: err.code ?? null,
                    autoCorrect: reportedAutoCorrect(result),
'
       to    = '                    cause: err.code ?? null,
' }

    @{ name = 'a directory refused as an existing file, shadowing the host classifier'
       file = 'src/word/document-author.mjs'
       from  = '        return (await stat(docPath)).isFile();'
       to    = '        await stat(docPath);
        return true;' }

    @{ name = 'an unreadable file reported as a file that was never written'
       file = 'src/word/document-author.mjs'
       from  = '            if (cause === "ENOENT") {'
       to    = '            if (true) {' }

    # additionalProperties: false enforces nothing on this host (issue #28), so
    # the only thing that can refuse an unknown argument is validateSpec -- and
    # only if the handler hands it one rather than picking `blocks` out.
    @{ name = 'handler picks blocks out, so an unknown argument is dropped instead of refused'
       file = 'extension.mjs'
       from  = '        const { path: _path, ...spec } = args ?? {};'
       to    = '        const spec = { blocks: args?.blocks };' }
)

# Three ways a mutant can fail to be evidence, and they need different fixes, so
# they are counted separately. Folding them together is how "26/26" starts
# reading as strength while coverage falls:
#
#   SURVIVED  the defect was introduced and no test noticed.   -> weak test
#   MISSING   the anchor no longer exists, so nothing was       -> stale anchor
#             introduced at all. The commonest cause is an
#             ordinary edit to the file under test moving the
#             line; the mutant then silently stops testing
#             anything while still being counted.
#   AMBIGUOUS the anchor appears more than once, so the mutant  -> narrow anchor
#             changes more than its name claims and a kill
#             cannot be attributed to the defect named.
#
# All three are fatal. The distinction is in the reporting, because a reader
# told "survived" will go looking for a weak test, and for MISSING there isn't
# one -- the test is fine and the anchor moved.
$survived = @()
$missing = @()
$ambiguous = @()

# A mutant is only informative against a suite that is otherwise green. The first
# run of this script reported 16 of 16 killed while one test was throwing ENOENT
# on a mistyped path -- the suite was red before any mutant was applied, so every
# "kill" was the same pre-existing failure. Refuse to run at all in that state.
& node --test 'test/unit/create-intent.test.mjs' 'test/unit/document-author.test.mjs' *> $null
if ($LASTEXITCODE -ne 0) {
    Write-Host 'Baseline is red. Fix the suite before mutating; a red baseline kills every mutant.' -ForegroundColor Red
    exit 1
}

foreach ($m in $mutants) {
    $file = Join-Path $root $m.file
    $original = [IO.File]::ReadAllText($file)

    # Counted, not just tested for presence. `String.Replace` replaces every
    # occurrence, so an anchor that matches twice quietly widens the mutant
    # beyond the defect its name describes.
    $occurrences = ([regex]::Matches($original, [regex]::Escape($m.from))).Count
    if ($occurrences -eq 0) {
        Write-Host ("MISSING  {0} -- anchor not found in {1}; re-anchor it" -f $m.name, $m.file) -ForegroundColor Yellow
        $missing += $m.name
        continue
    }
    if ($occurrences -gt 1) {
        Write-Host ("AMBIGUOUS {0} -- anchor matches {1}x in {2}; narrow it" -f $m.name, $occurrences, $m.file) -ForegroundColor Yellow
        $ambiguous += $m.name
        continue
    }
    if ($m.from -eq $m.to) {
        Write-Host ("MISSING  {0} -- mutant is a no-op in {1}" -f $m.name, $m.file) -ForegroundColor Yellow
        $missing += $m.name
        continue
    }

    [IO.File]::WriteAllText($file, $original.Replace($m.from, $m.to))
    try {
        & node --test 'test/unit/create-intent.test.mjs' 'test/unit/document-author.test.mjs' *> $null
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

Write-Host ''
$broken = $survived.Count + $missing.Count + $ambiguous.Count
if ($broken -eq 0) {
    Write-Host ("All {0} mutants killed." -f $mutants.Count) -ForegroundColor Green
    exit 0
}
# Reported as "not evidence" rather than "survived": only the first group means
# a test is weak. The other two mean the mutant never ran, which is the failure
# that inflates the count instead of reducing it.
Write-Host ("{0} of {1} mutants are not evidence:" -f $broken, $mutants.Count) -ForegroundColor Red
if ($survived.Count -gt 0) {
    Write-Host "  survived (the defect was introduced and no test noticed):" -ForegroundColor Red
    $survived | ForEach-Object { Write-Host "    - $_" }
}
if ($missing.Count -gt 0) {
    Write-Host "  missing anchor (nothing was introduced; the test may be fine):" -ForegroundColor Yellow
    $missing | ForEach-Object { Write-Host "    - $_" }
}
if ($ambiguous.Count -gt 0) {
    Write-Host "  ambiguous anchor (more was changed than the name claims):" -ForegroundColor Yellow
    $ambiguous | ForEach-Object { Write-Host "    - $_" }
}
exit 1
