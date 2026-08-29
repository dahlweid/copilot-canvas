# Mutation check for the create_document unit tests.
#
# Not part of the suite: a test that cannot fail is invisible to a green run, so
# each defect below is reintroduced, the suite is run, and the run must go red.
# Anything reported as SURVIVED is a test that proves nothing.
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

    @{ name = 'exception type set top-level, dropped at the tool boundary'
       file = 'src/word/document-author.mjs'
       from  = '                { data: { exception: result.exception ?? null, detail: result.detail ?? null, leftBehind } },'
       to    = '                { exception: result.exception ?? null, detail: result.detail ?? null, leftBehind },' }

    @{ name = 'created flag set top-level, dropped at the tool boundary'
       file = 'src/word/document-author.mjs'
       from  = '                { data: { created: true, cause: err.code ?? null } },'
       to    = '                { created: true, cause: err.code ?? null },' }

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
       from  = '        if (await exists(docPath)) failFromStatus({ status: "file_exists" }, docPath);'
       to    = '        if (false) failFromStatus({ status: "file_exists" }, docPath);' }

    @{ name = 'autocorrect outcome reported as always suppressed'
       file = 'src/word/document-author.mjs'
       from  = '                suppressed: Boolean(result.autoCorrect?.suppressed),
                reason: result.autoCorrect?.reason ?? null,'
       to    = '                suppressed: true,
                reason: null,' }

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
)

$survived = @()

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
    if (-not $original.Contains($m.from)) {
        Write-Host ("MISSING  {0} -- anchor not found in {1}" -f $m.name, $m.file) -ForegroundColor Yellow
        $survived += $m.name
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
if ($survived.Count -eq 0) {
    Write-Host ("All {0} mutants killed." -f $mutants.Count) -ForegroundColor Green
    exit 0
}
Write-Host ("{0} of {1} mutants survived:" -f $survived.Count, $mutants.Count) -ForegroundColor Red
$survived | ForEach-Object { Write-Host "  - $_" }
exit 1
