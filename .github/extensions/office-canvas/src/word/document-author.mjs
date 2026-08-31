// create_document, end to end.
//
// Shape of a create, and why it is this shape:
//
//   validate spec -> refuse an existing file -> one Word operation
//                                                 -> confirm on disk -> read back
//
// **It will not overwrite.** There is no flag for it and adding one would be a
// mistake. `create_document` creating a document is the whole contract; replacing
// the contents of a file that already exists is what `edit_document` is for, and
// that path has a revision token, an on-disk snapshot and a revert behind it.
// Routing an overwrite through here would be the same destruction with none of
// those. So an existing path is refused, and the caller is told which tool to
// reach for.
//
// The refusal buys a second thing. An existing file may be open in Word, and this
// repo has now revised its model of what that means four times (kickoff, PR #22,
// PR #24, and the run below). Measured directly --
// spikes/isolation/probes/probe-word-share-mode.ps1 -- real Word takes *write*
// access and grants `FileShare::Read`. A request for write access against it
// therefore fails on Windows' first rule, requested access against the holder's
// granted share mode, whatever share mode the requester offers.
//
// That conclusion rests on the *share* column, not the access column, and an
// earlier draft of this comment cited only the latter: a holder taking write
// access while granting `ReadWrite` permits the very request this paragraph says
// is refused. Both columns are measured now, so the sentence carries the one it
// actually needs.
//
// Never opening an existing file for writing means none of it is load-bearing
// here: the only file this module ever writes is one that did not exist a moment
// ago.
//
// Autocorrect
// -----------
// Autocorrect is NOT suppressed, and this module no longer reports on it. It
// once did: the host disabled five settings around authoring, and the outcome
// crossed the tool boundary as a result field so a test could assert it. Both
// are gone. The settings persist for the user rather than being per-process,
// and they were measured not to affect any insertion path this code uses — see
// the block above `Initialize-Word` in `word-host.ps1`. What replaces the
// reported field is a bait assertion in `test/integration/create-smoke.mjs`,
// which checks the text itself rather than a claim about a setting.

import { stat } from "node:fs/promises";
import path from "node:path";

import { fileRevisionToken } from "../revision-token.mjs";
import { describeSpec, validateSpec } from "./create-intent.mjs";

/**
 * A typed create failure.
 *
 * Identical to `EditError` in `document-editor.mjs`, deliberately and by
 * adoption rather than by coincidence: `details` is copied onto `data` as well
 * as onto the error, and the `data` half is the one that reaches the agent,
 * because `toolFailure` renders only `code`, `message` and `data` into the text
 * the agent receives.
 *
 * This class previously did the `Object.assign` half and not the mirror, so a
 * detail reached the agent only if the *caller* remembered to nest it under
 * `data` by hand. All four call sites did remember, so nothing was broken and
 * no test was red. It was still the wrong shape, for the reason `EditError`'s
 * own comment gives about the bug it was extracted from: "a per-site fix could
 * not close the class; a constructor can." This was that per-site fix, written
 * after the constructor that supersedes it, and it made the natural mistake --
 * passing a flat detail, as every `EditError` call site does -- silently
 * unobservable rather than loud. It was found by asking whether the neighbours
 * had already solved this, not by a failure.
 *
 * Nine of the eleven error classes here take no details at all; of the two that
 * do, this was the only one whose contract differed. There is now one contract.
 */
export class CreateError extends Error {
    constructor(code, message, details = {}) {
        super(message);
        this.name = "CreateError";
        this.code = code;
        Object.assign(this, details);
        // Only when there is something to carry, so an error without details
        // does not arrive at the boundary with an empty object attached.
        if (details && Object.keys(details).length > 0) this.data = { ...details };
    }
}

/**
 * Wall-clock ceiling on one `create_document`, covering the authoring and the
 * read-back that confirms it.
 *
 * Same reasoning and same figure as `EDIT_BUDGET_MS`: measured authoring is
 * 120–420 ms warm (spikes/isolation/probes/probe-authoring-save.ps1) and a cold
 * Word start is ~4.5 s, so five minutes is not a performance limit. It exists so
 * that a wedged Word surfaces as an error the caller can act on rather than as
 * silence. Those measurements are typical-case on a quiet machine and are
 * explicitly not bounds, which is the reason the budget is not derived from them.
 */
const CREATE_BUDGET_MS = 300_000;

/**
 * The extensions this module will author into.
 *
 * Narrower than the set the reader opens, and deliberately so: `SaveAs2` is
 * called with wdFormatXMLDocument, so `.docx` is the only extension whose name
 * would match its contents. Writing that payload to a `.doc` or a `.dotx` would
 * produce a file whose extension lies about it.
 */
export const CREATABLE = new Set([".docx"]);

/** The creatable extensions as prose, for tool descriptions and messages. */
export const creatableList = () => [...CREATABLE].join(", ");

/**
 * Maps the host's structured status onto typed errors.
 *
 * The host reports failure as a `status` on a successful response rather than by
 * throwing, for the reason `document-editor.mjs` documents: the dispatch loop
 * collapses every exception into one `word_error` whose only distinguishing
 * feature is a message that is localized German on this machine.
 */
function failFromStatus(result, docPath) {
    const name = path.basename(docPath);
    switch (result.status) {
        case "invalid_path":
            throw new CreateError("invalid_path", "A document path is required.");
        case "path_is_directory":
            throw new CreateError("invalid_path", `${docPath} is a directory, not a file.`);
        case "file_exists":
            throw new CreateError(
                "file_exists",
                `${name} already exists. create_document will not overwrite a document; use edit_document to change ` +
                    `one, or choose a path that does not exist yet.`,
            );
        case "directory_not_found":
            throw new CreateError(
                "directory_not_found",
                `The folder ${result.directory ?? path.dirname(docPath)} does not exist. create_document writes into ` +
                    `an existing folder; it does not create one.`,
            );
        case "create_failed": {
            // What was observed, and nothing about why.
            //
            // The only thing established here is that a COM call raised. The
            // exception *type* is carried in `data` for a caller that wants to
            // discriminate, because the message is German on this machine and
            // matching on it is how a contract rots. The prose names no cause,
            // because none was distinguished.
            //
            // It also does not claim an outcome nobody checked. This used to end
            // "and no document was written", which the host could not promise:
            // its cleanup is a best-effort `Remove-Item` inside a swallowed
            // `catch`, so a failed delete left a partial document on disk while
            // the caller was told the opposite. The host now looks, and this
            // says whichever it saw.
            const leftBehind = result.leftBehind === true;
            throw new CreateError(
                "create_failed",
                `Word raised an error while authoring ${name}. ` +
                    (leftBehind
                        ? `A partial document was left at ${docPath} and could not be removed.`
                        : `No document was left behind.`),
                { exception: result.exception ?? null, detail: result.detail ?? null, leftBehind },
            );
        }
        default:
            throw new CreateError("create_failed", `Word reported '${result.status}' while creating ${name}.`);
    }
}

export class DocumentAuthor {
    #reader;
    #host;
    #log;

    constructor({ reader, host, log = () => {} }) {
        this.#reader = reader;
        this.#host = host;
        this.#log = log;
    }

    /** Normalizes and rejects anything this module will not author into. */
    static requireCreatable(docPath) {
        const ext = path.extname(docPath).toLowerCase();
        if (!CREATABLE.has(ext)) {
            throw new CreateError(
                "unsupported_type",
                ext
                    ? `Cannot create ${docPath}: create_document does not write ${ext} files. It writes ${creatableList()}.`
                    : `Cannot create ${docPath}: the path has no file extension. create_document writes ${creatableList()}.`,
            );
        }
        return docPath;
    }

    /**
     * Authors a document from a spec and returns it as it now stands on disk.
     *
     * No per-document lock, unlike the edit path. That lock exists because two
     * concurrent edits can both pass their preconditions against the same bytes
     * and both apply. Two concurrent creates cannot: the host's dispatch loop is
     * a single `ReadLine` switch, so the second `create` does not begin until
     * the first has returned, and its existence check then sees the finished
     * file. Serialization, not locking, is what closes that case.
     *
     * This comment used to say the host "re-checks immediately before
     * `SaveAs2`", and that what fitted in the remaining window was an overwrite
     * of a file created "microseconds earlier by a call the agent issued
     * itself". Both halves were wrong, and wrong in opposite directions. The
     * re-check was at the top of `Cmd-Create`, with a cold Word start and the
     * whole block build between it and the save, so the window was seconds
     * rather than microseconds. And the racing pair it reasoned about was the
     * one case that could not happen, while the case it did not mention -- any
     * other writer to that path, which `SaveAs2` overwrites without a prompt --
     * was the one the window actually exposed.
     *
     * The check now is where this comment always claimed it was, so the window
     * is the gap between two adjacent lines. It is still not zero: `SaveAs2`
     * has no create-exclusive mode, so a check is the only instrument available.
     */
    async create(docPath, rawSpec) {
        const spec = validateSpec(rawSpec);
        DocumentAuthor.requireCreatable(docPath);

        // One wall clock for the whole operation, not one per layer. Spent, not
        // restarted: a step that would begin with nothing left fails with the
        // reason rather than blocking on a Word that is not coming back.
        const deadline = Date.now() + CREATE_BUDGET_MS;
        const remaining = (label) => {
            const left = deadline - Date.now();
            if (left <= 0) {
                throw new CreateError(
                    "word_timeout",
                    `Creating ${path.basename(docPath)} exceeded its ${Math.round(CREATE_BUDGET_MS / 1000)}s budget ` +
                        `before ${label} could start.`,
                );
            }
            return left;
        };

        // Cheapest possible refusal, before Word is involved at all. The host
        // checks this again; that is not redundancy but the same non-atomicity
        // the edit path has, and the authoritative check is the later one --
        // which is also why this asks only about a regular file and lets every
        // other kind of entry through to be classified there.
        if (await isExistingFile(docPath)) failFromStatus({ status: "file_exists" }, docPath);

        const result = await this.#host.create({
            path: docPath,
            blocks: spec.blocks,
            timeoutMs: remaining("the authoring itself"),
        });

        if (result.status !== "created") failFromStatus(result, docPath);

        // Word said it saved. Confirm against the filesystem before saying so to
        // the caller: `SaveAs2` returning is not evidence a file exists, and
        // "created" with nothing on disk is the failure an agent cannot detect.
        //
        // Two outcomes, not one. `.catch(() => null)` used to collapse every
        // failure into "no file", and the message then asserted absence -- a
        // cause a failed read cannot establish. `fileRevisionToken` opens a read
        // stream, so a failure arrives as an errno: `ENOENT` genuinely means
        // nothing is there, while `EACCES`/`EPERM`/`EBUSY` mean the file could
        // not be read and may well exist. (Read streams are the context the
        // errno mapping in `document-reader.mjs` was measured in, so it applies
        // here directly rather than by transport.)
        //
        // They are split because they want opposite responses at the worst
        // possible moment -- Word has just claimed it saved. "Nothing on disk"
        // invites another attempt; "could not read it back" must not, because
        // re-authoring would write over a document that already exists. This is
        // the repo's own rule: split where the platform distinguishes.
        let revisionToken = null;
        let verifyFailure = null;
        try {
            revisionToken = await fileRevisionToken(docPath);
        } catch (err) {
            verifyFailure = err;
        }
        if (!revisionToken) {
            const cause = verifyFailure?.code ?? null;
            if (cause === "ENOENT") {
                throw new CreateError(
                    "create_not_persisted",
                    `Word reported ${path.basename(docPath)} as created but there is no file at that path.`,
                    { created: false, cause },
                );
            }
            // States what was observed and names no cause beyond the errno the
            // platform supplied. `created` is null, not false: whether a
            // document exists is exactly what could not be established.
            throw new CreateError(
                "create_unverified",
                `Word reported ${path.basename(docPath)} as created, but reading the file back to verify that failed ` +
                    `(${verifyFailure?.message ?? "no error was reported"}). Do not create it again — a document may ` +
                    `exist at that path, and re-authoring would overwrite it. Use read_document to find out.`,
                {
                    created: null,
                    cause,
                },
            );
        }

        // Complete, never paged, for the reason the edit path re-reads
        // completely: the caller holds no addresses at all yet, and handing back
        // the first page of a long document would leave it unable to address the
        // rest without another round trip.
        //
        // By this line the document exists on disk. So a failure here is a
        // failure to *describe* it, not to create it, and must not be reported as
        // though nothing had happened — an agent told "create failed" would retry
        // against a path that now exists and get `file_exists` for its trouble.
        let document;
        try {
            document = await this.#reader.read(docPath, { limit: 0, deadline });
        } catch (err) {
            throw new CreateError(
                "document_unreadable",
                `${path.basename(docPath)} was created and saved, but reading it back afterwards failed ` +
                    `(${err.message}). Do not repeat the call — the file exists. Read it with read_document to get ` +
                    `its addresses.`,
                {
                    created: true,
                    cause: err.code ?? null,
                },
            );
        }

        // The host polls for the file to be released rather than trusting
        // `Document.Close()`, and reports the answer. `released: false` means
        // something genuinely unusual, and the next write to this document will
        // fail its pre-flight — worth saying so, or that failure looks like a
        // fresh and unexplained lock rather than a consequence of this call.
        const lockReleased = result.released !== false;
        if (!lockReleased) {
            this.#log(
                `create_document: ${path.basename(docPath)} was still held ${result.releaseMs}ms after it was closed. ` +
                    `The document is saved; a follow-up edit may be refused until the handle goes.`,
            );
        }

        this.#log(
            `create_document: ${describeSpec(spec)} at ${path.basename(docPath)} ` +
                `(build ${result.buildMs}ms, save ${result.saveMs}ms, release ${result.releaseMs}ms, ` +
                `total ${result.totalMs}ms)`,
        );

        return {
            created: { path: docPath, description: describeSpec(spec), blocks: spec.blocks.length },
            // No predicted paragraph count is reported. An earlier version
            // returned one and it was wrong by construction: it modelled Word's
            // COM `Paragraphs.Count`, which counts a row-end mark per table row,
            // while every paragraph number the caller can act on — including
            // `document.paragraphCount` and every address in the map — comes
            // from OOXML, where row-end marks are not paragraphs. Measured on
            // the smoke fixture: the prediction said 18, the map says 14. Two
            // coordinate systems joined on a bare number is the same trap as
            // joining on a localized style name, and the authoritative count is
            // already in `document` below.
            tableCount: result.tableCount ?? 0,
            lockReleased,
            revisionToken,
            timings: {
                buildMs: result.buildMs ?? null,
                saveMs: result.saveMs ?? null,
                releaseMs: result.releaseMs ?? null,
                totalMs: result.totalMs ?? null,
            },
            document,
        };
    }
}

/**
 * Whether `docPath` names an existing *regular file*.
 *
 * `stat` resolves for any filesystem entry, so testing only that it resolved
 * made a directory take the `file_exists` refusal — telling the caller the file
 * already exists and to use `edit_document` on it, about a directory.
 *
 * The damage is not only the wrong answer. The host classifies this correctly
 * (`path_is_directory`, `word-host.ps1`), and this preflight runs *upstream* of
 * it, so a true-for-anything test does not merely answer badly: it shadows the
 * component that answers well, exactly as picking `blocks` out of the tool
 * arguments once shadowed the validator's refusal of unknown fields.
 *
 * So anything that is not a regular file falls through to the host on purpose.
 */
async function isExistingFile(docPath) {
    try {
        return (await stat(docPath)).isFile();
    } catch {
        return false;
    }
}
