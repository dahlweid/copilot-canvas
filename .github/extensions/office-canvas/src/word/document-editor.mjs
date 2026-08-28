// edit_document and revert_document, end to end.
//
// Shape of an edit, and why it is this shape:
//
//   token check -> read -> resolve address -> snapshot -> one Word operation
//                                                          -> confirm -> re-read
//
// **One edit per call, against one read.** An address is a coordinate, not a
// handle: it is minted from (heading path, normalized text, occurrence), so it
// is stable across the mutations that leave those three alone — appending,
// inserting elsewhere, rewriting one paragraph, Word re-splitting runs — but
// two edits move addresses that were valid a moment ago:
//
//   * Deleting one of several identically-worded paragraphs renumbers the rest.
//     The occurrence counter shifts, so the *second* duplicate inherits the
//     first one's address. That address still resolves. It just resolves to
//     different content, which is the failure mode you cannot see.
//   * Renaming a heading changes the heading path of everything beneath it, and
//     so every address under it, and its own.
//
// A batch API would therefore be quietly lying: it would accept a list of
// addresses minted before the first edit and apply the later ones to a document
// where they no longer mean what the caller meant. Refusing to batch is not a
// limitation to be lifted later, it is the honest interface for this addressing
// scheme. Callers that want several changes read, edit, and read again.
//
// The revision token is checked twice on purpose. Once before any Word work,
// because refusing a stale edit should cost a file hash and not a Word start;
// and once as part of the read, which closes the window between the two.
//
// Neither check closes the read->write window, and nothing about a token can:
// two calls that read the same bytes both hold a valid token and would both be
// authorized. Serialization is what closes that window, so every operation that
// writes a given document runs under a per-document lock — see `#withLock`.

import { open, rm, stat } from "node:fs/promises";
import path from "node:path";

import { fileRevisionToken, tokensMatch } from "../revision-token.mjs";
import { describeIntent, validateIntent } from "./edit-intent.mjs";
import { listSnapshots, revertToLatest, takeSnapshot } from "./snapshots.mjs";

export class EditError extends Error {
    constructor(code, message, details = {}) {
        super(message);
        this.name = "EditError";
        this.code = code;
        Object.assign(this, details);
    }
}

/**
 * Wall-clock ceiling on one `edit_document`, covering both reads and the edit.
 *
 * Five minutes is far longer than any measured edit -- 281 ms warm, 2.7 s on a
 * document carrying the mark of the web, and those are typical-case figures
 * rather than bounds -- so it is not a performance limit. It exists so a wedged
 * Word surfaces as an error the caller can act on instead of as silence.
 */
const EDIT_BUDGET_MS = 300_000;

/**
 * Maps the host's structured status onto typed errors.
 *
 * The host reports failure as a `status` field on a successful response rather
 * than by throwing, because the dispatch loop collapses every exception into a
 * single `word_error` whose only distinguishing feature is a message that is
 * localized German on this machine. Branching on that would mean a regex per
 * failure mode, which is how a contract rots.
 */
function failFromStatus(result, docPath, intent) {
    const name = path.basename(docPath);
    switch (result.status) {
        case "file_not_found":
            throw new EditError("file_not_found", `No such file: ${docPath}`);
        case "document_locked":
            // The shared `file_locked` code, but the meaning differs from the
            // read path's use of it, in both directions.
            //
            // For a read, `file_locked` means a holder *stricter than Word*,
            // because a document open in Word can still be copied -- so naming
            // Word would be actively wrong. For a write there is no such gap:
            // every write open fails while Word holds the file, so unlike the
            // read path a document open in Word *is* a possible cause here.
            //
            // It is still not a cause we may state. `Test-FileWritable` takes a
            // write handle, and a write handle cannot distinguish a sharing
            // violation from a denying ACL from a read-only attribute -- the
            // flag is collapsed by construction. So the message reports the
            // fact it actually established, that the file could not be opened
            // for writing, and offers the possibilities without asserting one.
            throw new EditError(
                "file_locked",
                `${name} could not be opened for writing. Another program may be holding it, or it may be protected ` +
                    `against writing.`,
            );
        case "open_failed":
            throw new EditError(
                "document_unreadable",
                `Word could not open ${name} for editing. It may be password-protected or corrupt.`,
                { detail: result.detail ?? null },
            );
        case "address_not_resolvable":
            throw new EditError(
                "address_not_resolvable",
                result.reason === "text_mismatch"
                    ? `${intent.address} no longer points at the text it was minted from — ${name} changed between the read and the edit. Read it again and use the new address.`
                    : `${intent.address} resolves to paragraph ${result.wordIndex}, but ${name} only has ${result.paragraphCount} paragraphs. Read it again.`,
                {
                    reason: result.reason ?? null,
                    expectedText: result.expectedText ?? null,
                    actualText: result.actualText ?? null,
                },
            );
        case "edit_had_no_effect":
            // Word neither threw nor changed anything. The known cause is a
            // delete inside a table cell: a cell must retain one paragraph, so
            // Word declines the deletion silently rather than raising. Reported
            // as its own status so the caller is not told an edit was applied
            // that was not.
            throw new EditError(
                "edit_failed",
                `Deleting paragraph ${result.wordIndex} of ${name} had no effect — it still has ` +
                    `${result.paragraphCount} paragraphs. A paragraph that is the only one in a table cell cannot be ` +
                    `deleted, because a cell must contain at least one; replace its text with an empty string instead.`,
                { paragraphCount: result.paragraphCount ?? null, expectedCount: result.expectedCount ?? null },
            );
        default:
            throw new EditError("edit_failed", `Word reported '${result.status}' while editing ${name}.`);
    }
}

export class DocumentEditor {
    #reader;
    #host;
    #snapshotRoot;
    #log;
    /** path (lowercased) -> tail of the promise chain writing that document. */
    #locks = new Map();

    constructor({ reader, host, snapshotRoot, log = () => {} }) {
        this.#reader = reader;
        this.#host = host;
        this.#snapshotRoot = snapshotRoot;
        this.#log = log;
    }

    /**
     * Serializes everything that writes one document.
     *
     * A revision token authorizes an edit but cannot serialize one: two calls
     * that read the same bytes both hold a valid token, and both would pass
     * every check. Since `replace_text` does not shift `wordIndex` and does not
     * touch another paragraph's text, two concurrent edits to different
     * paragraphs of the same document both satisfy their pre-mutation checks
     * and both apply -- while both snapshots hold the same pre-first-edit
     * bytes, so one revert silently undoes two edits and reports one.
     *
     * Copilot issues tool calls in parallel as a matter of course, so this is
     * reachable rather than theoretical. The lock is per document, keyed
     * case-insensitively because Windows paths are.
     */
    #withLock(docPath, fn) {
        const key = path.resolve(docPath).toLowerCase();
        const previous = this.#locks.get(key) ?? Promise.resolve();
        const run = previous.then(fn, fn);
        // The stored tail exists only to order the next caller, so it must not
        // reject -- an unhandled rejection here would be this operation's
        // failure resurfacing inside an unrelated later one.
        const tail = run.then(
            () => {},
            () => {},
        );
        this.#locks.set(key, tail);
        // Drop the entry once this is the last operation queued, so a long-lived
        // process does not accumulate one promise per document ever edited.
        tail.then(() => {
            if (this.#locks.get(key) === tail) this.#locks.delete(key);
        });
        return run;
    }

    async #requireFile(docPath) {
        let info;
        try {
            info = await stat(docPath);
        } catch {
            throw new EditError("file_not_found", `No such file: ${docPath}`);
        }
        if (!info.isFile()) throw new EditError("file_not_found", `Not a file: ${docPath}`);
        return info;
    }

    /**
     * Applies one intent and returns the document as it now stands.
     *
     * `revisionToken` is required, not optional. An edit without one is an edit
     * against a document the caller has not looked at, and the whole point of
     * the token is that "I read this, then changed it" is checkable.
     */
    async edit(docPath, rawIntent, { revisionToken } = {}) {
        return this.#withLock(docPath, () => this.#edit(docPath, rawIntent, { revisionToken }));
    }

    async #edit(docPath, rawIntent, { revisionToken } = {}) {
        const intent = validateIntent(rawIntent);
        await this.#requireFile(docPath);

        // One wall clock for the whole operation, not one per layer.
        //
        // An edit is three Word round trips -- read, edit, read -- and each of
        // them defaults to its own budget. Left alone those compose
        // multiplicatively: two 240 s reads either side of a 180 s host edit is
        // eleven minutes of silence from what the caller issued as a single
        // call. Nothing distinguishes that from a hang, and the agent on the
        // other end has no way to know it should stop waiting.
        //
        // So the budget is spent, not restarted. Each step gets what remains,
        // and a step that would start with nothing left fails immediately with
        // the reason rather than blocking on a Word that is not coming back.
        const deadline = Date.now() + EDIT_BUDGET_MS;
        const remaining = (label) => {
            const left = deadline - Date.now();
            if (left <= 0) {
                throw new EditError(
                    "word_timeout",
                    `Editing ${path.basename(docPath)} exceeded its ${Math.round(EDIT_BUDGET_MS / 1000)}s budget ` +
                        `before ${label} could start.`,
                );
            }
            return left;
        };

        // Cheapest possible refusal, before Word is involved at all.
        const current = await fileRevisionToken(docPath);
        if (!tokensMatch(revisionToken, current)) {
            throw new EditError(
                "stale_revision_token",
                `${path.basename(docPath)} has changed since you read it. Read it again and retry with the new token.`,
                { expectedRevisionToken: revisionToken ?? null, actualRevisionToken: current },
            );
        }

        // Full read, never paged: the address has to be resolvable wherever it
        // is in the document, and paging would hide paragraphs from resolution
        // while leaving their addresses perfectly valid. `limit: 0` is stated
        // rather than defaulted -- the tool boundary caps reads at a few hundred
        // paragraphs, and if that cap ever migrates into the reader, a silent
        // truncation here would surface as an unresolvable address on long
        // documents only.
        const before = await this.#reader.read(docPath, { limit: 0, deadline });

        if (!tokensMatch(before.revisionToken, current)) {
            throw new EditError(
                "stale_revision_token",
                `${path.basename(docPath)} changed while it was being read. Read it again and retry.`,
                { expectedRevisionToken: current, actualRevisionToken: before.revisionToken },
            );
        }

        // Layer 1 reports `writable` and leaves it at that, because a read never
        // touches the original. An edit does, so here it is a precondition.
        // It stays a *pre-flight* check rather than the only guard: detection
        // and open are not atomic, so the host re-checks and its open is
        // timeout-bounded regardless.
        if (before.writable === false) {
            throw new EditError(
                "file_locked",
                `${path.basename(docPath)} could not be opened for writing. Another program may be holding it, or it ` +
                    `may be protected against writing.`,
            );
        }

        const target = before.paragraphs.find((p) => p.address === intent.address);
        if (!target) {
            throw new EditError(
                "address_not_found",
                `No paragraph with address ${intent.address} in ${path.basename(docPath)}. Addresses come from read_document and are invalidated by an edit — read again and use a current one.`,
                { paragraphCount: before.paragraphCount },
            );
        }

        // On disk, before anything is touched, and one per operation. Word's own
        // undo cannot serve here: ADR 0005 closes the document at the end of the
        // operation and the in-process undo history dies with it.
        const snapshot = await takeSnapshot({
            root: this.#snapshotRoot,
            docPath,
            op: intent.op,
            token: current,
            description: describeIntent(intent),
            nonce: Math.random().toString(36).slice(2, 8),
        });

        let result;
        try {
            result = await this.#host.edit({
                path: docPath,
                // Word's own paragraph numbering, not the map's. They diverge:
                // Word counts an extra paragraph for each table row's
                // end-of-row mark, and does not count text-box paragraphs at
                // all. `wordIndex` is computed for that; `index` is not.
                wordIndex: target.wordIndex,
                expectedText: target.text,
                op: intent.op,
                text: intent.text ?? null,
                headingLevel: intent.headingLevel ?? null,
                timeoutMs: remaining("the edit itself"),
            });
        } catch (err) {
            // The host *threw*, so the outcome is unknown -- which is precisely
            // when the snapshot is the only copy of the pre-edit bytes. The two
            // other discard sites below are safe because the host affirmatively
            // reported that it did not touch the file; this one is their
            // opposite and must not be treated the same way.
            //
            // The reachable sequence is one this codebase already knows is
            // real: the request times out, the PowerShell child is killed, and
            // the reap runs `taskkill /F /T` while Word may be mid-Save on the
            // user's original. Discarding here would leave a half-written
            // document and a `revert_document` that answers `no_snapshot`.
            await this.#recoverFromUnknownOutcome({ docPath, snapshot, before: current, cause: err });
            throw err;
        }

        if (result.status !== "edited") {
            // Nothing changed, so the snapshot is a duplicate of the current
            // state. Keeping it would consume a revert step that undoes nothing.
            await this.#discardSnapshot(snapshot);
            failFromStatus(result, docPath, intent);
        }

        const afterToken = await fileRevisionToken(docPath);
        if (tokensMatch(afterToken, current)) {
            await this.#discardSnapshot(snapshot);
            throw new EditError(
                "edit_not_persisted",
                `Word reported the edit as applied but ${path.basename(docPath)} is byte-for-byte unchanged. Nothing was modified.`,
            );
        }

        // The verification the agent can actually see. It also re-mints every
        // address, which the caller needs: the ones it holds are now stale.
        //
        // Deliberately complete, where `read_document` pages. After an edit the
        // caller holds nothing it can still use, so handing back the first page
        // of a long document would leave it unable to address the rest without
        // another round trip -- and unable to see its own edit if that edit was
        // past the cap. Completeness is the point of this read.
        //
        // By this line the edit is applied and persisted -- proven, not assumed,
        // by the token comparison above. So a failure here is a failure to
        // *describe* the document, not to change it, and it must not be
        // reported as though the edit had not happened: an agent told "the edit
        // failed" would retry and apply it twice. The snapshot is deliberately
        // kept, because the edit it undoes is real.
        let after;
        try {
            after = await this.#reader.read(docPath, { limit: 0, deadline });
        } catch (err) {
            throw new EditError(
                "document_unreadable",
                `The edit was applied to ${path.basename(docPath)} and saved, but reading the document back afterwards ` +
                    `failed (${err.message}). Do not repeat the edit — it is already in the file. Read the document ` +
                    `again to get current addresses.`,
                { editApplied: true, snapshot: snapshot.name, cause: err.code ?? null },
            );
        }
        const touched =
            result.wordIndex > 0 ? (after.paragraphs.find((p) => p.wordIndex === result.wordIndex) ?? null) : null;

        // The host polls until the file is writable again rather than trusting
        // `Document.Close()` to have released it, and reports the answer.
        // Measured, it releases in 0-1 ms every time -- unlike `Application
        // .Quit()`, which returns seconds before the process actually exits.
        // So `released: false` means something genuinely unusual, and it is
        // worth saying so here: the next operation on this document will fail
        // its writable pre-flight, and without this line that would look like a
        // fresh and unexplained lock rather than a consequence of this edit.
        const lockReleased = result.released !== false;
        if (!lockReleased) {
            this.#log(
                `edit_document: ${path.basename(docPath)} was still held ${result.releaseMs}ms after the document was ` +
                    `closed. The edit is saved; a follow-up edit may be refused until the handle goes.`,
            );
        }

        this.#log(
            `edit_document: ${describeIntent(intent)} in ${path.basename(docPath)} ` +
                `(open ${result.openMs}ms, edit ${result.editMs}ms, save ${result.saveMs}ms, ` +
                `release ${result.releaseMs}ms, total ${result.totalMs}ms)`,
        );

        return {
            applied: { ...intent, description: describeIntent(intent) },
            paragraph: touched,
            page: result.page || null,
            protectedView: Boolean(result.protectedView),
            markOfTheWeb: Boolean(result.markOfTheWeb),
            lockReleased,
            previousRevisionToken: current,
            snapshot: { name: snapshot.name, takenAt: snapshot.takenAt },
            timings: {
                openMs: result.openMs ?? null,
                editMs: result.editMs ?? null,
                saveMs: result.saveMs ?? null,
                releaseMs: result.releaseMs ?? null,
                lockHeldMs: result.totalMs ?? null,
            },
            document: after,
        };
    }

    async #discardSnapshot(snapshot) {
        await rm(snapshot.file, { force: true }).catch(() => {});
        await rm(`${snapshot.file}.json`, { force: true }).catch(() => {});
    }

    /**
     * Called when the host threw, so we do not know whether the document was
     * written. Re-hashes to find out, and makes the outcome one of two states
     * the caller can reason about rather than an unknown one.
     *
     * Untouched -> the snapshot is a duplicate of what is on disk, so discard it
     * exactly as the affirmative "nothing changed" paths do, and let the
     * original error propagate unchanged.
     *
     * Changed -> the document was modified by an operation that then failed, so
     * it may be half-written. Restore the snapshot. If the restore also fails,
     * keep the snapshot and say its name in the error: the recovery point must
     * survive even when we cannot apply it, because it is the only copy.
     */
    async #recoverFromUnknownOutcome({ docPath, snapshot, before, cause }) {
        // Facts go on `data` as well as on the error itself. Only `code`,
        // `message` and `data` survive `asToolError` at the tool boundary, so a
        // top-level property here reaches this module's own tests and nothing
        // else -- which is exactly how this was wrong the first time: the
        // snapshot was correctly kept and its name never left the process.
        const record = (fields, note) => {
            Object.assign(cause, fields);
            cause.data = { ...(cause.data ?? {}), ...fields };
            // The message is what the model actually reads. Leaving it as the
            // bare host error would tell the agent its call timed out while
            // saying nothing about the state of the user's document.
            if (note) cause.message = `${cause.message} ${note}`;
        };

        const name = path.basename(docPath);

        let afterToken = null;
        try {
            afterToken = await fileRevisionToken(docPath);
        } catch {
            // Cannot even hash it -- keep the snapshot. Losing a recovery point
            // because we could not read the file is the worst available choice.
            record(
                { snapshot: snapshot.name, rolledBack: false },
                `${name} could not be re-read afterwards, so whether it was modified is unknown. A snapshot of the ` +
                    `state before this edit is kept as ${snapshot.name}; revert_document restores it.`,
            );
            return;
        }

        if (tokensMatch(afterToken, before)) {
            await this.#discardSnapshot(snapshot);
            record(
                { rolledBack: false, documentUnchanged: true },
                `${name} is byte-for-byte unchanged, so nothing was written and no recovery is needed.`,
            );
            return;
        }

        try {
            await revertToLatest({ root: this.#snapshotRoot, docPath });
            record(
                { rolledBack: true },
                `${name} had been modified before the failure and has been rolled back to its state before this edit.`,
            );
            this.#log(
                `edit_document: ${name} was modified by a failed operation and has been rolled back ` +
                    `from ${snapshot.name}`,
            );
        } catch (restoreErr) {
            record(
                { rolledBack: false, snapshot: snapshot.name },
                `${name} had been modified before the failure and could NOT be rolled back automatically ` +
                    `(${restoreErr.message}). It may be partially written. A snapshot of the state before this edit ` +
                    `is kept as ${snapshot.name}; revert_document restores it.`,
            );
            this.#log(
                `edit_document: ${name} was modified by a failed operation and could NOT be rolled ` +
                    `back (${restoreErr.message}); snapshot ${snapshot.name} kept`,
            );
        }
    }

    /**
     * Restores the newest snapshot and discards it, so repeated calls walk back
     * through the history rather than toggling between the last two states.
     *
     * The state that is reverted away is not kept. A "redo" snapshot would make
     * the next revert restore what was just undone, and an undo that alternates
     * is worse than one that only goes backwards.
     *
     * `revisionToken` is required, for the same reason `edit` requires it and
     * with more at stake. A revert overwrites the document with older bytes and
     * keeps no snapshot of what it destroyed, so without the check the sequence
     * "agent edits, user works in Word for twenty minutes, agent reverts" ends
     * with the user's work gone and nothing to recover it from. `edit_document`
     * refuses exactly that situation; it would be incoherent for the
     * destructive operation to be the lenient one. The agent always has a
     * token, because `edit_document` returns one.
     */
    async revert(docPath, { revisionToken } = {}) {
        return this.#withLock(docPath, () => this.#revert(docPath, { revisionToken }));
    }

    async #revert(docPath, { revisionToken } = {}) {
        await this.#requireFile(docPath);

        const current = await fileRevisionToken(docPath);
        if (!tokensMatch(revisionToken, current)) {
            throw new EditError(
                "stale_revision_token",
                `${path.basename(docPath)} has changed since you read it. Read it again and retry with the new token. ` +
                    `A revert overwrites the document and keeps no copy of what it replaces, so it will not run against ` +
                    `bytes you have not seen.`,
                { expectedRevisionToken: revisionToken ?? null, actualRevisionToken: current },
            );
        }

        const history = await listSnapshots({ root: this.#snapshotRoot, docPath });
        if (history.length === 0) {
            throw new EditError(
                "no_snapshot",
                `There is no snapshot of ${path.basename(docPath)} to revert to. Snapshots are taken by edit_document, one per edit.`,
            );
        }

        // A revert writes the original, so it needs the same exclusive access an
        // edit does. Word is not involved, so this is a plain write-handle probe
        // rather than a host round trip.
        if (!(await isWritable(docPath))) {
            throw new EditError(
                "file_locked",
                `${path.basename(docPath)} could not be opened for writing, so it cannot be reverted. Another program ` +
                    `may be holding it, or it may be protected against writing.`,
            );
        }

        const { restored, remaining } = await revertToLatest({ root: this.#snapshotRoot, docPath });
        const after = await this.#reader.read(docPath, { limit: 0 });

        this.#log(`revert_document: restored ${restored.name} over ${path.basename(docPath)}`);

        return {
            restored: {
                name: restored.name,
                op: restored.op ?? null,
                description: restored.description ?? null,
                takenAt: restored.takenAt ?? null,
                revisionToken: restored.revisionToken ?? null,
            },
            revertedFromRevisionToken: current,
            snapshotsRemaining: remaining,
            document: after,
        };
    }

    history(docPath) {
        return listSnapshots({ root: this.#snapshotRoot, docPath });
    }
}

/**
 * Whether the file can be opened for writing right now.
 *
 * `'r+'` on Windows fails while another process holds the file without sharing
 * write access, which is exactly the condition that matters. It is a probe, not
 * a guarantee: the file can be taken in the moment between this returning and
 * the write starting, which is why nothing downstream relies on it alone.
 */
async function isWritable(filePath) {
    let handle;
    try {
        handle = await open(filePath, "r+");
        return true;
    } catch {
        return false;
    } finally {
        await handle?.close().catch(() => {});
    }
}
