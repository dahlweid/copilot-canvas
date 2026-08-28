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
            throw new EditError(
                "document_locked",
                `${name} is open in another program, so it cannot be edited. Close it there and try again.`,
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
        default:
            throw new EditError("edit_failed", `Word reported '${result.status}' while editing ${name}.`);
    }
}

export class DocumentEditor {
    #reader;
    #host;
    #snapshotRoot;
    #log;

    constructor({ reader, host, snapshotRoot, log = () => {} }) {
        this.#reader = reader;
        this.#host = host;
        this.#snapshotRoot = snapshotRoot;
        this.#log = log;
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
        const intent = validateIntent(rawIntent);
        await this.#requireFile(docPath);

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
        // while leaving their addresses perfectly valid.
        const before = await this.#reader.read(docPath);

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
                "document_locked",
                `${path.basename(docPath)} is open in another program, so it cannot be edited. Close it there and try again.`,
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
            });
        } catch (err) {
            await this.#discardSnapshot(snapshot);
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
        const after = await this.#reader.read(docPath);
        const touched =
            result.wordIndex > 0 ? (after.paragraphs.find((p) => p.wordIndex === result.wordIndex) ?? null) : null;

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
     * Restores the newest snapshot and discards it, so repeated calls walk back
     * through the history rather than toggling between the last two states.
     *
     * The state that is reverted away is not kept. A "redo" snapshot would make
     * the next revert restore what was just undone, and an undo that alternates
     * is worse than one that only goes backwards.
     */
    async revert(docPath, { revisionToken = null } = {}) {
        await this.#requireFile(docPath);

        const current = await fileRevisionToken(docPath);
        if (revisionToken !== null && !tokensMatch(revisionToken, current)) {
            throw new EditError(
                "stale_revision_token",
                `${path.basename(docPath)} has changed since you read it. Read it again and retry with the new token.`,
                { expectedRevisionToken: revisionToken, actualRevisionToken: current },
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
                "document_locked",
                `${path.basename(docPath)} is open in another program, so it cannot be reverted. Close it there and try again.`,
            );
        }

        const { restored, remaining } = await revertToLatest({ root: this.#snapshotRoot, docPath });
        const after = await this.#reader.read(docPath);

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
