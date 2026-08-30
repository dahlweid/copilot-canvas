// Watches one document for changes.
//
// Watching the containing directory rather than the file itself is deliberate:
// scripts and editors usually *replace* a document (write a temp file, then
// rename over the original), which silently detaches a file-level watch.
//
// A regenerating script also writes in several steps, so a raw event is not a
// safe signal to re-render on. We wait until size and mtime hold steady before
// reporting a change; otherwise Word would open a half-written file.

import { watch } from "node:fs";
import { stat } from "node:fs/promises";
import path from "node:path";

const SETTLE_POLL_MS = 200;
const SETTLE_STABLE_ROUNDS = 2;
const SETTLE_TIMEOUT_MS = 30_000;

export class FileWatcher {
    #watcher = null;
    #timer = null;
    #settling = false;
    #pendingWhileSettling = false;
    // Bumped on every write to `lastSeen`. A delivery that fails needs to know
    // whether the field is still the one it set, and it cannot ask the value:
    // `acknowledge()` records the fingerprint of the file as it stands, which
    // during a delivery is the very fingerprint being delivered. Comparing
    // values would call an acknowledgement "untouched" and roll back over it --
    // which is the echo loop `acknowledge` exists to prevent.
    #seenGeneration = 0;

    /**
     * @param {object} options
     * @param {(info: {path: string}) => void | Promise<void>} options.onChange fired once the
     *   file is stable. **Awaited, and a rejection means the change was not consumed**: the
     *   fingerprint is rolled back so the next directory event reports it again. Returning a
     *   promise also serializes delivery -- no second change is dispatched while one is in
     *   flight -- which is what keeps a slow consumer from being handed overlapping work.
     * @param {(message: string) => void} [options.log]
     * @param {number} [options.debounceMs]
     */
    constructor({ onChange, log = () => {}, debounceMs = 300 }) {
        this.onChange = onChange;
        this.log = log;
        this.debounceMs = debounceMs;
        this.filePath = null;
        this.lastSeen = null;
    }

    async watch(filePath) {
        if (this.filePath === filePath && this.#watcher) return;
        this.close();
        this.filePath = filePath;
        this.#markSeen(await this.#fingerprint());

        const dir = path.dirname(filePath);
        const base = path.basename(filePath).toLowerCase();
        try {
            this.#watcher = watch(dir, { persistent: false }, (_event, filename) => {
                if (!filename) return this.#schedule();
                const name = filename.toString().toLowerCase();
                // `~$name.docx` is Word's own owner file; it churns constantly and
                // never means the document itself changed.
                if (name.startsWith("~$")) return;
                if (name === base) this.#schedule();
            });
            this.#watcher.on("error", (err) => this.log(`watch: ${err.message}`));
        } catch (err) {
            this.log(`watch: could not watch ${dir}: ${err.message}`);
        }
    }

    async #fingerprint() {
        try {
            const info = await stat(this.filePath);
            return `${info.mtimeMs}|${info.size}`;
        } catch {
            return null;
        }
    }

    #schedule() {
        if (this.#settling) {
            // The writer is still going; re-settle once the current pass ends.
            this.#pendingWhileSettling = true;
            return;
        }
        clearTimeout(this.#timer);
        this.#timer = setTimeout(() => this.#settle(), this.debounceMs);
    }

    async #settle() {
        this.#settling = true;
        // The file this pass is about. `watch()` can be called for another
        // document while this one is still settling -- and now also while its
        // change is being consumed, which is a much longer window -- and a
        // fingerprint of document B must never be recorded as A's, nor
        // dispatched under A's path.
        const filePath = this.filePath;
        try {
            const deadline = Date.now() + SETTLE_TIMEOUT_MS;
            let previous = await this.#fingerprint();
            let stableRounds = 0;

            while (Date.now() < deadline) {
                await new Promise((r) => setTimeout(r, SETTLE_POLL_MS));
                const current = await this.#fingerprint();
                if (current !== null && current === previous) {
                    if (++stableRounds >= SETTLE_STABLE_ROUNDS) break;
                } else {
                    stableRounds = 0;
                }
                previous = current;
            }

            if (this.filePath !== filePath) return; // we were re-pointed mid-pass
            if (previous === null) return; // deleted or still unreadable
            if (previous === this.lastSeen) return; // touched, but identical
            await this.#deliver(filePath, previous);
        } catch (err) {
            this.log(`watch: settle failed: ${err.message}`);
        } finally {
            this.#settling = false;
            if (this.#pendingWhileSettling) {
                this.#pendingWhileSettling = false;
                // Only if we are still watching. `close()` can land while a
                // change is being consumed, and re-arming the timer here would
                // resurrect a watcher the owner has already let go of.
                if (this.#watcher) this.#schedule();
            }
        }
    }

    /**
     * Hands one change to the consumer, and records it as seen only if that
     * succeeds.
     *
     * Marking it seen *before* the call is deliberate and is not the bug it
     * looks like: the consumer re-renders, which touches the file, and an echo
     * of that must not be mistaken for a new change. What was wrong was leaving
     * it marked when the consumer *failed*. A refresh that threw -- a lock held
     * for the moment the settle happened to land in -- consumed the edit that
     * triggered it, so the panel had nothing left to re-fire on and stayed in
     * `error` even once the file became readable again. That was #67.
     *
     * The rollback is conditional on the generation because `acknowledge()`
     * shares this field: our own save records the fingerprint it just wrote so
     * the echo of it does not loop. If that ran while this delivery was failing,
     * restoring the older value would re-report our own write -- the loop
     * `acknowledge` exists to prevent.
     */
    async #deliver(filePath, fingerprint) {
        const seenBefore = this.lastSeen;
        const generation = this.#markSeen(fingerprint);
        try {
            await this.onChange({ path: filePath });
        } catch (err) {
            if (this.filePath === filePath && this.#seenGeneration === generation) {
                this.#markSeen(seenBefore);
            }
            this.log(`watch: change was not consumed (${err.message}); it stays pending`);
        }
    }

    /** The one place `lastSeen` is written, so a rollback can tell it was. */
    #markSeen(fingerprint) {
        this.lastSeen = fingerprint;
        this.#seenGeneration += 1;
        return this.#seenGeneration;
    }

    /** Records the current fingerprint as already-seen, so our own reads don't loop. */
    async acknowledge() {
        this.#markSeen(await this.#fingerprint());
    }

    close() {
        clearTimeout(this.#timer);
        this.#timer = null;
        if (this.#watcher) {
            try {
                this.#watcher.close();
            } catch {
                /* already closed */
            }
            this.#watcher = null;
        }
    }
}
