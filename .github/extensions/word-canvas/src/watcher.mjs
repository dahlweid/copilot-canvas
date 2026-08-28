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

    /**
     * @param {object} options
     * @param {(info: {path: string}) => void} options.onChange fired once the file is stable
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
        this.lastSeen = await this.#fingerprint();

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

            if (previous === null) return; // deleted or still unreadable
            if (previous === this.lastSeen) return; // touched, but identical
            this.lastSeen = previous;
            this.onChange({ path: this.filePath });
        } catch (err) {
            this.log(`watch: settle failed: ${err.message}`);
        } finally {
            this.#settling = false;
            if (this.#pendingWhileSettling) {
                this.#pendingWhileSettling = false;
                this.#schedule();
            }
        }
    }

    /** Records the current fingerprint as already-seen, so our own reads don't loop. */
    async acknowledge() {
        this.lastSeen = await this.#fingerprint();
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
