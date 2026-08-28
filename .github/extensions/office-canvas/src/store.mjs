// Durable, user-scoped storage for the canvas.
//
// Recents and rendered PDFs follow the user across sessions and projects, so
// they live under $COPILOT_HOME (not in the repo, and not keyed by instance id).

import { mkdir, readFile, writeFile } from "node:fs/promises";
import { homedir } from "node:os";
import path from "node:path";

const EXTENSION_NAME = "office-canvas";
const MAX_RECENTS = 15;

export function copilotHome() {
    return process.env.COPILOT_HOME?.trim() || path.join(homedir(), ".copilot");
}

export function artifactsRoot() {
    return path.join(copilotHome(), "extensions", EXTENSION_NAME, "artifacts");
}

const recentsFile = () => path.join(artifactsRoot(), "recents.json");

export async function readRecents() {
    try {
        const raw = await readFile(recentsFile(), "utf8");
        const parsed = JSON.parse(raw);
        return Array.isArray(parsed) ? parsed.filter((e) => e && typeof e.path === "string") : [];
    } catch {
        return [];
    }
}

/** Most-recent-first, de-duplicated by path, capped. */
export async function addRecent({ path: docPath, name, pageCount }) {
    const entries = await readRecents();
    const key = process.platform === "win32" ? docPath.toLowerCase() : docPath;
    const next = [
        { path: docPath, name, pageCount: pageCount ?? 0, openedAt: new Date().toISOString() },
        ...entries.filter((e) => (process.platform === "win32" ? e.path.toLowerCase() : e.path) !== key),
    ].slice(0, MAX_RECENTS);
    try {
        await mkdir(artifactsRoot(), { recursive: true });
        await writeFile(recentsFile(), JSON.stringify(next, null, 2), "utf8");
    } catch {
        /* a missing recents list is not worth failing an open over */
    }
    return next;
}
