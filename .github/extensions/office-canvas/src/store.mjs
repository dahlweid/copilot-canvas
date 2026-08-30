// Durable, user-scoped storage for the canvas.
//
// Rendered PDFs follow the user across sessions and projects, so they live under
// $COPILOT_HOME (not in the repo, and not keyed by instance id).
//
// A recents list lived here too, until #69. It existed to feed the picker, and
// the picker went because a click-to-open list is user-driven document
// selection -- the one job this panel gives to Copilot. Nothing else ever read
// it, so it left with its only reader rather than becoming a file written for
// no one.

import { homedir } from "node:os";
import path from "node:path";

const EXTENSION_NAME = "office-canvas";

export function copilotHome() {
    return process.env.COPILOT_HOME?.trim() || path.join(homedir(), ".copilot");
}

export function artifactsRoot() {
    return path.join(copilotHome(), "extensions", EXTENSION_NAME, "artifacts");
}
