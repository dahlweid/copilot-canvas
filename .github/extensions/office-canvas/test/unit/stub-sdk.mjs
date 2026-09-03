// Enough of `@github/copilot-sdk/extension` to import the real extension.mjs.
//
// The SDK is auto-resolved by the runtime and is not present on disk here --
// `import("@github/copilot-sdk/extension")` fails with ERR_MODULE_NOT_FOUND --
// so a test that wants to drive the extension's own wiring has to supply it.
// This is deliberately a stand-in for the transport only: every canvas, tool
// and lifecycle handler under test is the committed one.
//
// Installed with `module.registerHooks`, which is synchronous and in-thread, so
// no extra flag is needed on the `node --test` command line the gate runs.

export class CanvasError extends Error {
    constructor(code, message) {
        super(message);
        this.name = "CanvasError";
        this.code = code;
    }
}

export class DocumentError extends Error {
    constructor(code, message) {
        super(message);
        this.name = "DocumentError";
        this.code = code;
    }
}

/** What the extension handed to `joinSession`, for a test to drive. */
export const joined = { canvases: [], tools: [], hooks: {} };

/** The SDK returns a canvas handle; the definition itself is what we need. */
export const createCanvas = (definition) => definition;

export const logged = [];

/**
 * The session object the extension captured from `joinSession`, so a test can
 * drive it after the extension has imported.
 *
 * The extension reads two distinct things off this object, live, on every
 * `resolveInputPath` call, and a test controls each independently:
 *
 *   * `session.workspacePath` -- still read by the canvas `open` action to seed
 *     the document picker's scan root (out of #158's scope). `setWorkspacePath`
 *     drives it; the default is `null`.
 *   * `session.rpc.metadata.snapshot()` -- the async host call whose
 *     `workingDirectory` a relative path now resolves against (#158).
 *     `setWorkingDirectory` sets what it returns; `failMetadataSnapshot` makes it
 *     reject, which is how the no-fallback refusal path is exercised. The default
 *     `workingDirectory` is `null`, i.e. the field-guard case (a snapshot that
 *     answers without a usable absolute working directory).
 */
let currentSession = null;

// Snapshot behaviour, module-level so setters reach it before joinSession has
// run and so it survives across the single import-time joinSession call.
const metadata = { workingDirectory: null, snapshotError: null };

/** Sets `session.workspacePath` on the object the extension holds. */
export function setWorkspacePath(workspacePath) {
    if (currentSession) currentSession.workspacePath = workspacePath;
}

/** What `session.rpc.metadata.snapshot()` reports as the working directory. */
export function setWorkingDirectory(workingDirectory) {
    metadata.workingDirectory = workingDirectory;
    metadata.snapshotError = null;
}

/** Makes `session.rpc.metadata.snapshot()` reject, to drive the no-fallback path. */
export function failMetadataSnapshot(error = new Error("metadata snapshot unavailable")) {
    metadata.snapshotError = error;
}

/** Restores the default snapshot state (no working directory, no error). */
export function resetMetadata() {
    metadata.workingDirectory = null;
    metadata.snapshotError = null;
}

export async function joinSession(options = {}) {
    joined.canvases = options.canvases ?? [];
    joined.tools = options.tools ?? [];
    joined.hooks = options.hooks ?? {};
    currentSession = {
        workspacePath: null,
        rpc: {
            metadata: {
                async snapshot() {
                    if (metadata.snapshotError) throw metadata.snapshotError;
                    return {
                        workspacePath: currentSession.workspacePath,
                        workingDirectory: metadata.workingDirectory,
                    };
                },
            },
        },
        async log(message) {
            logged.push(message);
        },
    };
    return currentSession;
}
