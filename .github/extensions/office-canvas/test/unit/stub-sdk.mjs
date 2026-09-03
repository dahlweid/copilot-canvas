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
 * drive `session.workspacePath` after the extension has imported. The extension
 * reads `session?.workspacePath` live on every `resolveInputPath` call, so
 * mutating this object's property is what a workspace-relative resolution test
 * needs -- the default stays `null`, which is the no-workspace case (#158).
 */
let currentSession = null;

/** Sets `session.workspacePath` on the object the extension holds. */
export function setWorkspacePath(workspacePath) {
    if (currentSession) currentSession.workspacePath = workspacePath;
}

export async function joinSession(options = {}) {
    joined.canvases = options.canvases ?? [];
    joined.tools = options.tools ?? [];
    joined.hooks = options.hooks ?? {};
    currentSession = {
        workspacePath: null,
        async log(message) {
            logged.push(message);
        },
    };
    return currentSession;
}
