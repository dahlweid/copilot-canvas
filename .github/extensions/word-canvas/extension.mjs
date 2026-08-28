// word-canvas — renders .docx files page-accurately in a side canvas.
//
// A hidden, automation-owned Microsoft Word instance is the rendering engine:
// it exports the document to PDF, which the canvas serves over loopback to the
// panel's native PDF viewer. Word also stays available as a live document model
// for outline, search and text queries.
//
// Read-only by design: the user's file is never opened by Word directly, only
// an unblocked temp copy of it.

import path from "node:path";
import { createCanvas, CanvasError, joinSession } from "@github/copilot-sdk/extension";
import { RenderCache, DocumentError, normalizeDocPath } from "./src/render-cache.mjs";
import { ViewerInstance } from "./src/server.mjs";
import { artifactsRoot } from "./src/store.mjs";

/** instanceId -> ViewerInstance */
const instances = new Map();

let session = null;
let cache = null;

// `session.log` rejects asynchronously (e.g. on an unsupported level); an
// unhandled rejection would take the whole extension process down.
const log = (message, level = "info", { ephemeral = level === "info" } = {}) => {
    void session?.log(`[word-canvas] ${message}`, { level, ephemeral })?.catch(() => {});
};

function getCache() {
    if (!cache) {
        cache = new RenderCache({ cacheRoot: artifactsRoot(), log: (m) => log(m) });
    }
    return cache;
}

/** Maps our typed errors onto canvas error codes the agent can act on. */
function asCanvasError(err) {
    if (err instanceof CanvasError) return err;
    if (err instanceof DocumentError) return new CanvasError(err.code, err.message);
    if (err?.code === "word_unavailable") {
        return new CanvasError(
            "word_unavailable",
            `${err.message} This canvas requires Microsoft Word to be installed on this machine.`,
        );
    }
    return new CanvasError("word_canvas_error", err?.message ?? "Unknown error");
}

function requireInstance(instanceId) {
    const instance = instances.get(instanceId);
    if (!instance) {
        throw new CanvasError("canvas_not_open", `No open Word canvas with instance id '${instanceId}'.`);
    }
    return instance;
}

function requireDocument(instance) {
    if (!instance.doc) {
        throw new CanvasError(
            "no_document",
            "No document is open in this canvas. Call open_document with a .docx path first.",
        );
    }
    return instance.doc;
}

/** Resolves a user- or agent-supplied path against the workspace. */
function resolveInputPath(input) {
    if (path.isAbsolute(input) || !session?.workspacePath) return normalizeDocPath(input);
    return normalizeDocPath(path.join(session.workspacePath, input));
}

async function run(fn) {
    try {
        return await fn();
    } catch (err) {
        throw asCanvasError(err);
    }
}

// --- canvas ----------------------------------------------------------------

const wordCanvas = createCanvas({
    id: "word-doc",
    displayName: "Word document",
    description:
        "Displays a Word (.docx) document with page-accurate rendering, plus outline, search and text extraction. Read-only.",
    inputSchema: {
        type: "object",
        properties: {
            path: {
                type: "string",
                description:
                    "Absolute or workspace-relative path to a Word document (.docx, .docm, .doc, .rtf). Omit to open the canvas on its document picker.",
            },
        },
        additionalProperties: false,
    },

    open: async (ctx) =>
        run(async () => {
            let instance = instances.get(ctx.instanceId);
            if (!instance) {
                instance = new ViewerInstance({
                    cache: getCache(),
                    instanceId: ctx.instanceId,
                    workspacePath: session?.workspacePath,
                    log: (m) => log(m),
                });
                await instance.start();
                instances.set(ctx.instanceId, instance);
            }

            // Re-opening the same panel (rehydrate, reload, focus) must not
            // disturb what is already shown, so only act on a *different* path.
            if (ctx.input?.path) {
                const wanted = resolveInputPath(ctx.input.path);
                if (instance.doc?.path !== wanted) await instance.openDocument(wanted);
            }

            return {
                url: instance.url,
                title: instance.doc ? instance.doc.name : "Word document",
                status: instance.doc
                    ? `${instance.doc.pageCount} ${instance.doc.pageCount === 1 ? "page" : "pages"}`
                    : "No document open",
            };
        }),

    onClose: async (ctx) => {
        const instance = instances.get(ctx.instanceId);
        if (!instance) return;
        instances.delete(ctx.instanceId);
        const docPath = instance.doc?.path ?? null;
        await instance.close();

        if (instances.size === 0) {
            // Nothing left to render for: let Word go rather than leaving a
            // hidden process running for the rest of the session.
            log("last canvas closed, shutting Word down");
            await cache?.dispose().catch(() => {});
            cache = null;
        } else if (docPath && ![...instances.values()].some((i) => i.doc?.path === docPath)) {
            await cache?.close(docPath).catch(() => {});
        }
    },

    actions: [
        {
            name: "open_document",
            description: "Opens a Word document in this canvas, replacing whatever is currently shown.",
            inputSchema: {
                type: "object",
                properties: {
                    path: {
                        type: "string",
                        description: "Absolute or workspace-relative path to the document.",
                    },
                },
                required: ["path"],
                additionalProperties: false,
            },
            handler: async (ctx) =>
                run(async () => {
                    const instance = requireInstance(ctx.instanceId);
                    const doc = await instance.openDocument(resolveInputPath(ctx.input.path));
                    return { opened: true, ...doc };
                }),
        },
        {
            name: "go_to_page",
            description: "Scrolls the canvas to a page of the open document.",
            inputSchema: {
                type: "object",
                properties: { page: { type: "integer", minimum: 1 } },
                required: ["page"],
                additionalProperties: false,
            },
            handler: async (ctx) =>
                run(async () => {
                    const instance = requireInstance(ctx.instanceId);
                    const doc = requireDocument(instance);
                    if (ctx.input.page > doc.pageCount) {
                        throw new CanvasError(
                            "page_out_of_range",
                            `Page ${ctx.input.page} is past the end; the document has ${doc.pageCount} pages.`,
                        );
                    }
                    return { page: instance.goToPage(ctx.input.page), pageCount: doc.pageCount };
                }),
        },
        {
            name: "get_outline",
            description: "Returns the heading outline of the open document, with page numbers.",
            inputSchema: {
                type: "object",
                properties: { limit: { type: "integer", minimum: 1, maximum: 2000 } },
                additionalProperties: false,
            },
            handler: async (ctx) =>
                run(async () => {
                    const doc = requireDocument(requireInstance(ctx.instanceId));
                    return getCache().outline(doc.path, { limit: ctx.input?.limit });
                }),
        },
        {
            name: "search",
            description:
                "Searches the open document and returns matches with page numbers and surrounding text.",
            inputSchema: {
                type: "object",
                properties: {
                    query: { type: "string", minLength: 1 },
                    limit: { type: "integer", minimum: 1, maximum: 500 },
                    matchCase: { type: "boolean" },
                    wholeWord: { type: "boolean" },
                },
                required: ["query"],
                additionalProperties: false,
            },
            handler: async (ctx) =>
                run(async () => {
                    const doc = requireDocument(requireInstance(ctx.instanceId));
                    return getCache().search(doc.path, ctx.input.query, {
                        limit: ctx.input.limit ?? 100,
                        matchCase: ctx.input.matchCase === true,
                        wholeWord: ctx.input.wholeWord === true,
                    });
                }),
        },
        {
            name: "get_text",
            description:
                "Extracts the text of the open document, optionally limited to a page range. Use this to read the document rather than inferring it from the outline.",
            inputSchema: {
                type: "object",
                properties: {
                    fromPage: { type: "integer", minimum: 1 },
                    toPage: { type: "integer", minimum: 1 },
                },
                additionalProperties: false,
            },
            handler: async (ctx) =>
                run(async () => {
                    const doc = requireDocument(requireInstance(ctx.instanceId));
                    return getCache().text(doc.path, {
                        fromPage: ctx.input?.fromPage,
                        toPage: ctx.input?.toPage,
                    });
                }),
        },
        {
            name: "refresh",
            description:
                "Re-reads the document from disk and re-renders it. Use after a script or tool has rewritten the file.",
            inputSchema: {
                type: "object",
                properties: { force: { type: "boolean" } },
                additionalProperties: false,
            },
            handler: async (ctx) =>
                run(async () => {
                    const instance = requireInstance(ctx.instanceId);
                    requireDocument(instance);
                    return instance.refresh({ force: ctx.input?.force === true });
                }),
        },
        {
            name: "get_info",
            description: "Returns metadata about the document currently shown in this canvas.",
            handler: async (ctx) =>
                run(async () => {
                    const instance = requireInstance(ctx.instanceId);
                    if (!instance.doc) return { open: false, wordVersion: cache?.wordVersion ?? null };
                    const info = await getCache().info(instance.doc.path);
                    return {
                        open: true,
                        path: instance.doc.path,
                        currentPage: instance.lastPage,
                        wordVersion: cache?.wordVersion ?? null,
                        ...info,
                    };
                }),
        },
    ],
});

// --- lifecycle -------------------------------------------------------------

let shuttingDown = false;

function reapOnExit() {
    // Synchronous last resort: async teardown does not get to run on exit, and a
    // hidden Word left behind is invisible to the user and never cleaned up.
    try {
        cache?.reap();
    } catch {
        /* nothing more can be done at this point */
    }
}

async function shutdown(signal) {
    if (shuttingDown) return;
    shuttingDown = true;
    try {
        await Promise.all([...instances.values()].map((i) => i.close().catch(() => {})));
        instances.clear();
        await cache?.dispose().catch(() => {});
    } finally {
        reapOnExit();
        if (signal) process.exit(0);
    }
}

process.on("exit", reapOnExit);
process.on("SIGTERM", () => void shutdown("SIGTERM"));
process.on("SIGINT", () => void shutdown("SIGINT"));

// stderr goes to the extension log file; a stray rejection should show up there
// rather than killing the provider and taking every open canvas with it.
process.on("unhandledRejection", (reason) => {
    process.stderr.write(`[word-canvas] unhandled rejection: ${reason?.stack ?? reason}\n`);
});

session = await joinSession({
    canvases: [wordCanvas],
    hooks: {
        onSessionEnd: async () => {
            await shutdown(null);
        },
    },
});
