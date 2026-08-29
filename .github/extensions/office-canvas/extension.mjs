// office-canvas — Word document tools, plus a canvas that renders .docx files
// page-accurately in a side panel.
//
// The tools are the product; the canvas is how the user watches. `read_document`
// is callable with no canvas open.
//
// A hidden, automation-owned Microsoft Word instance does the work: it exports
// the document to PDF, which the canvas serves over loopback to the panel's
// native PDF viewer, and it hands back WordprocessingML for structural reads.
//
// Reads are read-only by design: the user's file is never opened by Word for a
// read, only an unblocked temp copy of it. Edits are not, and cannot be -- an
// edit must touch the original, so it opens it directly for one operation and
// closes it again (ADR 0005), with an on-disk snapshot taken first.
//
// This header said the file is *never* opened directly, full stop. That was
// true until `edit_document` shipped in this same file, and it is the kind of
// line a reader trusts precisely because it sits at the top and reads like a
// guarantee.

import path from "node:path";
import { createCanvas, CanvasError, joinSession } from "@github/copilot-sdk/extension";
import { RenderCache, DocumentError, normalizeDocPath, supportedList } from "./src/render-cache.mjs";
import { ViewerInstance } from "./src/server.mjs";
import { artifactsRoot } from "./src/store.mjs";
import { createIdleShutdown } from "./src/word-lifecycle.mjs";
import { normalizeReadArgs, DEFAULT_READ_LIMIT, MAX_READ_LIMIT } from "./src/word/read-args.mjs";
import { MAX_HEADING_LEVEL, MIN_HEADING_LEVEL, OPERATION_HELP, OPERATION_NAMES } from "./src/word/edit-intent.mjs";
import { asToolError } from "./src/tool-error.mjs";
import {
    BLOCK_HELP,
    BLOCK_KINDS,
    MAX_BLOCKS,
    MAX_LIST_ITEMS,
    MAX_TABLE_COLUMNS,
    MAX_TABLE_ROWS,
    MIN_BLOCK_HEADING_LEVEL,
} from "./src/word/create-intent.mjs";
import { creatableList } from "./src/word/document-author.mjs";

/** instanceId -> ViewerInstance */
const instances = new Map();

let session = null;
let cache = null;

// A tool call can arrive with no canvas open, so Word may be started purely to
// serve one. Nothing would then ever shut it down -- the canvas path releases
// Word when the last panel closes, and there is no panel. An idle timer gives
// repeated reads a warm instance without leaving a hidden Word running for the
// rest of the session, and an in-flight counter keeps it from firing under a
// tool call that is still running. See src/word-lifecycle.mjs.
const IDLE_SHUTDOWN_MS = 60_000;

// `session.log` rejects asynchronously (e.g. on an unsupported level); an
// unhandled rejection would take the whole extension process down.
const log = (message, level = "info", { ephemeral = level === "info" } = {}) => {
    void session?.log(`[office-canvas] ${message}`, { level, ephemeral })?.catch(() => {});
};

const lifecycle = createIdleShutdown({
    idleMs: IDLE_SHUTDOWN_MS,
    isDisplaying: () => instances.size > 0,
    dispose: async () => {
        const idle = cache;
        cache = null;
        await idle?.dispose();
    },
    log: (m) => log(m),
});

/**
 * Wraps anything that needs Word alive for its duration.
 *
 * Every tool handler goes through this rather than remembering to schedule an
 * idle shutdown in its own `finally` -- that was the original shape, and it is
 * the shape that carried the race.
 */
const withWordWork = (fn) => lifecycle.run(fn);

function getCache() {
    lifecycle.cancel();
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

/**
 * Tools are not canvas actions, so `CanvasError` means nothing to a tool caller.
 * The code is folded into the message instead, because that is what the agent
 * actually reads when deciding what to do next.
 *
 * Lives in `./src/tool-error.mjs` so it can be tested directly: this module is
 * not importable from a test, and the defects this boundary causes are exactly
 * the ones invisible from one layer beneath it.
 */

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
                    `Absolute or workspace-relative path to a Word document (${supportedList()}). ` +
                    `Omit to open the canvas on its document picker.`,
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
            // hidden process running for the rest of the session -- unless a
            // tool call is still using it, in which case closing the panel must
            // not tear down the instance under it. The idle timer picks it up.
            if (lifecycle.busy) {
                log("last canvas closed while a tool call is running; deferring Word shutdown");
                lifecycle.schedule();
                return;
            }
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

// --- tools -----------------------------------------------------------------

// A document with no limit returns every paragraph -- text, style id, heading
// path and a 12-hex address each -- straight into the agent's context on the
// first call. Paging is free here because addresses are minted across the whole
// document regardless, and the response says how many were withheld.
//
// The default and the validation that protects it live in src/word/read-args.mjs,
// so the description below and the handler cannot state different numbers.

const readDocumentTool = {
    name: "read_document",
    description: [
        "Reads a Word document and returns its structure map — every paragraph with its text,",
        "resolved style, heading path and a stable address — together with a revision token for",
        "the file. Cite an address to say which paragraph an edit applies to, and present the",
        "token so the edit can be refused if the document changed underneath. Does not need a",
        "canvas to be open, and does not open one.",
    ].join(" "),
    parameters: {
        type: "object",
        properties: {
            path: {
                type: "string",
                description: `Absolute or workspace-relative path to a Word document (${supportedList()}).`,
            },
            limit: {
                type: "integer",
                minimum: 1,
                // Declared from the same constant the validator enforces: a
                // declared bound the runtime does not check is a promise the
                // contract cannot keep.
                maximum: MAX_READ_LIMIT,
                description:
                    `Maximum paragraphs to return (default ${DEFAULT_READ_LIMIT}). Addresses are always minted across the whole document, so paging never changes one; the response reports paragraphCount and truncated.`,
            },
            offset: {
                type: "integer",
                minimum: 0,
                description: "Index of the first paragraph to return, 0-based. Use with limit to page a long document.",
            },
        },
        required: ["path"],
        additionalProperties: false,
    },
    handler: async (args) => {
        // Validated before any Word work is entered. This is placement rather
        // than a rescue: nothing on this path started Word for a bad argument
        // before either, because `resolveInputPath` already threw ahead of
        // `getCache()`. Keeping the checks together and ahead of `withWordWork`
        // is what stops that from being an accident of evaluation order as more
        // validation arrives -- L2's edit path will have considerably more.
        let paging;
        let docPath;
        try {
            paging = normalizeReadArgs(args);
            docPath = resolveInputPath(args?.path);
        } catch (err) {
            throw asToolError(err);
        }

        return withWordWork(async () => {
            try {
                return await getCache().readStructure(docPath, paging);
            } catch (err) {
                throw asToolError(err);
            }
        });
    },
};

// Every bound below is imported from src/word/create-intent.mjs, which is also
// what the runtime enforces. A description that restates a limit is a second
// copy of it, and this repo has shipped the drifted version of that three times
// in three pull requests — a range hardcoded in prose, the constant moved, and
// the model told something false with no test able to notice.

const createBlockSchema = {
    type: "object",
    properties: {
        kind: {
            type: "string",
            enum: BLOCK_KINDS,
            description: `What this block is. ${BLOCK_HELP}`,
        },
        level: {
            type: "integer",
            minimum: MIN_BLOCK_HEADING_LEVEL,
            maximum: MAX_HEADING_LEVEL,
            description: `heading only: the heading depth, ${MIN_BLOCK_HEADING_LEVEL} being the top level.`,
        },
        text: {
            type: "string",
            description:
                "heading and paragraph: the text. One paragraph — line breaks are refused, because a second paragraph is a second block.",
        },
        ordered: {
            type: "boolean",
            description: "list only: true numbers the items; omitted or false bullets them.",
        },
        items: {
            type: "array",
            items: { type: "string" },
            maxItems: MAX_LIST_ITEMS,
            description: `list only: the items, one paragraph each. At most ${MAX_LIST_ITEMS}.`,
        },
        rows: {
            type: "array",
            items: { type: "array", items: { type: "string" }, maxItems: MAX_TABLE_COLUMNS },
            maxItems: MAX_TABLE_ROWS,
            description: `table only: the cells, row by row. The table must be rectangular — every row the same length — and at most ${MAX_TABLE_ROWS} rows by ${MAX_TABLE_COLUMNS} columns.`,
        },
        headerRow: {
            type: "boolean",
            description:
                "table only: true makes the first row a bold heading row that repeats if the table crosses a page break.",
        },
    },
    required: ["kind"],
    additionalProperties: false,
};

const createDocumentTool = {
    name: "create_document",
    description: [
        "Creates a new Word document from an ordered list of blocks — headings, paragraphs, bulleted or",
        "numbered lists and tables — and returns its structure map and revision token, the same ones",
        "read_document would give, so it can be edited straight away without reading it first.",
        "It will not overwrite: a path that already exists is refused, because replacing a document is",
        "what edit_document is for and that path has a revision token, a snapshot and a revert behind it.",
        "Text is written verbatim — Word's autocorrect is switched off on the instance that authors it,",
        "so straight quotes stay straight and nothing is capitalised or substituted on the way in.",
    ].join(" "),
    parameters: {
        type: "object",
        properties: {
            path: {
                type: "string",
                description: `Absolute or workspace-relative path to create (${creatableList()}). The folder must already exist, and the file must not.`,
            },
            blocks: {
                type: "array",
                minItems: 1,
                maxItems: MAX_BLOCKS,
                items: createBlockSchema,
                description: `The document's content, in order. At most ${MAX_BLOCKS} blocks. ${BLOCK_HELP}`,
            },
        },
        required: ["path", "blocks"],
        additionalProperties: false,
    },
    handler: async (args) => {
        // Path resolution runs before any Word work is entered, matching
        // read_document: a malformed path should cost a string operation, not a
        // cold Word start.
        let docPath;
        try {
            docPath = resolveInputPath(args?.path);
        } catch (err) {
            throw asToolError(err);
        }

        return withWordWork(async () => {
            try {
                return await getCache().createDocument(docPath, { blocks: args?.blocks });
            } catch (err) {
                throw asToolError(err);
            }
        });
    },
};

const editDocumentTool = {
    name: "edit_document",
    description: [
        "Applies one change to a Word document, in place, and returns the document as it stands",
        "afterwards. Requires the address of the paragraph to change and the revision token from",
        "read_document; the edit is refused if the file changed since that read. A snapshot is taken",
        "first, so revert_document can undo it.",
        "One edit per call, and the addresses you hold are invalidated by it: deleting one of several",
        "identically-worded paragraphs renumbers the others, and renaming a heading moves every address",
        "beneath it. The result carries a fresh map and token — use those for the next edit.",
    ].join(" "),
    parameters: {
        type: "object",
        properties: {
            path: {
                type: "string",
                description: `Absolute or workspace-relative path to a Word document (${supportedList()}).`,
            },
            revisionToken: {
                type: "string",
                description:
                    "The revisionToken from the read_document call that produced the address. The edit is refused if the file has changed since.",
            },
            op: {
                type: "string",
                enum: OPERATION_NAMES,
                description: `What to do to the paragraph. ${OPERATION_HELP}`,
            },
            address: {
                type: "string",
                description: "Address of the paragraph to change, from read_document (looks like 'p:0123456789ab').",
            },
            text: {
                type: "string",
                description:
                    "The new text, for replace_text and the insert operations. One paragraph: line breaks are refused, because a second paragraph would move the addresses after it.",
            },
            headingLevel: {
                type: "integer",
                minimum: MIN_HEADING_LEVEL,
                maximum: MAX_HEADING_LEVEL,
                description: `Heading level: ${MIN_HEADING_LEVEL + 1}–${MAX_HEADING_LEVEL} for a heading, ${MIN_HEADING_LEVEL} for body text. Required by set_heading_level; optional on an insert, which otherwise follows the style Word would use itself.`,
            },
        },
        required: ["path", "revisionToken", "op", "address"],
        additionalProperties: false,
    },
    handler: async (args) =>
        withWordWork(async () => {
            try {
                const intent = { op: args?.op, address: args?.address };
                if (args?.text !== undefined) intent.text = args.text;
                if (args?.headingLevel !== undefined) intent.headingLevel = args.headingLevel;

                const result = await getCache().editDocument(resolveInputPath(args?.path), intent, {
                    revisionToken: args?.revisionToken,
                });
                await refreshCanvasesFor(result.document.path);
                return result;
            } catch (err) {
                throw asToolError(err);
            }
        }),
};

const revertDocumentTool = {
    name: "revert_document",
    description: [
        "Undoes the most recent edit_document change to a Word document by restoring the snapshot taken",
        "before it, and returns the document as it stands afterwards. Repeated calls step further back",
        "through the edit history; each restored snapshot is consumed, so this only moves backwards.",
        "Word's own undo cannot be used — the document is closed at the end of every edit and the undo",
        "history goes with it.",
        "Requires the revisionToken you last saw for the document, so a revert cannot silently discard",
        "changes made by someone else since.",
    ].join(" "),
    parameters: {
        type: "object",
        properties: {
            path: {
                type: "string",
                description: "Absolute or workspace-relative path to the Word document to revert.",
            },
            revisionToken: {
                type: "string",
                description:
                    "Required. The revisionToken from the read_document or edit_document call whose result you are " +
                    "reverting. The revert is refused unless the file still matches it. This is not optional: a revert " +
                    "overwrites the document with older bytes and keeps no copy of what it replaced, so running it " +
                    "against a document someone has changed since you looked destroys that work permanently.",
            },
        },
        required: ["path", "revisionToken"],
        additionalProperties: false,
    },
    handler: async (args) =>
        withWordWork(async () => {
            try {
                const result = await getCache().revertDocument(resolveInputPath(args?.path), {
                    revisionToken: args?.revisionToken,
                });
                await refreshCanvasesFor(result.document.path);
                return result;
            } catch (err) {
                throw asToolError(err);
            }
        }),
};

/**
 * Re-renders any open canvas showing a document we have just changed.
 *
 * Best-effort by design: the edit has already been made and verified by a
 * re-read, so a canvas that fails to refresh is a stale picture, not a failed
 * edit, and must not turn a successful edit into an error.
 */
async function refreshCanvasesFor(docPath) {
    const target = identityFor(docPath);
    for (const instance of instances.values()) {
        if (!instance.doc || identityFor(instance.doc.path) !== target) continue;
        try {
            await instance.refresh({ force: true });
        } catch (err) {
            log(`could not refresh the canvas after an edit: ${err?.message ?? err}`, "warning");
        }
    }
}

const identityFor = (p) => (process.platform === "win32" ? String(p).toLowerCase() : String(p));

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
    lifecycle.cancel();
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
    process.stderr.write(`[office-canvas] unhandled rejection: ${reason?.stack ?? reason}\n`);
});

session = await joinSession({
    canvases: [wordCanvas],
    tools: [createDocumentTool, readDocumentTool, editDocumentTool, revertDocumentTool],
    hooks: {
        onSessionEnd: async () => {
            await shutdown(null);
        },
    },
});
