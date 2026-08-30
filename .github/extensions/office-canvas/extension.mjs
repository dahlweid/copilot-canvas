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
import { createRenderCacheSlot } from "./src/render-cache-slot.mjs";
import { normalizeReadArgs, DEFAULT_READ_LIMIT, MAX_READ_LIMIT } from "./src/word/read-args.mjs";
import { MAX_HEADING_LEVEL, MIN_HEADING_LEVEL, OPERATION_HELP, OPERATION_NAMES } from "./src/word/edit-intent.mjs";
import { toolFailure } from "./src/tool-error.mjs";
import { changeRecordFrom } from "./src/change-record.mjs";
import {
    BLOCK_HELP,
    BLOCK_KINDS,
    fieldUsage,
    MAX_BLOCKS,
    MAX_LIST_ITEMS,
    MAX_TABLE_COLUMNS,
    MAX_TABLE_ROWS,
    MIN_BLOCK_HEADING_LEVEL,
    MIN_BLOCKS,
    MIN_LIST_ITEMS,
    MIN_TABLE_COLUMNS,
    MIN_TABLE_ROWS,
    MAX_TEXT_LENGTH,
} from "./src/word/create-intent.mjs";
import { creatableList } from "./src/word/document-author.mjs";

/** instanceId -> ViewerInstance */
const instances = new Map();

let session = null;

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

// The one live RenderCache, and therefore the one hidden Word. Held in a slot
// rather than a bare `let` because the slot is what makes a disposal invisible
// to the next caller: it empties synchronously and awaits the teardown after,
// so nothing is ever handed a cache that is already being disposed. That was
// #61 -- the whole Word surface stayed down after one shutdown.
const caches = createRenderCacheSlot({
    create: () => new RenderCache({ cacheRoot: artifactsRoot(), log: (m) => log(m) }),
    log: (m) => log(m),
});

const lifecycle = createIdleShutdown({
    idleMs: IDLE_SHUTDOWN_MS,
    isDisplaying: () => instances.size > 0,
    dispose: () => caches.dispose(),
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
    return caches.get();
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

/**
 * The canvas-action half of #45, which is a *different* channel from the tool
 * half and fails differently.
 *
 * Measured in the same session as the tool arms (spikes/tool-errors/): a canvas
 * action that throws reaches the agent as `Canvas operation failed: Error:
 * <message>`. The message survives -- so unlike a tool, throwing from a canvas
 * action is not a dead channel and this path is not rewritten. But the `code`
 * does not survive: a `CanvasError("MARK-D-CODE", "MARK-D-MESSAGE")` arrived
 * carrying `MARK-D-MESSAGE` and no trace of `MARK-D-CODE`.
 *
 * So the code is folded into the message here, for the same reason and by the
 * same rule as the tool surface: the only field that reaches the agent is the
 * one it has to be in. `CanvasError.code` is still set, because the field is
 * the SDK's contract and the extension's own tests and callers read it -- it is
 * simply not something the agent can see.
 */
const withCodeInMessage = (err) => new CanvasError(err.code, `${err.code}: ${err.message}`);

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

const describePathArg = (value) => (typeof value === "string" ? "an empty string" : `a ${typeof value}`);

/**
 * Resolves a user- or agent-supplied path against the workspace.
 *
 * The guard is the whole point of this function's shape (#38). Every tool
 * declares `required: ["path"]`, and measured in #28 this host enforces no
 * schema keyword before dispatch -- not `required`, not `type` -- so the
 * handler is the only defence there is. Without the guard the first thing a
 * missing `path` met was `path.isAbsolute(undefined)`, and Node refused the
 * call with its own `ERR_INVALID_ARG_TYPE`: measured, `ERR_INVALID_ARG_TYPE:
 * The "path" argument must be of type string. Received undefined`, from all
 * four tools. That is a refusal nobody chose, authored by Node about its own
 * argument contract, so it cannot name the tool's parameter or say what to
 * send instead.
 *
 * Empty and whitespace-only strings are refused here too, rather than left to
 * `normalizeDocPath`. They never reached it: with a workspace set,
 * `path.join(workspacePath, "")` is the workspace directory -- measured, and
 * the same for `"   "` once `normalizeDocPath` trims -- so a blank `path`
 * resolved to a *directory* and was only ever caught downstream, by a code
 * naming something else. The canvas is unaffected: `open` tests
 * `ctx.input?.path` for truthiness before calling this, so an omitted path
 * still opens the canvas on its empty state.
 *
 * `invalid_path` rather than a new code: it is already the extension-surface
 * code for an unusable path, and it is what `normalizeDocPath` answers an
 * empty string with today, so every spelling of "no usable path" now gives one
 * answer.
 */
function resolveInputPath(input) {
    if (input === undefined || input === null) {
        throw new DocumentError(
            "invalid_path",
            "`path` is required: give an absolute path to the document, or one relative to the workspace.",
        );
    }
    if (typeof input !== "string" || input.trim() === "") {
        throw new DocumentError(
            "invalid_path",
            `\`path\` must be a non-empty string — an absolute path to the document, or one relative to the ` +
                `workspace. Received ${describePathArg(input)}.`,
        );
    }
    if (path.isAbsolute(input) || !session?.workspacePath) return normalizeDocPath(input);
    return normalizeDocPath(path.join(session.workspacePath, input));
}

async function run(fn) {
    try {
        return await fn();
    } catch (err) {
        throw withCodeInMessage(asCanvasError(err));
    }
}

/**
 * Tool failures travel as *results*, never as exceptions.
 *
 * Measured: a throw from a tool handler reaches the agent as the string
 * `Tool execution failed`, with the message, the code and the data bag all
 * discarded en route — see `src/tool-error.mjs` for the two hops and
 * `spikes/tool-errors/` for the run. A returned `ToolResultObject` arrives
 * intact.
 *
 * Applied here, at the registration site, rather than inside each handler. Both
 * work today; only this one keeps working. A tool added later inherits the
 * legible channel by being registered at all, which is the difference between a
 * property and a convention — and #45's predecessor was exactly a convention
 * that one call site kept and the rest did not.
 */
const reportingFailures = (tool) => ({
    ...tool,
    handler: async (args) => {
        try {
            return await tool.handler(args);
        } catch (err) {
            return toolFailure(tool.name, err);
        }
    },
});

// --- canvas ----------------------------------------------------------------

/*
 * No `icon` here, and its absence is a measurement rather than an oversight
 * (#68). The Word mark reached the document name and the Open in Word button;
 * the canvas *tab* was the third placement asked for and it is not reachable
 * from a Node extension.
 *
 * `spikes/word-icon/probes/probe-icon-sources.mjs` measures why, on the SDK
 * bundled with CLI 1.0.80: `Canvas` builds `this.declaration` from an explicit
 * five-field literal -- `id, displayName, description, inputSchema, actions` --
 * with no spread, so an `icon` option is dropped before the declaration exists,
 * whatever its value. The string `icon` appears nowhere in the SDK's runtime
 * JavaScript; the only occurrences are in the generated wire typings, where
 * `DiscoveredCanvas.icon` is documented as a host-local PNG path that nothing on
 * this side populates.
 *
 * The one route that might work -- committing a PNG for the host to find -- is
 * the one the issue forbids. Re-run arm 3 of that probe against a newer SDK
 * before concluding this is still closed.
 */

const wordCanvas = createCanvas({
    id: "word-doc",
    displayName: "Word document",
    // "Read-only" used to close this sentence, and it was the one string every
    // agent saw before deciding what to do with a .docx. It described the panel
    // and was read as the extension: the product that ships create_document,
    // read_document, edit_document and revert_document announced itself as
    // incapable of editing. The surface really is display-only (ADR 0002), so
    // the fix is to say who edits rather than to drop the fact.
    description: [
        "Displays a Word (.docx) document with page-accurate rendering, plus outline, search and text extraction.",
        "The document is read and changed by the agent through this extension's tools — read_document,",
        "create_document, edit_document and revert_document — which drive the installed Word, and the canvas",
        "re-renders after an edit and states what changed: the operation, the page when one was reported, and a",
        "highlight on the changed text where there is text left to highlight. The panel itself is display only —",
        "the user does not type into it, they ask Copilot for the change.",
    ].join(" "),
    inputSchema: {
        type: "object",
        properties: {
            path: {
                type: "string",
                description:
                    `Absolute or workspace-relative path to a Word document (${supportedList()}). ` +
                    `Omit to open the canvas empty; it then tells the user to ask for a document, ` +
                    `and this canvas's open_document action puts one in it.`,
            },
        },
        additionalProperties: false,
    },

    open: async (ctx) =>
        run(async () => {
            let instance = instances.get(ctx.instanceId);
            if (!instance) {
                instance = new ViewerInstance({
                    // A resolver, not a reference: a panel outlives Word being
                    // shut down, and a captured cache is how a canvas ended up
                    // holding a host that had been disposed (#61).
                    cache: () => caches.get(),
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
            await caches.dispose();
        } else if (docPath && ![...instances.values()].some((i) => i.doc?.path === docPath)) {
            await caches.peek()?.close(docPath).catch(() => {});
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
                    if (!instance.doc) return { open: false, wordVersion: caches.peek()?.wordVersion ?? null };
                    const info = await getCache().info(instance.doc.path);
                    return {
                        open: true,
                        path: instance.doc.path,
                        currentPage: instance.lastPage,
                        wordVersion: caches.peek()?.wordVersion ?? null,
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
        "Use this whenever a Word document has to be read — to answer a question about it, summarize, quote or",
        "review it, or find the paragraph to change — and always before edit_document, which cannot run without",
        "the address and the revision token this call mints. Do not convert the file to another format first,",
        "and do not parse its OOXML or reach for a document library: this opens the document in the installed",
        "Word and reports what the file itself stores — each paragraph's style id and name, its heading level and",
        "the heading path it sits under, whether it is inside a table, and its list id and level. A conversion to",
        "text or Markdown flattens most of that away, and none of it gives you an address to edit against. Word",
        "is only ever handed a copy, never the original, so a document already open in Word can still be read.",
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
        //
        // Nothing is caught here. `reportingFailures` catches at the
        // registration site, so both the validation throws and anything the
        // Word work raises become the same legible failure result; a catch here
        // as well would only wrap the code into the message twice.
        const paging = normalizeReadArgs(args);
        const docPath = resolveInputPath(args?.path);

        return withWordWork(() => getCache().readStructure(docPath, paging));
    },
};

// Every bound below is imported from src/word/create-intent.mjs, which is also
// what the runtime enforces. A description that restates a limit is a second
// copy of it, and this repo has shipped the drifted version of that three times
// in three pull requests — a range hardcoded in prose, the constant moved, and
// the model told something false with no test able to notice.
//
// The schema declares every bound `validateSpec` enforces except one: the rule
// that a paragraph may not contain a line or paragraph break. That is
// `/[\r\n\v\f\u0007]/` in `requireText`, and JSON Schema could carry it as a
// `pattern`. It deliberately does not. A `pattern` subtly wrong in the other
// direction would reject legal text at the schema layer, which is worse than
// the gap it closes, and a negative character-class is not something a model
// reliably satisfies by construction. Each affected description states the rule
// in prose instead, and the runtime refusal names it. Recorded as a decision so
// the next audit finds a reason here rather than an oversight.

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
            description: `${fieldUsage("level")} The heading depth, ${MIN_BLOCK_HEADING_LEVEL} being the top level.`,
        },
        text: {
            type: "string",
            maxLength: MAX_TEXT_LENGTH,
            description: `${fieldUsage("text")} The text. One paragraph — line breaks are refused, because a second paragraph is a second block.`,
        },
        ordered: {
            type: "boolean",
            description: `${fieldUsage("ordered")} True numbers the items; omitted or false bullets them.`,
        },
        items: {
            type: "array",
            items: { type: "string", maxLength: MAX_TEXT_LENGTH },
            minItems: MIN_LIST_ITEMS,
            maxItems: MAX_LIST_ITEMS,
            description: `${fieldUsage("items")} The items, one paragraph each. From ${MIN_LIST_ITEMS} to ${MAX_LIST_ITEMS}.`,
        },
        rows: {
            type: "array",
            items: {
                type: "array",
                items: { type: "string", maxLength: MAX_TEXT_LENGTH },
                minItems: MIN_TABLE_COLUMNS,
                maxItems: MAX_TABLE_COLUMNS,
            },
            minItems: MIN_TABLE_ROWS,
            maxItems: MAX_TABLE_ROWS,
            description: `${fieldUsage("rows")} The cells, row by row. The table must be rectangular — every row the same length — and from ${MIN_TABLE_ROWS}×${MIN_TABLE_COLUMNS} to ${MAX_TABLE_ROWS}×${MAX_TABLE_COLUMNS}.`,
        },
        headerRow: {
            type: "boolean",
            description: `${fieldUsage("headerRow")} True makes the first row a bold heading row that repeats if the table crosses a page break.`,
        },
    },
    // Only `kind` is unconditionally required; what each kind additionally
    // requires is derived into the field descriptions above by `fieldUsage`,
    // because JSON Schema `required` cannot be made conditional on `kind` here.
    // `edit_document` carries the same limitation the same way.
    required: ["kind"],
    additionalProperties: false,
};

const createDocumentTool = {
    name: "create_document",
    description: [
        "Use this when a new Word document is wanted — a report, a letter, minutes, notes — rather than writing",
        "OOXML, zipping a package by hand, or using a document library. Word itself authors the file, so its",
        "headings, list numbering and tables are Word's own rather than markup approximating them, and no part",
        "of the package has to be modelled by hand.",
        "Creates a new Word document from an ordered list of blocks — headings, paragraphs, bulleted or",
        "numbered lists and tables — and returns its structure map and revision token, the same ones",
        "read_document would give, so it can be edited straight away without reading it first.",
        "It will not overwrite: a path that already exists is refused, because replacing a document is",
        "what edit_document is for and that path has a revision token, a snapshot and a revert behind it.",
        "Text is written verbatim. Autocorrect is switched off first on a Word this tool started, so",
        "straight quotes stay straight and nothing is capitalised or substituted on the way in; if it",
        "attached to a Word you already had running, that instance's settings are left alone instead,",
        "because they are yours. The result's `autoCorrect` field reports which of the two happened.",
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
                minItems: MIN_BLOCKS,
                maxItems: MAX_BLOCKS,
                items: createBlockSchema,
                description: `The document's content, in order. From ${MIN_BLOCKS} to ${MAX_BLOCKS} blocks. ${BLOCK_HELP}`,
            },
        },
        required: ["path", "blocks"],
        additionalProperties: false,
    },
    handler: async (args) => {
        // Path resolution runs before any Word work is entered, matching
        // read_document: a malformed path should cost a string operation, not a
        // cold Word start. Uncaught here; see `reportingFailures`.
        const docPath = resolveInputPath(args?.path);

        // Everything except `path` is forwarded, rather than `{ blocks }` being
        // picked out.
        //
        // This looks like the looser choice and is the stricter one. `validateSpec`
        // already refuses a spec field nobody implements -- that refusal exists
        // because `title` was once accepted here and silently ignored by the host,
        // which is the same defect as a half-applied block. Picking `blocks` out
        // put that refusal out of reach: an unknown argument never reached the
        // validator, so `additionalProperties: false` above was decorative and the
        // caller was told nothing.
        //
        // It is decorative in a second way too, measured by the coordinator on
        // this CLI host (issue #28): the host validates `parameters` before
        // dispatch not at all -- `required`, `enum`, `type`, `minItems`,
        // `maximum` and `additionalProperties` were each violated and every one
        // reached the handler. So the schema is documentation for the model and
        // `validateSpec` is the only enforcement there is. Anything declared up
        // there and not checked in code is a promise nothing keeps.
        const { path: _path, ...spec } = args ?? {};

        return withWordWork(() => getCache().createDocument(docPath, spec));
    },
};

const editDocumentTool = {
    name: "edit_document",
    description: [
        "Use this to change an existing Word document — reword a paragraph, insert one, delete one, or make a",
        "paragraph a heading or body text — including when the user asks for the change in prose. Read the",
        "document with read_document first, then edit it here. Do not rewrite the file with a document library",
        "or hand-edited OOXML: Word applies the change to the document itself, so whatever the edit does not",
        "touch is left to Word rather than to someone's model of the file format, and the revision token below",
        "means an edit built on a stale read is refused rather than applied over someone else's change.",
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
            const intent = { op: args?.op, address: args?.address };
            if (args?.text !== undefined) intent.text = args.text;
            if (args?.headingLevel !== undefined) intent.headingLevel = args.headingLevel;

            const result = await getCache().editDocument(resolveInputPath(args?.path), intent, {
                revisionToken: args?.revisionToken,
            });
            await refreshCanvasesFor(result.document.path, changeRecordFrom(result));
            return result;
        }),
};

const revertDocumentTool = {
    name: "revert_document",
    description: [
        "Use this when an edit turned out wrong or the user asks to undo it or put the document back. Prefer it",
        "to editing the old text back in: the snapshot is a byte copy of the file taken before that edit and it",
        "is put back whole, by rename rather than by rewriting the document, so what returns is the document",
        "exactly as it was — whereas a compensating edit is a further edit and restores only the text it retypes.",
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
            const result = await getCache().revertDocument(resolveInputPath(args?.path), {
                revisionToken: args?.revisionToken,
            });
            // A revert undoes the edit the overlay was describing, so the
            // record is cleared rather than replaced. The restored text is
            // not a change the agent made -- highlighting it would tell the
            // user the opposite of what happened.
            await refreshCanvasesFor(result.document.path, null);
            return result;
        }),
};

/**
 * Re-renders any open canvas showing a document we have just changed, and tells
 * it what changed so the overlay can mark it.
 *
 * Best-effort by design: the edit has already been made and verified by a
 * re-read, so a canvas that fails to refresh is a stale picture, not a failed
 * edit, and must not turn a successful edit into an error.
 *
 * `change` is passed explicitly as `null` by operations that invalidate the
 * previous overlay without producing a new one -- a revert undoes the very edit
 * the last record described, so leaving that record up would point the user at
 * a change that no longer exists.
 */
async function refreshCanvasesFor(docPath, change) {
    const target = identityFor(docPath);
    for (const instance of instances.values()) {
        if (!instance.doc || identityFor(instance.doc.path) !== target) continue;
        try {
            await instance.refresh({ force: true, change });
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
        caches.peek()?.reap();
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
        await caches.dispose();
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
    tools: [createDocumentTool, readDocumentTool, editDocumentTool, revertDocumentTool].map(reportingFailures),
    hooks: {
        onSessionEnd: async () => {
            await shutdown(null);
        },
    },
});
