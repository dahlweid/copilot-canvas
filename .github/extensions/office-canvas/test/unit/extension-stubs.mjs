// Makes the committed `extension.mjs` importable from a test, and nothing more.
//
// `extension.mjs` imports `@github/copilot-sdk/extension`, which the runtime
// auto-resolves and which is not on disk here, so importing it raises
// ERR_MODULE_NOT_FOUND. `module.registerHooks` supplies a stand-in for that
// transport and swaps `RenderCache` for one that never starts Word. Everything
// else under test -- every canvas handler, every tool handler, the slot and the
// lifecycle -- is the committed code.
//
// This lives in its own module because two test files need it and a copied test
// helper is a second thing to keep true. Importing it is enough: the hooks are
// installed by its module body, `registerHooks` is synchronous and in-thread
// (so `node --test` needs no extra flag), and both callers reach `extension.mjs`
// through a dynamic import that runs later.

import { registerHooks } from "node:module";

const SDK_SPECIFIER = "@github/copilot-sdk/extension";

export const stubSdkUrl = new URL("./stub-sdk.mjs", import.meta.url).href;
export const stubCacheUrl = new URL("./stub-render-cache.mjs", import.meta.url).href;

registerHooks({
    resolve(specifier, context, nextResolve) {
        if (specifier === SDK_SPECIFIER) return { url: stubSdkUrl, shortCircuit: true };
        const resolved = nextResolve(specifier, context);
        // Only extension.mjs's own import is swapped. src/server.mjs imports the
        // same module for `DocumentError` and `normalizeDocPath`, and the stub
        // re-exports them from the real one -- redirecting those too would be a
        // cycle.
        if (
            resolved.url.endsWith("/src/render-cache.mjs") &&
            String(context.parentURL ?? "").endsWith("/extension.mjs")
        ) {
            return { ...resolved, url: stubCacheUrl, shortCircuit: true };
        }
        return resolved;
    },
});

/** Imports the committed extension against the stubs and returns the stub SDK. */
export async function loadExtension() {
    const sdk = await import(stubSdkUrl);
    await import("../../extension.mjs");
    return sdk;
}
