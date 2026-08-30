// Viewer front-end. Talks to the extension over plain HTTP; there is no
// privileged bridge from an extension canvas iframe to the host.

import { PdfView } from "./pdf-view.mjs";
import { describeChangeBanner } from "./change-wording.mjs";
import { monitorConnection } from "./connection-status.mjs";
import { copyOutcome, describeDocument } from "./doc-identity.mjs";
import { showWordMark } from "./word-mark.mjs";

/**
 * What the viewer calls itself.
 *
 * The header bar names the open *document*, so this is the only name the
 * product has while one is open. It is duplicated in `index.html` -- once in
 * `<title>`, so a load with no script still says what this is, and once in the
 * picker's `<h1>` -- and `ui-contract.test.mjs` requires all three to agree.
 */
const PRODUCT = "Word Document Viewer";

const $ = (id) => document.getElementById(id);

const el = {
    bar: $("bar"),
    docName: $("docName"),
    docMeta: $("docMeta"),
    copyPath: $("copyPath"),
    docNameMark: $("docNameMark"),
    status: $("status"),
    sidebar: $("sidebar"),
    outline: $("outline"),
    searchInput: $("searchInput"),
    searchResults: $("searchResults"),
    viewer: $("viewer"),
    pages: $("pages"),
    emptyState: $("emptyState"),
    engineNote: $("engineNote"),
    toggleSidebar: $("toggleSidebar"),
    reload: $("reload"),
    openInWord: $("openInWord"),
    changeBar: $("changeBar"),
    changeText: $("changeText"),
    jumpToChange: $("jumpToChange"),
    dismissChange: $("dismissChange"),
};

let state = null;
let currentPage = 1;
let outlineLoadedFor = null;

/**
 * The rendered document. Reporting the visible page back to the host is how a
 * reopened canvas restores position, and it is now driven by scrolling rather
 * than by an explicit jump, so it is throttled to the page actually changing.
 */
const view = new PdfView(el.pages, {
    onPageChange: (page) => {
        if (page === currentPage) return;
        currentPage = page;
        api("/api/page", { method: "POST", body: JSON.stringify({ page }) }).catch(() => {});
    },
});

/** The URL currently loaded into the view, so a redundant reload is skipped. */
let loadedPdfUrl = null;

/**
 * The change the reader has dismissed, if any.
 *
 * Held as a value derived from what the record *says*, not as the record: the
 * state arrives over SSE, so every frame carries a fresh object and identity
 * comparison would un-dismiss the same change on the next push. Keyed this way,
 * dismissing hides one change and the next one still announces itself.
 */
let dismissedChange = null;

const changeKey = (record) =>
    record ? `${record.at}\u0000${record.op}\u0000${record.page}\u0000${record.text ?? ""}` : null;

/**
 * Puts the current change -- or nothing -- in front of the reader.
 *
 * The banner and the marker are driven together from one place. Splitting them
 * would let the page carry a highlight with nothing explaining it, which is the
 * state a reader cannot act on.
 */
function showChange() {
    const record = state?.change ?? null;
    const dismissed = record !== null && changeKey(record) === dismissedChange;

    if (!record || dismissed) {
        view.setChange(null);
        el.changeBar.hidden = true;
        return;
    }

    view.setChange(record);
    const banner = describeChangeBanner(record);
    el.changeText.textContent = banner.text;
    el.jumpToChange.hidden = !banner.jumpable;
    el.changeBar.hidden = false;
}

// --- helpers ---------------------------------------------------------------

async function api(path, options = {}) {
    const res = await fetch(path, {
        headers: options.body ? { "Content-Type": "application/json" } : undefined,
        ...options,
    });
    const body = await res.json().catch(() => ({}));
    if (!res.ok) {
        const err = new Error(body?.error?.message ?? `Request failed (${res.status})`);
        err.code = body?.error?.code;
        throw err;
    }
    return body;
}

function setStatus(text, { error = false, busy = false } = {}) {
    if (!text) {
        el.status.hidden = true;
        el.status.textContent = "";
        return;
    }
    el.status.hidden = false;
    el.status.classList.toggle("error", error);
    el.status.innerHTML = "";
    if (busy) el.status.append(Object.assign(document.createElement("span"), { className: "spinner" }));
    el.status.append(document.createTextNode(text));
}

/**
 * Loads the current render into the view, restoring the reader's position.
 *
 * The PDF URL carries the render key, so it changes exactly when the document
 * does. Comparing against it is what stops every unrelated state push from
 * tearing down and re-parsing a document that has not moved.
 */
async function showPdf(page = currentPage, { force = false } = {}) {
    if (!state?.pdfUrl) return;
    if (!force && loadedPdfUrl === state.pdfUrl) {
        showChange();
        return;
    }
    loadedPdfUrl = state.pdfUrl;
    try {
        await view.load(state.pdfUrl);
        showChange();
        if (page > 1) view.goToPage(page);
    } catch (err) {
        loadedPdfUrl = null;
        setStatus(`Could not display the document: ${err.message}`, { error: true });
    }
}

function goToPage(page) {
    currentPage = Math.max(1, Math.floor(page));
    view.goToPage(currentPage);
    api("/api/page", { method: "POST", body: JSON.stringify({ page: currentPage }) }).catch(() => {});
}

// --- rendering -------------------------------------------------------------

function render() {
    const hasDoc = Boolean(state?.doc && state.status !== "error");
    // The empty state is the no-document state and nothing else. It offers no
    // control: the canvas is driven by the agent, which opens documents through
    // `open_canvas` and the `open_document` action, so the screen names that
    // rather than asking the user to select a file (#69).
    const showEmpty = !state?.doc;

    el.bar.hidden = !state?.doc;
    el.emptyState.classList.toggle("visible", showEmpty);
    el.viewer.classList.toggle("visible", !showEmpty && hasDoc);
    el.sidebar.hidden = showEmpty || el.sidebar.dataset.open !== "true";

    if (state?.doc) {
        const doc = state.doc;
        // The filename, not the docx `title` property. The bar identifies *this
        // document*, and a metadata title is neither reliably present nor
        // reliably about the file -- `demo.docx` carries "Word Canvas Fixture",
        // which read as the product naming itself after an internal fixture.
        // The title is still worth showing, so it goes to the meta line among
        // the other document properties, where it reads as one.
        el.docName.textContent = doc.name;

        // The absolute path, which used to have a row of its own (#71). Its
        // tooltip repeated its text verbatim, so the tooltip bought nothing and
        // the row cost a line. Now the tooltip is on the name, where it says
        // something the name does not -- and because a `title` is not
        // keyboard-reachable, the same path is also the accessible name of the
        // copy button beside it, which is a tab stop.
        const identity = describeDocument(doc);
        el.docName.title = identity.nameTitle;
        el.copyPath.title = identity.copyTitle;
        el.copyPath.setAttribute("aria-label", identity.copyLabel);
        el.copyPath.disabled = !identity.canCopy;

        const bits = [`${doc.pageCount} ${doc.pageCount === 1 ? "page" : "pages"}`];
        if (doc.wordCount) bits.push(`${doc.wordCount.toLocaleString()} words`);
        if (doc.title?.trim()) bits.push(doc.title.trim());
        if (doc.author?.trim()) bits.push(doc.author);
        el.docMeta.textContent = bits.join(" · ");
        document.title = `${doc.name} — ${PRODUCT}`;
    } else {
        document.title = PRODUCT;
    }

    if (state?.status === "error" && state.error) {
        setStatus(state.error.message, { error: true });
    } else if (state?.status === "opening") {
        setStatus("Opening in Word…", { busy: true });
    } else {
        setStatus(null);
    }

    if (!showEmpty && hasDoc) showPdf(currentPage);
}

function entryButton({ label, page, level, onClick }) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "entry";
    if (level) button.dataset.level = String(level);

    const labelSpan = document.createElement("span");
    labelSpan.className = "label";
    labelSpan.textContent = label;
    button.append(labelSpan);

    if (page) {
        const pageSpan = document.createElement("span");
        pageSpan.className = "page";
        pageSpan.textContent = String(page);
        button.append(pageSpan);
    }
    button.addEventListener("click", onClick);
    return button;
}

async function loadOutline() {
    if (!state?.doc) return;
    if (outlineLoadedFor === state.doc.key) return;
    el.outline.innerHTML = "";
    el.outline.append(
        Object.assign(document.createElement("div"), { className: "panel-title", textContent: "Loading…" }),
    );
    try {
        const { headings = [] } = await api("/api/outline?limit=500");
        outlineLoadedFor = state.doc.key;
        el.outline.innerHTML = "";
        el.outline.append(
            Object.assign(document.createElement("div"), {
                className: "panel-title",
                textContent: headings.length ? "Outline" : "No headings in this document",
            }),
        );
        for (const heading of headings) {
            el.outline.append(
                entryButton({
                    label: heading.text,
                    page: heading.page,
                    level: heading.level,
                    onClick: () => goToPage(heading.page),
                }),
            );
        }
    } catch (err) {
        el.outline.innerHTML = "";
        el.outline.append(
            Object.assign(document.createElement("div"), {
                className: "panel-title",
                textContent: err.message,
            }),
        );
    }
}

let searchTimer = null;
function scheduleSearch() {
    clearTimeout(searchTimer);
    searchTimer = setTimeout(runSearch, 250);
}

async function runSearch() {
    const query = el.searchInput.value.trim();
    if (!query) {
        el.searchResults.hidden = true;
        el.searchResults.innerHTML = "";
        return;
    }
    el.searchResults.hidden = false;
    el.searchResults.innerHTML = "";
    el.searchResults.append(
        Object.assign(document.createElement("div"), { className: "panel-title", textContent: "Searching…" }),
    );
    try {
        const { hits = [], truncated } = await api(`/api/search?q=${encodeURIComponent(query)}&limit=100`);
        el.searchResults.innerHTML = "";
        el.searchResults.append(
            Object.assign(document.createElement("div"), {
                className: "panel-title",
                textContent: hits.length
                    ? `${hits.length}${truncated ? "+" : ""} ${hits.length === 1 ? "match" : "matches"}`
                    : "No matches",
            }),
        );
        for (const hit of hits) {
            el.searchResults.append(
                entryButton({
                    label: hit.snippet || query,
                    page: hit.page,
                    onClick: () => goToPage(hit.page),
                }),
            );
        }
    } catch (err) {
        el.searchResults.innerHTML = "";
        el.searchResults.append(
            Object.assign(document.createElement("div"), {
                className: "panel-title",
                textContent: err.message,
            }),
        );
    }
}

function applyState(next, { forceReload = false } = {}) {
    const previousKey = state?.doc?.key ?? null;
    state = next;
    if (state?.lastPage && !forceReload) currentPage = state.lastPage;
    if (state?.wordVersion) {
        el.engineNote.textContent = `Rendered by Microsoft Word ${state.wordVersion} for page-accurate layout.`;
    }
    render();
    if (state?.doc && state.doc.key !== previousKey) {
        outlineLoadedFor = null;
        if (el.sidebar.dataset.open === "true") loadOutline();
        // Deliberately no `showPdf` here. `render()` above has already loaded
        // it -- synchronously as far as `loadedPdfUrl`, so a second call was
        // only ever a no-op -- and it loads under the condition this branch does
        // not test: `state.status !== "error"`. A document whose refresh failed
        // terminally keeps its last-known `doc` and a `pdfUrl` naming a render
        // that has since been evicted, so a panel re-opened onto that instance
        // fetched it, failed, and replaced the host's typed message ("…is held
        // by another program") with "Could not display the document" -- an error
        // about something the reader was never being shown, in place of the one
        // they could act on. Found by `app.test.mjs` once #76 made this file
        // executable.
    }
}

// --- events ----------------------------------------------------------------

function connect() {
    const source = new EventSource("/events");
    source.addEventListener("state", (event) => applyState(JSON.parse(event.data)));
    source.addEventListener("rendering", () => setStatus("Document changed, re-rendering…", { busy: true }));
    source.addEventListener("reloaded", (event) => {
        const payload = JSON.parse(event.data);
        applyState(payload);
        currentPage = payload.restorePage ?? currentPage;
        showPdf(currentPage);
        outlineLoadedFor = null;
        if (el.sidebar.dataset.open === "true") loadOutline();
        if (el.searchInput.value.trim()) runSearch();
        setStatus("Reloaded from disk.");
        setTimeout(() => setStatus(null), 2500);
    });
    source.addEventListener("goto", (event) => {
        const { page } = JSON.parse(event.data);
        currentPage = page;
        view.goToPage(page);
    });
    // Everything the panel says about the connection itself lives in
    // `connection-status.mjs`, which owns `onopen`/`onerror` from here. #66: the
    // handler this replaced claimed "Reconnecting…" for as long as `readyState`
    // was CONNECTING, which `EventSource` keeps it at forever.
    monitorConnection(source, { setStatus });
}

el.toggleSidebar.addEventListener("click", () => {
    const open = el.sidebar.dataset.open !== "true";
    el.sidebar.dataset.open = String(open);
    el.toggleSidebar.classList.toggle("active", open);
    el.toggleSidebar.setAttribute("aria-pressed", String(open));
    render();
    if (open) {
        loadOutline();
        el.searchInput.focus();
    }
});

el.copyPath.addEventListener("click", async () => {
    // The path the button announces, not `state` read afresh: what is copied is
    // then provably the thing its accessible name just said.
    const path = el.copyPath.title;
    let failure = null;
    try {
        await navigator.clipboard.writeText(path);
    } catch (err) {
        // Includes `navigator.clipboard` being absent, which throws a TypeError
        // here rather than rejecting. Outside a secure context it simply is not
        // there, and the panel is not always one.
        failure = err;
    }
    const outcome = copyOutcome(failure);
    setStatus(outcome.text, { error: outcome.error });
    setTimeout(() => setStatus(null), 2500);
});

el.reload.addEventListener("click", async () => {
    setStatus("Re-rendering…", { busy: true });
    try {
        await api("/api/refresh", { method: "POST", body: JSON.stringify({ force: true }) });
        // The state push that follows carries the new render key, and the
        // `reloaded` event reloads the view. Nothing to do here but say so.
        setStatus("Reloaded from disk.");
        setTimeout(() => setStatus(null), 2500);
    } catch (err) {
        setStatus(err.message, { error: true });
    }
});

el.openInWord.addEventListener("click", async () => {
    try {
        await api("/api/open-in-word", { method: "POST" });
        setStatus("Opening in Microsoft Word…");
        setTimeout(() => setStatus(null), 2500);
    } catch (err) {
        setStatus(err.message, { error: true });
    }
});

el.jumpToChange.addEventListener("click", () => {
    const page = state?.change?.page;
    if (page) goToPage(page);
});

el.dismissChange.addEventListener("click", () => {
    dismissedChange = changeKey(state?.change ?? null);
    showChange();
});

el.searchInput.addEventListener("input", scheduleSearch);
el.searchInput.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
        event.preventDefault();
        clearTimeout(searchTimer);
        runSearch();
    }
});

// The real Word mark, if this machine has one (#68). Wired at load and never
// again: the answer cannot change while the panel lives.
//
// One placement, beside the document's name. The Open in Word button had a
// second one until #87, which made that button's appearance depend on whether
// the mark could be extracted here; it now always draws its own glyph.
//
// Deliberately not awaited and deliberately not reported. It decorates a bar
// that is already drawn, so nothing here is allowed to delay or fail the
// startup below.
showWordMark({ img: el.docNameMark });

api("/api/state")
    .then((initial) => applyState(initial))
    .catch((err) => setStatus(err.message, { error: true }))
    .finally(connect);
