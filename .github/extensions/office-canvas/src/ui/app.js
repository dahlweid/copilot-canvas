// Viewer front-end. Talks to the extension over plain HTTP; there is no
// privileged bridge from an extension canvas iframe to the host.

import { PdfView } from "./pdf-view.mjs";
import { describeChange } from "./change-wording.mjs";

const $ = (id) => document.getElementById(id);

const el = {
    bar: $("bar"),
    docName: $("docName"),
    docMeta: $("docMeta"),
    docPath: $("docPath"),
    status: $("status"),
    sidebar: $("sidebar"),
    outline: $("outline"),
    searchInput: $("searchInput"),
    searchResults: $("searchResults"),
    viewer: $("viewer"),
    pages: $("pages"),
    picker: $("picker"),
    pathForm: $("pathForm"),
    pathInput: $("pathInput"),
    pickerError: $("pickerError"),
    recents: $("recents"),
    recentsBlock: $("recentsBlock"),
    workspaceDocs: $("workspaceDocs"),
    workspaceBlock: $("workspaceBlock"),
    engineNote: $("engineNote"),
    toggleSidebar: $("toggleSidebar"),
    reload: $("reload"),
    openInWord: $("openInWord"),
    changeDoc: $("changeDoc"),
    changeBar: $("changeBar"),
    changeText: $("changeText"),
    jumpToChange: $("jumpToChange"),
    dismissChange: $("dismissChange"),
};

let state = null;
let currentPage = 1;
let outlineLoadedFor = null;
let forcePicker = false;

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
    el.changeText.textContent = record.locatable
        ? `${describeChange(record)} — page ${record.page}`
        : `${describeChange(record)} — page ${record.page}, marked but not highlighted`;
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
    const showPicker = forcePicker || !state?.doc;

    el.bar.hidden = !state?.doc;
    el.picker.classList.toggle("visible", showPicker);
    el.viewer.classList.toggle("visible", !showPicker && hasDoc);
    el.sidebar.hidden = showPicker || el.sidebar.dataset.open !== "true";

    if (state?.doc) {
        const doc = state.doc;
        el.docName.textContent = doc.title?.trim() ? `${doc.title}` : doc.name;
        const bits = [`${doc.pageCount} ${doc.pageCount === 1 ? "page" : "pages"}`];
        if (doc.wordCount) bits.push(`${doc.wordCount.toLocaleString()} words`);
        if (doc.author?.trim()) bits.push(doc.author);
        el.docMeta.textContent = bits.join(" · ");
        el.docPath.textContent = doc.path;
        el.docPath.title = doc.path;
        document.title = doc.name;
    }

    if (state?.status === "error" && state.error) {
        setStatus(state.error.message, { error: true });
    } else if (state?.status === "opening") {
        setStatus("Opening in Word…", { busy: true });
    } else {
        setStatus(null);
    }

    if (!showPicker && hasDoc) showPdf(currentPage);
}

function entryButton({ label, page, snippet, level, onClick }) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "entry";
    if (level) button.dataset.level = String(level);

    const labelSpan = document.createElement("span");
    labelSpan.className = "label";
    labelSpan.textContent = label;
    if (snippet) {
        const s = document.createElement("span");
        s.className = "snippet";
        s.textContent = snippet;
        labelSpan.append(s);
    }
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

async function loadPicker() {
    try {
        const { workspaceDocs = [], recents = [], workspacePath } = await api("/api/browse");

        el.recentsBlock.hidden = recents.length === 0;
        el.recents.innerHTML = "";
        for (const entry of recents) {
            const li = document.createElement("li");
            li.append(
                entryButton({
                    label: entry.name,
                    snippet: entry.path,
                    onClick: () => open(entry.path),
                }),
            );
            el.recents.append(li);
        }

        el.workspaceBlock.hidden = workspaceDocs.length === 0;
        if (workspacePath) {
            el.workspaceBlock.querySelector("h2").textContent = `In ${workspacePath}`;
        }
        el.workspaceDocs.innerHTML = "";
        for (const doc of workspaceDocs) {
            const li = document.createElement("li");
            li.append(
                entryButton({
                    label: doc.name,
                    snippet: doc.relative,
                    onClick: () => open(doc.path),
                }),
            );
            el.workspaceDocs.append(li);
        }
    } catch (err) {
        el.pickerError.hidden = false;
        el.pickerError.textContent = err.message;
    }
}

async function open(docPath) {
    el.pickerError.hidden = true;
    setStatus("Opening in Word…", { busy: true });
    try {
        const { state: next } = await api("/api/open", {
            method: "POST",
            body: JSON.stringify({ path: docPath }),
        });
        forcePicker = false;
        currentPage = 1;
        applyState(next, { forceReload: true });
    } catch (err) {
        el.pickerError.hidden = false;
        el.pickerError.textContent = err.message;
        setStatus(err.message, { error: true });
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
        showPdf(currentPage);
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
    source.onerror = () => {
        // EventSource reconnects on its own; surface it only if it lingers.
        setTimeout(() => {
            if (source.readyState === EventSource.CONNECTING) setStatus("Reconnecting…", { busy: true });
        }, 1500);
    };
}

el.pathForm.addEventListener("submit", (event) => {
    event.preventDefault();
    const value = el.pathInput.value.trim();
    if (value) open(value);
});

el.toggleSidebar.addEventListener("click", () => {
    const open = el.sidebar.dataset.open !== "true";
    el.sidebar.dataset.open = String(open);
    el.toggleSidebar.classList.toggle("active", open);
    render();
    if (open) {
        loadOutline();
        el.searchInput.focus();
    }
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

el.changeDoc.addEventListener("click", () => {
    forcePicker = true;
    loadPicker();
    render();
    el.pathInput.focus();
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

api("/api/state")
    .then((initial) => {
        applyState(initial);
        if (!initial.doc) loadPicker();
    })
    .catch((err) => setStatus(err.message, { error: true }))
    .finally(connect);
