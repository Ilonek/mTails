/* Bridge access, document lifecycle, autosave and print/PDF workflows.
 *
 * Every bridge slot answers with JSON.  Chromium renders PDFs and print jobs
 * asynchronously on the native side and stops delivering QWebChannel signals
 * while it owns the page, so those operations poll a status slot instead of
 * waiting for a callback.
 */

import { el, showDialog, toast } from "./ui.js";
import { t } from "./i18n.js";
import { renderPrintPages, clearPrint, pageMetrics } from "./print.js";

let bridge = null;
let mock = false;
let info = {};
let autosaveTimer = null;

const EXT = { xlsx: ".xlsx", csv: ".csv", tsv: ".tsv" };

function parse(value) {
  if (typeof value !== "string") return value;
  try {
    return JSON.parse(value);
  } catch {
    return value;
  }
}

function call(name, ...args) {
  return new Promise((resolve, reject) => {
    if (!bridge) {
      reject(new Error(t("The file bridge is not available.")));
      return;
    }
    let settled = false;
    const timer = setTimeout(() => {
      if (!settled) {
        settled = true;
        reject(new Error(t("The operation timed out.")));
      }
    }, 300000);
    try {
      bridge[name](...args, (result) => {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        resolve(parse(result));
      });
    } catch (error) {
      settled = true;
      clearTimeout(timer);
      reject(error);
    }
  });
}

function notify(name, ...args) {
  if (!bridge) return;
  try {
    bridge[name](...args);
  } catch {
    /* the window may be closing */
  }
}

function waitForBridgeResult(name, timeoutMs = 90000) {
  return new Promise((resolve) => {
    if (!bridge) {
      resolve({ ok: false, error: t("The print bridge is not available.") });
      return;
    }
    const started = Date.now();
    const tick = () => {
      if (!bridge) {
        resolve({ ok: false, error: t("The print bridge is not available.") });
        return;
      }
      try {
        bridge[name]((value) => {
          const data = parse(value);
          if (data) resolve(data);
          else if (Date.now() - started > timeoutMs) resolve({ ok: false, error: t("The operation timed out.") });
          else setTimeout(tick, 200);
        });
      } catch (error) {
        resolve({ ok: false, error: String(error) });
      }
    };
    tick();
  });
}

function makeMock() {
  const fonts = ["Calibri", "Carlito", "Liberation Serif", "Liberation Sans", "DejaVu Sans", "Noto Sans", "Georgia", "Arial", "Times New Roman", "Courier New"];
  return {
    bootstrap(callback) {
      callback(JSON.stringify({
        version: "1.0.0",
        platform: "browser",
        fonts,
        recent: [],
        settings: {
          theme: "light", autosave: true, autosaveInterval: 120, autorecover: true,
          zoom: 100, uiLanguage: "system", defaultFont: "Calibri", defaultSize: 11,
          showGridlines: true, showHeadings: true, showFormulaBar: true, defaultSheets: 1,
        },
        defaultFont: "Calibri",
        defaultSize: 11,
        startupFile: "",
        autosaves: [],
        home: "",
        separator: "/",
      }));
    },
    openDocument(_path, callback) {
      callback(JSON.stringify({ ok: false, canceled: true }));
    },
    openPath(_path, callback) {
      callback(JSON.stringify({ ok: false, error: "Native file access is only available in the desktop app." }));
    },
    saveDocument(_json, _fmt, _suggested, _current, _saveAs, callback) {
      callback(JSON.stringify({ ok: false, error: "Native file access is only available in the desktop app." }));
    },
    exportDocument(_json, _fmt, _suggested, callback) {
      callback(JSON.stringify({ ok: false, error: "Native file access is only available in the desktop app." }));
    },
    choosePdfPath(_suggested, callback) {
      callback(JSON.stringify({ ok: false, canceled: true }));
    },
    renderPdf(_path, _meta, callback) {
      callback(JSON.stringify({ ok: false, error: "PDF export needs the desktop app." }));
    },
    pdfStatus(callback) {
      callback(JSON.stringify({ ok: false }));
    },
    printDocument(_meta, callback) {
      callback(JSON.stringify({ ok: false, error: "Printing needs the desktop app." }));
    },
    printStatus(callback) {
      callback(JSON.stringify({ ok: false }));
    },
    autosave(_json, _id, _name, _source, callback) {
      callback(JSON.stringify({ ok: false }));
    },
    loadAutosave(_id, callback) {
      callback(JSON.stringify({ ok: false }));
    },
    discardAutosave(_id, callback) {
      callback(JSON.stringify({ ok: true }));
    },
    clearAutosave(_id, callback) {
      callback(JSON.stringify({ ok: true }));
    },
    getRecent(callback) {
      callback(JSON.stringify([]));
    },
    setSetting() {},
    setTitle() {},
    setModified() {},
    setDocumentPath() {},
    requestClose() {},
    closeNow() {},
    openExternal(url) {
      window.open(url, "_blank", "noopener");
    },
  };
}

export async function initBridge() {
  if (window.QWebChannel && window.qt && window.qt.webChannelTransport) {
    try {
      await new Promise((resolve, reject) => {
        const timer = setTimeout(() => reject(new Error("web channel timeout")), 2500);
        new window.QWebChannel(window.qt.webChannelTransport, (channel) => {
          clearTimeout(timer);
          bridge = channel.objects.bridge || null;
          if (bridge) resolve();
          else reject(new Error("bridge object missing"));
        });
      });
    } catch {
      bridge = null;
    }
  }
  if (!bridge) {
    bridge = makeMock();
    mock = true;
  }
  return bridge;
}

export function isMock() {
  return mock;
}

export function getBridge() {
  return bridge;
}

export async function bootstrap() {
  try {
    info = await call("bootstrap");
  } catch {
    info = {};
  }
  return info;
}

export function getInfo() {
  return info;
}

function hashPath(path) {
  let hash = 2166136261;
  for (let index = 0; index < path.length; index += 1) {
    hash ^= path.charCodeAt(index);
    hash = Math.imul(hash, 16777619);
  }
  return (hash >>> 0).toString(36);
}

export function docIdFor(path) {
  return path ? "book-" + hashPath(path) : "untitled-" + Math.random().toString(36).slice(2, 10);
}

export function suggestedName(title, format) {
  const base = String(title || "Workbook").replace(/[\\/:*?"<>|]+/g, "");
  return base + (EXT[format] || ".xlsx");
}

function extensionFor(path) {
  const match = /\.([^.\\/]+)$/.exec(path || "");
  return match ? match[1].toLowerCase() : "";
}

export function formatForPath(path, fallback = "xlsx") {
  const ext = extensionFor(path);
  if (ext === "xlsm") return "xlsx";
  if (ext === "txt") return "tsv";
  return EXT[ext] ? ext : fallback;
}

/* --------------------------------------------------------------- loading */
export function applyModel(app, model, meta = {}) {
  const title = String(meta.title || model.title || meta.name || "Workbook1").replace(/\.(xlsx|xlsm|csv|tsv)$/i, "");
  app.loadModel(model, {
    path: meta.path || "",
    title,
    format: meta.format || (meta.path ? formatForPath(meta.path) : "xlsx"),
    docId: meta.docId || docIdFor(meta.path || ""),
    dirty: !!meta.dirty,
  });
  return true;
}

export async function openFile(app, path = "") {
  if (!(await confirmSaveIfNeeded(app))) return false;
  try {
    const result = await call("openDocument", path);
    if (!result || result.canceled) return false;
    if (!result.ok) {
      showDialog({
        title: "Unable to open the workbook",
        size: "narrow",
        body: el("p", { text: result.error || "Unknown error", style: "margin:0;font-size:12.5px;line-height:1.6" }),
        buttons: [{ label: "Close", kind: "primary" }],
      });
      return false;
    }
    applyModel(app, result.model, {
      path: result.path,
      // The file name wins: workbook properties often keep the original
      // "Workbook1" title after a Save As, which would rename the tab.
      title: result.name || result.model?.title,
      format: result.format || formatForPath(result.path),
    });
    toast(t("Opened {name}", { name: result.name }), "success");
    return true;
  } catch (error) {
    toast(t("Could not open the file: {error}", { error: error.message }), "error");
    return false;
  }
}

export async function openRecent(app, path) {
  return openFile(app, path);
}

/* ---------------------------------------------------------------- saving */
export async function confirmSaveIfNeeded(app) {
  if (!app.state.dirty) return true;
  return new Promise((resolve) => {
    let done = false;
    showDialog({
      title: "Save changes?",
      size: "narrow",
      body: el("p", {
        text: t("Do you want to save your changes to “{name}” before continuing?", { name: app.state.title || "Workbook1" }),
        style: "margin:0;font-size:12.5px;line-height:1.6;color:var(--text-soft)",
      }),
      buttons: [
        { label: "Cancel", kind: "ghost", onClick: () => { done = true; resolve(false); } },
        { label: "Don't Save", kind: "ghost", onClick: () => { done = true; resolve(true); } },
        {
          label: "Save",
          kind: "primary",
          onClick: async () => {
            done = true;
            resolve(await saveFile(app, {}));
          },
        },
      ],
      onClose: () => { if (!done) resolve(false); },
    });
  });
}

export async function saveFile(app, { saveAs = false, format = "" } = {}) {
  const fmt = format || app.state.format || "xlsx";
  const oldId = app.state.docId;
  try {
    const result = await call(
      "saveDocument",
      JSON.stringify(app.model),
      fmt,
      suggestedName(app.state.title, fmt),
      app.state.path,
      !!saveAs,
    );
    if (!result || result.canceled) return false;
    if (!result.ok) {
      toast(t(result.error || "The workbook could not be saved."), "error");
      return false;
    }
    app.state.path = result.path;
    app.state.format = result.format || formatForPath(result.path, fmt);
    app.state.title = result.title || app.state.title;
    app.state.docId = docIdFor(result.path);
    app.model.title = app.state.title;
    app.setDirty(false);
    app.afterSave(result);
    if (oldId && oldId !== app.state.docId) call("clearAutosave", oldId).catch(() => {});
    call("clearAutosave", app.state.docId).catch(() => {});
    toast(t("Saved {name}", { name: result.name }), "success");
    return true;
  } catch (error) {
    toast(t("Could not save: {error}", { error: error.message }), "error");
    return false;
  }
}

export function saveAsFormat(app, format) {
  return saveFile(app, { saveAs: true, format });
}

export async function exportAs(app, format, targetPath = "") {
  if (format === "pdf") return exportPdf(app, targetPath);
  try {
    const result = targetPath
      ? await call("exportTo", JSON.stringify(app.model), format, targetPath)
      : await call(
        "exportDocument",
        JSON.stringify(app.model),
        format,
        suggestedName(app.state.title, format),
      );
    if (!result || result.canceled) return false;
    if (!result.ok) {
      toast(t(result.error || "Export failed."), "error");
      return false;
    }
    toast(t("Exported {name}", { name: result.name }), "success");
    return true;
  } catch (error) {
    toast(t("Export failed: {error}", { error: error.message }), "error");
    return false;
  }
}

/* ------------------------------------------------------------ print / pdf */
async function waitForPaint() {
  await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
  await new Promise((resolve) => setTimeout(resolve, 160));
}

function preparePrint(app) {
  renderPrintPages(app.model, { allSheets: true });
  document.body.classList.add("printing");
}

function finishPrint() {
  document.body.classList.remove("printing");
  clearPrint();
}

export async function exportPdf(app, targetPath = "") {
  try {
    const target = targetPath ? { path: targetPath } : await call("choosePdfPath", suggestedName(app.state.title, "pdf"));
    if (!target || target.canceled) return false;
    toast(t("Preparing PDF…"));
    preparePrint(app);
    await waitForPaint();
    const first = app.model.sheets[app.model.activeSheet] || app.model.sheets[0];
    const metrics = printMetricsFor(first);
    const started = await call("renderPdf", target.path, JSON.stringify({ w: metrics.w, h: metrics.h }));
    if (!started || started.ok === false) {
      finishPrint();
      toast(t(started?.error || "PDF export failed."), "error");
      return false;
    }
    const result = await waitForBridgeResult("pdfStatus", 90000);
    finishPrint();
    if (!result || !result.ok) {
      toast(t(result?.error || "PDF export failed."), "error");
      return false;
    }
    toast(t("PDF exported to {path}", { path: target.path }), "success");
    return true;
  } catch (error) {
    finishPrint();
    toast(t("PDF export failed: {error}", { error: error.message }), "error");
    return false;
  }
}

export async function printFile(app) {
  try {
    const first = app.model.sheets[app.model.activeSheet] || app.model.sheets[0];
    const metrics = printMetricsFor(first);
    preparePrint(app);
    await waitForPaint();
    const started = await call("printDocument", JSON.stringify({ w: metrics.w, h: metrics.h }));
    if (started && started.canceled) {
      finishPrint();
      return false;
    }
    if (!started || started.ok === false) {
      finishPrint();
      toast(t(started?.error || "Printing failed."), "error");
      return false;
    }
    const result = await waitForBridgeResult("printStatus", 120000);
    finishPrint();
    if (!result || result.ok === false) {
      toast(t(result?.error || "Printing failed."), "error");
      return false;
    }
    return true;
  } catch (error) {
    finishPrint();
    toast(t("Printing failed: {error}", { error: error.message }), "error");
    return false;
  }
}

/* Local import to avoid a circular dependency at module load. */
function printMetricsFor(sheet) {
  const metrics = pageMetrics(sheet?.page || {});
  return { w: metrics.w, h: metrics.h };
}

/* --------------------------------------------------------------- autosave */
export function startAutosave(app) {
  clearInterval(autosaveTimer);
  autosaveTimer = null;
  const enabled = app.settings.autosave !== false;
  const indicator = document.getElementById("autosave-indicator");
  if (indicator) {
    indicator.classList.toggle("off", !enabled);
    indicator.title = t(enabled ? "Autosave is on" : "Autosave is off");
    const label = indicator.querySelector("span");
    if (label) label.textContent = t(enabled ? "Autosave" : "Autosave off");
  }
  if (!enabled) return;
  const seconds = Math.max(30, Number(app.settings.autosaveInterval) || 120);
  autosaveTimer = setInterval(() => autosaveTick(app), seconds * 1000);
}

export async function autosaveTick(app) {
  if (!app.state.dirty || mock) return;
  const indicator = document.getElementById("autosave-indicator");
  indicator?.classList.add("saving");
  try {
    const result = await call("autosave", JSON.stringify(app.model), app.state.docId, app.state.title, app.state.path);
    if (result?.ok) {
      indicator?.classList.remove("failure");
      indicator?.setAttribute("data-saved", String(Date.now()));
    }
  } catch {
    indicator?.classList.add("failure");
  } finally {
    indicator?.classList.remove("saving");
  }
}

export async function recoverAutosaves(app, autosaves) {
  if (!autosaves || !autosaves.length || mock) return;
  let closeDialog = null;
  const list = el("div", { class: "recovery-list" });
  for (const item of autosaves) {
    const card = el("div", { class: "recovery-card" }, [
      el("div", { class: "recovery-icon" }, [
        el("img", { src: "../assets/mtails.png", alt: "mTails" }),
      ]),
      el("div", { class: "recovery-meta" }, [
        el("strong", { text: item.name || t("Untitled workbook"), i18n: false }),
        el("span", {
          text: (item.source || t("Unsaved workbook")) + " \u00B7 " + new Date((item.savedAt || 0) * 1000).toLocaleString(),
          i18n: false,
        }),
      ]),
      el("button", { class: "btn primary", type: "button", text: "Recover" }),
    ]);
    card.querySelector("button").addEventListener("click", async () => {
      try {
        const result = await call("loadAutosave", item.id);
        if (!result || !result.ok) {
          toast(t(result?.error || "The recovery file could not be read."), "error");
          return;
        }
        applyModel(app, result.model, {
          title: result.name || item.name,
          path: result.source || "",
          format: formatForPath(result.source || "", "xlsx"),
          docId: item.id,
          dirty: true,
        });
        toast(t("Recovered {name}", { name: item.name || t("Workbook") }), "success");
        closeDialog?.();
      } catch (error) {
        toast(t("Recovery failed: {error}", { error: error.message }), "error");
      }
    });
    list.append(card);
  }
  showDialog({
    title: "Recover unsaved workbooks",
    body: el("div", {}, [
      el("p", {
        class: "dialog-note",
        text: "mTails found workbooks that were not saved before the application closed. Recover one to continue where you left off.",
      }),
      list,
    ]),
    buttons: [
      {
        label: "Discard all",
        kind: "ghost",
        onClick: () => { for (const item of autosaves) call("discardAutosave", item.id).catch(() => {}); },
      },
      { label: "Not now", kind: "ghost" },
    ],
    onOpen: ({ close }) => { closeDialog = close; },
  });
}

export async function discardAllAutosaves(autosaves) {
  for (const item of autosaves || []) {
    try {
      await call("discardAutosave", item.id);
    } catch {
      /* ignore */
    }
  }
}

/* ------------------------------------------------------------------- misc */
export async function refreshRecent() {
  try {
    return await call("getRecent");
  } catch {
    return [];
  }
}

export function setSetting(app, key, value) {
  app.settings[key] = value;
  notify("setSetting", key, JSON.stringify(value));
}

export function closeWindow() {
  notify("requestClose");
}

export function closeNow() {
  notify("closeNow");
}

export function openExternal(url) {
  if (!url) return;
  if (mock) {
    window.open(url, "_blank", "noopener");
    return;
  }
  notify("openExternal", url);
}

export async function saveBeforeClose(app) {
  if (!app.state.dirty) {
    closeNow();
    return;
  }
  const ok = await saveFile(app, {});
  if (ok || !app.state.dirty) closeNow();
}
