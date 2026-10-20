/* mTails application shell: workbook identity, history, chrome and backstage.
 *
 * app.js is the only module that owns the document state.  The grid reports
 * every mutation through onChange(); the ribbon, dialogs and backstage call
 * app.mutate() or the structural helpers, so each change reaches the same
 * revision counter and the same undo stack.  A revision is only ever advanced
 * by a recorded change, which keeps the dirty marker and the bridge title in
 * step without cloning the model on every keystroke.
 */

import {
  $, $$, el, toast, closeMenus, closeAllDialogs, openMenu, openContextMenu, showDialog,
  confirmDialog, promptDialog, field, numberInput, formatDate, clamp,
} from "./ui.js";
import {
  t, tp, initI18n, currentLanguage,
} from "./i18n.js";
import {
  createWorkbook, cloneModel, normalizeModel, addSheet, removeSheet, moveSheet,
  renameSheet, usedRange, workbookStats, commentsAt, addComment, deleteComment,
  clearCell, cellAt, mergeAt, uniqueSheetName,
} from "./model.js";
import {
  FormulaEngine, FUNCTION_NAMES, functionArity as formulaArity,
} from "./formula.js";
import { formatValue, parseUserInput } from "./format.js";
import { Grid } from "./grid.js";
import { rawEditText } from "./grid-base.js";
import { Ribbon } from "./ribbon.js";
import { Commands } from "./commands.js";
import * as file from "./file.js";
import * as dlg from "./dialogs.js";
import { TEMPLATES, templateById } from "./templates.js";
import {
  renderPrintPages, renderPreview, printSummary, computePageBreaks, clearPrint,
} from "./print.js";
import { icon } from "./icons.js";
import {
  addr, parseAddr, parseRange, rangeText, splitSheetRef, quoteSheet,
  indexToCol, MAX_ROWS, MAX_COLS,
} from "./units.js";

const APP_NAME = "mTails";
const APP_VERSION = "1.0.0";
const MAX_HISTORY = 80;
const ZOOM_STEPS = [40, 50, 60, 70, 75, 80, 90, 100, 110, 120, 130, 140, 150, 160, 175, 200];
const LANGUAGES = ["system", "en", "pl"];
const PANELS = ["new", "open", "info", "save", "saveas", "print", "export", "options", "close"];

/* --------------------------------------------------------------- utilities */
function bind(selector, event, handler, root = document) {
  const node = root.querySelector(selector);
  if (node) node.addEventListener(event, handler);
  return node;
}

let idCounter = 0;
function freshId(prefix) {
  idCounter += 1;
  return prefix + "-" + Date.now().toString(36) + "-" + idCounter;
}

function escapeRegExp(text) {
  return String(text).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function numberText(value) {
  if (!Number.isFinite(value)) return "";
  const text = formatValue(value, "General");
  return text === undefined || text === null ? String(value) : String(text);
}

function selectionArea(range) {
  return (range.r2 - range.r1 + 1) * (range.c2 - range.c1 + 1);
}

function visibleSheets(model) {
  return model.sheets.filter((sheet) => !sheet.hidden);
}

/* ------------------------------------------------------------------ class */
class App {
  constructor() {
    this.info = {};
    this.errors = [];
    this.fonts = [];
    this.recent = [];
    this.autosaves = [];
    this.defaultFont = "Calibri";
    this.defaultSize = 11;
    this.settings = {
      theme: "light",
      autosave: true,
      autosaveInterval: 120,
      autorecover: true,
      zoom: 100,
      uiLanguage: "system",
      defaultFont: "Calibri",
      defaultSize: 11,
      showGridlines: true,
      showHeadings: true,
      showFormulaBar: true,
      defaultSheets: 1,
      spellcheck: true,
      proofingLanguage: "en-US",
      customStyles: [],
    };
    this.state = {
      path: "",
      format: "xlsx",
      title: "Workbook1",
      docId: file.docIdFor(""),
      dirty: false,
      view: "normal",
      backstage: null,
      modifiedAt: Date.now(),
    };
    this.model = createWorkbook({ sheets: 1 });
    this.engine = new FormulaEngine(this.model);
    this.grid = null;
    this.ribbon = null;
    this.commands = null;
    this.selection = null;
    this.revision = 0;
    this.savedRevision = 0;
    this.undoStack = [];
    this.redoStack = [];
    this.editText = null;
    this.tabsSignature = "";
    this.lastShell = "";
    this.statusTimer = 0;
    this.palette = null;
    this.findState = { open: false, replace: false, term: "", count: 0, matches: [] };
    this.booted = false;
    this.layoutTimer = 0;
    this.file = this.buildFileApi();
  }

  buildFileApi() {
    return {
      openFile: (path = "") => file.openFile(this, path),
      openRecent: (path) => file.openRecent(this, path),
      saveFile: (options = {}) => file.saveFile(this, options || {}),
      saveAsFormat: (format) => file.saveAsFormat(this, format),
      exportAs: (format, path = "") => file.exportAs(this, format, path),
      exportPdf: (path = "") => file.exportPdf(this, path),
      printFile: () => file.printFile(this),
      closeWindow: () => file.closeWindow(),
      closeNow: () => file.closeNow(),
      openExternal: (url) => file.openExternal(url),
      saveBeforeClose: () => file.saveBeforeClose(this),
      startAutosave: () => file.startAutosave(this),
      refreshRecent: () => file.refreshRecent(),
      setSetting: (key, value) => file.setSetting(this, key, value),
      confirmSaveIfNeeded: () => file.confirmSaveIfNeeded(this),
      applyModel: (model, meta) => file.applyModel(this, model, meta),
      getBridge: () => file.getBridge(),
      isMock: () => file.isMock(),
      recoverAutosaves: (list) => file.recoverAutosaves(this, list),
      discardAllAutosaves: (list) => file.discardAllAutosaves(list),
      suggestedName: (title, format) => file.suggestedName(title, format),
      formatForPath: (path, fallback) => file.formatForPath(path, fallback),
    };
  }

  /* ------------------------------------------------------------- lifecycle */
  async init() {
    document.getElementById("app")?.classList.add("loading");
    await file.initBridge();
    this.info = (await file.bootstrap()) || {};
    this.fonts = Array.isArray(this.info.fonts) ? this.info.fonts.slice() : [];
    this.settings = { ...this.settings, ...(this.info.settings || {}) };
    this.settings.customStyles = Array.isArray(this.settings.customStyles)
      ? this.settings.customStyles
      : [];
    this.defaultFont = this.settings.defaultFont || this.info.defaultFont || "Calibri";
    this.defaultSize = clamp(
      Math.round(Number(this.settings.defaultSize) || Number(this.info.defaultSize) || 11),
      6,
      72,
    );
    this.settings.defaultFont = this.defaultFont;
    this.settings.defaultSize = this.defaultSize;
    if (!this.fonts.includes(this.defaultFont)) this.fonts.unshift(this.defaultFont);
    this.recent = Array.isArray(this.info.recent) ? this.info.recent : [];
    this.autosaves = Array.isArray(this.info.autosaves) ? this.info.autosaves : [];
    this.state.docId = file.docIdFor("");
    this.model = createWorkbook({ sheets: Math.max(1, Number(this.settings.defaultSheets) || 1) });
    this.applyDefaults(this.model, true);
    this.engine = new FormulaEngine(this.model);
    initI18n(this.settings.uiLanguage || "system");
    this.applyTheme(this.settings.theme || "light");
    if (this.settings.showFormulaBar === false) $("#formula-bar")?.classList.add("collapsed");
    this.buildGrid();
    this.buildChrome();
    this.renderAll();
    this.grid.renderNow();
    document.getElementById("app")?.classList.remove("loading");
    this.booted = true;
    if (this.info.startupFile) {
      const opened = await file.openFile(this, this.info.startupFile);
      if (!opened) toast(t("The startup workbook could not be opened."), "error");
    }
    if (this.autosaves.length && this.settings.autorecover !== false) {
      file.recoverAutosaves(this, this.autosaves);
    }
    file.startAutosave(this);
    this.syncShell();
  }

  applyDefaults(model, force = false) {
    const gridlines = this.settings.showGridlines !== false;
    const headings = this.settings.showHeadings !== false;
    const zoom = clamp(Number(this.settings.zoom) || 100, 40, 200);
    for (const sheet of model.sheets) {
      if (force || sheet.showGridlines === undefined) sheet.showGridlines = gridlines;
      if (force || sheet.showHeadings === undefined) sheet.showHeadings = headings;
      if (force || !Number(sheet.zoom)) sheet.zoom = zoom;
    }
  }

  buildGrid() {
    this.grid = new Grid({
      model: this.model,
      engine: this.engine,
      defaultFont: this.defaultFont,
      defaultSize: this.defaultSize,
      captureState: () => this.captureState(),
      onSelection: (info) => this.onSelection(info),
      onChange: (event) => this.onChange(event),
      onStatus: (text) => this.onStatus(text),
      onUndo: (shift) => (shift ? this.redo() : this.undo()),
      onRedo: () => this.redo(),
      onFind: () => this.openFind(false),
      onGoTo: () => dlg.goToDialog(this),
      onShortcut: (kind) => this.onShortcut(kind),
      onHyperlink: (url, cell) => this.onHyperlink(url, cell),
      onContextMenu: (x, y) => this.onContextMenu(x, y),
      onCommentRequest: (address) => this.onCommentRequest(address),
      onFilterRequest: (col, anchor) => dlg.filterPopover(this, col, anchor),
      onEditStart: (info) => this.onEditStart(info),
      onEditInput: (value, address) => this.onEditInput(value, address),
      onEditEnd: () => this.onEditEnd(),
      onZoomChange: (value) => this.onZoomChange(value),
      onOverlayActivate: (item, kind) => this.onOverlayActivate(item, kind),
      onOverlayContextMenu: (x, y, item, kind) => this.onOverlayContextMenu(x, y, item, kind),
      onEscape: () => this.handleEscape(),
    });
    this.grid.defaultFont = this.defaultFont;
    this.grid.defaultSize = this.defaultSize;
    this.grid.mount();
  }

  buildChrome() {
    this.ribbon = new Ribbon(this);
    this.ribbon.mount();
    this.commands = new Commands(this);
    this.buildQat();
    this.bindTitlebar();
    this.bindFormulaBar();
    this.bindStatusBar();
    this.bindSheetTabs();
    this.bindComments();
    this.bindFind();
    this.bindBackstage();
    this.bindSearch();
    this.bindImageInput();
    this.bindKeyboard();
    this.renderTabs();
    this.renderComments();
    this.renderTemplates();
    this.renderRecent();
    this.renderRecovery();
    this.updateShell();
  }

  renderAll() {
    this.renderTabs();
    this.renderComments();
    this.renderTemplates();
    this.renderRecent();
    this.renderRecovery();
    this.buildQat();
    this.syncFormulaBar();
    this.updateStatus();
    this.updateShell();
    const sheet = this.grid.sheet();
    this.ribbon?.update({
      ...(this.selection || this.grid.selectionInfo()),
      merged: !!mergeAt(sheet, this.grid.active.row, this.grid.active.col),
      showGridlines: sheet.showGridlines !== false,
      showHeadings: sheet.showHeadings !== false,
    });
  }

  /* --------------------------------------------------------------- command */
  execute(command, anchor = null, extra = {}) {
    if (!command) return false;
    try {
      if (this.commands?.execute(command, anchor, extra)) return true;
      if (command === "palette") {
        this.openPalette();
        return true;
      }
      if (command === "share") {
        this.showShare();
        return true;
      }
      return false;
    } catch (error) {
      this.recordError(error);
      toast(t("The command could not be completed."), "error");
      return false;
    }
  }

  recordError(error, source = "", line = 0) {
    const message = String(error?.message || error || "Unknown error");
    this.errors.push({ message, source: String(source || ""), line: Number(line) || 0 });
    if (this.errors.length > 20) this.errors.shift();
    if (typeof console !== "undefined" && console.error) {
      console.error("[mTails]", message, source || "", line || "");
    }
  }

  /* --------------------------------------------------------------- history */
  captureState() {
    return cloneModel(this.model);
  }

  mutate(kind, fn, { layout = false, recalc = true, changed = true } = {}) {
    const before = this.captureState();
    const result = fn();
    if (recalc) this.grid.recalculate();
    this.grid.afterMutation(kind, before, { layout, changed });
    return result;
  }

  recordChange(kind, before) {
    const previous = this.undoStack[this.undoStack.length - 1];
    if (previous && !previous.after) previous.after = before;
    this.revision += 1;
    this.undoStack.push({ kind, before, after: null, rev: this.revision });
    if (this.undoStack.length > MAX_HISTORY) this.undoStack.shift();
    this.redoStack.length = 0;
    this.state.modifiedAt = Date.now();
    this.syncDirty();
  }

  undo() {
    const entry = this.undoStack.pop();
    if (!entry) {
      toast(t("Nothing to undo."));
      return;
    }
    if (this.state.view !== "normal") this.setView("normal");
    if (!entry.after) entry.after = cloneModel(this.model);
    this.redoStack.push(entry);
    this.applySnapshot(entry.before);
    this.revision = Math.max(0, entry.rev - 1);
    this.syncDirty();
    this.afterHistoryRestore();
  }

  redo() {
    const entry = this.redoStack.pop();
    if (!entry) {
      toast(t("Nothing to redo."));
      return;
    }
    if (this.state.view !== "normal") this.setView("normal");
    this.undoStack.push(entry);
    this.applySnapshot(entry.after || cloneModel(this.model));
    this.revision = entry.rev;
    this.syncDirty();
    this.afterHistoryRestore();
  }

  applySnapshot(snapshot) {
    const index = clamp(Number(snapshot.activeSheet) || 0, 0, Math.max(0, (snapshot.sheets?.length || 1) - 1));
    this.model = normalizeModel(snapshot);
    this.model.activeSheet = index;
    this.engine.setModel(this.model);
    this.grid.setModel(this.model, this.engine);
    this.grid.epoch += 1;
    this.grid.buildLayout();
    this.grid.renderNow();
    this.renderTabs();
    this.renderComments();
    this.syncFormulaBar();
    this.updateStatus();
    this.refreshView();
    this.syncShell();
  }

  afterHistoryRestore() {
    this.renderTabs();
    this.renderComments();
    this.syncFormulaBar();
    this.syncHistoryButtons();
    this.updateStatus();
    this.refreshView();
    this.syncShell();
    this.ribbon?.update({
      ...(this.selection || this.grid.selectionInfo()),
      merged: !!mergeAt(this.grid.sheet(), this.grid.active.row, this.grid.active.col),
      showGridlines: this.grid.sheet().showGridlines !== false,
      showHeadings: this.grid.sheet().showHeadings !== false,
    });
    this.grid.els?.scroller.focus({ preventScroll: true });
  }

  syncDirty() {
    this.state.dirty = this.revision !== this.savedRevision;
    this.syncShell();
  }

  setDirty(dirty) {
    this.state.dirty = !!dirty;
    if (!dirty) this.savedRevision = this.revision;
    else if (this.savedRevision === this.revision) this.savedRevision = -1;
    this.syncShell();
  }

  /* -------------------------------------------------------------- document */
  loadModel(model, meta = {}) {
    const index = clamp(Number(model?.activeSheet) || 0, 0, Math.max(0, (model?.sheets?.length || 1) - 1));
    this.model = normalizeModel(model);
    this.model.activeSheet = index;
    const title = String(meta.title || this.model.title || "Workbook1")
      .replace(/\.(xlsx|xlsm|csv|tsv)$/i, "");
    this.model.title = title;
    if (!this.model.properties.title) this.model.properties.title = title;
    this.state.path = meta.path || "";
    this.state.format = meta.format || (this.state.path ? file.formatForPath(this.state.path, "xlsx") : "xlsx");
    this.state.title = title;
    this.state.docId = meta.docId || file.docIdFor(this.state.path);
    this.state.modifiedAt = Date.now();
    this.state.dirty = !!meta.dirty;
    this.engine.setModel(this.model);
    this.grid.setModel(this.model, this.engine);
    this.grid.epoch += 1;
    this.grid.buildLayout();
    this.grid.renderNow();
    this.revision = 0;
    this.savedRevision = meta.dirty ? -1 : 0;
    this.undoStack.length = 0;
    this.redoStack.length = 0;
    this.lastShell = "";
    this.tabsSignature = "";
    this.editText = null;
    this.commentDraft = null;
    this.closeFind();
    this.clearAudit();
    if (this.state.view !== "normal") this.setView("normal");
    this.renderAll();
    this.syncDirty();
    this.afterHistoryRestore();
  }

  async newWorkbook(templateId = "") {
    if (!(await file.confirmSaveIfNeeded(this))) return false;
    const template = templateById(templateId);
    let model;
    if (template) {
      model = normalizeModel(template.build());
    } else {
      model = createWorkbook({ sheets: Math.max(1, Number(this.settings.defaultSheets) || 1) });
      this.applyDefaults(model, true);
    }
    const title = template ? template.label : "Workbook1";
    file.applyModel(this, model, {
      title,
      path: "",
      format: "xlsx",
      docId: file.docIdFor(""),
      dirty: false,
    });
    this.state.title = title;
    this.model.title = title;
    this.closeBackstage();
    this.lastShell = "";
    this.syncShell();
    toast(template ? t("Created {name}", { name: title }) : t("New workbook"));
    return true;
  }

  async afterSave(result = {}) {
    if (result.path) this.state.path = result.path;
    if (result.format) this.state.format = result.format;
    this.state.modifiedAt = Date.now();
    this.model.title = this.state.title;
    this.lastShell = "";
    this.syncShell();
    this.renderRecent();
    this.renderSaveCard();
    this.renderInfo();
    try {
      this.recent = await file.refreshRecent();
      this.renderRecent();
    } catch {
      /* the bridge may be closing */
    }
  }

  syncShell() {
    const title = this.state.title || "Workbook1";
    const dirty = !!this.state.dirty;
    const key = title + "|" + dirty + "|" + this.state.path;
    if (key === this.lastShell) return;
    this.lastShell = key;
    const titleNode = $("#doc-title");
    if (titleNode) titleNode.textContent = title;
    const stateNode = $("#save-state");
    if (stateNode) {
      stateNode.textContent = t(dirty ? "Unsaved changes" : "Saved");
      stateNode.classList.toggle("dirty", dirty);
    }
    document.title = title + (dirty ? " *" : "") + " - " + APP_NAME;
    const bridge = file.getBridge();
    try {
      bridge?.setDocumentPath?.(this.state.path || "");
      bridge?.setTitle?.(title);
      bridge?.setModified?.(dirty);
    } catch {
      /* the window may be closing */
    }
  }

  updateShell() {
    const sheets = $("#sb-sheets");
    if (sheets) {
      const count = this.model.sheets.length;
      sheets.textContent = tp(count, "{count} sheet", "{count} sheets", { count });
    }
    const language = $("#sb-language");
    if (language) language.textContent = currentLanguage() === "pl" ? "Polski (PL)" : "English (US)";
    document.body.dataset.view = this.state.view || "normal";
    this.syncHistoryButtons();
    $("#sb-normal")?.classList.toggle("active", this.state.view === "normal");
    $("#sb-pagebreak")?.classList.toggle("active", this.state.view === "pagebreak");
    this.syncZoomUi();
    this.updateStatus();
  }

  syncTitle() {
    this.lastShell = "";
    this.syncShell();
  }

  /* -------------------------------------------------------------- chrome */
  buildQat() {
    const qat = $("#qat");
    if (!qat) return;
    qat.replaceChildren();
    const entries = [
      ["save", "Save", "save"],
      ["undo", "Undo", "undo"],
      ["redo", "Redo", "redo"],
      ["printPreview", "Print", "print"],
      ["exportPdf", "Export PDF", "exportPdf"],
      ["painter", "Format Painter", "formatPainter"],
      ["spellcheck", "Spelling", "spellcheck"],
    ];
    for (const [iconName, title, command] of entries) {
      const button = el("button", { type: "button", title, html: icon(iconName) });
      button.addEventListener("click", () => this.execute(command, button));
      qat.append(button);
    }
  }

  bindTitlebar() {
    bind("#btn-comments-top", "click", () => this.toggleComments());
    bind("#btn-share", "click", () => this.showShare());
    bind("#btn-account", "click", () => dlg.aboutDialog(this));
    bind("#autosave-indicator", "click", () => this.openBackstage("options"));
  }

  bindFormulaBar() {
    const input = $("#formula-input");
    const nameBox = $("#name-box");
    if (input) {
      input.addEventListener("focus", () => {
        if (this.editText === null) input.select();
      });
      input.addEventListener("keydown", (event) => {
        if (event.key === "Enter") {
          event.preventDefault();
          this.commitFormulaBar("down");
        } else if (event.key === "Tab") {
          event.preventDefault();
          this.commitFormulaBar(event.shiftKey ? "left" : "right");
        } else if (event.key === "Escape") {
          event.preventDefault();
          this.syncFormulaBar();
          this.grid.els?.scroller.focus({ preventScroll: true });
        }
      });
    }
    if (nameBox) {
      nameBox.addEventListener("focus", () => nameBox.select());
      nameBox.addEventListener("keydown", (event) => {
        if (event.key !== "Enter") return;
        event.preventDefault();
        const value = nameBox.value.trim();
        if (!this.selectReference(value)) {
          toast(t("Enter a cell reference such as B12 or A1:C10."));
          this.syncFormulaBar();
        }
        this.grid.els?.scroller.focus({ preventScroll: true });
      });
    }
    bind("#btn-fx", "click", () => dlg.functionDialog(this));
    bind("#formula-expand", "click", () => {
      const bar = $("#formula-bar");
      bar?.classList.toggle("expanded");
      this.settings.formulaBarExpanded = !!bar?.classList.contains("expanded");
    });
  }

  commitFormulaBar(move = "none") {
    const input = $("#formula-input");
    if (!input) return;
    const text = input.value;
    this.grid.setActiveValue(text, { move });
    this.grid.els?.scroller.focus({ preventScroll: true });
    this.syncFormulaBar();
  }

  bindStatusBar() {
    bind("#sb-normal", "click", () => this.setView("normal"));
    bind("#sb-pagebreak", "click", () => this.setView(this.state.view === "pagebreak" ? "normal" : "pagebreak"));
    bind("#sb-language", "click", () => this.cycleLanguage());
    bind("#sb-cell", "click", () => dlg.goToDialog(this));
    bind("#sb-stats", "click", () => this.copyStats());
    bind("#zoom-out", "click", () => this.stepZoom(-1));
    bind("#zoom-in", "click", () => this.stepZoom(1));
    bind("#zoom-label", "click", () => dlg.zoomDialog(this));
    const slider = $("#zoom-slider");
    if (slider) {
      slider.addEventListener("input", () => this.grid.setZoom(Number(slider.value) || 100));
    }
  }

  bindSheetTabs() {
    bind("#sheet-nav-back", "click", () => this.stepSheet(-1));
    bind("#sheet-nav-fwd", "click", () => this.stepSheet(1));
    bind("#sheet-add", "click", () => {
      this.insertSheet();
      toast(t("Sheet added."));
    });
    bind("#sheet-menu", "click", () => {
      const anchor = $("#sheet-menu");
      const rect = anchor?.getBoundingClientRect();
      this.openSheetMenu(rect ? rect.left : 40, rect ? rect.top : 40, this.model.activeSheet);
    });
  }

  bindComments() {
    bind("#comment-post", "click", () => this.postComment());
    bind("#comment-cancel", "click", () => {
      const input = $("#comment-input");
      if (input) input.value = "";
      this.commentDraft = null;
    });
    $$(".pane-close[data-close-pane='side']").forEach((node) => {
      node.addEventListener("click", () => this.toggleComments(false));
    });
    const input = $("#comment-input");
    if (input) {
      input.addEventListener("keydown", (event) => {
        if (event.key === "Enter" && !event.shiftKey) {
          event.preventDefault();
          this.postComment();
        }
      });
    }
  }

  bindFind() {
    const bar = $("#find-bar");
    const pane = $("#sheet-pane");
    if (bar && pane && bar.parentElement !== pane) pane.append(bar);
    const findInput = $("#find-input");
    if (findInput) {
      findInput.addEventListener("input", () => {
        clearTimeout(this.findTimer);
        this.findTimer = setTimeout(() => this.runFind(), 160);
      });
      findInput.addEventListener("keydown", (event) => {
        if (event.key === "Enter") {
          event.preventDefault();
          this.findStep(event.shiftKey ? -1 : 1);
        } else if (event.key === "Escape") {
          event.preventDefault();
          this.closeFind();
        }
      });
    }
    const replaceInput = $("#replace-input");
    if (replaceInput) {
      replaceInput.addEventListener("keydown", (event) => {
        if (event.key === "Enter") {
          event.preventDefault();
          if (event.shiftKey) this.replaceAll();
          else this.replaceOne();
        } else if (event.key === "Escape") {
          event.preventDefault();
          this.closeFind();
        }
      });
    }
    bind("#find-next", "click", () => this.findStep(1));
    bind("#find-prev", "click", () => this.findStep(-1));
    bind("#find-close", "click", () => this.closeFind());
    bind("#find-replace-toggle", "click", () => {
      const row = bar?.querySelector(".find-row.secondary");
      const open = row?.classList.contains("collapsed");
      row?.classList.toggle("collapsed", !open);
      $("#find-replace-toggle")?.classList.toggle("on", !!open);
      if (open) $("#replace-input")?.focus();
    });
    bind("#replace-one", "click", () => this.replaceOne());
    bind("#replace-all", "click", () => this.replaceAll());
  }

  bindBackstage() {
    bind("#bs-back", "click", () => this.closeBackstage());
    bind("#bs-nav", "click", (event) => {
      const button = event.target.closest("button[data-bs]");
      if (button) this.showBsPanel(button.dataset.bs);
    });
    bind("#bs-browse", "click", async () => {
      const opened = await file.openFile(this, "");
      if (opened) this.closeBackstage();
    });
    bind("#recovery-discard", "click", async () => {
      await file.discardAllAutosaves(this.autosaves);
      this.autosaves = [];
      this.renderRecovery();
      toast(t("Recovery snapshots discarded."));
    });
    bind("#recovery-later", "click", () => $("#recovery-box")?.classList.add("hidden"));
    bind("#bs-save", "click", () => this.file.saveFile());
    bind("#bs-saveas", "click", () => this.file.saveFile({ saveAs: true }));
    $$("#format-grid [data-format]").forEach((button) => {
      button.addEventListener("click", () => this.file.saveAsFormat(button.dataset.format));
    });
    $$("#backstage [data-export]").forEach((button) => {
      button.addEventListener("click", () => this.file.exportAs(button.dataset.export));
    });
    bind("#export-pdf", "click", () => this.file.exportPdf());
    bind("#bs-print", "click", () => this.file.printFile());
    bind("#bs-pdf", "click", () => this.file.exportPdf());
    bind("#bs-page-setup", "click", () => {
      dlg.pageSetupDialog(this);
      this.renderPrintPanel();
    });
    bind("#prop-save", "click", () => this.saveProperties());
    bind("#info-protect", "click", () => dlg.protectionDialog(this, "workbook"));
    bind("#opt-save", "click", () => this.saveOptions());
    bind("#print-orientation", "change", () => this.updatePageSetting("orientation", $("#print-orientation").value));
    bind("#print-size", "change", () => this.updatePageSetting("size", $("#print-size").value));
    bind("#bs-account", "click", () => dlg.aboutDialog(this));
    bind("#bs-close-doc", "click", () => this.closeDocument());
  }

  bindSearch() {
    const input = $("#command-search");
    if (!input) return;
    input.addEventListener("focus", () => this.openPalette(input.value));
    input.addEventListener("input", () => {
      if (this.palette) {
        this.palette.input.value = input.value;
        this.palette.input.dispatchEvent(new Event("input"));
      } else {
        this.openPalette(input.value);
      }
    });
    input.addEventListener("keydown", (event) => {
      if (event.key === "ArrowDown" && this.palette) {
        event.preventDefault();
        this.palette.input.focus();
      } else if (event.key === "Escape") {
        event.preventDefault();
        this.closePalette();
        input.blur();
      }
    });
  }

  bindImageInput() {
    const picker = $("#image-input");
    if (!picker) return;
    picker.addEventListener("change", () => {
      const picked = picker.files && picker.files[0];
      picker.value = "";
      if (!picked) return;
      if (picked.size > 8 * 1024 * 1024) {
        toast(t("Choose an image smaller than 8 MB."), "error");
        return;
      }
      const reader = new FileReader();
      reader.onload = () => {
        const src = String(reader.result || "");
        if (!src) return;
        const probe = new Image();
        probe.onload = () => {
          const ratio = probe.width / Math.max(1, probe.height);
          const width = Math.min(420, Math.max(120, probe.width || 200));
          this.addImage(src, width, Math.round(width / ratio));
        };
        probe.onerror = () => this.addImage(src, 200, 140);
        probe.src = src;
      };
      reader.readAsDataURL(picked);
    });
  }

  /* --------------------------------------------------- selection and state */
  onSelection(info) {
    this.selection = info;
    const sheet = this.grid?.sheet();
    if (!sheet) return;
    this.ribbon?.update({
      ...info,
      merged: !!mergeAt(sheet, this.grid.active.row, this.grid.active.col),
      showGridlines: sheet.showGridlines !== false,
      showHeadings: sheet.showHeadings !== false,
    });
    this.syncFormulaBar();
    this.updateStatus();
  }

  onChange(event = {}) {
    const kind = event.kind || "edit";
    const recordable = event.changed !== false && !!event.before && kind !== "zoom";
    if (recordable) this.recordChange(kind, event.before);
    this.renderTabs();
    this.renderComments();
    this.syncFormulaBar();
    this.updateStatus();
    this.updateShell();
    this.syncShell();
    if (this.state.view !== "normal") this.scheduleViewRefresh();
  }

  onStatus(text) {
    clearTimeout(this.statusTimer);
    const node = $("#sb-mode");
    if (!text) {
      this.statusMessage = "";
      this.updateStatus();
      return;
    }
    this.statusMessage = text;
    if (node) node.textContent = t(text);
    this.statusTimer = setTimeout(() => {
      this.statusMessage = "";
      this.updateStatus();
    }, 2400);
  }

  onEditStart(info = {}) {
    this.editText = info.text ?? "";
    const input = $("#formula-input");
    if (input) input.value = this.editText;
    const nameBox = $("#name-box");
    if (nameBox && info.address) nameBox.value = info.address;
    this.statusMessage = "";
    this.updateStatus();
  }

  onEditInput(value) {
    this.editText = String(value ?? "");
    const input = $("#formula-input");
    if (input && document.activeElement !== input) input.value = this.editText;
    const mode = $("#sb-mode");
    if (mode) mode.textContent = t("Enter");
  }

  onEditEnd() {
    this.editText = null;
    this.syncFormulaBar();
    this.updateStatus();
  }

  onZoomChange() {
    this.syncZoomUi();
  }

  syncFormulaBar() {
    if (!this.grid) return;
    const sheet = this.grid.sheet();
    const cell = sheet ? cellAt(sheet, this.grid.active.row, this.grid.active.col) : null;
    const nameBox = $("#name-box");
    if (nameBox && document.activeElement !== nameBox) nameBox.value = this.grid.activeAddress();
    const input = $("#formula-input");
    if (input && document.activeElement !== input) {
      input.value = this.editText !== null ? this.editText : rawEditText(cell);
    }
  }

  updateStatus() {
    const cellNode = $("#sb-cell");
    if (cellNode && this.grid) cellNode.textContent = this.grid.activeAddress();
    const mode = $("#sb-mode");
    if (mode && !this.statusMessage) {
      if (this.grid?.editing) mode.textContent = t("Editing");
      else if (this.state.view === "pagebreak") mode.textContent = t("Page Break Preview");
      else if (this.state.view === "layout") mode.textContent = t("Page Layout");
      else mode.textContent = t("Ready");
    }
    const statsNode = $("#sb-stats");
    if (!statsNode) return;
    const stats = this.grid?.selectionStats?.();
    if (!stats || !stats.count) {
      statsNode.textContent = "";
      return;
    }
    const parts = [];
    if (stats.numeric > 0) {
      parts.push(t("Sum:") + " " + numberText(stats.sum));
      parts.push(t("Average:") + " " + numberText(stats.average));
      if (stats.numeric > 1) {
        parts.push(t("Min:") + " " + numberText(stats.min));
        parts.push(t("Max:") + " " + numberText(stats.max));
      }
      parts.push(t("Numerical Count:") + " " + String(stats.numeric));
    } else {
      parts.push(t("Count:") + " " + String(stats.count));
    }
    statsNode.textContent = parts.join("    ");
  }

  syncZoomUi() {
    const zoom = this.grid ? Math.round(this.grid.zoom) : 100;
    const slider = $("#zoom-slider");
    if (slider && document.activeElement !== slider) slider.value = String(zoom);
    const label = $("#zoom-label");
    if (label) label.textContent = zoom + "%";
  }

  stepZoom(direction) {
    if (!this.grid) return;
    const current = Number(this.grid.zoom) || 100;
    let index = ZOOM_STEPS.findIndex((step) => step >= current);
    if (index < 0) index = ZOOM_STEPS.length - 1;
    if (ZOOM_STEPS[index] !== current) {
      if (direction < 0) index = Math.max(0, index - 1);
    } else {
      index = clamp(index + (direction > 0 ? 1 : -1), 0, ZOOM_STEPS.length - 1);
    }
    this.grid.setZoom(ZOOM_STEPS[index]);
  }

  copyStats() {
    const stats = this.grid?.selectionStats?.();
    if (!stats || !stats.count) {
      toast(t("Select cells with values first."));
      return;
    }
    const parts = [];
    if (stats.numeric > 0) {
      parts.push(t("Sum:") + " " + numberText(stats.sum));
      parts.push(t("Average:") + " " + numberText(stats.average));
      parts.push(t("Min:") + " " + numberText(stats.min));
      parts.push(t("Max:") + " " + numberText(stats.max));
    }
    parts.push(t("Count:") + " " + String(stats.count));
    this.copyText(parts.join("\t"));
    toast(t("Copied."));
  }

  /* -------------------------------------------------------------- sheet tabs */
  renderTabs() {
    const list = $("#sheet-tab-list");
    if (!list || !this.model) return;
    const signature = this.model.sheets
      .map((sheet) => (sheet.hidden ? "h" : "v") + ":" + sheet.name + ":" + (sheet.tabColor || ""))
      .join("|") + "#" + this.model.activeSheet;
    if (signature === this.tabsSignature) return;
    this.tabsSignature = signature;
    list.replaceChildren();
    this.model.sheets.forEach((sheet, index) => {
      if (sheet.hidden) return;
      const tab = el("button", {
        class: "sheet-tab" + (index === this.model.activeSheet ? " active" : ""),
        type: "button",
        title: sheet.name,
        draggable: "true",
        dataset: { index: String(index) },
        i18n: false,
      });
      if (sheet.tabColor) {
        tab.append(el("span", { class: "tab-dot", i18n: false, style: { background: sheet.tabColor } }));
      }
      tab.append(el("span", { class: "tab-name", text: sheet.name, i18n: false }));
      tab.addEventListener("click", () => this.activateSheet(index));
      tab.addEventListener("dblclick", (event) => {
        event.preventDefault();
        dlg.renameSheetDialog(this, index);
      });
      tab.addEventListener("contextmenu", (event) => {
        event.preventDefault();
        this.openSheetMenu(event.clientX, event.clientY, index);
      });
      tab.addEventListener("dragstart", (event) => {
        this.dragSheet = index;
        event.dataTransfer?.setData("text/plain", String(index));
        if (event.dataTransfer) event.dataTransfer.effectAllowed = "move";
        tab.classList.add("dragging");
      });
      tab.addEventListener("dragend", () => {
        this.dragSheet = null;
        this.dropTarget = null;
        this.clearTabDrop();
        tab.classList.remove("dragging");
      });
      tab.addEventListener("dragover", (event) => {
        if (this.dragSheet === null || this.dragSheet === undefined) return;
        event.preventDefault();
        const rect = tab.getBoundingClientRect();
        const after = event.clientX > rect.left + rect.width / 2;
        this.clearTabDrop();
        tab.classList.add(after ? "drop-after" : "drop-before");
        this.dropTarget = { index, after };
      });
      tab.addEventListener("drop", (event) => {
        event.preventDefault();
        this.finishTabDrop();
      });
      list.append(tab);
    });
    requestAnimationFrame(() => {
      const active = list.querySelector(".sheet-tab.active");
      active?.scrollIntoView({ block: "nearest", inline: "nearest" });
    });
  }

  clearTabDrop() {
    for (const node of $$("#sheet-tab-list .sheet-tab")) {
      node.classList.remove("drop-before", "drop-after");
    }
  }

  finishTabDrop() {
    const from = this.dragSheet;
    const target = this.dropTarget;
    this.clearTabDrop();
    this.dragSheet = null;
    this.dropTarget = null;
    if (from === null || from === undefined || !target) return;
    const before = target.index + (target.after ? 1 : 0);
    this.moveSheet(from, before, false);
  }

  sheetInsertIndex(from, before) {
    const count = this.model.sheets.length;
    const position = clamp(Number(before) || 0, 0, count);
    const index = position > from ? position - 1 : position;
    return clamp(index, 0, Math.max(0, count - 1));
  }

  activateSheet(index, { focus = true } = {}) {
    if (!this.grid) return;
    const target = clamp(Number(index) || 0, 0, this.model.sheets.length - 1);
    const sheet = this.model.sheets[target];
    if (!sheet || sheet.hidden) return;
    if (target === this.grid.sheetIndex) {
      if (focus) this.grid.els?.scroller.focus({ preventScroll: true });
      return;
    }
    this.model.activeSheet = target;
    this.grid.setSheet(target);
    this.tabsSignature = "";
    this.renderTabs();
    this.renderComments();
    this.syncFormulaBar();
    this.updateStatus();
    this.updateShell();
    this.refreshView();
    if (focus) this.grid.els?.scroller.focus({ preventScroll: true });
  }

  insertSheet(name = "", index = null) {
    const at = index === null || index === undefined
      ? this.model.sheets.length
      : clamp(Number(index) || 0, 0, this.model.sheets.length);
    const sheet = this.mutate("sheet-add", () => addSheet(this.model, {
      name: String(name || "").trim() || "Sheet",
      index: at,
    }), { layout: false });
    this.grid.setSheet(this.model.activeSheet);
    this.tabsSignature = "";
    this.renderTabs();
    this.renderComments();
    this.syncFormulaBar();
    this.updateStatus();
    this.updateShell();
    this.refreshView();
    this.grid.els?.scroller.focus({ preventScroll: true });
    return sheet;
  }

  removeSheet(index = null) {
    if (this.model.sheets.length <= 1) {
      toast(t("Cannot delete the last sheet."));
      return false;
    }
    const target = clamp(
      Number(index === null ? this.model.activeSheet : index) || 0,
      0,
      this.model.sheets.length - 1,
    );
    this.mutate("sheet-delete", () => removeSheet(this.model, target), { layout: false });
    this.grid.setSheet(this.model.activeSheet);
    this.tabsSignature = "";
    this.renderTabs();
    this.renderComments();
    this.syncFormulaBar();
    this.updateStatus();
    this.updateShell();
    this.refreshView();
    return true;
  }

  renameSheet(index, name) {
    const target = clamp(Number(index) || 0, 0, this.model.sheets.length - 1);
    const sheet = this.model.sheets[target];
    if (!sheet) return { ok: false, error: "The sheet does not exist." };
    const oldName = sheet.name;
    const before = this.captureState();
    const result = renameSheet(this.model, target, name);
    if (!result || result.ok === false) return result || { ok: false };
    if (result.name !== oldName) {
      this.rewriteSheetReferences(oldName, result.name);
      this.grid.recalculate();
      this.grid.afterMutation("sheet-rename", before, { layout: false });
    }
    this.tabsSignature = "";
    this.renderTabs();
    this.syncShell();
    return result;
  }

  rewriteSheetReferences(from, to) {
    if (!from || !to || from === to) return;
    const quotedFrom = "'" + String(from).replace(/'/g, "''") + "'!";
    const quotedTo = "'" + String(to).replace(/'/g, "''") + "'!";
    const bare = new RegExp("(^|[^A-Za-z0-9_.'])" + escapeRegExp(from) + "!", "g");
    const replace = (text) => {
      if (typeof text !== "string" || !text) return text;
      let next = text.split(quotedFrom).join(quotedTo);
      next = next.replace(bare, (match, prefix) => prefix + quoteSheet(to) + "!");
      return next;
    };
    for (const sheet of this.model.sheets) {
      for (const cell of Object.values(sheet.cells || {})) {
        if (cell && cell.f) cell.f = replace(cell.f);
      }
      for (const rule of sheet.conditional || []) {
        if (rule.formula) rule.formula = replace(rule.formula);
      }
      for (const validation of sheet.validations || []) {
        for (const key of ["formula1", "formula2"]) {
          if (validation[key]) validation[key] = replace(validation[key]);
        }
      }
    }
    for (const [key, value] of Object.entries(this.model.names || {})) {
      if (typeof value === "string") this.model.names[key] = replace(value);
    }
  }

  moveSheet(from, to, copy = false) {
    const count = this.model.sheets.length;
    const source = clamp(Number(from) || 0, 0, count - 1);
    if (copy) {
      const clone = JSON.parse(JSON.stringify(this.model.sheets[source]));
      clone.name = uniqueSheetName(this.model, this.model.sheets[source].name);
      clone.hidden = false;
      for (const item of [...(clone.charts || []), ...(clone.images || []), ...(clone.textBoxes || [])]) {
        item.id = freshId("copy");
      }
      clone.comments = (clone.comments || []).map((comment, index) => ({ ...comment, id: index + 1 }));
      const at = clamp(Number(to) || 0, 0, count);
      this.mutate("sheet-copy", () => {
        this.model.sheets.splice(at, 0, clone);
        this.model.activeSheet = at;
      }, { layout: false });
    } else {
      const insertAt = this.sheetInsertIndex(source, to);
      if (insertAt === source) return;
      this.mutate("sheet-move", () => moveSheet(this.model, source, insertAt), { layout: false });
    }
    this.grid.setSheet(this.model.activeSheet);
    this.tabsSignature = "";
    this.renderTabs();
    this.renderComments();
    this.syncFormulaBar();
    this.updateStatus();
    this.updateShell();
    this.refreshView();
  }

  duplicateSheet(index = null) {
    const source = Number(index === null ? this.model.activeSheet : index) || 0;
    this.moveSheet(source, source + 1, true);
    const sheet = this.model.sheets[this.model.activeSheet];
    toast(t("Sheet copied."));
    return sheet;
  }

  setSheetHidden(index, hidden = true) {
    const target = clamp(Number(index) || 0, 0, this.model.sheets.length - 1);
    const sheet = this.model.sheets[target];
    if (!sheet) return;
    const visible = this.model.sheets.filter((item) => !item.hidden).length;
    if (hidden && visible <= 1 && !sheet.hidden) {
      toast(t("A workbook must contain at least one visible sheet."));
      return;
    }
    this.mutate("sheet-hide", () => {
      sheet.hidden = !!hidden;
      if (sheet.hidden && this.model.activeSheet === target) {
        const next = this.model.sheets.findIndex((item, position) => position !== target && !item.hidden);
        if (next >= 0) this.model.activeSheet = next;
      }
    }, { layout: false });
    this.grid.setSheet(this.model.activeSheet);
    this.tabsSignature = "";
    this.renderTabs();
    this.renderComments();
    this.syncFormulaBar();
    this.updateStatus();
    this.updateShell();
    this.refreshView();
  }

  setSheetColor(index, color = "") {
    const target = clamp(Number(index) || 0, 0, this.model.sheets.length - 1);
    const sheet = this.model.sheets[target];
    if (!sheet) return;
    this.mutate("sheet-color", () => {
      sheet.tabColor = String(color || "");
    }, { layout: false, recalc: false });
    this.tabsSignature = "";
    this.renderTabs();
  }

  openSheetMenu(x, y, index = null) {
    const target = clamp(Number(index === null ? this.model.activeSheet : index) || 0, 0, this.model.sheets.length - 1);
    const sheet = this.model.sheets[target];
    if (!sheet) return;
    const items = [
      { icon: "newWorkbook", label: t("Insert Sheet"), onClick: () => dlg.insertSheetDialog(this) },
      { icon: "trash", label: t("Delete Sheet"), disabled: this.model.sheets.length <= 1, onClick: () => dlg.deleteSheetDialog(this, target) },
      { label: t("Rename Sheet"), onClick: () => dlg.renameSheetDialog(this, target) },
      { label: t("Move or Copy…"), onClick: () => dlg.moveSheetDialog(this, target) },
      { icon: "copy", label: t("Duplicate Sheet"), onClick: () => this.duplicateSheet(target) },
      { type: "sep" },
      {
        label: t("Hide"),
        disabled: sheet.hidden || this.model.sheets.filter((item) => !item.hidden).length <= 1,
        onClick: () => this.setSheetHidden(target, true),
      },
      {
        label: t("Unhide…"),
        disabled: !this.model.sheets.some((item) => item.hidden),
        onClick: () => dlg.unhideSheetDialog(this),
      },
      { label: t("Tab Color…"), onClick: () => dlg.sheetColorDialog(this, target) },
      { type: "sep" },
      { icon: "protect", label: t("Protect Sheet…"), onClick: () => dlg.protectionDialog(this, "sheet") },
    ];
    openContextMenu(x, y, items);
  }

  stepSheet(direction) {
    const visible = this.model.sheets
      .map((sheet, index) => ({ sheet, index }))
      .filter((entry) => !entry.sheet.hidden);
    if (!visible.length) return;
    const current = visible.findIndex((entry) => entry.index === this.model.activeSheet);
    const next = (current + (direction >= 0 ? 1 : -1) + visible.length) % visible.length;
    this.activateSheet(visible[next].index, { focus: false });
  }

  /* --------------------------------------------------------------- comments */
  sortedComments() {
    const sheet = this.grid?.sheet();
    if (!sheet) return [];
    return (sheet.comments || []).slice().sort((left, right) => {
      const a = parseAddr(left.addr) || { row: 0, col: 0 };
      const b = parseAddr(right.addr) || { row: 0, col: 0 };
      return a.row - b.row || a.col - b.col;
    });
  }

  renderComments() {
    const list = $("#comments-list");
    if (!list) return;
    const comments = this.sortedComments();
    list.replaceChildren();
    const top = $("#btn-comments-top");
    top?.classList.toggle("has-comments", comments.length > 0);
    if (!comments.length) {
      list.append(el("div", { class: "comment-empty", text: t("No comments yet.") }));
      return;
    }
    const active = this.grid.activeAddress();
    for (const comment of comments) {
      const card = el("div", {
        class: "comment-card" + (comment.addr === active ? " active" : ""),
        dataset: { addr: comment.addr },
      });
      const author = String(comment.author || "mTails");
      const initials = author.split(/\s+/).map((part) => part.charAt(0)).join("").slice(0, 2).toUpperCase() || "MT";
      const address = el("button", {
        class: "cc-cell",
        type: "button",
        text: comment.addr,
        title: t("Go To"),
        i18n: false,
      });
      address.addEventListener("click", () => {
        const parsed = parseAddr(comment.addr);
        if (!parsed) return;
        this.grid.selectCell(parsed.row, parsed.col);
        this.grid.scrollToCell(parsed.row, parsed.col, { center: true });
      });
      card.append(el("div", { class: "comment-meta" }, [
        el("span", { class: "comment-avatar", text: initials, i18n: false }),
        el("span", { class: "comment-author", text: author, i18n: false }),
        address,
        el("span", { class: "comment-date", text: formatDate(comment.created), i18n: false }),
      ]));
      card.append(el("div", { class: "comment-text", text: comment.text, i18n: false }));
      const actions = el("div", { class: "comment-actions-row" });
      const edit = el("button", { type: "button", text: t("Edit comment") });
      edit.addEventListener("click", () => {
        const input = $("#comment-input");
        this.commentDraft = comment.addr;
        if (input) {
          input.value = comment.text;
          input.focus();
          input.select();
        }
      });
      const remove = el("button", { type: "button", text: t("Delete Comment") });
      remove.addEventListener("click", () => this.removeComment(comment));
      actions.append(edit, remove);
      card.append(actions);
      card.addEventListener("click", (event) => {
        if (event.target.closest("button")) return;
        const parsed = parseAddr(comment.addr);
        if (!parsed) return;
        this.commentDraft = comment.addr;
        this.grid.selectCell(parsed.row, parsed.col);
        this.grid.scrollToCell(parsed.row, parsed.col, { center: true });
      });
      list.append(card);
    }
  }

  toggleComments(force) {
    const pane = $("#side-pane");
    if (!pane) return false;
    const open = force === undefined ? pane.classList.contains("hidden") : !!force;
    pane.classList.toggle("hidden", !open);
    $("#btn-comments-top")?.classList.toggle("active", open);
    if (open) {
      this.renderComments();
      const drafts = this.commentDraft || this.grid?.activeAddress();
      if (drafts) this.highlightComment(drafts);
    }
    return open;
  }

  highlightComment(address) {
    const list = $("#comments-list");
    if (!list) return;
    let target = null;
    for (const card of list.querySelectorAll(".comment-card")) {
      const match = card.dataset.addr === address;
      card.classList.toggle("active", match);
      if (match) target = card;
    }
    target?.scrollIntoView({ block: "nearest" });
  }

  newComment() {
    const address = this.grid?.activeAddress();
    if (!address) return;
    this.commentDraft = address;
    this.toggleComments(true);
    const input = $("#comment-input");
    if (input) {
      input.value = "";
      input.focus();
    }
    this.highlightComment(address);
  }

  onCommentRequest(address) {
    if (!address) return;
    this.commentDraft = address;
    this.toggleComments(true);
    this.highlightComment(address);
  }

  postComment() {
    const input = $("#comment-input");
    const text = input ? input.value.trim() : "";
    const address = this.commentDraft || this.grid?.activeAddress();
    if (!address) return;
    if (!text) {
      toast(t("Comment cannot be empty."), "error");
      return;
    }
    const sheet = this.grid.sheet();
    const existing = commentsAt(sheet, address)[0];
    this.mutate("comment", () => {
      addComment(sheet, address, text, this.model.properties.author || "mTails");
    }, { layout: true, recalc: false });
    if (input) input.value = "";
    this.commentDraft = null;
    this.renderComments();
    toast(t(existing ? "Comment updated." : "Comment added."));
  }

  removeComment(comment) {
    if (!comment) return;
    const sheet = this.grid.sheet();
    this.mutate("comment", () => {
      deleteComment(sheet, comment.id);
    }, { layout: true, recalc: false });
    this.renderComments();
    toast(t("Comment deleted."));
  }

  cycleComment(direction = 1) {
    const comments = this.sortedComments();
    if (!comments.length) {
      toast(t("No comments yet."));
      return;
    }
    const current = comments.findIndex((comment) => comment.addr === this.grid.activeAddress());
    const step = direction >= 0 ? 1 : -1;
    const index = current < 0
      ? (step > 0 ? 0 : comments.length - 1)
      : (current + step + comments.length) % comments.length;
    const target = comments[index];
    const parsed = parseAddr(target.addr);
    if (parsed) {
      this.grid.selectCell(parsed.row, parsed.col);
      this.grid.scrollToCell(parsed.row, parsed.col, { center: true });
    }
    this.commentDraft = target.addr;
    this.toggleComments(true);
    this.highlightComment(target.addr);
  }

  /* ------------------------------------------------------------------- find */
  openFind(replace = false) {
    const bar = $("#find-bar");
    if (!bar) return;
    bar.classList.remove("hidden");
    this.findState.open = true;
    this.findState.replace = !!replace;
    const row = bar.querySelector(".find-row.secondary");
    row?.classList.toggle("collapsed", !replace);
    $("#find-replace-toggle")?.classList.toggle("on", !!replace);
    const findInput = $("#find-input");
    if (findInput) {
      if (!findInput.value) {
        const info = this.grid.selectionInfo();
        if (info.cell && !info.cell.f && typeof info.cell.v === "string") {
          findInput.value = info.cell.v;
        }
      }
      findInput.focus();
      findInput.select();
    }
    this.runFind();
    if (replace) $("#replace-input")?.focus();
  }

  runFind({ select = false } = {}) {
    if (!this.grid) return;
    const input = $("#find-input");
    const term = input ? input.value : "";
    this.findState.term = term;
    if (!term) {
      this.grid.clearSearch();
      this.findState.count = 0;
      this.findState.matches = [];
      this.setFindCount(0);
      return;
    }
    const result = this.grid.search(term, { select, forward: true });
    this.findState.count = result.count;
    this.findState.matches = result.matches || [];
    this.setFindCount(result.count, result.address);
  }

  findStep(direction = 1) {
    if (!this.grid) return;
    const input = $("#find-input");
    const term = (input && input.value) || this.findState.term || "";
    if (!term) return;
    const result = this.grid.search(term, { select: true, forward: direction >= 0 });
    this.findState.term = term;
    this.findState.count = result.count;
    this.findState.matches = result.matches || [];
    this.setFindCount(result.count, result.address);
    if (!result.count) toast(t("No results"));
  }

  setFindCount(count) {
    const node = $("#find-count");
    if (!node) return;
    if (!count) node.textContent = t("0 results");
    else node.textContent = tp(count, "{count} result", "{count} results", { count });
  }

  replaceOne() {
    if (!this.grid) return;
    const term = this.findState.term || $("#find-input")?.value || "";
    const replacement = $("#replace-input")?.value ?? "";
    if (!term) {
      toast(t("Enter the text to find."));
      return;
    }
    const sheet = this.grid.sheet();
    const cell = cellAt(sheet, this.grid.active.row, this.grid.active.col);
    const raw = rawEditText(cell);
    const pattern = new RegExp(escapeRegExp(term), "i");
    if (raw && pattern.test(raw)) {
      const next = raw.replace(new RegExp(escapeRegExp(term), "gi"), () => replacement);
      if (next !== raw) {
        this.grid.setActiveValue(next, { move: "none" });
        toast(tp(1, "Replaced {count} occurrence.", "Replaced {count} occurrences.", { count: 1 }));
      }
    }
    this.runFind();
    this.findStep(1);
  }

  replaceAll() {
    if (!this.grid) return;
    const term = this.findState.term || $("#find-input")?.value || "";
    const replacement = $("#replace-input")?.value ?? "";
    if (!term) {
      toast(t("Enter the text to find."));
      return;
    }
    const sheet = this.grid.sheet();
    const result = this.grid.search(term, { select: false });
    const matches = (result.matches || []).slice();
    if (!matches.length) {
      this.setFindCount(0);
      toast(t("No results"));
      return;
    }
    let replaced = 0;
    this.mutate("replace", () => {
      for (const match of matches) {
        const cell = cellAt(sheet, match.row, match.col);
        const raw = rawEditText(cell);
        if (!raw) continue;
        const next = raw.replace(new RegExp(escapeRegExp(term), "gi"), () => replacement);
        if (next === raw) continue;
        if (this.grid.applyText(match.address, next)) replaced += 1;
      }
    }, { layout: false });
    this.runFind();
    toast(tp(replaced, "Replaced {count} occurrence.", "Replaced {count} occurrences.", { count: replaced }));
  }

  closeFind() {
    const bar = $("#find-bar");
    bar?.classList.add("hidden");
    this.findState.open = false;
    this.grid?.clearSearch();
    this.grid?.els?.scroller.focus({ preventScroll: true });
  }

  /* ------------------------------------------------------------------ views */
  setView(view) {
    const next = ["normal", "pagebreak", "layout"].includes(view) ? view : "normal";
    this.state.view = next;
    this.refreshView();
    this.updateShell();
  }

  ensureLayoutLayer() {
    let layer = $("#page-layout-layer");
    if (layer) return layer;
    const area = $("#sheet-area");
    if (!area) return null;
    layer = el("div", { id: "page-layout-layer", class: "hidden", i18n: false });
    area.append(layer);
    return layer;
  }

  renderLayoutView() {
    const layer = this.ensureLayoutLayer();
    if (!layer || !this.grid) return;
    layer.replaceChildren();
    const rendered = renderPrintPages(this.model, {
      root: layer,
      allSheets: false,
      sheetIndex: this.grid.sheetIndex,
    });
    if (!rendered || rendered.ok === false) {
      layer.append(el("div", { class: "palette-empty", text: t("Nothing to print.") }));
      return;
    }
    const ptToPx = 96 / 72;
    const available = Math.max(320, layer.clientWidth - 44);
    for (const page of Array.from(layer.querySelectorAll(".print-page"))) {
      const width = (parseFloat(page.style.width) || 595) * ptToPx;
      const height = (parseFloat(page.style.height) || 842) * ptToPx;
      const factor = Math.min(1, Math.max(0.2, available / Math.max(1, width)));
      const holder = el("div", {
        class: "layout-holder",
        i18n: false,
        style: {
          position: "relative",
          width: Math.round(width * factor) + "px",
          height: Math.round(height * factor) + "px",
        },
      });
      page.classList.add("layout-page");
      page.style.position = "absolute";
      page.style.left = "0";
      page.style.top = "0";
      page.style.transformOrigin = "top left";
      page.style.transform = `scale(${factor})`;
      holder.append(page);
      layer.append(holder);
    }
    layer.prepend(el("div", { class: "layout-hint", text: t("Page Layout"), i18n: false }));
  }

  scheduleViewRefresh() {
    clearTimeout(this.layoutTimer);
    this.layoutTimer = setTimeout(() => {
      if (this.state.view === "layout") this.renderLayoutView();
      else if (this.state.view === "pagebreak") {
        this.grid.pageBreaks = computePageBreaks(this.model, this.grid.sheetIndex);
        this.grid.renderNow();
      }
    }, 220);
  }

  refreshView() {
    if (!this.grid) return;
    const view = this.state.view || "normal";
    const layer = this.ensureLayoutLayer();
    document.body.dataset.view = view;
    if (layer) layer.classList.toggle("hidden", view !== "layout");
    if (view === "layout") {
      this.grid.pageBreaks = null;
      this.renderLayoutView();
    } else if (view === "pagebreak") {
      this.grid.pageBreaks = computePageBreaks(this.model, this.grid.sheetIndex);
      this.grid.renderNow();
    } else {
      this.grid.pageBreaks = null;
      layer?.replaceChildren();
      this.grid.renderNow();
    }
    this.ribbon?.setActive("view.normal", view === "normal");
    this.ribbon?.setActive("view.pagebreak", view === "pagebreak");
    this.ribbon?.setActive("view.pageLayout", view === "layout");
    $("#sb-normal")?.classList.toggle("active", view === "normal");
    $("#sb-pagebreak")?.classList.toggle("active", view === "pagebreak");
  }

  /* ------------------------------------------------------------- navigation */
  selectReference(text) {
    const raw = String(text || "").trim();
    if (!raw || !this.grid) return false;
    const parsed = splitSheetRef(raw);
    let index = this.grid.sheetIndex;
    if (parsed.sheet) {
      const found = this.model.sheets.findIndex(
        (sheet) => sheet.name.toLowerCase() === parsed.sheet.toLowerCase(),
      );
      if (found < 0) return false;
      index = found;
      if (found !== this.grid.sheetIndex) this.activateSheet(found, { focus: false });
    }
    const range = parseRange(parsed.ref);
    if (!range) return false;
    this.grid.selectRange(range, { reveal: true });
    this.grid.els?.scroller.focus({ preventScroll: true });
    return true;
  }

  goToSpecial(kind) {
    if (!this.grid) return;
    const sheet = this.grid.sheet();
    const used = usedRange(sheet) || { r1: 1, c1: 1, r2: 1, c2: 1 };
    const active = { ...this.grid.active };
    const filled = (row, col) => {
      const cell = cellAt(sheet, row, col);
      return !!(cell && (cell.f || (cell.v !== undefined && cell.v !== null && cell.v !== "")));
    };
    let found = null;
    if (kind === "comments") {
      const first = this.sortedComments().find((comment) => parseAddr(comment.addr));
      if (first) found = parseAddr(first.addr);
    } else if (kind === "lastCell") {
      found = { row: used.r2, col: used.c2 };
    } else if (kind === "currentRegion") {
      let r1 = active.row;
      let r2 = active.row;
      let c1 = active.col;
      let c2 = active.col;
      if (filled(active.row, active.col)) {
        let changed = true;
        let guard = 0;
        while (changed && guard < 5000) {
          changed = false;
          guard += 1;
          if (r1 > 1) {
            let hit = false;
            for (let col = c1; col <= c2 && !hit; col += 1) hit = filled(r1 - 1, col);
            if (hit) { r1 -= 1; changed = true; }
          }
          if (r2 < used.r2) {
            let hit = false;
            for (let col = c1; col <= c2 && !hit; col += 1) hit = filled(r2 + 1, col);
            if (hit) { r2 += 1; changed = true; }
          }
          if (c1 > 1) {
            let hit = false;
            for (let row = r1; row <= r2 && !hit; row += 1) hit = filled(row, c1 - 1);
            if (hit) { c1 -= 1; changed = true; }
          }
          if (c2 < used.c2) {
            let hit = false;
            for (let row = r1; row <= r2 && !hit; row += 1) hit = filled(row, c2 + 1);
            if (hit) { c2 += 1; changed = true; }
          }
        }
      }
      this.grid.selectRange({ r1, c1, r2, c2 }, { reveal: true });
      return;
    } else {
      for (let row = used.r1; row <= used.r2 && !found; row += 1) {
        for (let col = used.c1; col <= used.c2; col += 1) {
          const cell = cellAt(sheet, row, col);
          const isFormula = !!cell?.f;
          const hasValue = !!(cell && cell.v !== undefined && cell.v !== null && cell.v !== "");
          if (kind === "formulas" && isFormula) { found = { row, col }; break; }
          if (kind === "constants" && !isFormula && hasValue) { found = { row, col }; break; }
          if (kind === "blanks" && !hasValue && !cell?.s && !cell?.h) { found = { row, col }; break; }
        }
      }
    }
    if (!found) {
      toast(t("No cells found."));
      return;
    }
    this.grid.selectCell(found.row, found.col);
    this.grid.scrollToCell(found.row, found.col, { center: true });
  }

  traceArrows(rows, mode = "precedents") {
    if (!this.grid) return;
    const active = { row: this.grid.active.row, col: this.grid.active.col };
    const links = [];
    for (const row of rows || []) {
      const range = parseRange(row.label || "");
      if (!range) continue;
      if (mode === "dependents") {
        links.push({
          sheet: row.sheet,
          row1: active.row,
          col1: active.col,
          row2: range.r1,
          col2: range.c1,
        });
      } else {
        links.push({
          sheet: row.sheet,
          row1: range.r1,
          col1: range.c1,
          row2: active.row,
          col2: active.col,
        });
      }
    }
    this.grid.auditLinks = links;
    this.grid.renderNow();
  }

  clearAudit() {
    if (!this.grid) return;
    this.grid.auditLinks = [];
    this.grid.renderNow();
  }

  onHyperlink(url) {
    const link = String(url || "").trim();
    if (!link) return;
    if (link.startsWith("#")) {
      this.selectReference(link.slice(1));
      return;
    }
    if (/^(https?:|mailto:)/i.test(link)) {
      confirmDialog(t("Open link?"), link, t("Open")).then((ok) => {
        if (ok) file.openExternal(link);
      });
      return;
    }
    this.selectReference(link);
  }

  onShortcut(kind) {
    switch (kind) {
      case "format-cells":
        dlg.formatCellsDialog(this);
        break;
      case "format-b":
        this.execute("bold");
        break;
      case "format-i":
        this.execute("italic");
        break;
      case "format-u":
        this.execute("underline");
        break;
      case "hyperlink":
        dlg.hyperlinkDialog(this);
        break;
      case "autosum":
        this.execute("autosum:SUM");
        break;
      default:
        break;
    }
  }

  onContextMenu(x, y) {
    const items = [
      { icon: "cut", label: t("Cut"), shortcut: "Ctrl+X", onClick: () => this.execute("cut") },
      { icon: "copy", label: t("Copy"), shortcut: "Ctrl+C", onClick: () => this.execute("copy") },
      { icon: "paste", label: t("Paste"), shortcut: "Ctrl+V", onClick: () => this.execute("paste") },
      { icon: "pasteSpecial", label: t("Paste Special…"), onClick: () => this.execute("pasteSpecial") },
      { type: "sep" },
      { icon: "insertRow", label: t("Insert"), onClick: () => this.execute("insert.rows") },
      { icon: "deleteRow", label: t("Delete"), onClick: () => this.execute("delete.rows") },
      { type: "sep" },
      { icon: "formatCells", label: t("Format Cells…"), onClick: () => dlg.formatCellsDialog(this) },
      { icon: "link", label: t("Link…"), onClick: () => dlg.hyperlinkDialog(this) },
      { icon: "comment", label: t("New Comment"), onClick: () => this.newComment() },
      { icon: "trash", label: t("Delete Comment"), onClick: () => this.cycleComment(1) },
      { type: "sep" },
      { icon: "clearContents", label: t("Clear Contents"), shortcut: "Delete", onClick: () => this.execute("clear:contents") },
      { icon: "find", label: t("Find…"), shortcut: "Ctrl+F", onClick: () => this.openFind(false) },
    ];
    openContextMenu(x, y, items);
  }

  /* --------------------------------------------------------------- overlays */
  onOverlayActivate(item, kind) {
    if (!item) return;
    if (kind === "chart") dlg.chartDialog(this, item.type || "column", item);
    else if (kind === "text") dlg.textBoxDialog(this, item);
    else if (kind === "image") this.editImage(item);
  }

  onOverlayContextMenu(x, y, item, kind) {
    if (!item) return;
    const items = [
      {
        icon: kind === "chart" ? "chartDialog" : kind === "text" ? "textBox" : "picture",
        label: t("Edit"),
        onClick: () => this.onOverlayActivate(item, kind),
      },
      { type: "sep" },
      { icon: "trash", label: t("Delete"), onClick: () => this.deleteOverlay(item, kind) },
    ];
    openContextMenu(x, y, items);
  }

  deleteOverlay(item, kind) {
    if (!item) return;
    const sheet = this.grid.sheet();
    this.mutate("drawing", () => {
      if (kind === "chart") sheet.charts = (sheet.charts || []).filter((entry) => entry.id !== item.id);
      else if (kind === "image") sheet.images = (sheet.images || []).filter((entry) => entry.id !== item.id);
      else sheet.textBoxes = (sheet.textBoxes || []).filter((entry) => entry.id !== item.id);
      if (this.grid.selectedOverlay === item.id) this.grid.selectedOverlay = null;
    }, { layout: false, recalc: false });
  }

  editImage(item) {
    const width = numberInput(item.w || 200, { min: "24", max: "2000" });
    const height = numberInput(item.h || 140, { min: "24", max: "2000" });
    showDialog({
      title: t("Picture"),
      size: "narrow",
      body: el("div", {}, [
        el("div", { class: "dialog-grid" }, [field(t("Width"), width), field(t("Height"), height)]),
        el("p", { class: "dialog-note", text: t("Drag the picture to move it, or use the corner handle to resize it.") }),
      ]),
      buttons: [
        { label: t("Delete"), kind: "ghost", left: true, onClick: () => this.deleteOverlay(item, "image") },
        { label: t("Cancel"), kind: "ghost" },
        {
          label: t("OK"),
          kind: "primary",
          onClick: () => {
            this.mutate("drawing", () => {
              item.w = clamp(Number(width.value) || item.w || 200, 24, 2000);
              item.h = clamp(Number(height.value) || item.h || 140, 24, 2000);
            }, { layout: false, recalc: false });
          },
        },
      ],
    });
  }

  overlayAnchor() {
    return {
      x: Math.round((this.grid.colPrefix[this.grid.active.col - 1] || 0) / this.grid.scale) + 8,
      y: Math.round((this.grid.rowPrefix[this.grid.active.row - 1] || 0) / this.grid.scale) + 8,
    };
  }

  addImage(src, width = 200, height = 140) {
    if (!src) return null;
    const sheet = this.grid.sheet();
    const anchor = this.overlayAnchor();
    const item = {
      id: freshId("img"),
      src,
      x: anchor.x,
      y: anchor.y,
      w: clamp(Number(width) || 200, 24, 2000),
      h: clamp(Number(height) || 140, 24, 2000),
    };
    this.mutate("drawing", () => {
      if (!Array.isArray(sheet.images)) sheet.images = [];
      sheet.images.push(item);
    }, { layout: false, recalc: false });
    toast(t("Picture inserted."));
    return item;
  }

  addTextBox(text, { w = 220, h = 90 } = {}) {
    const label = String(text || "").trim();
    if (!label) return null;
    const sheet = this.grid.sheet();
    const anchor = this.overlayAnchor();
    const item = {
      id: freshId("box"),
      text: label,
      x: anchor.x,
      y: anchor.y,
      w: clamp(Number(w) || 220, 40, 1200),
      h: clamp(Number(h) || 90, 24, 800),
      size: 12,
    };
    this.mutate("textbox", () => {
      if (!Array.isArray(sheet.textBoxes)) sheet.textBoxes = [];
      sheet.textBoxes.push(item);
    }, { layout: false, recalc: false });
    return item;
  }

  /* ---------------------------------------------------------- cell tooling */
  insertFunction(name) {
    const fn = String(name || "").trim().toUpperCase();
    if (!fn) return;
    closeAllDialogs();
    this.recentFunctionsList = this.recentFunctionsList || [];
    this.recentFunctionsList = [fn, ...this.recentFunctionsList.filter((entry) => entry !== fn)].slice(0, 8);
    this.grid.beginEdit("=" + fn + "(", { selectAll: false });
  }

  recentFunctions() {
    return this.recentFunctionsList && this.recentFunctionsList.length
      ? this.recentFunctionsList.slice()
      : ["SUM", "AVERAGE", "IF", "COUNT", "MAX", "MIN", "VLOOKUP", "TEXT", "CONCAT"];
  }

  functionsByCategory(category) {
    return dlg.functionsByCategory(category);
  }

  functionArity(name) {
    try {
      return formulaArity(name);
    } catch {
      return { min: 0, max: 255 };
    }
  }

  definedNames() {
    return Object.keys(this.model.names || {});
  }

  insertName(name) {
    const key = String(name || "");
    const reference = this.model.names?.[key];
    if (!reference) return;
    const input = $("#formula-input");
    if (input && document.activeElement === input) {
      const start = input.selectionStart ?? input.value.length;
      const end = input.selectionEnd ?? start;
      input.value = input.value.slice(0, start) + key + input.value.slice(end);
      const caret = start + key.length;
      input.setSelectionRange(caret, caret);
      return;
    }
    this.selectReference(reference);
  }

  insertIntoCell(text) {
    const value = String(text ?? "");
    if (!value) return;
    const active = document.activeElement;
    if (active && (active.tagName === "INPUT" || active.tagName === "TEXTAREA")) {
      const inEditor = active.id === "formula-input" || active.closest?.(".cell-editor");
      if (inEditor) {
        const start = active.selectionStart ?? active.value.length;
        const end = active.selectionEnd ?? start;
        active.value = active.value.slice(0, start) + value + active.value.slice(end);
        const caret = start + value.length;
        active.setSelectionRange?.(caret, caret);
        return;
      }
    }
    if (!this.grid) return;
    const cell = cellAt(this.grid.sheet(), this.grid.active.row, this.grid.active.col);
    this.grid.setActiveValue(rawEditText(cell) + value, { move: "none" });
  }

  pasteSpecial(mode) {
    this.commands.paste(mode || "all");
  }

  setPainterCursor(active) {
    this.painterActive = !!active;
    document.body.classList.toggle("painter-active", !!active);
  }

  /* ------------------------------------------------------------ calculation */
  calcMode() {
    return this.model.settings?.calcMode === "manual" ? "manual" : "auto";
  }

  setCalcMode(mode) {
    const value = mode === "manual" ? "manual" : "auto";
    if (this.calcMode() !== value) {
      this.mutate("calc-mode", () => {
        this.model.settings.calcMode = value;
      }, { layout: false, recalc: false });
    }
    if (value === "auto") {
      this.grid.recalculate(true);
      this.grid.renderNow();
      this.grid.notifySelection();
      toast(t("Calculated."));
    } else {
      toast(t("Manual calculation is on. Press F9 to recalculate."));
    }
  }

  calculateNow() {
    if (!this.grid) return;
    this.grid.recalculate(true);
    this.grid.renderNow();
    this.grid.notifySelection();
    toast(t("Calculated."));
  }

  calculateSheet() {
    this.calculateNow();
  }

  toggleShowFormulas() {
    const next = !this.model.settings?.showFormulas;
    this.mutate("view", () => {
      this.model.settings.showFormulas = next;
    }, { layout: true, recalc: false });
    this.ribbon?.setActive("showFormulas", next);
  }

  toggleFormulaBar() {
    const bar = $("#formula-bar");
    if (!bar) return;
    const collapsed = !bar.classList.contains("collapsed");
    this.setFormulaBar(!collapsed);
  }

  setFormulaBar(visible) {
    const bar = $("#formula-bar");
    const show = visible !== false;
    bar?.classList.toggle("collapsed", !show);
    this.settings.showFormulaBar = show;
    const checkbox = $("#opt-formula-bar");
    if (checkbox) checkbox.checked = show;
    file.setSetting(this, "showFormulaBar", show);
  }

  setSpelling(enabled, language) {
    this.settings.spellcheck = !!enabled;
    if (language) this.settings.proofingLanguage = String(language);
    if (this.grid) this.grid.spellcheck = !!enabled;
    file.setSetting(this, "spellcheck", !!enabled);
    file.setSetting(this, "proofingLanguage", this.settings.proofingLanguage || "en-US");
  }

  /* ------------------------------------------------------------ cell styles */
  customStyles() {
    return (this.settings.customStyles || []).filter((preset) => preset && preset.id && preset.label);
  }

  addCustomStyle(label, style) {
    const base = String(label || "").trim();
    if (!base) return null;
    const stem = "custom-" + (base.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "") || "style");
    const used = new Set(this.customStyles().map((preset) => preset.id));
    let id = stem;
    let suffix = 1;
    while (used.has(id)) {
      suffix += 1;
      id = stem + "-" + suffix;
    }
    const preset = { id, label: base, style: { ...(style || {}) } };
    this.settings.customStyles = [...this.customStyles(), preset];
    file.setSetting(this, "customStyles", this.settings.customStyles);
    return preset;
  }

  /* ------------------------------------------------------------ open files */
  async openFile(path = "") {
    const opened = await file.openFile(this, path);
    if (opened) this.closeBackstage();
    return opened;
  }

  /* -------------------------------------------------------------- backstage */
  openBackstage(panel = "new") {
    const node = $("#backstage");
    if (!node) return;
    this.closePalette();
    closeMenus();
    node.classList.remove("hidden");
    node.setAttribute("aria-hidden", "false");
    document.body.classList.add("backstage-open");
    this.showBsPanel(panel);
    if (this.state.backstage === "open") {
      file.refreshRecent().then((list) => {
        this.recent = Array.isArray(list) ? list : [];
        if (this.state.backstage === "open") this.renderRecent();
      }).catch(() => {});
    }
  }

  closeBackstage() {
    const node = $("#backstage");
    if (node) {
      node.classList.add("hidden");
      node.setAttribute("aria-hidden", "true");
    }
    document.body.classList.remove("backstage-open");
    this.state.backstage = null;
    if (this.booted) this.grid?.els?.scroller?.focus({ preventScroll: true });
  }

  showBsPanel(panel = "new") {
    const name = PANELS.includes(panel) ? panel : "new";
    this.state.backstage = name;
    $$("#bs-nav button[data-bs]").forEach((button) => {
      button.classList.toggle("active", button.dataset.bs === name);
    });
    $$("#backstage section[data-panel]").forEach((section) => {
      section.hidden = section.dataset.panel !== name;
    });
    const main = $("#bs-main");
    if (main) main.scrollTop = 0;
    if (name === "new") this.renderTemplates();
    else if (name === "open") { this.renderRecent(); this.renderRecovery(); }
    else if (name === "info") this.renderInfo();
    else if (name === "save") this.renderSaveCard();
    else if (name === "print") this.renderPrintPanel();
    else if (name === "options") this.renderOptions();
  }

  renderTemplates() {
    const grid = $("#template-grid");
    if (!grid) return;
    grid.replaceChildren();
    for (const template of TEMPLATES) {
      const thumb = el("div", { class: "template-thumb " + (template.accent || "accent") });
      thumb.append(el("div", { class: "tt-title", text: template.thumb?.title || template.label }));
      if (template.thumb) {
        const table = el("table");
        if (Array.isArray(template.thumb.head) && template.thumb.head.length) {
          table.append(el("thead", {}, [
            el("tr", {}, template.thumb.head.map((value) => el("th", { text: String(value) }))),
          ]));
        }
        if (Array.isArray(template.thumb.rows) && template.thumb.rows.length) {
          table.append(el("tbody", {}, template.thumb.rows.map((row) => el(
            "tr",
            {},
            (row || []).map((value) => el("td", { text: String(value) })),
          ))));
        }
        thumb.append(table);
      }
      const card = el("button", { class: "template-card", type: "button" }, [
        thumb,
        el("strong", { text: template.label }),
        el("span", { text: template.description || "" }),
      ]);
      card.addEventListener("click", () => { this.newWorkbook(template.id); });
      grid.append(card);
    }
  }

  renderRecent() {
    const list = $("#recent-list");
    if (!list) return;
    list.replaceChildren();
    const items = Array.isArray(this.recent) ? this.recent : [];
    if (!items.length) {
      list.append(el("p", {
        class: "muted",
        text: "No recent workbooks yet. Open a file to add it to this list.",
      }));
      return;
    }
    for (const item of items.slice(0, 12)) {
      const path = String(item.path || "");
      const name = String(item.name || path.split(/[\\/]/).pop() || path);
      const ext = (/[.]([a-z0-9]+)$/i.exec(name) || [])[1] || "xlsx";
      const button = el("button", {
        class: "recent-item" + (item.exists === false ? " missing" : ""),
        type: "button",
        title: path,
      }, [
        el("div", { class: "recent-icon " + ext.slice(0, 4).toLowerCase(), i18n: false }, [
          el("span", { text: ext.slice(0, 4).toUpperCase(), i18n: false }),
        ]),
        el("div", { class: "recent-meta" }, [
          el("div", { class: "recent-name", text: name, i18n: false }),
          el("div", { class: "recent-path", text: path, i18n: false }),
        ]),
      ]);
      button.addEventListener("click", () => { this.openFile(path); });
      list.append(button);
    }
  }

  renderRecovery() {
    const box = $("#recovery-box");
    const list = $("#recovery-list");
    if (!box || !list) return;
    const items = Array.isArray(this.autosaves) ? this.autosaves : [];
    box.classList.toggle("hidden", items.length === 0);
    list.replaceChildren();
    for (const item of items) {
      const card = el("div", { class: "recovery-card" }, [
        el("div", { class: "rc-meta" }, [
          el("div", { class: "rc-name", text: item.name || "Untitled workbook" }),
          el("div", {
            class: "rc-info",
            i18n: false,
            text: [item.source || t("Unsaved workbook"), formatDate((item.savedAt || 0) * 1000)]
              .filter(Boolean).join(" \u00B7 "),
          }),
        ]),
        el("button", {
          class: "btn primary",
          type: "button",
          text: "Recover",
          onClick: () => this.recoverAutosave(item),
        }),
      ]);
      list.append(card);
    }
  }

  async recoverAutosave(item) {
    if (!item) return;
    if (!(await file.confirmSaveIfNeeded(this))) return;
    const bridge = file.getBridge();
    if (!bridge) {
      toast(t("The file bridge is not available."), "error");
      return;
    }
    const result = await new Promise((resolve) => {
      let settled = false;
      const finish = (value) => {
        if (settled) return;
        settled = true;
        resolve(value);
      };
      setTimeout(() => finish(null), 20000);
      try {
        bridge.loadAutosave(item.id, (value) => {
          try {
            finish(typeof value === "string" ? JSON.parse(value) : value);
          } catch {
            finish(null);
          }
        });
      } catch {
        finish(null);
      }
    });
    if (!result || !result.ok) {
      toast(t(result?.error || "The recovery file could not be read."), "error");
      return;
    }
    file.applyModel(this, result.model, {
      title: result.name || item.name,
      path: result.source || "",
      format: file.formatForPath(result.source || "", "xlsx"),
      docId: item.id,
      dirty: true,
    });
    this.autosaves = (this.autosaves || []).filter((entry) => entry.id !== item.id);
    this.renderRecovery();
    this.closeBackstage();
    toast(t("Recovered workbook restored."), "success");
  }

  renderInfo() {
    const stats = workbookStats(this.model);
    const setText = (id, value) => {
      const node = document.getElementById(id);
      if (node) node.textContent = String(value);
    };
    setText("info-name", this.state.title || "Workbook1");
    setText("info-location", this.state.path || t("Not saved yet"));
    setText("info-sheets", stats.sheets);
    setText("info-cells", stats.cells);
    setText("info-formulas", stats.formulas);
    setText("info-modified", formatDate(this.state.modifiedAt));
    const props = this.model.properties || {};
    for (const [id, value] of [
      ["prop-title", props.title], ["prop-author", props.author],
      ["prop-subject", props.subject], ["prop-keywords", props.keywords],
    ]) {
      const input = document.getElementById(id);
      if (input) input.value = value || "";
    }
    const protect = $("#info-protect");
    if (protect) {
      protect.textContent = this.model.protection?.protected
        ? t("Unprotect workbook")
        : t("Protect workbook");
    }
  }

  saveProperties() {
    const read = (id) => String(document.getElementById(id)?.value || "").trim();
    const title = read("prop-title");
    this.mutate("properties", () => {
      this.model.properties.title = title;
      this.model.properties.author = read("prop-author");
      this.model.properties.subject = read("prop-subject");
      this.model.properties.keywords = read("prop-keywords");
      if (title) {
        this.model.title = title;
        this.state.title = title;
      }
    }, { layout: false, recalc: false });
    this.lastShell = "";
    this.syncShell();
    this.renderInfo();
    this.renderSaveCard();
    toast(t("Properties saved."));
  }

  renderSaveCard() {
    const format = this.state.format || "xlsx";
    const name = file.suggestedName(this.state.title || "Workbook1", format);
    const title = $("#save-card-title");
    if (title) title.textContent = name;
    const path = $("#save-card-path");
    if (path) {
      path.textContent = this.state.path || t("Choose a folder and file name.");
    }
  }

  renderPrintPanel() {
    const sheet = this.model.sheets[this.model.activeSheet];
    if (!sheet) return;
    const name = $("#print-doc-name");
    if (name) name.textContent = this.state.title || "Workbook1";
    const summary = $("#print-summary");
    if (summary) summary.textContent = printSummary(this.model, this.model.activeSheet);
    const orientation = $("#print-orientation");
    if (orientation) orientation.value = sheet.page?.orientation === "landscape" ? "landscape" : "portrait";
    const size = $("#print-size");
    if (size) size.value = sheet.page?.paper || sheet.page?.size || "A4";
    renderPreview($("#print-preview"), this.model, { sheetIndex: this.model.activeSheet });
  }

  renderOptions() {
    const check = (id, value) => {
      const node = document.getElementById(id);
      if (node) node.checked = !!value;
    };
    const pick = (id, value) => {
      const node = document.getElementById(id);
      if (node) node.value = String(value);
    };
    check("opt-autosave", this.settings.autosave !== false);
    check("opt-autorecover", this.settings.autorecover !== false);
    pick("opt-autosave-interval", this.settings.autosaveInterval || 120);
    pick("opt-ui-language", this.settings.uiLanguage || "system");
    pick("opt-theme", this.settings.theme === "dark" ? "dark" : "light");
    check("opt-gridlines", this.settings.showGridlines !== false);
    check("opt-headings", this.settings.showHeadings !== false);
    check("opt-formula-bar", this.settings.showFormulaBar !== false);
    pick("opt-default-sheets", this.settings.defaultSheets || 1);
    const fontSelect = $("#opt-font");
    if (fontSelect) {
      const fonts = this.fonts.length ? this.fonts.slice() : [this.defaultFont, "Calibri", "Arial"];
      if (!fonts.includes(this.defaultFont)) fonts.unshift(this.defaultFont);
      fontSelect.replaceChildren(...fonts.map((name) => el("option", {
        value: name, text: name, i18n: false,
      })));
      fontSelect.value = this.defaultFont;
    }
    const sizeSelect = $("#opt-font-size");
    if (sizeSelect) {
      const sizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36];
      if (!sizes.includes(this.defaultSize)) sizes.push(this.defaultSize);
      sizes.sort((left, right) => left - right);
      sizeSelect.replaceChildren(...sizes.map((value) => el("option", {
        value: String(value), text: String(value), i18n: false,
      })));
      sizeSelect.value = String(this.defaultSize);
    }
  }

  saveOptions() {
    const value = (id) => String(document.getElementById(id)?.value ?? "");
    const checked = (id) => !!document.getElementById(id)?.checked;
    const interval = clamp(Math.round(Number(value("opt-autosave-interval")) || 120), 15, 3600);
    const language = LANGUAGES.includes(value("opt-ui-language")) ? value("opt-ui-language") : "system";
    const theme = value("opt-theme") === "dark" ? "dark" : "light";
    const font = String(value("opt-font") || this.defaultFont || "Calibri");
    const size = clamp(Math.round(Number(value("opt-font-size")) || 11), 6, 72);
    const gridlines = checked("opt-gridlines");
    const headings = checked("opt-headings");
    const formulaBar = checked("opt-formula-bar");
    const sheets = clamp(Math.round(Number(value("opt-default-sheets")) || 1), 1, 24);
    this.settings.autosave = checked("opt-autosave");
    this.settings.autorecover = checked("opt-autorecover");
    this.settings.autosaveInterval = interval;
    this.settings.defaultSheets = sheets;
    this.settings.showGridlines = gridlines;
    this.settings.showHeadings = headings;
    this.settings.defaultFont = font;
    this.settings.defaultSize = size;
    this.defaultFont = font;
    this.defaultSize = size;
    this.grid.defaultFont = font;
    this.grid.defaultSize = size;
    for (const key of [
      "autosave", "autorecover", "autosaveInterval", "defaultSheets",
      "showGridlines", "showHeadings", "defaultFont", "defaultSize",
    ]) {
      file.setSetting(this, key, this.settings[key]);
    }
    this.applyTheme(theme);
    file.setSetting(this, "theme", theme);
    this.setFormulaBar(formulaBar);
    this.applySheetDisplay(gridlines, headings);
    if (this.settings.uiLanguage !== language) this.setLanguage(language);
    file.startAutosave(this);
    this.renderOptions();
    toast(t("Options saved."));
  }

  applySheetDisplay(gridlines, headings) {
    const sheet = this.grid?.sheet();
    if (!sheet) return;
    if (sheet.showGridlines === gridlines && sheet.showHeadings === headings) return;
    this.mutate("display", () => {
      for (const item of this.model.sheets) {
        item.showGridlines = gridlines;
        item.showHeadings = headings;
      }
    }, { layout: false, recalc: false });
    this.ribbon?.update({
      ...this.grid.selectionInfo(),
      showGridlines: gridlines,
      showHeadings: headings,
    });
  }

  updatePageSetting(key, value) {
    const sheet = this.model.sheets[this.model.activeSheet];
    if (!sheet) return;
    this.mutate("page-setup", () => {
      sheet.page = sheet.page || {};
      if (key === "orientation") sheet.page.orientation = value === "landscape" ? "landscape" : "portrait";
      else if (key === "size" || key === "paper") sheet.page.paper = String(value || "A4");
      else sheet.page[key] = value;
    }, { layout: false, recalc: false });
    this.renderPrintPanel();
    this.scheduleViewRefresh();
  }

  async closeDocument() {
    const ok = await file.confirmSaveIfNeeded(this);
    if (!ok) return;
    this.closeBackstage();
    file.closeWindow();
  }

  showShare() {
    const path = this.state.path || "";
    const body = el("div", {}, [
      el("p", {
        class: "dialog-note",
        text: path
          ? "Share this workbook by sending a copy or by pointing to the folder below."
          : "Save the workbook to share its location.",
      }),
      el("input", {
        type: "text",
        value: path || this.state.title || "Workbook1",
        readonly: "readonly",
        i18n: false,
        style: { width: "100%", padding: "6px 8px" },
      }),
    ]);
    showDialog({
      title: "Share",
      size: "narrow",
      body,
      buttons: [
        {
          label: "Copy Path",
          kind: "primary",
          onClick: () => {
            this.copyText(path || this.state.title || "Workbook1");
          },
        },
        { label: "Close", kind: "ghost" },
      ],
    });
  }

  async copyText(text) {
    const value = String(text ?? "");
    try {
      await navigator.clipboard.writeText(value);
    } catch {
      const area = el("textarea", { i18n: false });
      area.value = value;
      area.style.cssText = "position:fixed;left:-9999px;top:0;opacity:0";
      document.body.append(area);
      area.select();
      try { document.execCommand("copy"); } catch { /* clipboard unavailable */ }
      area.remove();
    }
    toast(t("Copied."), "success");
  }

  /* --------------------------------------------------------------- settings */
  syncHistoryButtons() {
    this.ribbon?.setDisabled("undo", this.undoStack.length === 0);
    this.ribbon?.setDisabled("redo", this.redoStack.length === 0);
  }

  applyTheme(theme) {
    const value = theme === "dark" ? "dark" : "light";
    this.settings.theme = value;
    document.body.dataset.theme = value;
    document.documentElement.dataset.theme = value;
    const select = $("#opt-theme");
    if (select) select.value = value;
    this.ribbon?.setActive("theme.toggle", value === "dark");
  }

  toggleTheme() {
    const next = this.settings.theme === "dark" ? "light" : "dark";
    this.applyTheme(next);
    file.setSetting(this, "theme", next);
    toast(t("Theme changed."));
  }

  setLanguage(language) {
    const value = LANGUAGES.includes(language) ? language : "system";
    this.settings.uiLanguage = value;
    file.setSetting(this, "uiLanguage", value);
    initI18n(value);
    this.updateShell();
    this.renderTabs();
    this.renderComments();
    this.renderTemplates();
    this.renderRecent();
    this.renderRecovery();
    this.renderOptions();
    this.syncFormulaBar();
    this.updateStatus();
    const sheet = this.grid.sheet();
    this.ribbon?.update({
      ...(this.selection || this.grid.selectionInfo()),
      merged: !!mergeAt(sheet, this.grid.active.row, this.grid.active.col),
      showGridlines: sheet.showGridlines !== false,
      showHeadings: sheet.showHeadings !== false,
    });
    this.lastShell = "";
    this.syncShell();
  }

  cycleLanguage() {
    const next = currentLanguage() === "pl" ? "en" : "pl";
    this.setLanguage(next);
    toast(next === "pl" ? "Interfejs: polski" : "Interface: English");
  }

  /* ---------------------------------------------------------------- palette */
  commandItems() {
    const rows = [
      ["new", "New workbook", "File", "newWorkbook"],
      ["open", "Open…", "File", "open"],
      ["save", "Save", "File", "save"],
      ["saveAs", "Save As…", "File", "save"],
      ["print", "Print", "File", "print"],
      ["exportPdf", "Export PDF", "File", "print"],
      ["pageSetup", "Page Setup", "Page Layout", "pageSetup"],
      ["close", "Close", "File", "close"],
      ["undo", "Undo", "Home", "undo"],
      ["redo", "Redo", "Home", "redo"],
      ["cut", "Cut", "Home", "cut"],
      ["copy", "Copy", "Home", "copy"],
      ["paste", "Paste", "Home", "paste"],
      ["formatPainter", "Format Painter", "Home", "copy"],
      ["bold", "Bold", "Home", "bold"],
      ["italic", "Italic", "Home", "italic"],
      ["underline", "Underline", "Home", "underline"],
      ["strike", "Strikethrough", "Home", "strike"],
      ["wrap", "Wrap Text", "Home", "mergeCells"],
      ["mergeCenter", "Merge & Center", "Home", "mergeCells"],
      ["numformat:$#,##0.00", "Currency Format", "Home", "currency"],
      ["numformat:0%", "Percent Style", "Home", "percent"],
      ["numformat:0.00", "Number Format", "Home", "numberFormat"],
      ["formatCells", "Format Cells…", "Home", "formatCells"],
      ["conditionalDialog", "Conditional Formatting", "Home", "conditional"],
      ["cellStyleDialog", "Cell Styles", "Home", "cellStyle"],
      ["insert.picture", "Insert Picture", "Insert", "picture"],
      ["insert.textBox", "Insert Text Box", "Insert", "textBox"],
      ["insert.chart:column", "Insert Column Chart", "Insert", "chartColumn"],
      ["insert.chart:line", "Insert Line Chart", "Insert", "chartLine"],
      ["insert.chart:pie", "Insert Pie Chart", "Insert", "chartPie"],
      ["insert.table", "Insert Table", "Insert", "table"],
      ["insertFunction:SUM", "Insert Function", "Formulas", "sum"],
      ["autosum:SUM", "AutoSum", "Formulas", "sum"],
      ["calc.now", "Calculate Now", "Formulas", "calculator"],
      ["calc.sheet", "Calculate Sheet", "Formulas", "calculator"],
      ["showFormulas", "Show Formulas", "Formulas", "showFormulas"],
      ["name.manager", "Name Manager", "Formulas", "names"],
      ["sortDialog", "Sort", "Data", "sortAZ"],
      ["filter.toggle", "Filter", "Data", "filter"],
      ["textToColumns", "Text to Columns", "Data", "table"],
      ["removeDuplicates", "Remove Duplicates", "Data", "table"],
      ["validation", "Data Validation", "Data", "validation"],
      ["freeze.panes", "Freeze Panes", "View", "freezePanes"],
      ["view.normal", "Normal View", "View", "viewNormal"],
      ["view.pagebreak", "Page Break Preview", "View", "viewPageBreak"],
      ["view.pageLayout", "Page Layout View", "View", "viewPageLayout"],
      ["gridlines", "Gridlines", "View", "gridlines"],
      ["headings", "Headings", "View", "headings"],
      ["comment.new", "New Comment", "Review", "comment"],
      ["comment.show", "Show Comments", "Review", "comment"],
      ["spellcheck", "Spelling", "Review", "spellcheck"],
      ["protect.sheet", "Protect Sheet", "Review", "protect"],
      ["insert.headerFooter", "Header & Footer", "Page Layout", "pageSetup"],
      ["shortcuts", "Keyboard Shortcuts", "Help", "help"],
      ["about", "About mTails", "Help", "info"],
    ];
    return rows.map(([cmd, label, tab, iconName]) => ({ cmd, label, tab, icon: iconName }));
  }

  openPalette(initial = "") {
    const layer = $("#menu-layer");
    if (!layer) return;
    if (this.palette) {
      const next = String(initial ?? "");
      if (next !== this.palette.input.value) {
        this.palette.input.value = next;
        this.palette.input.dispatchEvent(new Event("input"));
      }
      this.palette.input.focus();
      this.palette.input.select();
      return;
    }
    closeMenus();
    const input = el("input", {
      type: "text",
      placeholder: "Search commands",
      autocomplete: "off",
      spellcheck: "false",
      value: String(initial || ""),
    });
    const list = el("div", { class: "palette-list" });
    const box = el("div", { class: "palette" }, [input, list]);
    box.style.pointerEvents = "auto";
    layer.append(box);
    const state = {
      box, input, list, items: [], filtered: [], active: 0, outside: null,
    };
    this.palette = state;
    const run = (item) => {
      if (!item) return;
      this.closePalette();
      this.grid?.els?.scroller?.focus({ preventScroll: true });
      this.execute(item.cmd);
    };
    const highlight = () => {
      Array.from(list.children).forEach((node, index) => {
        node.classList?.toggle("active", index === state.active);
      });
      list.children[state.active]?.scrollIntoView({ block: "nearest" });
    };
    const render = () => {
      const term = input.value.trim().toLowerCase();
      const items = this.commandItems().filter((item) => (
        !term
        || item.label.toLowerCase().includes(term)
        || item.cmd.toLowerCase().includes(term)
        || item.tab.toLowerCase().includes(term)
      )).slice(0, 40);
      state.filtered = items;
      state.active = 0;
      list.replaceChildren();
      if (!items.length) {
        list.append(el("div", { class: "palette-empty", text: "No commands match your search." }));
        return;
      }
      items.forEach((item, index) => {
        const button = el("button", {
          class: "palette-item" + (index === 0 ? " active" : ""),
          type: "button",
        }, [
          el("span", { class: "pi-icon", html: icon(item.icon || "dot"), i18n: false }),
          el("span", { text: item.label }),
          el("span", { class: "pi-tab", text: item.tab }),
        ]);
        button.addEventListener("click", () => run(item));
        list.append(button);
      });
    };
    input.addEventListener("input", render);
    input.addEventListener("keydown", (event) => {
      if (event.key === "ArrowDown") {
        event.preventDefault();
        if (state.filtered.length) {
          state.active = (state.active + 1) % state.filtered.length;
          highlight();
        }
      } else if (event.key === "ArrowUp") {
        event.preventDefault();
        if (state.filtered.length) {
          state.active = (state.active - 1 + state.filtered.length) % state.filtered.length;
          highlight();
        }
      } else if (event.key === "Enter") {
        event.preventDefault();
        run(state.filtered[state.active]);
      } else if (event.key === "Escape") {
        event.preventDefault();
        this.closePalette();
      } else if (event.key === "Tab") {
        event.preventDefault();
        const step = event.shiftKey ? -1 : 1;
        if (state.filtered.length) {
          state.active = (state.active + step + state.filtered.length) % state.filtered.length;
          highlight();
        }
      }
    });
    state.outside = (event) => {
      if (!state.box.contains(event.target)) this.closePalette();
    };
    document.addEventListener("mousedown", state.outside, true);
    const mirror = $("#command-search");
    if (mirror) mirror.value = String(initial || "");
    render();
    requestAnimationFrame(() => {
      input.focus();
      input.select();
    });
  }

  closePalette() {
    const state = this.palette;
    if (!state) return;
    this.palette = null;
    if (state.outside) document.removeEventListener("mousedown", state.outside, true);
    state.box.remove();
    const mirror = $("#command-search");
    if (mirror) mirror.value = "";
  }

  /* --------------------------------------------------------------- keyboard */
  bindKeyboard() {
    document.addEventListener("keydown", (event) => {
      if (event.defaultPrevented) return;
      if (this.palette) return;
      if (document.getElementById("dialog-layer")?.classList.contains("open")) return;
      const ctrl = event.ctrlKey || event.metaKey;
      const key = event.key;
      const lower = key.toLowerCase();
      if (ctrl && !event.altKey) {
        if (lower === "s") {
          event.preventDefault();
          this.file.saveFile({ saveAs: event.shiftKey });
          return;
        }
        if (lower === "o") {
          event.preventDefault();
          this.openFile();
          return;
        }
        if (lower === "n") {
          event.preventDefault();
          this.newWorkbook();
          return;
        }
        if (lower === "p") {
          event.preventDefault();
          this.file.printFile();
          return;
        }
        if (lower === "h") {
          event.preventDefault();
          this.openFind(true);
          return;
        }
        if (key === "`") {
          event.preventDefault();
          this.toggleShowFormulas();
          return;
        }
      }
      if (event.altKey && !ctrl && lower === "q") {
        event.preventDefault();
        this.openPalette();
        return;
      }
      const target = event.target;
      const fieldTarget = target && (
        target.tagName === "INPUT" || target.tagName === "TEXTAREA"
        || target.tagName === "SELECT" || target.isContentEditable
      );
      if (fieldTarget) return;
      if (key === "F9") {
        event.preventDefault();
        this.calculateNow();
        return;
      }
      if (key === "F1") {
        event.preventDefault();
        dlg.shortcutsDialog();
        return;
      }
      if (key === "Escape") this.handleEscape();
    });
  }

  handleEscape() {
    if (this.palette) {
      this.closePalette();
      return;
    }
    if (this.state.backstage) {
      this.closeBackstage();
      return;
    }
    if (this.findState.open) {
      this.closeFind();
      return;
    }
    if (document.body.classList.contains("painter-active")) {
      if (this.commands) this.commands.painter = null;
      this.setPainterCursor(false);
      return;
    }
    if (this.commentDraft) {
      this.commentDraft = null;
      const input = $("#comment-input");
      if (input) input.value = "";
    }
  }

  /* ----------------------------------------------------------- ribbon host */
  isFiltered() {
    const sheet = this.grid?.sheet();
    return !!sheet?.filters?.range;
  }

  filterRange() {
    return this.grid?.sheet()?.filters?.range || "";
  }

  freezeState() {
    return {
      rows: this.grid?.freezeRows || 0,
      cols: this.grid?.freezeCols || 0,
    };
  }

  /* ------------------------------------------------------------------ smoke */
  smoke() {
    const errors = this.errors.slice();
    let formulaOk = false;
    try {
      const sheet = this.grid?.sheet();
      const value = sheet ? this.engine.evaluate("=SUM(1,2,3)", sheet) : null;
      formulaOk = Number(value) === 6;
    } catch (error) {
      errors.push({ message: String(error?.message || error), source: "smoke", line: 0 });
    }
    const stats = this.model ? workbookStats(this.model) : { sheets: 0, cells: 0, formulas: 0 };
    let mounted = false;
    try {
      mounted = !!this.grid?.els?.viewport?.isConnected;
    } catch {
      mounted = false;
    }
    return {
      ok: !!this.booted && mounted && formulaOk && errors.length === 0,
      version: APP_VERSION,
      language: currentLanguage(),
      title: this.state.title,
      view: this.state.view,
      sheets: stats.sheets,
      cells: stats.cells,
      formulas: stats.formulas,
      formula: formulaOk,
      dirty: !!this.state.dirty,
      errors,
    };
  }
}

/* -------------------------------------------------------------------- boot */
const app = new App();
window.app = app;
window.mtails = {
  version: APP_VERSION,
  app,
  get model() { return app.model; },
  get state() { return app.state; },
  get settings() { return app.settings; },
  evaluate: (formula) => app.engine.evaluate(formula, app.grid.sheet()),
  execute: (command) => app.execute(command),
  newWorkbook: (template) => app.newWorkbook(template),
  openBackstage: (panel) => app.openBackstage(panel),
  closeBackstage: () => app.closeBackstage(),
  setLanguage: (language) => app.setLanguage(language),
  revision: () => app.revision,
  dirty: () => !!app.state.dirty,
  getErrors: () => app.errors.slice(),
};
window.mTails = window.mtails;
window.__smoke = () => app.smoke();
window.__saveBeforeClose = () => app.file.saveBeforeClose();

window.addEventListener("error", (event) => {
  if (event.target && event.target !== window && !event.message) return;
  app.recordError(event.error || event.message, event.filename || "", event.lineno || 0);
});
window.addEventListener("unhandledrejection", (event) => {
  const reason = event.reason;
  app.recordError(reason?.stack || reason?.message || reason, "promise", 0);
});

app.init().catch((error) => {
  app.recordError(error, "init", 0);
  document.getElementById("app")?.classList.remove("loading");
  toast(t("The application could not be started."), "error");
});
