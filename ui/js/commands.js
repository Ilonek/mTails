/* Command router: turns ribbon ids, menu items and keyboard shortcuts into
 * workbook operations.  The router owns no state beyond the armed format
 * painter; everything else is delegated to the grid, the model or a dialog.
 *
 * Ribbon ids use a few namespaced prefixes so one button can carry a payload:
 *   numformat:<code>   borders:<preset>   fill:<direction|none|colour>
 *   fontColor:<hex>    cellStyle:<id>     formatTable:<index>
 *   insert.chart:<t>   insert.shape:<s>   insertFunction:<name>
 *   insertName:<name>  conditional:<rule> colorScale:<variant>  autosum:<FN>
 */

import { $, el, toast, showDialog, field, segmented, colorMenu } from "./ui.js";
import {
  cellAt,
  ensureCell,
  usedRange,
  setRowHidden,
  setColHidden,
  addMerge,
  deleteCellComments,
} from "./model.js";
import { addr, parseRange, rangeText } from "./units.js";
import { formatCell } from "./format.js";
import * as dlg from "./dialogs.js";
import { STYLE_PRESETS, CHART_TYPES } from "./ribbon.js";

const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 28, 32, 36, 48, 72];

function clone(value) {
  return value === undefined ? undefined : JSON.parse(JSON.stringify(value));
}

function selectionArea(app) {
  const grid = app.grid;
  if (grid.sel.r1 !== grid.sel.r2 || grid.sel.c1 !== grid.sel.c2) return { ...grid.sel };
  return usedRange(grid.sheet()) || { ...grid.sel };
}

function selectionOrUsed(app) {
  return selectionArea(app);
}

/* --------------------------------------------------------------- clipboard */
function legacyCopy(text) {
  const area = el("textarea", { i18n: false });
  area.value = text;
  area.style.cssText = "position:fixed;left:-9999px;top:0;opacity:0";
  document.body.append(area);
  area.select();
  let ok = false;
  try {
    ok = document.execCommand("copy");
  } catch {
    ok = false;
  }
  area.remove();
  return ok;
}

async function writeClipboard(text) {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch {
    return legacyCopy(text);
  }
}

async function readClipboard() {
  try {
    const text = await navigator.clipboard.readText();
    if (text) return text;
  } catch {
    /* fall through to the paste helper */
  }
  return "";
}

/* Adjust the number of decimal places in a format code without touching the
   literal parts, so "$#,##0.00" becomes "$#,##0.0" and not "$#,##0". */
function adjustDecimals(code, delta) {
  const source = String(code || "General");
  if (source === "General" || source === "@" || source === "") {
    return delta > 0 ? "0.00" : "0";
  }
  const masked = source.replace(/"[^"]*"/g, (text) => "\u0000".repeat(text.length));
  const fraction = /\.([0#]+)/.exec(masked);
  if (fraction) {
    const decimals = fraction[1];
    const count = decimals.length + delta;
    if (count <= 0) {
      return source.slice(0, fraction.index) + source.slice(fraction.index + fraction[0].length);
    }
    const next = delta > 0 ? decimals + "0" : decimals.slice(0, -1);
    return source.slice(0, fraction.index + 1) + next + source.slice(fraction.index + fraction[0].length);
  }
  if (delta > 0) {
    const last = /[0#](?=[^0#]*$)/.exec(masked);
    if (last) return source.slice(0, last.index + 1) + ".0" + source.slice(last.index + 1);
  }
  return source;
}

function changeDecimals(app, delta) {
  const style = app.grid.currentStyle();
  const next = adjustDecimals(style.numFmt, delta);
  if (next && next !== style.numFmt) app.grid.setNumberFormat(next);
}

function nextFontSize(current, direction) {
  const size = Number(current) || 11;
  if (direction > 0) return FONT_SIZES.find((value) => value > size) || FONT_SIZES[FONT_SIZES.length - 1];
  return [...FONT_SIZES].reverse().find((value) => value < size) || FONT_SIZES[0];
}

function clearSelectionStyle(app) {
  const grid = app.grid;
  const sheet = grid.sheet();
  const before = app.captureState();
  for (const address of grid.selectedAddresses()) {
    const cell = sheet.cells[address];
    if (!cell) continue;
    delete cell.s;
    if (cell.v === undefined && !cell.f && !cell.h) delete sheet.cells[address];
  }
  app.grid.afterMutation("style", before, { layout: true });
}

function applyStyleToSelection(app, style, { replace = false } = {}) {
  const grid = app.grid;
  const sheet = grid.sheet();
  const before = app.captureState();
  for (const address of grid.selectedAddresses()) {
    const cell = ensureCell(sheet, address);
    if (replace) {
      if (style && Object.keys(style).length) cell.s = clone(style);
      else delete cell.s;
    } else {
      cell.s = { ...(cell.s || {}), ...clone(style || {}) };
    }
    if (!cell.s && cell.v === undefined && !cell.f && !cell.h) delete sheet.cells[address];
  }
  grid.afterMutation("style", before, { layout: true });
}

function shapeSvg(shape) {
  const open = '<svg xmlns="http://www.w3.org/2000/svg" width="160" height="96" viewBox="0 0 160 96">';
  const close = "</svg>";
  if (shape === "ellipse") {
    return open + '<ellipse cx="80" cy="48" rx="66" ry="40" fill="#D9E2F3" stroke="#2F5597" stroke-width="2"/>' + close;
  }
  if (shape === "line") {
    return open + '<line x1="14" y1="82" x2="146" y2="14" stroke="#2F5597" stroke-width="3" stroke-linecap="round"/>' + close;
  }
  if (shape === "round") {
    return open + '<rect x="12" y="12" width="136" height="72" rx="14" fill="#D9E2F3" stroke="#2F5597" stroke-width="2"/>' + close;
  }
  return open + '<rect x="12" y="12" width="136" height="72" fill="#D9E2F3" stroke="#2F5597" stroke-width="2"/>' + close;
}

const FEEDBACK_TEXT = "mTails is a local-first spreadsheet: there is no account, "
  + "no telemetry and no cloud service. Feedback is handled by the team that "
  + "installed the application. Include the mTails version and, if possible, a "
  + "copy of the workbook so the problem can be reproduced.";

export class Commands {
  constructor(app) {
    this.app = app;
    this.painter = null;
  }

  /* --------------------------------------------------------------- router */
  execute(command, anchor = null, extra = {}) {
    if (!command) return false;
    const app = this.app;
    const grid = app.grid;

    if (command.startsWith("numformat:")) {
      const code = command.slice("numformat:".length);
      grid.setNumberFormat(!code || code === "General" ? null : code);
      return true;
    }
    if (command.startsWith("borders:")) {
      const preset = command.slice("borders:".length);
      const spec = preset === "outsideThick"
        ? { style: "thick", color: app.ribbon?.borderColor || "#000000" }
        : { style: app.ribbon?.borderStyle || "thin", color: app.ribbon?.borderColor || "#000000" };
      grid.applyBorders(preset === "outsideThick" ? "outside" : preset, spec);
      return true;
    }
    if (command.startsWith("fill:")) {
      const value = command.slice("fill:".length);
      if (["down", "up", "left", "right"].includes(value)) grid.fillSelection(value);
      else grid.setStyle({ fill: value && value !== "none" ? value : null });
      return true;
    }
    if (command.startsWith("fontColor:")) {
      const color = command.slice("fontColor:".length);
      grid.setStyle({ color: color || null });
      return true;
    }
    if (command.startsWith("cellStyle:")) {
      this.applyCellStyle(command.slice("cellStyle:".length));
      return true;
    }
    if (command.startsWith("formatTable:")) {
      const index = Number(command.slice("formatTable:".length)) || 0;
      dlg.applyTableStyle(app, index, selectionOrUsed(app), true);
      toast("Table style applied.");
      return true;
    }
    if (command.startsWith("insert.chart:")) {
      this.insertChart(command.slice("insert.chart:".length));
      return true;
    }
    if (command.startsWith("insert.shape:")) {
      const shape = command.slice("insert.shape:".length);
      app.addImage("data:image/svg+xml;charset=utf-8," + encodeURIComponent(shapeSvg(shape)), 160, 96);
      toast("Shape inserted.");
      return true;
    }
    if (command.startsWith("insertFunction:")) {
      app.insertFunction(command.slice("insertFunction:".length));
      return true;
    }
    if (command.startsWith("insertName:")) {
      app.insertName(command.slice("insertName:".length));
      return true;
    }
    if (command.startsWith("autosum:")) {
      this.autoSum(command.slice("autosum:".length));
      return true;
    }
    if (command.startsWith("colorScale:")) {
      this.applyColorScale(command.slice("colorScale:".length));
      return true;
    }
    if (command.startsWith("conditional:")) {
      const preset = command.slice("conditional:".length);
      if (preset === "clear") {
        const sheet = grid.sheet();
        app.mutate("conditional", () => { sheet.conditional = []; }, { layout: false, recalc: false });
        toast("Rules cleared.");
      } else {
        dlg.conditionalDialog(app, preset);
      }
      return true;
    }

    switch (command) {
      /* ---------------------------------------------------------- clipboard */
      case "copy":
        this.copy();
        return true;
      case "cut":
        this.cut();
        return true;
      case "paste":
        this.paste("all");
        return true;
      case "pasteValues":
        this.paste("values");
        return true;
      case "pasteFormulas":
        this.paste("formulas");
        return true;
      case "pasteFormats":
        this.paste("formats");
        return true;
      case "pasteTranspose":
        this.paste("transpose");
        return true;
      case "pasteSpecial":
        dlg.pasteSpecialDialog(app, (mode) => this.paste(mode));
        return true;
      case "formatPainter":
        this.applyFormatPainter();
        return true;

      /* ------------------------------------------------------------ history */
      case "undo":
        app.undo();
        return true;
      case "redo":
        app.redo();
        return true;

      /* --------------------------------------------------------------- font */
      case "bold":
        grid.toggleStyle("b", true);
        return true;
      case "italic":
        grid.toggleStyle("i", true);
        return true;
      case "underline":
        grid.toggleStyle("u", true);
        return true;
      case "strike":
        grid.toggleStyle("strike", true);
        return true;
      case "font.name": {
        const value = String(extra.value ?? "").trim();
        if (!value) return true;
        grid.setStyle({ font: value.slice(0, 40) });
        return true;
      }
      case "font.size": {
        const value = Number(extra.value);
        if (!Number.isFinite(value) || value < 4 || value > 300) {
          toast("Enter a font size between 4 and 300.");
          return true;
        }
        grid.setStyle({ size: Math.round(value * 10) / 10 });
        return true;
      }
      case "font.grow":
        grid.setStyle({ size: nextFontSize(grid.currentStyle().size, 1) });
        return true;
      case "font.shrink":
        grid.setStyle({ size: nextFontSize(grid.currentStyle().size, -1) });
        return true;
      case "fillColor":
        this.openFillMenu(anchor);
        return true;
      case "fontColor":
        this.openFontColorMenu(anchor);
        return true;

      /* ---------------------------------------------------------- alignment */
      case "align.left":
      case "align.center":
      case "align.right":
        grid.setAlignment(command.slice("align.".length));
        return true;
      case "valign.top":
      case "valign.middle":
      case "valign.bottom":
        grid.setVerticalAlignment(command.slice("valign.".length));
        return true;
      case "wrap":
        grid.toggleStyle("wrap", true);
        return true;
      case "indent.increase":
        grid.setStyle({ indent: Math.min(15, (Number(grid.currentStyle().indent) || 0) + 1) });
        return true;
      case "indent.decrease": {
        const indent = Math.max(0, (Number(grid.currentStyle().indent) || 0) - 1);
        grid.setStyle({ indent: indent || null });
        return true;
      }
      case "mergeCenter":
        this.mergeCenter();
        return true;
      case "mergeAcross":
        this.mergeAcross();
        return true;
      case "mergeCells":
        grid.mergeCells();
        return true;
      case "unmerge":
        grid.unmergeCells();
        return true;
      case "textOrientation":
        this.textOrientationDialog();
        return true;

      /* ------------------------------------------------------------- number */
      case "decimals.increase":
        changeDecimals(app, 1);
        return true;
      case "decimals.decrease":
        changeDecimals(app, -1);
        return true;

      /* ------------------------------------------------------------- styles */
      case "cellStyleDialog":
        dlg.newCellStyleDialog(app);
        return true;
      case "conditionalDialog":
        dlg.conditionalDialog(app);
        return true;
      case "conditionalRules":
        dlg.conditionalRulesDialog(app);
        return true;

      /* -------------------------------------------------------------- cells */
      case "insert.rows":
        grid.insertRow(1);
        return true;
      case "insert.cols":
        grid.insertColumn(1);
        return true;
      case "insert.cells":
        this.insertCellsDialog();
        return true;
      case "delete.rows":
        grid.deleteRowSelection();
        return true;
      case "delete.cols":
        grid.deleteColumnSelection();
        return true;
      case "delete.cells":
        this.deleteCellsDialog();
        return true;
      case "rowHeight":
        dlg.rowHeightDialog(app);
        return true;
      case "colWidth":
        dlg.columnWidthDialog(app);
        return true;
      case "autoFit.cols":
        for (let col = grid.sel.c1; col <= grid.sel.c2; col += 1) grid.autoFitColumn(col);
        return true;
      case "autoFit.rows":
        for (let row = grid.sel.r1; row <= grid.sel.r2; row += 1) grid.autoFitRow(row);
        return true;
      case "hide.rows":
        this.setRowsHidden(true);
        return true;
      case "unhide.rows":
        this.setRowsHidden(false);
        return true;
      case "hide.cols":
        this.setColsHidden(true);
        return true;
      case "unhide.cols":
        this.setColsHidden(false);
        return true;
      case "formatCells":
        dlg.formatCellsDialog(app, extra.tab || "number");
        return true;

      /* ------------------------------------------------------------- editing */
      case "fillSeries":
        dlg.fillSeriesDialog(app);
        return true;
      case "clear:all":
        grid.clearSelection("all");
        return true;
      case "clear:formats":
        grid.clearSelection("formats");
        return true;
      case "clear:contents":
        grid.clearSelection("contents");
        return true;
      case "clear:comments":
        this.clearComments();
        return true;
      case "sort.asc":
        this.sortByActiveColumn("asc");
        return true;
      case "sort.desc":
        this.sortByActiveColumn("desc");
        return true;
      case "sortDialog":
        dlg.sortDialog(app);
        return true;
      case "filter.toggle":
        this.toggleFilter();
        return true;
      case "filter.clear":
        if (grid.clearFilter()) toast("Filter cleared.");
        return true;
      case "filter.reapply":
        this.reapplyFilter();
        return true;
      case "filter.byValue":
        this.filterByValue();
        return true;
      case "find":
        app.openFind(false);
        return true;
      case "replace":
        app.openFind(true);
        return true;
      case "goto":
      case "gotoSpecial":
        dlg.goToDialog(app);
        return true;
      case "selectAll":
        grid.selectAll();
        return true;

      /* ------------------------------------------------------------- insert */
      case "insert.table":
        dlg.tableDialog(app);
        return true;
      case "insert.picture":
        $("#image-input")?.click();
        return true;
      case "insert.icons":
        dlg.iconsDialog(app);
        return true;
      case "insert.symbol":
        dlg.symbolDialog(app);
        return true;
      case "insert.textBox":
        dlg.textBoxDialog(app);
        return true;
      case "insert.headerFooter":
        dlg.headerFooterDialog(app);
        return true;
      case "insert.link":
        dlg.hyperlinkDialog(app);
        return true;
      case "insert.sparkline":
        dlg.sparklineDialog(app);
        return true;
      case "comment.new":
        app.newComment();
        return true;
      case "chartDialog":
        dlg.chartDialog(app);
        return true;

      /* ----------------------------------------------------------- formulas */
      case "functionDialog":
        dlg.functionDialog(app);
        return true;
      case "name.manager":
        dlg.nameManagerDialog(app);
        return true;
      case "name.define":
        dlg.defineNameDialog(app);
        return true;
      case "audit.precedents":
        dlg.auditDialog(app, "precedents");
        return true;
      case "audit.dependents":
        dlg.auditDialog(app, "dependents");
        return true;
      case "audit.remove":
        app.clearAudit();
        toast("Arrows cleared.");
        return true;
      case "showFormulas":
        app.toggleShowFormulas();
        return true;
      case "calc.auto":
        app.setCalcMode("auto");
        return true;
      case "calc.manual":
        app.setCalcMode("manual");
        return true;
      case "calc.now":
        app.calculateNow();
        return true;
      case "calc.sheet":
        app.calculateSheet();
        return true;

      /* --------------------------------------------------------------- data */
      case "textToColumns":
        dlg.textToColumnsDialog(app);
        return true;
      case "flashFill":
        this.flashFill();
        return true;
      case "removeDuplicates":
        dlg.removeDuplicatesDialog(app);
        return true;
      case "validation":
        dlg.validationDialog(app);
        return true;
      case "subtotal":
        dlg.subtotalDialog(app);
        return true;
      case "group.rows":
        this.setRowsHidden(true, "Rows grouped.");
        return true;
      case "ungroup.rows":
        this.setRowsHidden(false, "Rows ungrouped.");
        return true;

      /* ------------------------------------------------------------- review */
      case "comment.delete":
        if (grid.deleteCommentAt()) toast("Comment deleted.");
        else toast("The active cell has no comment.");
        return true;
      case "comment.prev":
        app.cycleComment(-1);
        return true;
      case "comment.next":
        app.cycleComment(1);
        return true;
      case "comment.show":
        app.toggleComments();
        return true;
      case "protect.sheet":
        dlg.protectionDialog(app, "sheet");
        return true;
      case "protect.workbook":
        dlg.protectionDialog(app, "workbook");
        return true;
      case "spellcheck":
        dlg.spellcheckDialog(app);
        return true;
      case "workbook.stats":
        dlg.statsDialog(app);
        return true;
      case "accessibility":
        dlg.accessibilityDialog(app);
        return true;

      /* --------------------------------------------------------------- view */
      case "view.normal":
        app.setView("normal");
        return true;
      case "view.pagebreak":
        app.setView("pagebreak");
        return true;
      case "view.pageLayout":
        app.setView("layout");
        return true;
      case "gridlines":
        grid.toggleGridlines();
        return true;
      case "headings":
        grid.toggleHeadings();
        return true;
      case "formulaBar":
        app.toggleFormulaBar();
        return true;
      case "theme.toggle":
        app.toggleTheme();
        return true;
      case "zoom":
        dlg.zoomDialog(app);
        return true;
      case "zoom100":
        grid.setZoom(100);
        return true;
      case "zoomSelection":
        this.zoomToSelection();
        return true;
      case "freeze.panes":
        grid.freezePanes();
        return true;
      case "freeze.top":
        grid.setFreeze(1, 0);
        return true;
      case "freeze.first":
        grid.setFreeze(0, 1);
        return true;
      case "unfreeze":
        grid.unfreezePanes();
        return true;
      case "splitPanes":
        this.splitPanes();
        return true;
      case "sheet.hide":
        if (app.model.sheets.length <= 1) toast("A workbook must contain at least one visible sheet.");
        else app.setSheetHidden(app.model.activeSheet, true);
        return true;
      case "sheet.unhide":
        dlg.unhideSheetDialog(app);
        return true;

      /* -------------------------------------------------------------- sheets */
      case "sheet.add":
        app.insertSheet();
        return true;
      case "sheet.delete":
        dlg.deleteSheetDialog(app);
        return true;
      case "sheet.rename":
        dlg.renameSheetDialog(app);
        return true;
      case "sheet.duplicate":
        app.duplicateSheet();
        return true;
      case "sheet.move":
        dlg.moveSheetDialog(app);
        return true;
      case "sheet.color":
        dlg.sheetColorDialog(app, app.model.activeSheet);
        return true;

      /* ---------------------------------------------------------------- help */
      case "help":
        dlg.helpDialog(app);
        return true;
      case "shortcuts":
        dlg.shortcutsDialog();
        return true;
      case "whatsNew":
        dlg.whatsNewDialog();
        return true;
      case "about":
        dlg.aboutDialog(app);
        return true;
      case "feedback":
        this.feedbackDialog();
        return true;

      /* ------------------------------------------------------------- shell */
      case "backstage":
        app.openBackstage("new");
        return true;
      case "save":
        app.file.saveFile();
        return true;
      case "saveAs":
        app.file.saveFile({ saveAs: true });
        return true;
      case "open":
        app.openFile();
        return true;
      case "new":
        app.newWorkbook();
        return true;
      case "print":
        app.file.printFile();
        return true;
      case "exportPdf":
        app.file.exportPdf();
        return true;
      case "pageSetup":
        dlg.pageSetupDialog(app);
        return true;
      case "close":
        app.file.closeWindow();
        return true;

      default:
        return false;
    }
  }

  /* -------------------------------------------------------------- clipboard */
  copy() {
    const data = this.app.grid.copySelection();
    writeClipboard(data.text);
    this.app.grid.els?.scroller.focus({ preventScroll: true });
    toast("Copied.");
  }

  cut() {
    const data = this.app.grid.cutSelection();
    writeClipboard(data.text);
    this.app.grid.els?.scroller.focus({ preventScroll: true });
    toast("Cut.");
  }

  async paste(mode = "all") {
    const grid = this.app.grid;
    if (grid.clipboard?.text) {
      grid.pasteText(grid.clipboard.text, { mode });
      return;
    }
    const text = await readClipboard();
    if (!text) {
      if (mode === "formats") toast("The clipboard does not contain cell formatting.");
      else this.pasteFallback(mode);
      return;
    }
    if (mode === "formats") {
      toast("The clipboard does not contain cell formatting.");
      return;
    }
    grid.pasteText(text, { mode });
  }

  pasteFallback(mode = "all") {
    const area = el("textarea", { rows: 6, placeholder: "Paste here with Ctrl+V", i18n: false });
    area.style.cssText = "width:100%;border:1px solid var(--border-strong);border-radius:4px;padding:7px 8px;font:12px var(--font-ui);resize:vertical";
    const dialog = showDialog({
      title: "Paste",
      size: "narrow",
      body: el("div", {}, [
        el("p", { class: "dialog-note", text: "The system clipboard could not be read directly. Paste the copied cells into this box and choose OK." }),
        area,
      ]),
      buttons: [
        { label: "Cancel", kind: "ghost" },
        {
          label: "OK", kind: "primary", onClick: () => {
            if (!area.value) return false;
            this.app.grid.pasteText(area.value, { mode });
            return true;
          },
        },
      ],
      onOpen: () => setTimeout(() => area.focus(), 40),
    });
    return dialog;
  }

  applyFormatPainter() {
    const app = this.app;
    const grid = app.grid;
    if (!this.painter) {
      this.painter = clone(grid.currentStyle()) || {};
      app.setPainterCursor?.(true);
      toast("Format copied. Select the destination and click Format Painter again.");
      return;
    }
    const style = this.painter;
    this.painter = null;
    app.setPainterCursor?.(false);
    applyStyleToSelection(app, style, { replace: true });
    toast("Formatting applied.");
  }

  /* ------------------------------------------------------------------ font */
  openFillMenu(anchor) {
    const target = anchor || $("#grid-viewport");
    colorMenu(target, {
      title: "Theme Colors",
      allowNone: true,
      onPick: (value) => this.app.grid.setStyle({ fill: value || null }),
    });
  }

  openFontColorMenu(anchor) {
    const target = anchor || $("#grid-viewport");
    colorMenu(target, {
      title: "Theme Colors",
      onPick: (value) => this.app.grid.setStyle({ color: value || null }),
    });
  }

  /* ------------------------------------------------------------- alignment */
  mergeCenter() {
    const grid = this.app.grid;
    const sheet = grid.sheet();
    const merged = (sheet.merges || []).some((text) => {
      const range = parseRange(text);
      return range && !(range.r2 < grid.sel.r1 || range.r1 > grid.sel.r2 || range.c2 < grid.sel.c1 || range.c1 > grid.sel.c2);
    });
    if (merged) {
      grid.unmergeCells();
      return;
    }
    if (grid.sel.r1 === grid.sel.r2 && grid.sel.c1 === grid.sel.c2) {
      grid.setAlignment("center");
      return;
    }
    if (grid.mergeCells()) grid.setAlignment("center");
  }

  mergeAcross() {
    const app = this.app;
    const grid = app.grid;
    const sheet = grid.sheet();
    const range = { ...grid.sel };
    if (range.c1 === range.c2) {
      toast("Select two or more columns to merge across.");
      return;
    }
    app.mutate("merge", () => {
      for (let row = range.r1; row <= range.r2; row += 1) {
        addMerge(sheet, { r1: row, c1: range.c1, r2: row, c2: range.c2 });
      }
    }, { layout: true, recalc: false });
  }

  textOrientationDialog() {
    const app = this.app;
    const grid = app.grid;
    let angle = Number(grid.currentStyle().rotate) || 0;
    const values = [0, 45, 90, -45];
    const control = segmented(
      values.map((value) => ({ value, label: value === 0 ? "Horizontal" : value + "°" })),
      angle,
      (value) => { angle = Number(value) || 0; },
    );
    showDialog({
      title: "Orientation",
      size: "narrow",
      body: el("div", {}, [
        field("Text rotation", control),
        el("p", { class: "dialog-note", text: "Rotated text keeps the cell size; increase the row height if the text is clipped." }),
      ]),
      buttons: [
        { label: "Cancel", kind: "ghost" },
        { label: "OK", kind: "primary", onClick: () => grid.setStyle({ rotate: angle || null }) },
      ],
    });
  }

  /* ----------------------------------------------------------------- cells */
  insertCellsDialog() {
    const app = this.app;
    const grid = app.grid;
    let direction = "down";
    const body = el("div", {}, [
      el("div", { class: "dialog-section-title", text: "Insert" }),
      el("div", { class: "check-row" }, [
        el("input", { type: "radio", name: "insert-direction", checked: true, onchange: () => { direction = "down"; } }),
        el("span", { text: "Shift cells down" }),
      ]),
      el("div", { class: "check-row" }, [
        el("input", { type: "radio", name: "insert-direction", onchange: () => { direction = "right"; } }),
        el("span", { text: "Shift cells right" }),
      ]),
      el("p", { class: "dialog-note", text: "The new cells take the size of the current selection." }),
    ]);
    showDialog({
      title: "Insert Cells",
      size: "narrow",
      body,
      buttons: [
        { label: "Cancel", kind: "ghost" },
        {
          label: "OK", kind: "primary", onClick: () => {
            if (direction === "down") grid.insertRow(grid.sel.r2 - grid.sel.r1 + 1);
            else grid.insertColumn(grid.sel.c2 - grid.sel.c1 + 1);
            toast(direction === "down" ? "Cells shifted down." : "Cells shifted right.");
          },
        },
      ],
    });
  }

  deleteCellsDialog() {
    const app = this.app;
    const grid = this.app.grid;
    let direction = "up";
    const body = el("div", {}, [
      el("div", { class: "dialog-section-title", text: "Delete" }),
      el("div", { class: "check-row" }, [
        el("input", { type: "radio", name: "delete-direction", checked: true, onchange: () => { direction = "up"; } }),
        el("span", { text: "Shift cells up" }),
      ]),
      el("div", { class: "check-row" }, [
        el("input", { type: "radio", name: "delete-direction", onchange: () => { direction = "left"; } }),
        el("span", { text: "Shift cells left" }),
      ]),
      el("p", { class: "dialog-note", text: "The selected cells are removed and the remaining cells move into their place." }),
    ]);
    showDialog({
      title: "Delete Cells",
      size: "narrow",
      body,
      buttons: [
        { label: "Cancel", kind: "ghost" },
        {
          label: "OK", kind: "primary", onClick: () => {
            if (direction === "up") grid.deleteRowSelection();
            else grid.deleteColumnSelection();
            toast(direction === "up" ? "Cells shifted up." : "Cells shifted left.");
          },
        },
      ],
    });
  }

  setRowsHidden(hidden, message = "") {
    const app = this.app;
    const grid = app.grid;
    const sheet = grid.sheet();
    const range = grid.sel;
    app.mutate(hidden ? "hide-rows" : "show-rows", () => {
      for (let row = range.r1; row <= range.r2; row += 1) setRowHidden(sheet, row, hidden);
    }, { layout: true, recalc: false });
    if (message) toast(message);
    else toast(hidden ? "Rows hidden." : "Rows shown.");
  }

  setColsHidden(hidden) {
    const app = this.app;
    const grid = app.grid;
    const sheet = grid.sheet();
    const range = grid.sel;
    app.mutate(hidden ? "hide-cols" : "show-cols", () => {
      for (let col = range.c1; col <= range.c2; col += 1) setColHidden(sheet, col, hidden);
    }, { layout: true, recalc: false });
    toast(hidden ? "Columns hidden." : "Columns shown.");
  }

  /* --------------------------------------------------------------- editing */
  clearComments() {
    const app = this.app;
    const grid = app.grid;
    const sheet = grid.sheet();
    let removed = 0;
    app.mutate("clear-comments", () => {
      for (const address of grid.selectedAddresses()) {
        if (deleteCellComments(sheet, address)) removed += 1;
      }
    }, { layout: true, recalc: false });
    toast(removed ? "Comments removed." : "The selection has no comments.");
  }

  sortByActiveColumn(order) {
    const app = this.app;
    const grid = app.grid;
    const sheet = grid.sheet();
    const header = checkHeaderRow(sheet, grid.sel);
    const ok = grid.applySort([{ col: grid.active.col, order }], { hasHeader: header });
    if (!ok) toast("There is nothing to sort in this range.");
    else toast(order === "asc" ? "Sorted A to Z." : "Sorted Z to A.");
  }

  toggleFilter() {
    const app = this.app;
    const grid = app.grid;
    const sheet = grid.sheet();
    if (sheet.filters) {
      grid.clearFilter();
      toast("Filter removed.");
      return;
    }
    const range = selectionOrUsed(app);
    if (!range || (range.r1 === range.r2 && range.c1 === range.c2)) {
      toast("Select a range with a header row first.");
      return;
    }
    grid.applyFilter(range, []);
    toast("Filter added.");
  }

  reapplyFilter() {
    const grid = this.app.grid;
    const filters = grid.sheet().filters;
    if (!filters) {
      toast("This sheet has no filter.");
      return;
    }
    const range = parseRange(filters.range);
    if (!range) return;
    grid.applyFilter(range, filters.criteria || []);
    toast("Filter reapplied.");
  }

  filterByValue() {
    const grid = this.app.grid;
    const sheet = grid.sheet();
    const cell = cellAt(sheet, grid.active.row, grid.active.col);
    const value = cell ? String(formatCell(cell).text ?? "") : "";
    const existing = sheet.filters?.range ? parseRange(sheet.filters.range) : null;
    const range = existing || selectionOrUsed(this.app);
    if (!range || value === "") {
      toast("Select a cell with a value to filter by.");
      return;
    }
    if (grid.active.col < range.c1 || grid.active.col > range.c2) {
      toast("The active cell is outside the filter range.");
      return;
    }
    grid.applyFilter(range, [{ col: grid.active.col, values: [value] }]);
    toast('Filtered by "' + value + '".');
  }

  /* ---------------------------------------------------------------- insert */
  insertChart(type) {
    const app = this.app;
    const grid = app.grid;
    const sheet = grid.sheet();
    if (!CHART_TYPES.includes(type)) type = "column";
    const multi = grid.sel.r1 !== grid.sel.r2 || grid.sel.c1 !== grid.sel.c2;
    const range = multi ? { ...grid.sel } : usedRange(sheet);
    if (!range || (range.r1 === range.r2 && range.c1 === range.c2)) {
      dlg.chartDialog(app, type);
      return;
    }
    const chart = {
      id: "chart-" + Date.now().toString(36),
      type,
      range: rangeText(range),
      title: "",
      headerRow: true,
      headerCol: false,
      x: Math.round((grid.colPrefix[range.c2] || 0) / grid.scale) + 18,
      y: Math.round((grid.rowPrefix[range.r1 - 1] || 0) / grid.scale) + 8,
      w: 440,
      h: 270,
    };
    app.mutate("chart", () => { sheet.charts.push(chart); }, { layout: false, recalc: false });
    toast("Chart inserted.");
  }

  /* -------------------------------------------------------------- formulas */
  autoSum(name) {
    const app = this.app;
    const grid = app.grid;
    const sheet = grid.sheet();
    const row = grid.active.row;
    const col = grid.active.col;
    let start = null;
    let end = null;
    for (let above = row - 1; above >= 1; above -= 1) {
      const cell = cellAt(sheet, above, col);
      if (cell && (cell.v !== undefined || cell.f)) {
        start = above;
        break;
      }
    }
    if (start !== null) {
      end = start;
      while (start > 1) {
        const cell = cellAt(sheet, start - 1, col);
        if (!cell || (cell.v === undefined && !cell.f)) break;
        start -= 1;
      }
      // Do not swallow a header or a total row: stop at the first run.
      const range = start === end ? addr(start, col) : addr(start, col) + ":" + addr(end, col);
      grid.setActiveValue("=" + name + "(" + range + ")", { move: "down" });
      return;
    }
    for (let left = col - 1; left >= 1; left -= 1) {
      const cell = cellAt(sheet, row, left);
      if (cell && (cell.v !== undefined || cell.f)) {
        start = left;
        break;
      }
    }
    if (start !== null) {
      end = start;
      while (start > 1) {
        const cell = cellAt(sheet, row, start - 1);
        if (!cell || (cell.v === undefined && !cell.f)) break;
        start -= 1;
      }
      const range = start === end ? addr(row, start) : addr(row, start) + ":" + addr(row, end);
      grid.setActiveValue("=" + name + "(" + range + ")", { move: "down" });
      return;
    }
    grid.editActive("=" + name + "()");
  }

  applyColorScale(variant) {
    const app = this.app;
    const grid = app.grid;
    const range = selectionOrUsed(app);
    const palettes = {
      "3": ["#63BE7B", "#FFEB84", "#F8696B"],
      "3r": ["#F8696B", "#FFEB84", "#63BE7B"],
      "3b": ["#5A8AC6", "#FFFFFF", "#F8696B"],
    };
    const rule = {
      range: rangeText(range),
      type: "colorScale",
      colors: palettes[variant] || palettes["3"],
    };
    const sheet = grid.sheet();
    app.mutate("conditional", () => {
      sheet.conditional = sheet.conditional || [];
      sheet.conditional.push(rule);
    }, { layout: false, recalc: false });
    toast("Color scale applied.");
  }

  applyCellStyle(id) {
    const app = this.app;
    const grid = app.grid;
    if (id === "normal") {
      clearSelectionStyle(app);
      return;
    }
    const preset = STYLE_PRESETS.find((item) => item.id === id);
    const custom = preset ? null : (app.customStyles() || []).find((item) => item.id === id);
    const style = preset?.style || custom?.style;
    if (!style) return;
    grid.setStyle(clone(style));
  }

  /* ------------------------------------------------------------------ data */
  flashFill() {
    const app = this.app;
    const grid = app.grid;
    const sheet = grid.sheet();
    const used = usedRange(sheet);
    if (!used) {
      toast("There is no data to flash fill.");
      return;
    }
    let range = { ...grid.sel };
    if (range.c1 === range.c2 && range.r1 === range.r2) {
      range = { r1: used.r1, r2: used.r2, c1: grid.active.col, c2: grid.active.col };
    }
    const targetCol = range.c1;
    const sourceCol = targetCol - 1;
    if (sourceCol < 1) {
      toast("Flash Fill needs a source column to the left.");
      return;
    }
    if (range.r2 - range.r1 < 1) {
      toast("Select at least two rows to fill.");
      return;
    }
    const sample = cellAt(sheet, range.r1, targetCol)?.v;
    const source = cellAt(sheet, range.r1, sourceCol)?.v;
    let transform = "digits";
    if (sample !== undefined && source !== undefined) {
      const left = String(source);
      const right = String(sample);
      if (right === left.toUpperCase() && right !== left) transform = "upper";
      else if (right === left.toLowerCase() && right !== left) transform = "lower";
      else if (right === left.split(/\s+/)[0]) transform = "first";
      else if (right === left.split(/\s+/).slice(-1)[0]) transform = "last";
      else if (right === (left.match(/\d+(?:[.,]\d+)?/) || [""])[0]) transform = "digits";
      else if (right === left.replace(/[^\p{L}\s]/gu, "").trim()) transform = "letters";
      else transform = "upper";
    }
    let filled = 0;
    app.mutate("flash-fill", () => {
      for (let row = range.r1; row <= range.r2; row += 1) {
        const leftValue = cellAt(sheet, row, sourceCol)?.v;
        if (leftValue === undefined || leftValue === null || leftValue === "") continue;
        const text = String(leftValue);
        let value;
        if (transform === "upper") value = text.toUpperCase();
        else if (transform === "lower") value = text.toLowerCase();
        else if (transform === "first") value = text.split(/\s+/)[0] || "";
        else if (transform === "last") value = text.split(/\s+/).slice(-1)[0] || "";
        else if (transform === "letters") value = text.replace(/[^\p{L}\s]/gu, "").trim();
        else value = (text.match(/\d+(?:[.,]\d+)?/) || [""])[0];
        if (!value) continue;
        const cell = ensureCell(sheet, addr(row, targetCol));
        cell.v = value;
        delete cell.f;
        filled += 1;
      }
    }, { layout: false });
    toast(filled ? "Flash Fill filled " + filled + " cells." : "Flash Fill could not find a pattern.");
  }

  /* ------------------------------------------------------------------ view */
  zoomToSelection() {
    const grid = this.app.grid;
    const range = grid.sel;
    const width = grid.colPrefix[range.c2] - grid.colPrefix[range.c1 - 1];
    const height = grid.rowPrefix[range.r2] - grid.rowPrefix[range.r1 - 1];
    if (!(width > 0) || !(height > 0)) {
      toast("Select a range first.");
      return;
    }
    const value = Math.round(Math.min((grid.viewportWidth - 12) / width, (grid.viewportHeight - 12) / height) * 100);
    grid.setZoom(Math.max(40, Math.min(200, value)));
  }

  splitPanes() {
    const grid = this.app.grid;
    const sheet = grid.sheet();
    if (sheet.freeze?.rows || sheet.freeze?.cols) {
      grid.unfreezePanes();
      return;
    }
    grid.setFreeze(grid.active.row - 1, grid.active.col - 1);
    if (grid.active.row > 1 || grid.active.col > 1) toast("Panes split at " + grid.activeAddress() + ".");
    else toast("Select a cell below or to the right of the split.");
  }

  /* ------------------------------------------------------------------ help */
  feedbackDialog() {
    showDialog({
      title: "Send Feedback",
      size: "narrow",
      body: el("div", {}, [
        el("p", { class: "dialog-note", text: FEEDBACK_TEXT, style: "margin:0" }),
      ]),
      buttons: [{ label: "Close", kind: "primary" }],
    });
  }
}

function checkHeaderRow(sheet, range) {
  const first = [];
  const second = [];
  for (let col = range.c1; col <= range.c2; col += 1) {
    const top = cellAt(sheet, range.r1, col);
    const next = cellAt(sheet, range.r1 + 1, col);
    first.push(top?.v);
    second.push(next?.v);
  }
  if (!first.some((value) => value !== undefined && value !== "")) return false;
  return first.every((value) => typeof value === "string")
    && second.some((value) => typeof value === "number");
}
