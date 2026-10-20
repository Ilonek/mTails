#!/usr/bin/env python3
"""Automated end-to-end UI test for mTails.

Boots the real Qt/Chromium application, drives the workbook through
``window.app``, saves an XLSX, reopens it, exports CSV and PDF, exercises
autosave recovery and reports a JSON summary prefixed with ``UITEST:``.

Run with:  python3 tests/ui_test.py [output-dir]
"""

from __future__ import annotations

import json
import os
import sys
import time

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

OUT_DIR = os.path.abspath(sys.argv[1]) if len(sys.argv) > 1 else os.path.join(BASE_DIR, "tests", "out")
os.makedirs(OUT_DIR, exist_ok=True)

# Keep settings, recent files and recovery snapshots out of the real profile.
os.environ["XDG_DATA_HOME"] = os.path.join(OUT_DIR, "settings")

from PySide6.QtCore import QCoreApplication, QTimer  # noqa: E402
from PySide6.QtWidgets import QApplication  # noqa: E402

from mtails.app import MainWindow  # noqa: E402

JS_TEST = r"""
(async () => {
  const out = { checks: {}, tabs: {}, dialogs: {}, notes: [] };
  const errors = [];
  window.addEventListener("error", (event) => errors.push(String(event.message)));
  window.addEventListener("unhandledrejection", (event) => errors.push("rejection: " + String(event.reason)));
  const q = (selector) => document.querySelector(selector);
  const qa = (selector) => Array.from(document.querySelectorAll(selector));
  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
  const check = (name, ok, got) => {
    out.checks[name] = ok ? true : { got: got === undefined ? "false" : String(got) };
    return ok;
  };
  const stage = (name) => {
    out.stage = name;
    window.__stage = name;
  };
  const colIndex = (letters) => letters.split("").reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0);
  const splitAddr = (address) => {
    const match = /^([A-Z]+)(\d+)$/.exec(String(address || "").toUpperCase());
    return match ? { col: colIndex(match[1]), row: Number(match[2]) } : null;
  };

  let A = null;
  try {
    for (let attempt = 0; attempt < 160; attempt += 1) {
      A = window.app;
      if (A && A.booted) break;
      await sleep(100);
    }
    if (!A || !A.booted) throw new Error("the application did not boot");
    window.__uitest = null;
    A.setLanguage("en");
    await sleep(200);

    const grid = A.grid;
    const put = (address, text) => {
      const parsed = splitAddr(address);
      if (!parsed) throw new Error("bad address " + address);
      grid.selectCell(parsed.row, parsed.col);
      grid.setActiveValue(String(text), { move: "none" });
    };
    const raw = (address) => grid.sheet().cells[address] || null;
    const valueOf = (address, sheetIndex = grid.sheetIndex) => {
      const parsed = splitAddr(address);
      if (!parsed) return null;
      return A.engine.cellValue(A.model.sheets[sheetIndex], parsed.row, parsed.col);
    };
    const numeric = (address) => {
      const value = valueOf(address);
      return typeof value === "number" ? value : NaN;
    };
    const bridge = A.file.getBridge();
    const bridgeCall = (name, ...args) => new Promise((resolve) => {
      try {
        bridge[name](...args, (value) => {
          try { resolve(typeof value === "string" ? JSON.parse(value) : value); }
          catch { resolve({ ok: false, error: "invalid bridge payload" }); }
        });
      } catch (error) {
        resolve({ ok: false, error: String(error) });
      }
    });

    /* ------------------------------------------------------------ ribbon */
    stage("ribbon");
    for (const tab of qa(".rtab[data-tab]")) {
      tab.click();
      await sleep(35);
      out.tabs[tab.dataset.tab] = {
        label: tab.textContent.trim(),
        groups: qa("#ribbon-panels .rgroup").length,
        buttons: qa("#ribbon-panels .rbtn").length,
        titled: qa("#ribbon-panels .rgroup-title").filter((node) => node.textContent.trim()).length,
      };
    }
    q('.rtab[data-tab="home"]').click();
    await sleep(60);
    check("tabCount", Object.keys(out.tabs).length === 8, Object.keys(out.tabs).join(","));
    check("allTabsRender", Object.values(out.tabs).every((tab) => tab.groups > 0 && tab.buttons > 0));
    check("allGroupsTitled", Object.values(out.tabs).every((tab) => tab.titled === tab.groups));
    const homeTab = out.tabs.home || { buttons: 0 };
    check("homeButtons", homeTab.buttons >= 25, homeTab.buttons);

    const ribbonButtons = qa("#ribbon-panels .rbtn");
    const stuck = ribbonButtons.filter((node) => node.disabled)
      .map((node) => (node.title || "").split(" (")[0]);
    check("ribbonEnabled", ribbonButtons.length >= 25
      && stuck.every((name) => name === "Undo" || name === "Redo"),
      ribbonButtons.length + " buttons, disabled: " + stuck.join(","));
    grid.selectCell(1, 1);
    const boldButton = ribbonButtons.find((node) => (node.title || "").startsWith("Bold"));
    if (boldButton) boldButton.click();
    await sleep(40);
    check("ribbonClick", !!raw("A1")?.s?.b,
      boldButton ? "disabled=" + boldButton.disabled : "missing");
    if (boldButton) boldButton.click();
    await sleep(30);

    const themeBefore = document.body.dataset.theme || "light";
    q('.rtab[data-tab="view"]').click();
    await sleep(40);
    const themeButton = qa("#ribbon-panels .rbtn")
      .find((node) => (node.title || "").startsWith("Switch between"));
    if (themeButton) themeButton.click();
    await sleep(50);
    const themeAfter = document.body.dataset.theme;
    check("themeToggle", themeAfter === (themeBefore === "dark" ? "light" : "dark")
      && !!themeButton && themeButton.classList.contains("active") === (themeAfter === "dark"),
      themeBefore + "->" + themeAfter);
    if (themeButton) themeButton.click();
    await sleep(30);
    q('.rtab[data-tab="home"]').click();
    await sleep(40);

    let hovered = 0;
    for (const tab of qa(".rtab[data-tab]")) {
      tab.click();
      await sleep(20);
      for (const button of qa("#ribbon-panels .rbtn")) {
        button.dispatchEvent(new MouseEvent("mouseover", { bubbles: true }));
        hovered += 1;
      }
    }
    out.notes.push("hovered " + hovered + " ribbon buttons");
    q('.rtab[data-tab="home"]').click();
    await sleep(40);

    /* ----------------------------------------------------------- dialogs */
    stage("dialogs");
    A.insertSheet("DialogData");
    await sleep(50);
    const dialogCommands = {
      formatCells: "formatCells",
      rowHeight: "rowHeight",
      columnWidth: "colWidth",
      fillSeries: "fillSeries",
      sortDialog: "sortDialog",
      goTo: "goto",
      tableDialog: "insert.table",
      iconsDialog: "insert.icons",
      symbolDialog: "insert.symbol",
      textBoxDialog: "insert.textBox",
      headerFooterDialog: "insert.headerFooter",
      hyperlinkDialog: "insert.link",
      sparklineDialog: "insert.sparkline",
      chartDialog: "chartDialog",
      functionDialog: "functionDialog",
      nameManager: "name.manager",
      defineName: "name.define",
      auditPrecedents: "audit.precedents",
      auditDependents: "audit.dependents",
      textToColumns: "textToColumns",
      removeDuplicates: "removeDuplicates",
      validation: "validation",
      subtotal: "subtotal",
      cellStyle: "cellStyleDialog",
      conditional: "conditionalDialog",
      conditionalRules: "conditionalRules",
      protectSheet: "protect.sheet",
      protectWorkbook: "protect.workbook",
      spellcheck: "spellcheck",
      workbookStats: "workbook.stats",
      accessibility: "accessibility",
      zoom: "zoom",
      pageSetup: "pageSetup",
      help: "help",
      shortcuts: "shortcuts",
      whatsNew: "whatsNew",
      about: "about",
      renameSheet: "sheet.rename",
      deleteSheet: "sheet.delete",
      moveSheet: "sheet.move",
      sheetColor: "sheet.color",
      pasteSpecial: "pasteSpecial",
    };
    for (const [name, command] of Object.entries(dialogCommands)) {
      const before = qa("#dialog-layer .dialog").length;
      let error = "";
      try {
        if (name === "auditPrecedents") put("D1", "=SUM(B1:C1)");
        if (name === "auditPrecedents" || name === "auditDependents") {
          const target = splitAddr(name === "auditPrecedents" ? "D1" : "B1");
          grid.selectCell(target.row, target.col);
        }
        if (!A.execute(command)) throw new Error("command not handled");
      } catch (failure) {
        error = String(failure);
      }
      await sleep(25);
      const dialogs = qa("#dialog-layer .dialog");
      const opened = dialogs.length - before;
      const title = dialogs.length ? dialogs[dialogs.length - 1].querySelector(".dialog-head h2").textContent.trim() : "";
      for (let guard = 0; guard < 4 && qa("#dialog-layer .dialog").length > before; guard += 1) {
        const open = qa("#dialog-layer .dialog");
        open[open.length - 1].querySelector(".dialog-close")?.click();
        await sleep(15);
      }
      const closed = qa("#dialog-layer .dialog").length === before;
      out.dialogs[name] = { opened, title, closed, error };
    }
    const dialogTotal = Object.keys(out.dialogs).length;
    check("dialogCount", dialogTotal >= 30, dialogTotal);
    check("dialogsOpen", Object.values(out.dialogs).every((entry) => entry.opened === 1 && entry.closed && !entry.error),
      Object.entries(out.dialogs).filter(([, entry]) => entry.opened !== 1 || !entry.closed || entry.error)
        .map(([name]) => name).join(","));
    A.execute("audit.remove");
    A.activateSheet(0);
    A.removeSheet(1);
    A.setDirty(false);
    await sleep(50);

    /* --------------------------------------------------------- backstage */
    stage("backstage");
    A.openBackstage("new");
    await sleep(80);
    check("backstageOpen", !q("#backstage").classList.contains("hidden"));
    check("templateCount", qa("#template-grid .template-card").length >= 6, qa("#template-grid .template-card").length);
    const panelState = {};
    for (const panel of ["new", "open", "info", "save", "saveas", "print", "export", "options", "close"]) {
      A.showBsPanel(panel);
      await sleep(35);
      const section = q(`#backstage section[data-panel="${panel}"]`);
      panelState[panel] = !!section && !section.classList.contains("hidden");
    }
    check("backstagePanels", Object.values(panelState).every(Boolean), JSON.stringify(panelState));
    A.showBsPanel("info");
    await sleep(30);
    check("infoStats", /\d/.test(q("#info-sheets")?.textContent || ""), q("#info-sheets")?.textContent);
    A.showBsPanel("print");
    await sleep(90);
    check("printPreview", !!q("#print-preview .print-page"));
    A.showBsPanel("options");
    await sleep(50);
    check("optionsRendered", qa("#opt-font option").length > 3 && !!q("#opt-save"));
    A.showBsPanel("new");
    await sleep(30);
    q("#template-grid .template-card:nth-child(2)")?.click();
    await sleep(700);
    check("templateApplied", Object.keys(A.model.sheets[0].cells || {}).length > 0
      && A.state.title === "Monthly budget",
      Object.keys(A.model.sheets[0].cells || {}).length + "/" + A.state.title);
    await A.newWorkbook();
    await sleep(120);
    check("blankWorkbook", Object.keys(A.model.sheets[0].cells || {}).length === 0);
    A.closeBackstage();
    check("backstageClosed", q("#backstage").classList.contains("hidden"));

    /* ----------------------------------------------------------- palette */
    stage("palette");
    A.openPalette();
    await sleep(50);
    const paletteItems = qa("#menu-layer .palette-item").length;
    check("paletteOpen", !!q("#menu-layer .palette") && paletteItems > 30, paletteItems);
    const paletteInput = q("#menu-layer .palette input");
    if (paletteInput) {
      paletteInput.value = "chart";
      paletteInput.dispatchEvent(new Event("input", { bubbles: true }));
      await sleep(40);
    }
    const filtered = qa("#menu-layer .palette-item").length;
    check("paletteFilter", filtered >= 1 && filtered < paletteItems, filtered);
    A.closePalette();
    check("paletteClosed", !q("#menu-layer .palette"));

    /* -------------------------------------------------------- data entry */
    stage("data");
    grid.selectCell(1, 1);
    const headers = [
      ["A1", "Product"], ["B1", "Region"], ["C1", "Units"], ["D1", "Price"], ["E1", "Total"],
      ["A2", "Widget"], ["B2", "North"], ["C2", "12"], ["D2", "4.5"], ["E2", "=C2*D2"],
      ["A3", "Gadget"], ["B3", "South"], ["C3", "8"], ["D3", "7.25"], ["E3", "=C3*D3"],
      ["A4", "Gizmo"], ["B4", "North"], ["C4", "5"], ["D4", "10"], ["E4", "=C4*D4"],
      ["F1", "=SUM(E2:E4)"],
      ["F2", "=AVERAGE(E2:E4)"],
      ["F3", '=IF(C2>10,"High","Low")'],
      ["F4", '=VLOOKUP("Gizmo",A2:E4,5,FALSE)'],
      ["F5", '=INDEX(E2:E4,MATCH("Gizmo",A2:A4,0))'],
      ["F6", "=SUM($E$2:$E$4)"],
      ["G1", "=1/0"],
      ["G2", "=SUM(A1:A2)"],
      ["H1", "=H2"],
      ["H2", "=H1"],
    ];
    for (const [address, text] of headers) put(address, text);
    await sleep(80);
    check("entryText", raw("A1")?.v === "Product", raw("A1")?.v);
    check("entryNumber", raw("C2")?.v === 12, raw("C2")?.v);
    check("sum", Math.abs(numeric("F1") - 162) < 1e-9, numeric("F1"));
    check("average", Math.abs(numeric("F2") - 54) < 1e-9, numeric("F2"));
    check("ifText", valueOf("F3") === "High", valueOf("F3"));
    check("vlookup", Math.abs(numeric("F4") - 50) < 1e-9, numeric("F4"));
    check("indexMatch", Math.abs(numeric("F5") - 50) < 1e-9, numeric("F5"));
    check("absoluteRef", Math.abs(numeric("F6") - 162) < 1e-9, numeric("F6"));
    check("divisionError", raw("G1")?.v === "#DIV/0!", raw("G1")?.v);
    check("sumIgnoresText", numeric("G2") === 0, numeric("G2"));
    check("circularError", raw("H1")?.v === "#CIRC!" && raw("H2")?.v === "#CIRC!",
      (raw("H1")?.v || "") + "/" + (raw("H2")?.v || ""));
    check("dirtyAfterEdits", A.state.dirty === true);

    /* ------------------------------------------------------------ sheets */
    stage("sheets");
    const firstName = A.model.sheets[0].name;
    A.insertSheet("Data");
    await sleep(60);
    check("sheetAdded", A.model.sheets.length === 2 && A.model.sheets[1].name === "Data",
      A.model.sheets.map((sheet) => sheet.name).join(","));
    put("A1", "20");
    put("B1", "=SUM(" + firstName + "!E2:E4)");
    await sleep(30);
    check("sheetReference", Math.abs(numeric("B1") - 162) < 1e-9, numeric("B1"));
    A.renameSheet(0, "Sales");
    A.renameSheet(1, "Metrics");
    await sleep(30);
    check("sheetRenamed", A.model.sheets.map((sheet) => sheet.name).join("|") === "Sales|Metrics",
      A.model.sheets.map((sheet) => sheet.name).join("|"));
    check("referenceRewritten", String(raw("B1")?.f || "").includes("Sales!"), raw("B1")?.f);
    A.activateSheet(0);
    await sleep(40);
    check("referenceStill165", Math.abs(numeric("F1") - 162) < 1e-9, numeric("F1"));
    A.moveSheet(1, 0);
    await sleep(30);
    check("sheetMoved", A.model.sheets.map((sheet) => sheet.name).join("|") === "Metrics|Sales",
      A.model.sheets.map((sheet) => sheet.name).join("|"));
    A.moveSheet(0, 2);
    await sleep(30);
    check("sheetMovedBack", A.model.sheets.map((sheet) => sheet.name).join("|") === "Sales|Metrics",
      A.model.sheets.map((sheet) => sheet.name).join("|"));
    A.duplicateSheet(1);
    await sleep(40);
    check("sheetDuplicated", A.model.sheets.length === 3
      && A.model.sheets[2].name !== "Metrics", A.model.sheets.map((sheet) => sheet.name).join("|"));
    A.removeSheet(2);
    await sleep(30);
    check("sheetRemoved", A.model.sheets.length === 2, A.model.sheets.length);
    A.setSheetColor(1, "#FFC000");
    check("sheetColor", A.model.sheets[1].tabColor === "#FFC000", A.model.sheets[1].tabColor);
    A.setSheetColor(1, "");
    // The sheet shuffle above leaves the moved sheet active, exactly like Excel.
    A.activateSheet(0);
    await sleep(40);
    check("activeAfterSheets", A.model.activeSheet === 0 && grid.sheetIndex === 0,
      A.model.activeSheet + "/" + grid.sheetIndex);

    /* -------------------------------------------------------- formatting */
    stage("formatting");
    grid.selectRange({ r1: 1, c1: 1, r2: 1, c2: 5 });
    grid.setStyle({ b: true, fill: "#DDEBF7", color: "#1F3864" });
    grid.setAlignment("center");
    grid.applyBorders("all", { style: "thin", color: "#000000" });
    grid.selectRange({ r1: 2, c1: 4, r2: 4, c2: 4 });
    grid.setNumberFormat("$#,##0.00");
    await sleep(30);
    const headerStyle = raw("A1")?.s || {};
    check("boldStyle", headerStyle.b === true, JSON.stringify(headerStyle));
    check("fillStyle", headerStyle.fill === "#DDEBF7", headerStyle.fill);
    check("alignment", headerStyle.align === "center", headerStyle.align);
    check("border", headerStyle.border?.t?.style === "thin", JSON.stringify(headerStyle.border || {}));
    check("numberFormat", String(raw("D2")?.s?.numFmt || "").includes("$"), raw("D2")?.s?.numFmt);
    grid.selectColumn(1);
    grid.setSelectedColumnWidth(140);
    await sleep(20);
    check("columnWidth", Number(A.model.sheets[0].cols?.A?.width || A.model.sheets[0].cols?.A?.w) === 140,
      JSON.stringify(A.model.sheets[0].cols?.A || {}));

    /* ---------------------------------------------------------- clipboard */
    stage("clipboard");
    grid.selectRange({ r1: 1, c1: 1, r2: 4, c2: 2 });
    const copied = grid.copySelection();
    check("clipboardText", copied.text.split("\n").length === 4 && copied.text.includes("Product"),
      copied.text.slice(0, 60));
    grid.selectCell(20, 1);
    grid.pasteText(copied.text);
    await sleep(30);
    check("pastedHeader", raw("A20")?.v === "Product", raw("A20")?.v);
    check("pastedCell", raw("B21")?.v === "Nord" || raw("B21")?.v === "North", raw("B21")?.v);

    /* --------------------------------------------------------- structure */
    stage("structure");
    grid.selectCell(4, 1);
    grid.insertRow(1, "above");
    await sleep(30);
    check("rowInserted", raw("A5")?.v === "Gizmo" && !raw("A4")?.v, (raw("A4")?.v || "") + "->" + (raw("A5")?.v || ""));
    A.undo();
    await sleep(30);
    check("rowInsertUndo", raw("A4")?.v === "Gizmo", raw("A4")?.v);
    grid.selectCell(1, 2);
    grid.insertColumn(1, "left");
    await sleep(30);
    check("columnInserted", raw("C1")?.v === "Region", raw("C1")?.v);
    A.undo();
    await sleep(30);
    check("columnInsertUndo", raw("B1")?.v === "Region", raw("B1")?.v);

    /* --------------------------------------------------------- sort/filter */
    stage("sort");
    grid.selectRange({ r1: 1, c1: 1, r2: 4, c2: 5 });
    check("sorted", grid.applySort([{ col: 3, order: "desc" }], { hasHeader: true }));
    await sleep(30);
    check("sortOrder", raw("C2")?.v === 12 && raw("C4")?.v === 5,
      (raw("C2")?.v || "") + "/" + (raw("C4")?.v || ""));
    const hidden = grid.applyFilter({ r1: 1, c1: 1, r2: 4, c2: 5 }, [{ col: 2, operator: "equals", value: "South" }]);
    await sleep(30);
    // Two of the three data rows are North, so filtering for South hides two.
    check("filterHidden", hidden === 2 && A.model.sheets[0].filters?.hidden?.length === 2, hidden);
    grid.clearFilter();
    await sleep(20);
    check("filterCleared", !A.model.sheets[0].filters);

    /* ------------------------------------------------------- find/replace */
    stage("find");
    A.openFind(true);
    await sleep(30);
    q("#find-input").value = "North";
    A.runFind();
    await sleep(30);
    check("findCount", A.findState.count >= 2, A.findState.count);
    q("#replace-input").value = "Nord";
    A.replaceAll();
    await sleep(40);
    let nord = 0;
    for (const cell of Object.values(A.grid.sheet().cells)) {
      if (cell && cell.v === "Nord") nord += 1;
    }
    check("replaceAll", nord >= 2, nord);
    A.closeFind();

    /* ----------------------------------------------------------- comments */
    stage("comments");
    grid.selectCell(2, 1);
    A.newComment();
    await sleep(30);
    q("#comment-input").value = "Automated test comment";
    A.postComment();
    await sleep(40);
    const commentCount = (A.grid.sheet().comments || []).length;
    check("commentAdded", commentCount >= 1, commentCount);
    check("commentPin", qa("#comment-pins .comment-pin").length >= 1, qa("#comment-pins .comment-pin").length);
    check("commentCard", qa("#comments-list .comment-card").length >= 1);
    A.toggleComments(false);

    /* ------------------------------------------------------ chart/objects */
    stage("chart");
    grid.selectRange({ r1: 1, c1: 1, r2: 4, c2: 3 });
    A.execute("insert.chart:column");
    await sleep(120);
    check("chartModel", (A.model.sheets[0].charts || []).length >= 1, (A.model.sheets[0].charts || []).length);
    check("chartRendered", !!q("#overlay-layer .sheet-chart .chart-body svg"));
    grid.selectCell(8, 7);
    A.addTextBox("Quarterly results", { w: 220, h: 80 });
    await sleep(60);
    check("textBox", !!q("#overlay-layer .sheet-textbox"));

    /* ------------------------------------------------------ views/freeze */
    stage("views");
    grid.selectCell(2, 2);
    grid.freezePanes();
    await sleep(40);
    check("freeze", A.model.sheets[0].freeze?.rows === 1 && A.model.sheets[0].freeze?.cols === 1,
      JSON.stringify(A.model.sheets[0].freeze || {}));
    A.execute("view.pagebreak");
    await sleep(150);
    check("pageBreakView", document.body.dataset.view === "pagebreak");
    A.execute("view.pageLayout");
    await sleep(200);
    check("pageLayoutView", document.body.dataset.view === "layout"
      && qa("#page-layout-layer .layout-holder").length >= 1);
    A.execute("view.normal");
    await sleep(80);
    check("normalView", document.body.dataset.view === "normal");
    A.execute("gridlines");
    await sleep(20);
    check("gridlinesOff", A.model.sheets[0].showGridlines === false);
    A.execute("gridlines");
    await sleep(20);
    check("gridlinesOn", A.model.sheets[0].showGridlines !== false);
    grid.setZoom(120);
    await sleep(20);
    check("zoom120", grid.zoom === 120, grid.zoom);
    A.execute("zoom100");
    await sleep(20);
    check("zoom100", grid.zoom === 100, grid.zoom);

    /* -------------------------------------------------------- undo/redo */
    stage("history:put");
    grid.selectCell(1, 10);
    for (let index = 1; index <= 12; index += 1) put("J" + index, String(index));
    await sleep(30);
    stage("history:undo");
    for (let index = 0; index < 10; index += 1) A.undo();
    await sleep(60);
    check("undoTen", raw("J2")?.v === 2 && !raw("J3")?.v,
      (raw("J2")?.v || "") + "/" + (raw("J3")?.v || ""));
    stage("history:redo");
    for (let index = 0; index < 10; index += 1) A.redo();
    await sleep(60);
    check("redoTen", raw("J12")?.v === 12, raw("J12")?.v);

    /* ---------------------------------------------------------- language */
    stage("language");
    const titleBefore = A.state.title;
    const sheetsBefore = A.model.sheets.map((sheet) => sheet.name).join("|");
    A.setLanguage("pl");
    await sleep(260);
    const polishLabel = q('.rtab[data-tab="home"]')?.textContent.trim() || "";
    check("polishRibbon", polishLabel !== "Home" && polishLabel.length > 0, polishLabel);
    check("polishKeepsContent", raw("A1")?.v === "Product" && A.state.title === titleBefore
      && A.model.sheets.map((sheet) => sheet.name).join("|") === sheetsBefore,
      (raw("A1")?.v || "") + "/" + A.state.title);
    check("polishStatus", (q("#sb-language")?.textContent || "").includes("Polski"), q("#sb-language")?.textContent);
    A.setLanguage("en");
    await sleep(260);
    check("englishRibbon", q('.rtab[data-tab="home"]')?.textContent.trim() === "Home",
      q('.rtab[data-tab="home"]')?.textContent);
    A.applyTheme("dark");
    await sleep(40);
    check("darkTheme", document.body.dataset.theme === "dark");
    A.applyTheme("light");
    await sleep(40);
    check("lightTheme", document.body.dataset.theme === "light");

    /* -------------------------------------------------------- print panel */
    stage("print");
    A.openBackstage("print");
    await sleep(120);
    check("printSummary", (q("#print-summary")?.textContent || "").length > 3, q("#print-summary")?.textContent);
    check("printPreviewFilled", !!q("#print-preview .print-page .print-table"));
    A.closeBackstage();

    /* ------------------------------------------------------- save/reopen */
    stage("save:write");
    const xlsxPath = "__OUT__/uitest.xlsx";
    A.state.path = xlsxPath;
    A.state.format = "xlsx";
    const saved = await A.file.saveFile({});
    await sleep(150);
    stage("save:reopen");
    check("xlsxSaved", saved === true, saved);
    check("savedClean", A.state.dirty === false);
    const reopened = await A.file.openFile(xlsxPath);
    await sleep(250);
    check("xlsxReopened", reopened === true);
    check("roundTripSheets", A.model.sheets.map((sheet) => sheet.name).join("|") === "Sales|Metrics",
      A.model.sheets.map((sheet) => sheet.name).join("|"));
    check("roundTripValue", raw("A1")?.v === "Product", raw("A1")?.v);
    check("roundTripFormula", Math.abs(numeric("F1") - 162) < 1e-9, numeric("F1"));
    check("roundTripStyle", raw("A1")?.s?.b === true, JSON.stringify(raw("A1")?.s || {}));
    check("roundTripComment", (A.model.sheets[0].comments || []).length >= 1,
      (A.model.sheets[0].comments || []).length);
    check("roundTripFreeze", A.model.sheets[0].freeze?.rows === 1 && A.model.sheets[0].freeze?.cols === 1,
      JSON.stringify(A.model.sheets[0].freeze || {}));
    check("roundTripTitle", A.state.title === "uitest" && A.state.path.endsWith("uitest.xlsx"),
      A.state.title + "/" + A.state.path);
    check("roundTripChart", (A.model.sheets[0].charts || []).length >= 1,
      (A.model.sheets[0].charts || []).length);

    /* -------------------------------------------------------------- export */
    stage("export:csv");
    const csvPath = "__OUT__/uitest.csv";
    const csv = await A.file.exportAs("csv", csvPath);
    await sleep(80);
    check("csvExported", csv === true);

    stage("export:pdf");
    const pdfPath = "__OUT__/uitest.pdf";
    const pdfStart = Date.now();
    const pdf = await A.file.exportPdf(pdfPath);
    out.pdf = { ok: pdf, elapsed: Date.now() - pdfStart };
    check("pdfExported", pdf === true);

    /* ----------------------------------------------------------- autosave */
    stage("autosave:write");
    const auto = await bridgeCall("autosave", JSON.stringify(A.model), "uitest-auto", "UITest", "");
    check("autosaveWritten", auto.ok === true, JSON.stringify(auto));
    const loaded = await bridgeCall("loadAutosave", "uitest-auto");
    check("autosaveLoaded", loaded.ok === true && loaded.model?.sheets?.[0]?.cells?.A1?.v === "Product",
      loaded.ok);
    A.autosaves = [{ id: "uitest-auto", name: "UITest", source: "", savedAt: Date.now() / 1000 }];
    A.renderRecovery();
    await sleep(40);
    check("recoveryCard", !!q("#recovery-box .recovery-card") && !q("#recovery-box").classList.contains("hidden"));
    stage("autosave:recover");
    await A.recoverAutosave(A.autosaves[0]);
    await sleep(200);
    check("recovered", A.state.title === "UITest" && A.state.dirty === true,
      A.state.title + "/" + A.state.dirty);
    stage("autosave:done");
    await bridgeCall("discardAutosave", "uitest-auto");
    stage("autosave:discarded");

    const smoke = window.__smoke ? window.__smoke() : { ok: false };
    check("smoke", smoke.ok === true && (smoke.errors || []).length === 0, JSON.stringify(smoke.errors || []));
    stage("smoke:done");

    out.ok = Object.values(out.checks).every((value) => value === true);
    out.failed = Object.keys(out.checks).filter((name) => out.checks[name] !== true);
    out.state = {
      title: A.state.title,
      dirty: A.state.dirty,
      sheets: A.model.sheets.map((sheet) => sheet.name),
      cells: Object.keys(A.model.sheets[0].cells || {}).length,
      charts: (A.model.sheets[0].charts || []).length,
      comments: (A.model.sheets[0].comments || []).length,
      revisions: A.revision,
    };
  } catch (error) {
    out.fatal = String(error);
    out.stack = String(error && error.stack || "");
  }
  out.errors = errors;
  window.__uitest = out;
})();
"""


def main() -> int:
    app = QApplication(sys.argv[:1])
    app.setApplicationName("mTailsTest")
    QCoreApplication.setApplicationName("mTailsTest")
    QCoreApplication.setOrganizationName("mTailsTest")

    window = MainWindow(startup_file="", smoke_path="")
    window.resize(1440, 900)
    window.show()

    state = {"started": False, "start": 0.0, "result": None, "stage": None, "busy": False}

    def finish(data: dict) -> None:
        if state["result"] is not None:
            return
        state["result"] = data

        # Grabbing the window inside a runJavaScript callback can stall the
        # event loop, so the capture runs from the main event loop instead.
        def complete() -> None:
            try:
                shot = os.path.join(OUT_DIR, "ui_test.png")
                window.grab().save(shot, "PNG")
                data["screenshot"] = shot
            except Exception as exc:  # noqa: BLE001
                data["screenshotError"] = str(exc)
            print("UITEST:" + json.dumps(data, ensure_ascii=False), flush=True)
            try:
                check_files(data)
            except Exception as exc:  # noqa: BLE001
                print("UITESTFILESERROR:" + str(exc), flush=True)
            # QtWebEngine occasionally keeps the process alive after quit(), so
            # the harness falls back to a hard exit once the report is out.
            QTimer.singleShot(2500, lambda: os._exit(0))
            QTimer.singleShot(200, app.quit)

        QTimer.singleShot(0, complete)

    def inspect() -> None:
        if state["result"] is not None or state["busy"]:
            return
        state["busy"] = True

        def receive(value) -> None:
            state["busy"] = False
            if not value:
                return
            try:
                payload = json.loads(value)
            except Exception:
                return
            stage = payload.get("stage") or ""
            if stage and stage != state["stage"]:
                state["stage"] = stage
                print("UITESTSTAGE:%s errors=%s dirty=%s title=%s" % (
                    stage, payload.get("errors"), payload.get("dirty"), payload.get("title")), flush=True)
            if payload.get("done") and payload.get("data"):
                finish(payload["data"])

        window.view.page().runJavaScript(
            "JSON.stringify({stage: window.__stage || '', errors: (window.app && window.app.errors || []).length, "
            "dirty: !!(window.app && window.app.state.dirty), title: (window.app && window.app.state.title) || '', "
            "done: !!window.__uitest, data: window.__uitest || null})",
            0,
            receive,
        )

    def run_test() -> None:
        if state["started"]:
            return
        state["started"] = True
        state["start"] = time.time()
        print("UITESTSTART", flush=True)
        window.view.page().runJavaScript(JS_TEST.replace("__OUT__", OUT_DIR.replace("\\", "/")))

    def timeout() -> None:
        if state["result"] is not None:
            return
        print("UITEST:" + json.dumps({"ok": False, "error": "timeout", "stage": state.get("stage")}), flush=True)
        app.quit()

    ticker = QTimer()
    ticker.timeout.connect(inspect)
    ticker.start(900)
    QTimer.singleShot(6000, inspect)
    QTimer.singleShot(270000, timeout)

    def check_files(data: dict) -> None:
        info = {}
        xlsx = os.path.join(OUT_DIR, "uitest.xlsx")
        csv = os.path.join(OUT_DIR, "uitest.csv")
        pdf = os.path.join(OUT_DIR, "uitest.pdf")
        info["xlsx"] = os.path.exists(xlsx)
        info["csv"] = os.path.exists(csv)
        info["pdf"] = os.path.exists(pdf)
        if info["xlsx"]:
            info["xlsxSize"] = os.path.getsize(xlsx)
            try:
                from mtails.xlsx_reader import read_xlsx

                model = read_xlsx(xlsx)
                sheets = model.get("sheets") or []
                info["xlsxSheets"] = [sheet.get("name") for sheet in sheets]
                if sheets:
                    sales = sheets[0]
                    cells = sales.get("cells") or {}
                    info["xlsxA1"] = (cells.get("A1") or {}).get("v")
                    info["xlsxF1"] = (cells.get("F1") or {}).get("v")
                    info["xlsxFormula"] = (cells.get("F1") or {}).get("f")
                    info["xlsxBold"] = bool(((cells.get("A1") or {}).get("s") or {}).get("b"))
                    info["xlsxComments"] = len(sales.get("comments") or [])
                    info["xlsxCharts"] = len(sales.get("charts") or [])
                    info["xlsxFreeze"] = sales.get("freeze")
            except Exception as exc:  # noqa: BLE001
                info["xlsxReadError"] = str(exc)
        if info["csv"]:
            info["csvSize"] = os.path.getsize(csv)
            with open(csv, "r", encoding="utf-8") as handle:
                text = handle.read()
            info["csvHasHeader"] = "Product" in text
            info["csvHasReplacement"] = "Nord" in text
            info["csvLines"] = len(text.splitlines())
        if info["pdf"]:
            info["pdfSize"] = os.path.getsize(pdf)
            with open(pdf, "rb") as handle:
                info["pdfMagic"] = handle.read(5).decode("latin-1")
        print("UITESTFILES:" + json.dumps(info, ensure_ascii=False), flush=True)

    window.view.loadFinished.connect(lambda ok: QTimer.singleShot(1200, run_test))
    QTimer.singleShot(2500, run_test)

    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
