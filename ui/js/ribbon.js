/* Ribbon: the eight Excel-style tabs above the formula bar.
 *
 * Every control is built from the same primitives (group/button/combo) and
 * routed through host.execute(commandId, anchor).  The ribbon owns no
 * workbook state; update() only reflects the current selection back into the
 * toggle buttons. */

import { $, el, openMenu, colorMenu, closeMenus } from "./ui.js";
import { icon } from "./icons.js";
import { FUNCTION_NAMES } from "./formula.js";

const TABS = [
  { id: "file", label: "File", file: true },
  { id: "home", label: "Home" },
  { id: "insert", label: "Insert" },
  { id: "formulas", label: "Formulas" },
  { id: "data", label: "Data" },
  { id: "review", label: "Review" },
  { id: "view", label: "View" },
  { id: "help", label: "Help" },
];

export const CHART_TYPES = ["column", "bar", "line", "pie", "area", "scatter", "doughnut"];

const STYLE_PRESETS = [
  { id: "normal", label: "Normal", style: {} },
  { id: "title", label: "Title", style: { b: true, size: 18, color: "#2F5496" } },
  { id: "heading1", label: "Heading 1", style: { b: true, size: 15, color: "#2F5496" } },
  { id: "heading2", label: "Heading 2", style: { b: true, size: 12, color: "#2F5496" } },
  { id: "total", label: "Total", style: { b: true, border: { t: { style: "thin", color: "#000000" }, b: { style: "double", color: "#000000" } } } },
  { id: "good", label: "Good", style: { fill: "#C6EFCE", color: "#006100" } },
  { id: "bad", label: "Bad", style: { fill: "#FFC7CE", color: "#9C0006" } },
  { id: "neutral", label: "Neutral", style: { fill: "#FFEB9C", color: "#9C6500" } },
  { id: "note", label: "Note", style: { fill: "#FFFFCC", border: {
    t: { style: "thin", color: "#B2B2B2" }, r: { style: "thin", color: "#B2B2B2" },
    b: { style: "thin", color: "#B2B2B2" }, l: { style: "thin", color: "#B2B2B2" },
  } } },
  { id: "accent1", label: "Accent 1", style: { fill: "#D9E2F3", color: "#1F3864" } },
  { id: "accent2", label: "Accent 2", style: { fill: "#FBE4D5", color: "#833C0B" } },
  { id: "accent3", label: "Accent 3", style: { fill: "#E2EFDA", color: "#375623" } },
];

function formatLabel(format) {
  if (!format || format === "General") return "General";
  if (/^[$€£¥]/.test(format) || /\[\$/.test(format)) return "Currency";
  if (/%/.test(format)) return "Percentage";
  if (/[ymdhs]/i.test(format) && !/[0#]/.test(format.replace(/[ymdhs]/gi, ""))) return "Date";
  if (/#,##0/.test(format) && /\.0+/.test(format)) return "Comma";
  if (/#,##0/.test(format)) return "Comma";
  if (/^0\.0+$/.test(format)) return "Number";
  if (/^0+$/.test(format)) return "Number";
  return format.length > 14 ? format.slice(0, 13) + "…" : format;
}

export class Ribbon {
  constructor(host) {
    this.host = host;
    this.activeTab = "home";
    this.panels = new Map();
    this.byCmd = new Map();
    this.fontCombo = null;
    this.sizeCombo = null;
    this.formatLabel = null;
    this.lastFontColor = "#C00000";
    this.lastFillColor = "#FFFF00";
    this.borderColor = "#000000";
    this.borderStyle = "thin";
    this.boundScroll = null;
    this.fonts = host.fonts || [];
    this.mounted = false;
  }

  /* --------------------------------------------------------------- mount */
  mount() {
    if (this.mounted) return;
    const strip = $("#tabstrip");
    const body = $("#ribbon-panels");
    if (!strip || !body) return;
    for (const tab of TABS) {
      const button = el("button", {
        class: "rtab" + (tab.file ? " file" : ""),
        type: "button",
        role: "tab",
        text: tab.label,
        dataset: { tab: tab.id },
      });
      button.addEventListener("click", () => {
        if (tab.file) this.host.execute("backstage", button);
        else this.setTab(tab.id);
      });
      strip.append(button);
    }
    body.replaceChildren();
    const left = $("#ribbon-left");
    const right = $("#ribbon-right");
    const updateArrows = () => {
      if (!left || !right) return;
      left.classList.toggle("visible", body.scrollLeft > 2);
      right.classList.toggle("visible", body.scrollLeft + body.clientWidth < body.scrollWidth - 2);
    };
    left?.addEventListener("click", () => body.scrollBy({ left: -220, behavior: "smooth" }));
    right?.addEventListener("click", () => body.scrollBy({ left: 220, behavior: "smooth" }));
    this.boundScroll = updateArrows;
    body.addEventListener("scroll", updateArrows, { passive: true });
    window.addEventListener("resize", updateArrows);
    setTimeout(updateArrows, 60);
    this.setTab("home");
    this.mounted = true;
  }

  setTab(id) {
    if (!TABS.some((tab) => tab.id === id) || id === "file") id = "home";
    this.activeTab = id;
    for (const button of $("#tabstrip")?.querySelectorAll(".rtab") || []) {
      button.classList.toggle("active", button.dataset.tab === id);
    }
    const body = $("#ribbon-panels");
    body.replaceChildren();
    const panel = this.panels.get(id) || this.buildPanel(id);
    body.append(panel);
    this.scheduleArrows();
  }

  scheduleArrows() {
    setTimeout(() => {
      const body = $("#ribbon-panels");
      if (!body) return;
      const left = $("#ribbon-left");
      const right = $("#ribbon-right");
      left?.classList.toggle("visible", body.scrollLeft > 2);
      right?.classList.toggle("visible", body.scrollLeft + body.clientWidth < body.scrollWidth - 2);
    }, 20);
  }

  /* ------------------------------------------------------------ primitives */
  group(title, children) {
    return el("div", { class: "rgroup" }, [
      el("div", { class: "rgroup-body" }, children),
      el("div", { class: "rgroup-title", text: title }),
    ]);
  }

  column(children, two = false) {
    return el("div", { class: "rcol" + (two ? " two" : "") }, children);
  }

  button({ iconName, label, cmd, title = "", kind = "", caret = false, active = false, disabled = false }) {
    const className = "rbtn" + (kind ? " " + kind : "") + (active ? " active" : "") + (caret ? " split" : "");
    const node = el("button", { class: className, type: "button", title: title || label, disabled }, [
      el("span", { class: "ricon", html: icon(iconName) }),
      el("span", { class: "rlabel", text: label }),
    ]);
    if (caret) node.append(el("span", { class: "caret", html: icon("chevronDown") }));
    node.addEventListener("click", (event) => {
      event.preventDefault();
      this.host.execute(cmd, node, { caret: !!event.target.closest(".caret") });
    });
    this.register(cmd, node);
    return node;
  }

  iconButton({ iconName, cmd, title, active = false }) {
    const node = el("button", { class: "rbtn icon" + (active ? " active" : ""), type: "button", title }, [
      el("span", { class: "ricon", html: icon(iconName) }),
    ]);
    node.addEventListener("click", () => this.host.execute(cmd, node));
    this.register(cmd, node);
    return node;
  }

  split({ iconName, label, cmd, menu, kind = "", title = "" }) {
    const node = this.button({ iconName, label, cmd, kind, title, caret: true });
    const caret = node.querySelector(".caret");
    if (caret) {
      caret.addEventListener("click", (event) => {
        event.stopPropagation();
        event.preventDefault();
        this.showMenu(node, menu);
      });
    }
    node.addEventListener("contextmenu", (event) => {
      event.preventDefault();
      event.stopPropagation();
      this.showMenu(node, menu);
    });
    return node;
  }

  menuButton({ iconName, label, menu, title = "", kind = "" }) {
    const node = this.button({ iconName, label, cmd: "", title, kind, caret: true });
    node.addEventListener("click", (event) => {
      event.preventDefault();
      this.showMenu(node, menu);
    });
    return node;
  }

  showMenu(anchor, items) {
    const list = typeof items === "function" ? items() : items;
    if (!list?.length) return;
    openMenu(anchor, list, { minWidth: 186 });
  }

  combo({ value = "", options, onChange, className = "", width = 0, placeholder = "" }) {
    const input = el("input", { type: "text", value, placeholder, spellcheck: "false" });
    const caret = el("button", { class: "combo-btn", type: "button", html: icon("chevronDown"), title: "Choose" });
    const wrap = el("div", { class: "rcombo " + className }, [input, caret]);
    if (width) wrap.style.width = width + "px";
    const pick = (item) => {
      input.value = item;
      onChange(item);
    };
    const open = () => {
      openMenu(caret, (options || []).map((item) => ({
        label: String(item),
        onClick: () => pick(String(item)),
      })), { minWidth: 130 });
    };
    caret.addEventListener("click", (event) => {
      event.preventDefault();
      open();
    });
    input.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        onChange(input.value);
        input.blur();
      } else if (event.key === "Escape") {
        input.blur();
      } else if (event.key === "ArrowDown") {
        event.preventDefault();
        open();
      }
      event.stopPropagation();
    });
    input.addEventListener("blur", () => {
      if (input.value !== value) onChange(input.value);
    });
    return { wrap, input };
  }

  colorButton({ iconName, cmd, title, color, onPick, allowNone = true }) {
    const node = el("button", { class: "rbtn color-btn", type: "button", title }, [
      el("span", { class: "ricon", html: icon(iconName) }),
      el("span", { class: "color-bar" }),
      el("span", { class: "caret", html: icon("chevronDown") }),
    ]);
    const bar = node.querySelector(".color-bar");
    bar.style.background = color;
    node.addEventListener("click", (event) => {
      event.preventDefault();
      if (event.target.closest(".caret")) {
        colorMenu(node, {
          title: "Theme Colors",
          allowNone,
          onPick: (value) => {
            if (value) {
              bar.style.background = value;
              onPick(value, true);
            } else {
              onPick("", true);
            }
          },
        });
      } else {
        onPick(color, false);
      }
    });
    this.register(cmd, node);
    return node;
  }

  register(cmd, node) {
    if (!cmd) return;
    const list = this.byCmd.get(cmd) || [];
    list.push(node);
    this.byCmd.set(cmd, list);
  }

  setActive(cmd, active) {
    for (const node of this.byCmd.get(cmd) || []) node.classList.toggle("active", !!active);
  }

  setDisabled(cmd, disabled) {
    for (const node of this.byCmd.get(cmd) || []) node.disabled = !!disabled;
  }

  /* -------------------------------------------------------------- panels */
  buildPanel(id) {
    let panel = null;
    if (id === "home") panel = this.buildHome();
    else if (id === "insert") panel = this.buildInsert();
    else if (id === "formulas") panel = this.buildFormulas();
    else if (id === "data") panel = this.buildData();
    else if (id === "review") panel = this.buildReview();
    else if (id === "view") panel = this.buildView();
    else if (id === "help") panel = this.buildHelp();
    else panel = el("div", { class: "ribbon-panel" });
    this.panels.set(id, panel);
    return panel;
  }

  pasteMenu() {
    return [
      { icon: "paste", label: "Paste", shortcut: "Ctrl+V", onClick: () => this.host.execute("paste") },
      { type: "sep" },
      { icon: "pasteValues", label: "Values", onClick: () => this.host.execute("pasteValues") },
      { icon: "copy", label: "Formulas", onClick: () => this.host.execute("pasteFormulas") },
      { icon: "painter", label: "Formatting", onClick: () => this.host.execute("pasteFormats") },
      { label: "Transpose", onClick: () => this.host.execute("pasteTranspose") },
      { type: "sep" },
      { icon: "pasteSpecial", label: "Paste Special…", shortcut: "Ctrl+Alt+V", onClick: () => this.host.execute("pasteSpecial") },
    ];
  }

  bordersMenu() {
    const apply = (preset) => this.host.execute("borders:" + preset);
    return [
      { icon: "borderBottom", label: "Bottom Border", onClick: () => apply("bottom") },
      { icon: "borderTop", label: "Top Border", onClick: () => apply("top") },
      { icon: "borderLeft", label: "Left Border", onClick: () => apply("left") },
      { icon: "borderRight", label: "Right Border", onClick: () => apply("right") },
      { type: "sep" },
      { icon: "borderNone", label: "No Border", onClick: () => apply("none") },
      { icon: "borderAll", label: "All Borders", onClick: () => apply("all") },
      { icon: "borderOutside", label: "Outside Borders", onClick: () => apply("outside") },
      { icon: "borderThick", label: "Thick Box Border", onClick: () => apply("outsideThick") },
      { type: "sep" },
      { icon: "borderStyle", label: "More Borders…", onClick: () => this.host.execute("formatCells", null, { tab: "border" }) },
    ];
  }

  mergeMenu() {
    return [
      { icon: "mergeCenter", label: "Merge & Center", onClick: () => this.host.execute("mergeCenter") },
      { icon: "mergeAcross", label: "Merge Across", onClick: () => this.host.execute("mergeAcross") },
      { icon: "mergeCells", label: "Merge Cells", onClick: () => this.host.execute("mergeCells") },
      { type: "sep" },
      { icon: "unmerge", label: "Unmerge Cells", onClick: () => this.host.execute("unmerge") },
    ];
  }

  numberMenu() {
    return [
      { type: "head", label: "Number" },
      { label: "General", onClick: () => this.host.execute("numformat:General") },
      { label: "Number", onClick: () => this.host.execute("numformat:#,##0.00") },
      { label: "Currency", onClick: () => this.host.execute("numformat:$#,##0.00") },
      { label: "Accounting", onClick: () => this.host.execute("numformat:_($* #,##0.00_)") },
      { label: "Short Date", onClick: () => this.host.execute("numformat:m/d/yyyy") },
      { label: "Long Date", onClick: () => this.host.execute("numformat:dddd, mmmm d, yyyy") },
      { label: "Time", onClick: () => this.host.execute("numformat:h:mm:ss AM/PM") },
      { label: "Percentage", onClick: () => this.host.execute("numformat:0.00%") },
      { label: "Fraction", onClick: () => this.host.execute("numformat:# ?/?") },
      { label: "Scientific", onClick: () => this.host.execute("numformat:0.00E+00") },
      { label: "Text", onClick: () => this.host.execute("numformat:@") },
      { type: "sep" },
      { label: "More Number Formats…", onClick: () => this.host.execute("formatCells", null, { tab: "number" }) },
    ];
  }

  autosumMenu() {
    return [
      { icon: "sum", label: "Sum", onClick: () => this.host.execute("autosum:SUM") },
      { icon: "average", label: "Average", onClick: () => this.host.execute("autosum:AVERAGE") },
      { icon: "countNumbers", label: "Count Numbers", onClick: () => this.host.execute("autosum:COUNT") },
      { icon: "max", label: "Max", onClick: () => this.host.execute("autosum:MAX") },
      { icon: "min", label: "Min", onClick: () => this.host.execute("autosum:MIN") },
      { type: "sep" },
      { icon: "moreFunctions", label: "More Functions…", onClick: () => this.host.execute("functionDialog") },
    ];
  }

  fillMenu() {
    return [
      { icon: "fillDown", label: "Down", onClick: () => this.host.execute("fill:down") },
      { icon: "fillRight", label: "Right", onClick: () => this.host.execute("fill:right") },
      { icon: "fillUp", label: "Up", onClick: () => this.host.execute("fill:up") },
      { icon: "fillLeft", label: "Left", onClick: () => this.host.execute("fill:left") },
      { type: "sep" },
      { icon: "series", label: "Series…", onClick: () => this.host.execute("fillSeries") },
    ];
  }

  clearMenu() {
    return [
      { icon: "clearAll", label: "Clear All", onClick: () => this.host.execute("clear:all") },
      { icon: "clearFormat", label: "Clear Formats", onClick: () => this.host.execute("clear:formats") },
      { icon: "clearContents", label: "Clear Contents", shortcut: "Delete", onClick: () => this.host.execute("clear:contents") },
      { icon: "clearComments", label: "Clear Comments", onClick: () => this.host.execute("clear:comments") },
    ];
  }

  sortMenu() {
    return [
      { icon: "sortAZ", label: "Sort A to Z", onClick: () => this.host.execute("sort.asc") },
      { icon: "sortZA", label: "Sort Z to A", onClick: () => this.host.execute("sort.desc") },
      { type: "sep" },
      { icon: "sortDialog", label: "Custom Sort…", onClick: () => this.host.execute("sortDialog") },
      { icon: "filter", label: "Filter", checked: !!this.host.isFiltered?.(), onClick: () => this.host.execute("filter.toggle") },
      { icon: "clearFilter", label: "Clear Filter", onClick: () => this.host.execute("filter.clear") },
    ];
  }

  findMenu() {
    return [
      { icon: "find", label: "Find…", shortcut: "Ctrl+F", onClick: () => this.host.execute("find") },
      { icon: "replace", label: "Replace…", shortcut: "Ctrl+H", onClick: () => this.host.execute("replace") },
      { type: "sep" },
      { icon: "goTo", label: "Go To…", shortcut: "Ctrl+G", onClick: () => this.host.execute("goto") },
      { label: "Go To Special…", onClick: () => this.host.execute("gotoSpecial") },
      { type: "sep" },
      { label: "Select All", shortcut: "Ctrl+A", onClick: () => this.host.execute("selectAll") },
    ];
  }

  cellStylesMenu() {
    const items = STYLE_PRESETS.map((preset) => ({
      label: preset.label,
      onClick: () => this.host.execute("cellStyle:" + preset.id),
    }));
    const custom = (this.host.customStyles?.() || []).map((preset) => ({
      label: preset.label,
      hint: "Custom",
      onClick: () => this.host.execute("cellStyle:" + preset.id),
    }));
    const tail = custom.length
      ? [{ type: "sep" }, { type: "head", label: "Custom" }, ...custom]
      : [];
    return [...items, ...tail, { type: "sep" }, { label: "New Cell Style…", onClick: () => this.host.execute("cellStyleDialog") }];
  }

  buildHome() {
    const panel = el("div", { class: "ribbon-panel" });
    const paste = this.split({ iconName: "paste", label: "Paste", cmd: "paste", menu: () => this.pasteMenu(), kind: "big" });
    const clipboard = this.group("Clipboard", [
      paste,
      this.column([
        this.button({ iconName: "cut", label: "Cut", cmd: "cut", kind: "compact" }),
        this.button({ iconName: "copy", label: "Copy", cmd: "copy", kind: "compact" }),
        this.button({ iconName: "painter", label: "Format Painter", cmd: "formatPainter", kind: "compact" }),
      ]),
    ]);
    const undo = this.group("Undo", [
      this.column([
        this.button({ iconName: "undo", label: "Undo", cmd: "undo", kind: "compact", title: "Undo (Ctrl+Z)" }),
        this.button({ iconName: "redo", label: "Redo", cmd: "redo", kind: "compact", title: "Redo (Ctrl+Y)" }),
      ]),
    ]);

    const font = this.combo({
      value: this.host.defaultFont || "Calibri",
      options: this.fonts,
      width: 126,
      className: "font",
      onChange: (value) => this.host.execute("font.name", null, { value }),
    });
    this.fontCombo = font.input;
    const size = this.combo({
      value: String(this.host.defaultSize || 11),
      options: [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 28, 32, 36, 48, 72],
      width: 58,
      className: "size",
      onChange: (value) => this.host.execute("font.size", null, { value }),
    });
    this.sizeCombo = size.input;

    const fonts = this.group("Font", [
      this.column([
        el("div", { class: "ritem-split" }, [font.wrap]),
        el("div", { class: "ritem-split", style: "margin-top:2px" }, [size.wrap]),
      ]),
      this.column([
        this.button({ iconName: "growFont", label: "", cmd: "font.grow", kind: "icon", title: "Increase Font Size" }),
        this.button({ iconName: "bold", label: "", cmd: "bold", kind: "icon", title: "Bold (Ctrl+B)" }),
        this.button({ iconName: "alignLeft", label: "", cmd: "align.left", kind: "icon", title: "Align Left" }),
      ]),
      this.column([
        this.button({ iconName: "shrinkFont", label: "", cmd: "font.shrink", kind: "icon", title: "Decrease Font Size" }),
        this.button({ iconName: "italic", label: "", cmd: "italic", kind: "icon", title: "Italic (Ctrl+I)" }),
        this.button({ iconName: "alignCenter", label: "", cmd: "align.center", kind: "icon", title: "Center" }),
      ]),
      this.column([
        this.button({ iconName: "underline", label: "", cmd: "underline", kind: "icon", title: "Underline (Ctrl+U)" }),
        this.button({ iconName: "strike", label: "", cmd: "strike", kind: "icon", title: "Strikethrough" }),
        this.button({ iconName: "alignRight", label: "", cmd: "align.right", kind: "icon", title: "Align Right" }),
      ]),
      this.column([
        this.colorButton({
          iconName: "fillColor", cmd: "fillColor", title: "Fill Color", color: this.lastFillColor,
          onPick: (value, fromMenu) => { this.lastFillColor = value || ""; this.host.execute("fill:" + (value || "none")); },
        }),
        this.colorButton({
          iconName: "fontColor", cmd: "fontColor", title: "Font Color", color: this.lastFontColor,
          onPick: (value) => { this.lastFontColor = value || "#000000"; this.host.execute("fontColor:" + (value || "#000000")); },
        }),
        this.menuButton({ iconName: "borders", label: "", menu: () => this.bordersMenu(), kind: "icon", title: "Borders" }),
      ]),
    ]);

    const alignment = this.group("Alignment", [
      this.column([
        this.button({ iconName: "alignTop", label: "", cmd: "valign.top", kind: "icon", title: "Top Align" }),
        this.button({ iconName: "alignMiddle", label: "", cmd: "valign.middle", kind: "icon", title: "Middle Align" }),
        this.button({ iconName: "alignBottom", label: "", cmd: "valign.bottom", kind: "icon", title: "Bottom Align" }),
      ]),
      this.column([
        this.button({ iconName: "orientation", label: "", cmd: "textOrientation", kind: "icon", title: "Orientation" }),
        this.button({ iconName: "wrapText", label: "", cmd: "wrap", kind: "icon", title: "Wrap Text" }),
        this.button({ iconName: "increaseIndent", label: "", cmd: "indent.increase", kind: "icon", title: "Increase Indent" }),
      ]),
      this.column([
        this.button({ iconName: "mergeCenter", label: "", cmd: "mergeCenter", kind: "icon", title: "Merge & Center" }),
        this.button({ iconName: "decreaseIndent", label: "", cmd: "indent.decrease", kind: "icon", title: "Decrease Indent" }),
        this.menuButton({ iconName: "mergeCells", label: "", menu: () => this.mergeMenu(), kind: "icon", title: "Merge Options" }),
      ]),
    ]);

    const numberLabel = el("span", { class: "rlabel", text: "General" });
    const numberButton = el("button", { class: "rbtn split", type: "button", title: "Number Format" }, [
      el("span", { class: "ricon", html: icon("numGeneral") }),
      numberLabel,
      el("span", { class: "caret", html: icon("chevronDown") }),
    ]);
    numberButton.addEventListener("click", (event) => {
      if (event.target.closest(".caret")) this.showMenu(numberButton, this.numberMenu());
      else this.host.execute("formatCells", null, { tab: "number" });
    });
    this.formatLabel = numberLabel;

    const number = this.group("Number", [
      numberButton,
      this.column([
        this.button({ iconName: "numCurrency", label: "", cmd: "numformat:$#,##0.00", kind: "icon", title: "Currency" }),
        this.button({ iconName: "numPercent", label: "", cmd: "numformat:0.00%", kind: "icon", title: "Percent" }),
        this.button({ iconName: "decimalsIncrease", label: "", cmd: "decimals.increase", kind: "icon", title: "Increase Decimal" }),
      ]),
      this.column([
        this.button({ iconName: "numComma", label: "", cmd: "numformat:#,##0.00", kind: "icon", title: "Comma Style" }),
        this.button({ iconName: "numDate", label: "", cmd: "numformat:m/d/yyyy", kind: "icon", title: "Short Date" }),
        this.button({ iconName: "decimalsDecrease", label: "", cmd: "decimals.decrease", kind: "icon", title: "Decrease Decimal" }),
      ]),
    ]);

    const styles = this.group("Styles", [
      this.menuButton({ iconName: "conditional", label: "Conditional Formatting", menu: () => this.conditionalMenu(), title: "Conditional Formatting" }),
      this.menuButton({ iconName: "formatTable", label: "Format as Table", menu: () => this.tableMenu(), title: "Format as Table" }),
      this.menuButton({ iconName: "cellStyles", label: "Cell Styles", menu: () => this.cellStylesMenu(), title: "Cell Styles" }),
    ]);

    const cells = this.group("Cells", [
      this.menuButton({ iconName: "insertCells", label: "Insert", menu: () => this.insertCellsMenu(), title: "Insert Cells" }),
      this.menuButton({ iconName: "deleteCells", label: "Delete", menu: () => this.deleteCellsMenu(), title: "Delete Cells" }),
      this.menuButton({ iconName: "formatCells", label: "Format", menu: () => this.formatCellsMenu(), title: "Format Cells" }),
    ]);

    const editing = this.group("Editing", [
      this.split({ iconName: "autoSum", label: "AutoSum", cmd: "autosum:SUM", menu: () => this.autosumMenu(), title: "AutoSum" }),
      this.split({ iconName: "fill", label: "Fill", cmd: "fill:down", menu: () => this.fillMenu(), title: "Fill" }),
      this.split({ iconName: "clear", label: "Clear", cmd: "clear:contents", menu: () => this.clearMenu(), title: "Clear" }),
      this.split({ iconName: "sortAZ", label: "Sort & Filter", cmd: "sort.asc", menu: () => this.sortMenu(), title: "Sort & Filter" }),
      this.split({ iconName: "find", label: "Find & Select", cmd: "find", menu: () => this.findMenu(), title: "Find & Select" }),
    ]);

    panel.append(clipboard, undo, fonts, alignment, number, styles, cells, editing);
    return panel;
  }

  insertCellsMenu() {
    return [
      { icon: "insertRow", label: "Insert Sheet Rows", onClick: () => this.host.execute("insert.rows") },
      { icon: "insertCol", label: "Insert Sheet Columns", onClick: () => this.host.execute("insert.cols") },
      { icon: "insertCells", label: "Insert Cells…", onClick: () => this.host.execute("insert.cells") },
      { type: "sep" },
      { icon: "newWorkbook", label: "Insert Sheet", onClick: () => this.host.execute("sheet.add") },
    ];
  }

  deleteCellsMenu() {
    return [
      { icon: "deleteRow", label: "Delete Sheet Rows", onClick: () => this.host.execute("delete.rows") },
      { icon: "deleteCol", label: "Delete Sheet Columns", onClick: () => this.host.execute("delete.cols") },
      { icon: "deleteCells", label: "Delete Cells…", onClick: () => this.host.execute("delete.cells") },
      { type: "sep" },
      { icon: "trash", label: "Delete Sheet", onClick: () => this.host.execute("sheet.delete") },
    ];
  }

  formatCellsMenu() {
    return [
      { icon: "rowHeight", label: "Row Height…", onClick: () => this.host.execute("rowHeight") },
      { icon: "colWidth", label: "Column Width…", onClick: () => this.host.execute("colWidth") },
      { icon: "autoFitCol", label: "AutoFit Column Width", onClick: () => this.host.execute("autoFit.cols") },
      { icon: "autoFitRow", label: "AutoFit Row Height", onClick: () => this.host.execute("autoFit.rows") },
      { type: "sep" },
      { icon: "hideRows", label: "Hide Rows", onClick: () => this.host.execute("hide.rows") },
      { icon: "hideCols", label: "Hide Columns", onClick: () => this.host.execute("hide.cols") },
      { icon: "unhideRows", label: "Unhide Rows", onClick: () => this.host.execute("unhide.rows") },
      { icon: "unhideCols", label: "Unhide Columns", onClick: () => this.host.execute("unhide.cols") },
      { type: "sep" },
      { icon: "formatCells", label: "Format Cells…", onClick: () => this.host.execute("formatCells") },
    ];
  }

  conditionalMenu() {
    return [
      {
        icon: "numGeneral", label: "Highlight Cells Rules", submenu: [
          { label: "Greater Than…", onClick: () => this.host.execute("conditional:greater") },
          { label: "Less Than…", onClick: () => this.host.execute("conditional:less") },
          { label: "Between…", onClick: () => this.host.execute("conditional:between") },
          { label: "Equal To…", onClick: () => this.host.execute("conditional:equal") },
          { label: "Text that Contains…", onClick: () => this.host.execute("conditional:contains") },
          { label: "A Date Occurring…", onClick: () => this.host.execute("conditional:date") },
        ],
      },
      {
        icon: "statistical", label: "Top/Bottom Rules", submenu: [
          { label: "Top 10 Items…", onClick: () => this.host.execute("conditional:top10") },
          { label: "Bottom 10 Items…", onClick: () => this.host.execute("conditional:bottom10") },
          { label: "Above Average", onClick: () => this.host.execute("conditional:aboveAverage") },
          { label: "Below Average", onClick: () => this.host.execute("conditional:belowAverage") },
        ],
      },
      { type: "sep" },
      { icon: "palette", label: "Color Scales", submenu: [
        { label: "Green - Yellow - Red", onClick: () => this.host.execute("colorScale:3") },
        { label: "Red - Yellow - Green", onClick: () => this.host.execute("colorScale:3r") },
        { label: "Blue - White - Red", onClick: () => this.host.execute("colorScale:3b") },
      ] },
      { icon: "conditional", label: "New Rule…", onClick: () => this.host.execute("conditionalDialog") },
      { icon: "clearFormat", label: "Clear Rules", onClick: () => this.host.execute("conditional:clear") },
      { label: "Manage Rules…", onClick: () => this.host.execute("conditionalRules") },
    ];
  }

  tableMenu() {
    const formats = ["Blue, Table Style Medium 2", "Orange, Table Style Medium 3", "Green, Table Style Medium 4", "Light Blue, Table Style Medium 5", "Light Gray, Table Style Medium 1"];
    return formats.map((label, index) => ({
      label,
      onClick: () => this.host.execute("formatTable:" + index),
    }));
  }

  buildInsert() {
    const panel = el("div", { class: "ribbon-panel" });
    const tables = this.group("Tables", [
      this.button({ iconName: "table", label: "Table", cmd: "insert.table", kind: "big", title: "Insert Table" }),
    ]);
    const illustrations = this.group("Illustrations", [
      this.button({ iconName: "picture", label: "Pictures", cmd: "insert.picture", kind: "big" }),
      this.menuButton({ iconName: "shapes", label: "Shapes", menu: () => this.shapeMenu(), title: "Shapes" }),
      this.button({ iconName: "icons", label: "Icons", cmd: "insert.icons" }),
    ]);
    const charts = this.group("Charts", [
      this.menuButton({ iconName: "chartColumn", label: "Column", menu: () => this.chartMenu("column"), title: "Column Chart" }),
      this.menuButton({ iconName: "chartBar", label: "Bar", menu: () => this.chartMenu("bar"), title: "Bar Chart" }),
      this.menuButton({ iconName: "chartLine", label: "Line", menu: () => this.chartMenu("line"), title: "Line Chart" }),
      this.menuButton({ iconName: "chartPie", label: "Pie", menu: () => this.chartMenu("pie"), title: "Pie Chart" }),
      this.menuButton({ iconName: "chartRecommended", label: "Recommended Charts", menu: () => this.chartMenu("column"), title: "Recommended Charts" }),
    ]);
    const links = this.group("Links", [
      this.button({ iconName: "link", label: "Link", cmd: "insert.link", kind: "big", title: "Insert Link" }),
    ]);
    const comments = this.group("Comments", [
      this.button({ iconName: "comment", label: "Comment", cmd: "comment.new", kind: "big", title: "New Comment" }),
    ]);
    const text = this.group("Text", [
      this.button({ iconName: "textBox", label: "Text Box", cmd: "insert.textBox", kind: "big" }),
      this.button({ iconName: "symbol", label: "Symbol", cmd: "insert.symbol", kind: "big" }),
      this.button({ iconName: "headerFooter", label: "Header & Footer", cmd: "insert.headerFooter" }),
    ]);
    const spark = this.group("Sparklines", [
      this.button({ iconName: "sparkline", label: "Line", cmd: "insert.sparkline", kind: "compact" }),
    ]);
    panel.append(tables, illustrations, charts, links, comments, text, spark);
    return panel;
  }

  chartMenu(type) {
    const entries = CHART_TYPES.map((chart) => ({
      icon: chart === "column" ? "chartColumn" : chart === "bar" ? "chartBar" : chart === "line" ? "chartLine" : chart === "pie" ? "chartPie" : chart === "area" ? "chartArea" : chart === "scatter" ? "chartScatter" : "chartPie",
      label: chart.charAt(0).toUpperCase() + chart.slice(1),
      onClick: () => this.host.execute("insert.chart:" + (type === "pie" && chart === "doughnut" ? "doughnut" : chart)),
    }));
    return [
      { type: "head", label: "Charts" },
      ...entries,
      { type: "sep" },
      { icon: "chartDialog", label: "More Chart Types…", onClick: () => this.host.execute("chartDialog") },
    ];
  }

  shapeMenu() {
    return [
      { label: "Rectangle", onClick: () => this.host.execute("insert.shape:rect") },
      { label: "Rounded Rectangle", onClick: () => this.host.execute("insert.shape:round") },
      { label: "Ellipse", onClick: () => this.host.execute("insert.shape:ellipse") },
      { label: "Line", onClick: () => this.host.execute("insert.shape:line") },
    ];
  }

  buildFormulas() {
    const panel = el("div", { class: "ribbon-panel" });
    const library = this.group("Function Library", [
      this.button({ iconName: "function", label: "Insert Function", cmd: "functionDialog", kind: "big", title: "Insert Function" }),
      this.split({ iconName: "autoSum", label: "AutoSum", cmd: "autosum:SUM", menu: () => this.autosumMenu() }),
      this.menuButton({ iconName: "recentlyUsed", label: "Recently Used", menu: () => this.recentFunctionsMenu(), title: "Recently Used Functions" }),
      this.menuButton({ iconName: "financial", label: "Financial", menu: () => this.categoryMenu("Financial"), title: "Financial Functions" }),
      this.menuButton({ iconName: "logical", label: "Logical", menu: () => this.categoryMenu("Logical"), title: "Logical Functions" }),
      this.menuButton({ iconName: "textFunction", label: "Text", menu: () => this.categoryMenu("Text"), title: "Text Functions" }),
      this.menuButton({ iconName: "dateTime", label: "Date & Time", menu: () => this.categoryMenu("Date & Time"), title: "Date & Time Functions" }),
      this.menuButton({ iconName: "lookup", label: "Lookup & Reference", menu: () => this.categoryMenu("Lookup & Reference"), title: "Lookup Functions" }),
      this.menuButton({ iconName: "mathTrig", label: "Math & Trig", menu: () => this.categoryMenu("Math & Trig"), title: "Math Functions" }),
      this.menuButton({ iconName: "statistical", label: "More Functions", menu: () => this.moreFunctionsMenu(), title: "More Functions" }),
    ]);
    const names = this.group("Defined Names", [
      this.button({ iconName: "nameManager", label: "Name Manager", cmd: "name.manager" }),
      this.button({ iconName: "defineName", label: "Define Name", cmd: "name.define" }),
      this.menuButton({ iconName: "useInFormula", label: "Use in Formula", menu: () => this.namesMenu(), title: "Use in Formula" }),
    ]);
    const audit = this.group("Formula Auditing", [
      this.button({ iconName: "tracePrecedents", label: "Trace Precedents", cmd: "audit.precedents" }),
      this.button({ iconName: "traceDependents", label: "Trace Dependents", cmd: "audit.dependents" }),
      this.button({ iconName: "removeArrows", label: "Remove Arrows", cmd: "audit.remove" }),
      this.button({ iconName: "showFormulas", label: "Show Formulas", cmd: "showFormulas" }),
    ]);
    const calculation = this.group("Calculation", [
      this.menuButton({ iconName: "calcOptions", label: "Calculation Options", menu: () => this.calcMenu(), title: "Calculation Options" }),
      this.button({ iconName: "calcNow", label: "Calculate Now", cmd: "calc.now", title: "Calculate Now (F9)" }),
      this.button({ iconName: "calcSheet", label: "Calculate Sheet", cmd: "calc.sheet", title: "Calculate Sheet (Shift+F9)" }),
    ]);
    panel.append(library, names, audit, calculation);
    return panel;
  }

  calcMenu() {
    const mode = this.host.calcMode?.() || "auto";
    return [
      { label: "Automatic", checked: mode === "auto", onClick: () => this.host.execute("calc.auto") },
      { label: "Manual", checked: mode === "manual", onClick: () => this.host.execute("calc.manual") },
      { type: "sep" },
      { icon: "calcNow", label: "Calculate Now", onClick: () => this.host.execute("calc.now") },
    ];
  }

  recentFunctionsMenu() {
    const names = this.host.recentFunctions?.() || ["SUM", "AVERAGE", "IF", "COUNT", "MAX", "MIN", "VLOOKUP", "TEXT", "CONCAT"];
    return names.map((name) => ({ label: name, onClick: () => this.host.execute("insertFunction:" + name) }));
  }

  categoryMenu(category) {
    const names = this.host.functionsByCategory?.(category) || FUNCTION_NAMES.slice(0, 24);
    return names.map((name) => ({ label: name, onClick: () => this.host.execute("insertFunction:" + name) }));
  }

  moreFunctionsMenu() {
    const categories = ["Statistical", "Engineering", "Information", "Compatibility", "Web", "Cube"];
    return categories.map((name) => ({
      label: name,
      submenu: (this.host.functionsByCategory?.(name) || []).map((func) => ({ label: func, onClick: () => this.host.execute("insertFunction:" + func) })),
    }));
  }

  namesMenu() {
    const names = this.host.definedNames?.() || [];
    if (!names.length) return [{ label: "No names defined", disabled: true }];
    return names.map((name) => ({
      label: name,
      onClick: () => this.host.execute("insertName:" + name),
    }));
  }

  buildData() {
    const panel = el("div", { class: "ribbon-panel" });
    const sortFilter = this.group("Sort & Filter", [
      this.split({ iconName: "sortAZ", label: "Sort", cmd: "sort.asc", menu: () => this.sortMenu(), kind: "big" }),
      this.split({ iconName: "filter", label: "Filter", cmd: "filter.toggle", menu: () => this.filterMenu(), kind: "big" }),
      this.button({ iconName: "clearFilter", label: "Clear", cmd: "filter.clear" }),
      this.button({ iconName: "sortDialog", label: "Advanced", cmd: "sortDialog" }),
    ]);
    const tools = this.group("Data Tools", [
      this.button({ iconName: "textToColumns", label: "Text to Columns", cmd: "textToColumns", kind: "big" }),
      this.button({ iconName: "flashFill", label: "Flash Fill", cmd: "flashFill", kind: "big" }),
      this.button({ iconName: "removeDuplicates", label: "Remove Duplicates", cmd: "removeDuplicates" }),
      this.button({ iconName: "validation", label: "Data Validation", cmd: "validation" }),
      this.button({ iconName: "subtotal", label: "Subtotal", cmd: "subtotal" }),
    ]);
    const outline = this.group("Outline", [
      this.button({ iconName: "group", label: "Group", cmd: "group.rows" }),
      this.button({ iconName: "ungroup", label: "Ungroup", cmd: "ungroup.rows" }),
    ]);
    panel.append(sortFilter, tools, outline);
    return panel;
  }

  filterMenu() {
    const range = this.host.filterRange?.();
    return [
      { label: "Toggle AutoFilter", onClick: () => this.host.execute("filter.toggle") },
      { label: "Clear Filter", onClick: () => this.host.execute("filter.clear") },
      { label: "Reapply", onClick: () => this.host.execute("filter.reapply") },
      { type: "sep" },
      { label: "Filter by Selected Cell's Value", disabled: !range, onClick: () => this.host.execute("filter.byValue") },
    ];
  }

  buildReview() {
    const panel = el("div", { class: "ribbon-panel" });
    const comments = this.group("Comments", [
      this.button({ iconName: "newComment", label: "New Comment", cmd: "comment.new", kind: "big" }),
      this.button({ iconName: "deleteComment", label: "Delete", cmd: "comment.delete" }),
      this.button({ iconName: "prevComment", label: "Previous", cmd: "comment.prev" }),
      this.button({ iconName: "nextComment", label: "Next", cmd: "comment.next" }),
      this.button({ iconName: "showComments", label: "Show Comments", cmd: "comment.show" }),
    ]);
    const changes = this.group("Changes", [
      this.button({ iconName: "protectSheet", label: "Protect Sheet", cmd: "protect.sheet", kind: "big" }),
      this.button({ iconName: "protectWorkbook", label: "Protect Workbook", cmd: "protect.workbook", kind: "big" }),
    ]);
    const proofing = this.group("Proofing", [
      this.button({ iconName: "spellcheck", label: "Spelling", cmd: "spellcheck" }),
      this.button({ iconName: "workbookStats", label: "Workbook Statistics", cmd: "workbook.stats" }),
    ]);
    const accessibility = this.group("Accessibility", [
      this.button({ iconName: "accessibility", label: "Check Accessibility", cmd: "accessibility", kind: "big" }),
    ]);
    panel.append(comments, changes, proofing, accessibility);
    return panel;
  }

  buildView() {
    const panel = el("div", { class: "ribbon-panel" });
    const views = this.group("Sheet View", [
      this.button({ iconName: "normalView", label: "Normal", cmd: "view.normal", kind: "big" }),
      this.button({ iconName: "pageBreakView", label: "Page Break Preview", cmd: "view.pagebreak", kind: "big" }),
      this.button({ iconName: "pageLayout", label: "Page Layout", cmd: "view.pageLayout", kind: "big" }),
    ]);
    const show = this.group("Show", [
      this.button({ iconName: "gridlines", label: "Gridlines", cmd: "gridlines", kind: "compact", active: true }),
      this.button({ iconName: "headings", label: "Headings", cmd: "headings", kind: "compact", active: true }),
      this.button({ iconName: "formulaBar", label: "Formula Bar", cmd: "formulaBar", kind: "compact", active: true }),
    ]);
    const themes = this.group("Themes", [
      this.button({
        iconName: "theme",
        label: "Dark Mode",
        cmd: "theme.toggle",
        kind: "big",
        title: "Switch between the light and dark theme",
        active: this.host.settings?.theme === "dark",
      }),
    ]);
    const zoom = this.group("Zoom", [
      this.button({ iconName: "zoomIn", label: "Zoom", cmd: "zoom", kind: "big" }),
      this.button({ iconName: "zoomFit", label: "100%", cmd: "zoom100", kind: "compact" }),
      this.button({ iconName: "zoomSelection", label: "Zoom to Selection", cmd: "zoomSelection", kind: "compact" }),
    ]);
    const freeze = this.group("Window", [
      this.menuButton({ iconName: "freezePanes", label: "Freeze Panes", menu: () => this.freezeMenu(), title: "Freeze Panes" }),
      this.button({ iconName: "splitPanes", label: "Split", cmd: "splitPanes" }),
    ]);
    const sheets = this.group("Sheet", [
      this.button({ iconName: "hideSheet", label: "Hide", cmd: "sheet.hide" }),
      this.button({ iconName: "unhideSheet", label: "Unhide", cmd: "sheet.unhide" }),
    ]);
    panel.append(views, show, themes, zoom, freeze, sheets);
    return panel;
  }

  freezeMenu() {
    const freeze = this.host.freezeState?.() || { rows: 0, cols: 0 };
    return [
      { label: "Freeze Panes", onClick: () => this.host.execute("freeze.panes") },
      { label: "Freeze Top Row", checked: freeze.rows === 1 && !freeze.cols, onClick: () => this.host.execute("freeze.top") },
      { label: "Freeze First Column", checked: freeze.cols === 1 && !freeze.rows, onClick: () => this.host.execute("freeze.first") },
      { type: "sep" },
      { label: "Unfreeze Panes", disabled: !freeze.rows && !freeze.cols, onClick: () => this.host.execute("unfreeze") },
    ];
  }

  buildHelp() {
    const panel = el("div", { class: "ribbon-panel" });
    const help = this.group("Help", [
      this.button({ iconName: "help", label: "Help", cmd: "help", kind: "big" }),
      this.button({ iconName: "shortcuts", label: "Keyboard Shortcuts", cmd: "shortcuts", kind: "big" }),
      this.button({ iconName: "whatsNew", label: "What's New", cmd: "whatsNew", kind: "big" }),
    ]);
    const about = this.group("About", [
      this.button({ iconName: "info", label: "About mTails", cmd: "about", kind: "big" }),
      this.button({ iconName: "feedback", label: "Send Feedback", cmd: "feedback", kind: "big" }),
    ]);
    panel.append(help, about);
    return panel;
  }

  /* --------------------------------------------------------------- update */
  update(info) {
    if (!info) return;
    const style = info.style || {};
    this.setActive("bold", !!style.b);
    this.setActive("italic", !!style.i);
    this.setActive("underline", !!style.u);
    this.setActive("strike", !!style.strike);
    this.setActive("wrap", !!style.wrap);
    this.setActive("align.left", (style.align || "general") === "left");
    this.setActive("align.center", style.align === "center");
    this.setActive("align.right", style.align === "right");
    this.setActive("valign.top", style.valign === "top");
    this.setActive("valign.middle", !style.valign || style.valign === "middle");
    this.setActive("valign.bottom", style.valign === "bottom");
    this.setActive("mergeCenter", !!info.merged);
    if (this.fontCombo && document.activeElement !== this.fontCombo) {
      this.fontCombo.value = style.font || this.host.defaultFont || "Calibri";
    }
    if (this.sizeCombo && document.activeElement !== this.sizeCombo) {
      this.sizeCombo.value = String(Number(style.size) || this.host.defaultSize || 11);
    }
    if (this.formatLabel) this.formatLabel.textContent = formatLabel(style.numFmt);
    const gridlinesButton = this.byCmd.get("gridlines") || [];
    for (const node of gridlinesButton) node.classList.toggle("active", info.showGridlines !== false);
    const headingsButton = this.byCmd.get("headings") || [];
    for (const node of headingsButton) node.classList.toggle("active", info.showHeadings !== false);
  }
}

export { formatLabel, STYLE_PRESETS };
