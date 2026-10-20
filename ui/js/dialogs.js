/* Modal dialogs: formatting, charts, validation, sorting, filters, names.
 *
 * Every dialog receives the application object and talks to the grid through
 * its public methods, so a dialog never mutates the model directly unless it
 * is a drawing or a named object that has no grid helper.  Dialogs return the
 * value produced by showDialog(); callers do not need to await them.
 */

import {
  $,
  el,
  field,
  numberInput,
  selectControl,
  checkRow,
  segmented,
  showDialog,
  confirmDialog,
  promptDialog,
  toast,
  colorMenu,
  openMenu,
  closeMenus,
  THEME_COLORS,
  STANDARD_COLORS,
} from "./ui.js";
import { icon } from "./icons.js";
import {
  STANDARD_FORMATS,
  formatCategory,
  formatPreview,
  formatValue,
  parseUserInput,
  resolveFormat,
} from "./format.js";
import {
  MAX_ROWS,
  MAX_COLS,
  addr,
  indexToCol,
  parseAddr,
  parseRange,
  rangeText,
  normalizeRange,
} from "./units.js";
import {
  usedRange,
  validName,
  cellAt,
  ensureCell,
  insertRows,
  deleteRows,
  sheetByName,
  workbookStats,
} from "./model.js";
import { FUNCTION_NAMES } from "./formula.js";
import { CHART_TYPES } from "./ribbon.js";
import { t } from "./i18n.js";

const clone = (value) => (value ? JSON.parse(JSON.stringify(value)) : null);
const rangeOf = (grid) => ({ ...grid.sel });
const asText = (value) => String(value ?? "");

function rangeLabel(range) {
  return rangeText(normalizeRange(range.r1, range.c1, range.r2, range.c2));
}

function columnOptions(app, range, withHeader = true) {
  const sheet = app.grid.sheet();
  const options = [];
  for (let col = range.c1; col <= range.c2; col += 1) {
    let label = indexToCol(col);
    if (withHeader) {
      const cell = cellAt(sheet, range.r1, col);
      const header = cell ? String(cell.v ?? "").trim() : "";
      if (header && header !== label) label += " — " + header.slice(0, 24);
    }
    options.push({ value: col, label });
  }
  return options;
}

function formatListFor(category) {
  const list = STANDARD_FORMATS.filter((item) => formatCategory(item.code) === category);
  if (list.length) return list;
  return STANDARD_FORMATS.filter((item) => item.label === "General");
}

function decimalsIn(code) {
  const match = /\.(0+)/.exec(String(code || ""));
  return match ? match[1].length : 0;
}

function withDecimals(code, places) {
  let text = String(code || "General");
  if (!text || text === "General" || text === "@") return text;
  if (/\.[0#?]+/.test(text)) {
    return text.replace(/\.([0#?]+)/, (match, digits) => {
      if (places <= 0) return "";
      const tail = digits.replace(/[0#?]/g, "0");
      return "." + (tail.slice(0, places) + "0".repeat(Math.max(0, places - tail.length)));
    });
  }
  if (places > 0) {
    const index = text.search(/[0#?]/);
    if (index >= 0) return text.slice(0, index) + "." + "0".repeat(places) + text.slice(index);
  }
  return text;
}

function withSeparator(code, enabled) {
  let text = String(code || "General");
  if (text === "General" || text === "@" || !/[0#]/.test(text)) return text;
  if (enabled) {
    if (/#,##0/.test(text) || /,#+/.test(text)) return text;
    return text.replace(/#{0,3}?0/, "#,##0");
  }
  return text.replace(/,/g, "");
}

function colorRow(label, value, onPick) {
  const sample = el("span", { class: "color-sample" });
  sample.style.background = value || "transparent";
  const button = el("button", { class: "btn", type: "button", text: "Color" });
  button.addEventListener("click", () => {
    colorMenu(button, {
      title: "Theme Colors",
      allowNone: true,
      onPick: (color) => {
        sample.style.background = color || "transparent";
        onPick(color);
      },
    });
  });
  const row = el("div", { class: "input-row" }, [sample, button]);
  return field(label, row);
}

function previewBorder(node, sides) {
  node.replaceChildren();
  for (const side of ["t", "b", "l", "r"]) {
    if (sides[side]) node.append(el("span", { class: side, i18n: false }));
  }
}

/* ------------------------------------------------------------ format cells */
export function formatCellsDialog(app, initialTab = "number") {
  const grid = app.grid;
  const base = clone(grid.currentStyle()) || {};
  const state = {
    tab: initialTab,
    format: base.numFmt || "General",
    align: base.align || "general",
    valign: base.valign || "middle",
    indent: Number(base.indent) || 0,
    wrap: !!base.wrap,
    merge: !!grid.mergeFor(grid.active.row, grid.active.col),
    font: base.font || app.settings.defaultFont,
    size: Number(base.size) || app.settings.defaultSize,
    b: !!base.b,
    i: !!base.i,
    u: base.u === true || base.u === "single" ? "single" : base.u === "double" ? "double" : "",
    strike: !!base.strike,
    color: base.color || "",
    fill: base.fill || "",
    border: base.border ? clone(base.border) : {},
    borderPreset: "",
    borderSpec: { style: "thin", color: "#000000" },
    locked: base.locked !== false,
    hidden: !!base.hidden,
  };
  const tabs = [
    ["number", "Number"], ["alignment", "Alignment"], ["font", "Font"],
    ["fill", "Fill"], ["border", "Border"], ["protection", "Protection"],
  ];
  const tabBar = el("div", { class: "dialog-tabs" });
  const host = el("div");
  const body = el("div", {}, [tabBar, host]);
  let sample = null;

  const drawTab = () => {
    for (const button of tabBar.querySelectorAll(".dialog-tab")) {
      button.classList.toggle("active", button.dataset.tab === state.tab);
    }
    host.replaceChildren();
    if (state.tab === "number") host.append(numberPanel());
    else if (state.tab === "alignment") host.append(alignmentPanel());
    else if (state.tab === "font") host.append(fontPanel());
    else if (state.tab === "fill") host.append(fillPanel());
    else if (state.tab === "border") host.append(borderPanel());
    else host.append(protectionPanel());
  };

  const refreshSample = () => {
    if (!sample) return;
    const code = state.format === "General" ? "General" : state.format;
    let value = 1234.567;
    if (formatCategory(code) === "date" || formatCategory(code) === "time") value = 45292.75;
    if (formatCategory(code) === "text") value = "mTails";
    try {
      sample.textContent = formatValue(value, resolveFormat(code));
    } catch {
      sample.textContent = "#VALUE!";
    }
  };

  const numberPanel = () => {
    const list = el("div", { class: "fmt-list" });
    for (const category of ["General", "Number", "Currency", "Accounting", "Date", "Time", "Percentage", "Fraction", "Scientific", "Text"]) {
      const probe = STANDARD_FORMATS.find((item) => item.label === category) || STANDARD_FORMATS[0];
      const item = el("div", { class: "list-item" + (formatCategory(state.format) === formatCategory(probe.code) ? " active" : "") }, [
        el("div", { text: category }),
        el("div", { class: "li-sub", text: formatPreview(probe.code) }),
      ]);
      item.addEventListener("click", () => {
        state.format = probe.code;
        drawTab();
        refreshSample();
      });
      list.append(item);
    }
    const sampleBox = el("div", { class: "fmt-preview" });
    sample = sampleBox;
    const decimal = numberInput(decimalsIn(state.format), { min: 0, max: 10 });
    decimal.style.width = "70px";
    decimal.addEventListener("input", () => {
      state.format = withDecimals(state.format, Number(decimal.value) || 0);
      code.value = state.format;
      refreshSample();
    });
    const separator = el("input", { type: "checkbox" });
    separator.checked = /,/.test(state.format);
    separator.addEventListener("change", () => {
      state.format = withSeparator(state.format, separator.checked);
      code.value = state.format;
      refreshSample();
    });
    const code = el("input", { type: "text", value: state.format, spellcheck: "false" });
    code.style.cssText = "width:100%;height:28px;border:1px solid var(--border-strong);border-radius:3px;padding:0 7px";
    code.addEventListener("input", () => {
      state.format = code.value.trim() || "General";
      refreshSample();
    });
    const right = el("div", {}, [
      el("div", { class: "dialog-section-title", text: "Sample" }),
      sampleBox,
      el("div", { class: "dialog-section-title", text: "Options" }),
      el("label", { class: "check-row" }, [separator, el("span", { text: "Use 1000 separator" })]),
      el("div", { class: "input-row" }, [
        el("span", { class: "unit", text: "Decimal places" }),
        decimal,
      ]),
      field("Type", code),
    ]);
    const wrap = el("div", { style: "display:grid;grid-template-columns:180px 1fr;gap:14px" }, [list, right]);
    queueMicrotask(refreshSample);
    return wrap;
  };

  const alignmentPanel = () => {
    const horizontal = selectControl([
      { value: "general", label: "General" }, { value: "left", label: "Left" },
      { value: "center", label: "Center" }, { value: "right", label: "Right" },
      { value: "justify", label: "Justify" },
    ], state.align);
    horizontal.addEventListener("change", () => { state.align = horizontal.value; });
    const vertical = selectControl([
      { value: "top", label: "Top" }, { value: "middle", label: "Middle" }, { value: "bottom", label: "Bottom" },
    ], state.valign);
    vertical.addEventListener("change", () => { state.valign = vertical.value; });
    const indent = numberInput(state.indent, { min: 0, max: 15 });
    indent.addEventListener("input", () => { state.indent = Number(indent.value) || 0; });
    const wrap = el("input", { type: "checkbox" });
    wrap.checked = state.wrap;
    wrap.addEventListener("change", () => { state.wrap = wrap.checked; });
    const merge = el("input", { type: "checkbox" });
    merge.checked = state.merge;
    merge.disabled = grid.sel.r1 === grid.sel.r2 && grid.sel.c1 === grid.sel.c2 && !state.merge;
    merge.addEventListener("change", () => { state.merge = merge.checked; });
    return el("div", { class: "dialog-grid" }, [
      field("Horizontal", horizontal),
      field("Vertical", vertical),
      field("Indent", indent),
      el("div", {}, [
        el("label", { class: "check-row" }, [wrap, el("span", { text: "Wrap text" })]),
        el("label", { class: "check-row" }, [merge, el("span", { text: "Merge cells" })]),
      ]),
    ]);
  };

  const fontPanel = () => {
    const fonts = el("div", { class: "font-list" });
    for (const name of app.fonts || []) {
      const item = el("div", { class: "list-item" + (name === state.font ? " active" : ""), text: name });
      item.style.fontFamily = `"${name}"`;
      item.addEventListener("click", () => {
        state.font = name;
        for (const other of fonts.children) other.classList.toggle("active", other === item);
      });
      fonts.append(item);
    }
    const styles = el("div", { class: "size-list" });
    for (const [label, flags] of [["Regular", [0, 0]], ["Italic", [0, 1]], ["Bold", [1, 0]], ["Bold Italic", [1, 1]]]) {
      const item = el("div", { class: "list-item" + (state.b === !!flags[0] && state.i === !!flags[1] ? " active" : ""), text: label });
      item.style.fontStyle = flags[1] ? "italic" : "normal";
      item.style.fontWeight = flags[0] ? "700" : "400";
      item.addEventListener("click", () => {
        state.b = !!flags[0];
        state.i = !!flags[1];
        for (const other of styles.children) {
          other.classList.toggle("active", other === item);
        }
      });
      styles.append(item);
    }
    const sizes = el("div", { class: "size-list" });
    for (const size of [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 40, 48, 56, 72]) {
      const item = el("div", { class: "list-item" + (Number(state.size) === size ? " active" : ""), text: String(size) });
      item.addEventListener("click", () => {
        state.size = size;
        for (const other of sizes.children) other.classList.toggle("active", other === item);
      });
      sizes.append(item);
    }
    const underline = selectControl([
      { value: "", label: "None" }, { value: "single", label: "Single" }, { value: "double", label: "Double" },
    ], state.u);
    underline.addEventListener("change", () => { state.u = underline.value; });
    const strike = el("input", { type: "checkbox" });
    strike.checked = state.strike;
    strike.addEventListener("change", () => { state.strike = strike.checked; });
    return el("div", { style: "display:grid;grid-template-columns:1.4fr 1fr 80px;gap:12px" }, [
      el("div", {}, [el("div", { class: "dialog-section-title", text: "Font" }), fonts]),
      el("div", {}, [el("div", { class: "dialog-section-title", text: "Font style" }), styles]),
      el("div", {}, [el("div", { class: "dialog-section-title", text: "Size" }), sizes]),
      el("div", { class: "full", style: "grid-column:1/-1;display:grid;grid-template-columns:1fr 1fr;gap:12px" }, [
        colorRow("Color", state.color, (value) => { state.color = value; }),
        el("div", {}, [
          field("Underline", underline),
          el("label", { class: "check-row" }, [strike, el("span", { text: "Strikethrough" })]),
        ]),
      ]),
    ]);
  };

  const fillPanel = () => {
    const summary = el("div", { class: "color-sample" });
    summary.style.background = state.fill || "transparent";
    const none = el("button", { class: "btn", type: "button", text: "None" });
    none.addEventListener("click", () => {
      state.fill = "";
      summary.style.background = "transparent";
    });
    const custom = el("input", { type: "color", value: state.fill || "#FFFF00" });
    custom.style.cssText = "width:44px;height:28px;border:1px solid var(--border);padding:1px;background:var(--surface)";
    custom.addEventListener("input", () => {
      state.fill = custom.value;
      summary.style.background = custom.value;
    });
    const palette = el("div", { class: "palette-grid", style: "padding:0" });
    for (const color of [...THEME_COLORS, ...STANDARD_COLORS]) {
      const swatch = el("button", { class: "swatch", type: "button", title: color });
      swatch.style.background = color;
      swatch.addEventListener("click", () => {
        state.fill = color;
        summary.style.background = color;
        custom.value = color;
      });
      palette.append(swatch);
    }
    return el("div", {}, [
      el("div", { class: "dialog-section-title", text: "Background Color" }),
      palette, el("div", { style: "height:10px" }),
      el("div", { class: "input-row" }, [summary, none, custom]),
      el("p", { class: "dialog-note", text: "The selected fill is applied to every cell in the selection." }),
    ]);
  };

  const borderPanel = () => {
    const preview = el("div", { class: "border-preview" });
    preview.style.width = "150px";
    preview.style.height = "96px";
    previewBorder(preview, state.border);
    const stylesRow = el("div", { class: "border-style-row" });
    for (const style of ["thin", "medium", "thick", "double", "dashed", "dotted"]) {
      const button = el("button", { type: "button", text: style });
      button.classList.toggle("active", state.borderSpec.style === style);
      button.addEventListener("click", () => {
        state.borderSpec.style = style;
        for (const other of stylesRow.children) other.classList.toggle("active", other === button);
      });
      stylesRow.append(button);
    }
    const color = el("span", { class: "color-sample" });
    color.style.background = state.borderSpec.color;
    const colorButton = el("button", { class: "btn", type: "button", text: "Color" });
    colorButton.addEventListener("click", () => {
      colorMenu(colorButton, {
        title: "Theme Colors",
        onPick: (value) => {
          state.borderSpec.color = value || "#000000";
          color.style.background = state.borderSpec.color;
        },
      });
    });
    const grid = el("div", { class: "border-grid" });
    const presets = [
      ["None", "none"], ["Outside", "outside"], ["Inside", "inside"],
      ["Top", "top"], ["Bottom", "bottom"], ["All", "all"],
      ["Left", "left"], ["Right", "right"],
    ];
    for (const [label, preset] of presets) {
      const button = el("button", { class: "border-cell", type: "button", text: label });
      button.addEventListener("click", () => {
        state.borderPreset = preset;
        state.border = {};
        if (preset === "all" || preset === "outside" || preset === "inside") {
          for (const side of ["t", "b", "l", "r"]) state.border[side] = { ...state.borderSpec };
        } else if (preset !== "none") {
          const key = preset === "top" ? "t" : preset === "bottom" ? "b" : preset === "left" ? "l" : "r";
          state.border[key] = { ...state.borderSpec };
        }
        previewBorder(preview, state.border);
        for (const other of grid.children) other.classList.toggle("active", other === button);
      });
      if (preset === "outside") button.classList.add("active");
      grid.append(button);
    }
    return el("div", { style: "display:grid;grid-template-columns:190px 1fr;gap:16px" }, [
      el("div", {}, [
        el("div", { class: "dialog-section-title", text: "Presets" }),
        grid,
      ]),
      el("div", {}, [
        el("div", { class: "dialog-section-title", text: "Line style" }),
        stylesRow,
        el("div", { style: "height:10px" }),
        el("div", { class: "input-row" }, [color, colorButton]),
        el("div", { class: "dialog-section-title", text: "Preview" }),
        el("div", { style: "display:flex;align-items:center;justify-content:center" }, [preview]),
      ]),
    ]);
  };

  const protectionPanel = () => {
    const locked = el("input", { type: "checkbox" });
    locked.checked = state.locked;
    locked.addEventListener("change", () => { state.locked = locked.checked; });
    const hidden = el("input", { type: "checkbox" });
    hidden.checked = state.hidden;
    hidden.addEventListener("change", () => { state.hidden = hidden.checked; });
    return el("div", {}, [
      el("label", { class: "check-row" }, [locked, el("span", { text: "Locked" })]),
      el("label", { class: "check-row" }, [hidden, el("span", { text: "Hidden" })]),
      el("p", { class: "dialog-note", text: "Locking takes effect after the sheet is protected with Review → Protect Sheet." }),
    ]);
  };

  for (const [id, label] of tabs) {
    const button = el("button", { class: "dialog-tab", type: "button", text: label, dataset: { tab: id } });
    button.addEventListener("click", () => {
      state.tab = id;
      drawTab();
    });
    tabBar.append(button);
  }

  return showDialog({
    title: "Format Cells",
    size: "wide",
    body,
    buttons: [
      { label: "Cancel", kind: "ghost" },
      {
        label: "OK",
        kind: "primary",
        onClick: () => {
          const patch = {
            numFmt: state.format === "General" ? null : state.format,
            align: state.align === "general" ? null : state.align,
            valign: state.valign === "middle" ? null : state.valign,
            indent: state.indent ? Number(state.indent) : null,
            wrap: state.wrap,
            font: state.font || null,
            size: state.size || null,
            b: state.b,
            i: state.i,
            u: state.u || null,
            strike: state.strike,
            color: state.color || null,
            fill: state.fill || null,
            locked: state.locked,
            hidden: state.hidden,
          };
          grid.setStyle(patch);
          if (state.borderPreset && state.borderPreset !== "outline") {
            grid.applyBorders(state.borderPreset === "outside" ? "outside" : state.borderPreset, state.borderSpec);
          } else if (Object.keys(state.border).length) {
            grid.setStyle({ border: {
              t: state.border.t || null, r: state.border.r || null,
              b: state.border.b || null, l: state.border.l || null,
            } });
          }
          if (state.merge) grid.mergeCells();
        },
      },
    ],
    onOpen: drawTab,
  });
}

/* ------------------------------------------------------------------ charts */
const PREVIEW_SHEET = (() => {
  const cells = {};
  const values = [
    ["Q1", 12, 8], ["Q2", 18, 11], ["Q3", 9, 14], ["Q4", 22, 16], ["Q5", 15, 10],
  ];
  values.forEach((row, rowIndex) => {
    row.forEach((value, colIndex) => {
      cells[addr(rowIndex + 1, colIndex + 1)] = { v: value };
    });
  });
  return { name: "Preview", cells };
})();

export function chartDialog(app, initialType = "column", existing = null) {
  const grid = app.grid;
  const sheet = grid.sheet();
  const used = usedRange(sheet);
  const initialRange = existing?.range || (grid.sel.r1 !== grid.sel.r2 || grid.sel.c1 !== grid.sel.c2
    ? rangeLabel(grid.sel)
    : used ? rangeLabel(used) : "A1:B5");
  const state = {
    type: existing?.type || initialType,
    range: initialRange,
    title: existing?.title || "Chart title",
    headerRow: existing ? !!existing.headerRow : true,
    headerCol: existing ? !!existing.headerCol : false,
  };
  const cards = el("div", { class: "chart-grid" });
  const preview = el("div", { class: "chart-dialog-preview" });
  const rangeInput = el("input", { type: "text", value: state.range, spellcheck: "false" });
  const titleInput = el("input", { type: "text", value: state.title });
  titleInput.style.cssText = "width:100%;height:28px;border:1px solid var(--border-strong);border-radius:3px;padding:0 7px";
  const headerRow = el("input", { type: "checkbox" });
  headerRow.checked = state.headerRow;
  const headerCol = el("input", { type: "checkbox" });
  headerCol.checked = state.headerCol;

  const redraw = () => {
    const chart = {
      type: state.type, range: state.range, title: state.title,
      headerRow: state.headerRow, headerCol: state.headerCol,
    };
    import("./grid-base.js").then(({ chartSvg }) => {
      preview.innerHTML = chartSvg(chart, sheet);
    });
    for (const card of cards.children) card.classList.toggle("active", card.dataset.type === state.type);
  };

  for (const type of CHART_TYPES) {
    const card = el("button", { class: "chart-card", type: "button", dataset: { type } });
    const thumb = el("div", { class: "chart-thumb", i18n: false });
    thumb.innerHTML = `<svg viewBox="0 0 320 180" preserveAspectRatio="none">${
      type === "pie" || type === "doughnut"
        ? '<circle cx="160" cy="90" r="62" fill="#4472C4"/><path d="M160 90 L160 28 A62 62 0 0 1 214 60 Z" fill="#ED7D31"/>'
        : '<rect x="40" y="100" width="26" height="50" fill="#4472C4"/><rect x="80" y="70" width="26" height="80" fill="#ED7D31"/><rect x="120" y="90" width="26" height="60" fill="#A5A5A5"/><rect x="160" y="45" width="26" height="105" fill="#FFC000"/><rect x="200" y="65" width="26" height="85" fill="#5B9BD5"/>'
    }</svg>`;
    card.append(thumb, el("span", { text: type === "doughnut" ? "Doughnut" : type.charAt(0).toUpperCase() + type.slice(1) }));
    card.addEventListener("click", () => {
      state.type = type;
      redraw();
    });
    cards.append(card);
  }

  rangeInput.addEventListener("input", () => {
    state.range = rangeInput.value.trim();
    redraw();
  });
  titleInput.addEventListener("input", () => { state.title = titleInput.value; redraw(); });
  headerRow.addEventListener("change", () => { state.headerRow = headerRow.checked; redraw(); });
  headerCol.addEventListener("change", () => { state.headerCol = headerCol.checked; redraw(); });

  const body = el("div", {}, [
    el("div", { class: "dialog-section-title", text: "Chart type" }),
    cards,
    el("div", { class: "dialog-section-title", text: "Data" }),
    el("div", { class: "chart-range-row" }, [
      rangeInput,
      el("label", { class: "check-row", style: "margin:0" }, [headerRow, el("span", { text: "First row contains labels" })]),
    ]),
    el("div", { style: "height:8px" }),
    el("label", { class: "check-row" }, [headerCol, el("span", { text: "First column contains labels" })]),
    field("Chart title", titleInput),
    el("div", { class: "chart-dialog-preview" }, [preview]),
  ]);
  redraw();

  return showDialog({
    title: existing ? "Chart" : "Insert Chart",
    size: "wide",
    body,
    buttons: [
      { label: "Cancel", kind: "ghost" },
      {
        label: existing ? "OK" : "Insert chart",
        kind: "primary",
        onClick: () => {
          const range = parseRange(state.range);
          if (!range) {
            toast("Select a data range first.");
            return false;
          }
          const chart = {
            id: existing?.id || "chart-" + Date.now().toString(36),
            type: state.type,
            range: rangeLabel(range),
            title: state.title,
            headerRow: state.headerRow,
            headerCol: state.headerCol,
            x: existing?.x ?? Math.round((grid.colPrefix[grid.active.col - 1] || 0) / grid.scale) + 8,
            y: existing?.y ?? Math.round((grid.rowPrefix[grid.active.row - 1] || 0) / grid.scale) + 8,
            w: existing?.w || 420,
            h: existing?.h || 260,
          };
          if (existing) {
            app.mutate("chart", () => Object.assign(existing, chart));
          } else {
            app.mutate("chart", () => { sheet.charts.push(chart); });
            toast("Chart inserted.");
          }
          return true;
        },
      },
    ],
  });
}

/* ----------------------------------------------------- conditional formatting */
const RULE_OPERATORS = [
  ["greaterThan", "Greater Than…"], ["lessThan", "Less Than…"], ["between", "Between…"],
  ["equal", "Equal To…"], ["notEqual", "Not Equal To…"],
];
const TEXT_OPERATORS = [
  ["containsText", "Text that Contains…"], ["notContains", "Text that Does Not Contain…"],
  ["beginsWith", "Text that Begins With…"], ["endsWith", "Text that Ends With…"],
];

function numericValues(app, range) {
  const sheet = app.grid.sheet();
  const values = [];
  const limit = Math.min((range.r2 - range.r1 + 1) * (range.c2 - range.c1 + 1), 100000);
  let count = 0;
  for (let row = range.r1; row <= range.r2 && count < limit; row += 1) {
    for (let col = range.c1; col <= range.c2 && count < limit; col += 1) {
      const cell = cellAt(sheet, row, col);
      if (cell && typeof cell.v === "number") values.push(cell.v);
      count += 1;
    }
  }
  return values;
}

function dateSerial(date) {
  const utc = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate());
  return Math.floor(utc / 86400000) + 25569;
}

export function conditionalDialog(app, preset = "greater") {
  const grid = app.grid;
  const range = rangeOf(grid);
  const state = {
    kind: TEXT_OPERATORS.some(([id]) => id === preset) ? "text"
      : preset === "date" ? "date"
      : preset === "top10" || preset === "bottom10" ? "rank"
      : preset === "aboveAverage" || preset === "belowAverage" ? "average"
      : "cell",
    operator: preset,
    value: "",
    value2: "",
    rank: 10,
    period: "today",
    fill: "#FFC7CE",
    color: "#9C0006",
    bold: false,
  };
  const body = el("div");
  const fillSample = el("span", { class: "color-sample" });
  fillSample.style.background = state.fill;
  const fontSample = el("span", { class: "color-sample" });
  fontSample.style.background = state.color;
  const boldBox = el("input", { type: "checkbox" });
  boldBox.checked = state.bold;
  boldBox.addEventListener("change", () => { state.bold = boldBox.checked; });

  const draw = () => {
    body.replaceChildren();
    const kindSelect = selectControl([
      { value: "cell", label: "Cell Value" },
      { value: "text", label: "Text" },
      { value: "date", label: "A Date Occurring" },
      { value: "rank", label: "Top / Bottom" },
      { value: "average", label: "Above / Below Average" },
    ], state.kind);
    kindSelect.addEventListener("change", () => {
      state.kind = kindSelect.value;
      state.operator = state.kind === "text" ? "containsText"
        : state.kind === "date" ? "date"
        : state.kind === "rank" ? "top10"
        : state.kind === "average" ? "aboveAverage" : "greaterThan";
      draw();
    });
    body.append(el("div", { class: "dialog-grid" }, [
      field("Rule type", kindSelect),
      field("Applies to", el("input", { type: "text", value: rangeLabel(range), disabled: true })),
    ]));

    if (state.kind === "cell") {
      const operator = selectControl(RULE_OPERATORS.map(([value, label]) => ({ value, label })), state.operator);
      operator.addEventListener("change", () => { state.operator = operator.value; draw(); });
      const first = el("input", { type: "text", value: state.value, placeholder: "Value" });
      first.addEventListener("input", () => { state.value = first.value; });
      const second = el("input", { type: "text", value: state.value2, placeholder: "And" });
      second.addEventListener("input", () => { state.value2 = second.value; });
      body.append(el("div", { class: "dialog-grid" }, [
        field("Operator", operator),
        field("Value", first),
        state.operator === "between" ? field("And", second) : el("div"),
      ]));
    } else if (state.kind === "text") {
      const operator = selectControl(TEXT_OPERATORS.map(([value, label]) => ({ value, label })), state.operator);
      operator.addEventListener("change", () => { state.operator = operator.value; });
      const first = el("input", { type: "text", value: state.value, placeholder: "Text" });
      first.addEventListener("input", () => { state.value = first.value; });
      body.append(el("div", { class: "dialog-grid" }, [
        field("Operator", operator),
        field("Text", first),
      ]));
    } else if (state.kind === "date") {
      const period = selectControl([
        { value: "today", label: "Today" }, { value: "yesterday", label: "Yesterday" },
        { value: "tomorrow", label: "Tomorrow" }, { value: "last7", label: "In the last 7 days" },
        { value: "thisMonth", label: "This month" }, { value: "thisYear", label: "This year" },
      ], state.period);
      period.addEventListener("change", () => { state.period = period.value; });
      body.append(field("Period", period));
    } else if (state.kind === "rank") {
      const rank = numberInput(state.rank, { min: 1, max: 1000 });
      rank.addEventListener("input", () => { state.rank = Number(rank.value) || 10; });
      const side = segmented([
        { value: "top10", label: "Top" }, { value: "bottom10", label: "Bottom" },
      ], state.operator === "bottom10" ? "bottom10" : "top10", (value) => { state.operator = value; });
      body.append(el("div", { class: "dialog-grid" }, [
        field("Show", side),
        field("Items", rank),
      ]));
    } else {
      const side = segmented([
        { value: "aboveAverage", label: "Above average" }, { value: "belowAverage", label: "Below average" },
      ], state.operator, (value) => { state.operator = value; });
      body.append(field("Show values that are", side));
    }

    const fillButton = el("button", { class: "btn", type: "button", text: "Fill" });
    fillButton.addEventListener("click", () => {
      colorMenu(fillButton, {
        title: "Theme Colors",
        onPick: (value) => {
          state.fill = value || "";
          fillSample.style.background = state.fill || "transparent";
        },
      });
    });
    const fontButton = el("button", { class: "btn", type: "button", text: "Font Color" });
    fontButton.addEventListener("click", () => {
      colorMenu(fontButton, {
        title: "Standard Colors",
        onPick: (value) => {
          state.color = value || "#000000";
          fontSample.style.background = state.color;
        },
      });
    });
    body.append(
      el("div", { class: "dialog-section-title", text: "Format" }),
      el("div", { class: "input-row" }, [fillSample, fillButton, fontSample, fontButton, el("label", { class: "check-row", style: "margin:0" }, [boldBox, el("span", { text: "Bold" })])]),
      el("p", { class: "dialog-note", text: "Rules are evaluated in order; the first matching rule wins." }),
    );
  };
  draw();

  return showDialog({
    title: "New Formatting Rule",
    size: "wide",
    body,
    buttons: [
      { label: "Cancel", kind: "ghost" },
      {
        label: "OK",
        kind: "primary",
        onClick: () => {
          const rule = {
            range: rangeLabel(range),
            type: "cellIs",
            operator: "greaterThan",
            formula1: state.value,
            formula2: state.value2,
            fill: state.fill,
            color: state.color,
            bold: state.bold,
          };
          if (state.kind === "text") {
            rule.operator = state.operator;
            rule.formula1 = state.value;
          } else if (state.kind === "date") {
            const today = new Date();
            const serial = dateSerial(today);
            const spans = {
              today: [serial, serial],
              yesterday: [serial - 1, serial - 1],
              tomorrow: [serial + 1, serial + 1],
              last7: [serial - 6, serial],
              thisMonth: [dateSerial(new Date(today.getFullYear(), today.getMonth(), 1)),
                dateSerial(new Date(today.getFullYear(), today.getMonth() + 1, 0))],
              thisYear: [dateSerial(new Date(today.getFullYear(), 0, 1)),
                dateSerial(new Date(today.getFullYear(), 11, 31))],
            };
            const [first, last] = spans[state.period] || spans.today;
            rule.operator = first === last ? "equal" : "between";
            rule.formula1 = String(first);
            rule.formula2 = String(last);
            rule.numFmt = "m/d/yyyy";
          } else if (state.kind === "rank") {
            const values = numericValues(app, range).sort((left, right) => left - right);
            if (!values.length) {
              toast("The chart has no numeric data.");
              return false;
            }
            const index = state.operator === "bottom10"
              ? Math.min(values.length - 1, state.rank - 1)
              : Math.max(0, values.length - state.rank);
            rule.operator = state.operator === "bottom10" ? "lessThanOrEqual" : "greaterThanOrEqual";
            rule.formula1 = String(values[index]);
          } else if (state.kind === "average") {
            const values = numericValues(app, range);
            if (!values.length) {
              toast("The chart has no numeric data.");
              return false;
            }
            const average = values.reduce((sum, value) => sum + value, 0) / values.length;
            rule.operator = state.operator === "belowAverage" ? "lessThan" : "greaterThan";
            rule.formula1 = String(average);
          } else {
            rule.operator = state.operator;
            if (state.operator === "between" && state.value2 === "") {
              toast("Enter the second value for the Between rule.");
              return false;
            }
          }
          app.mutate("conditional", () => {
            sheet().conditional = sheet().conditional || [];
            sheet().conditional.push(rule);
          });
          toast("Conditional format applied.");
          return true;
        },
      },
    ],
  });

  function sheet() {
    return grid.sheet();
  }
}

export function conditionalRulesDialog(app) {
  const sheet = app.grid.sheet();
  const list = el("div", { class: "fmt-list", style: "height:280px" });
  const draw = () => {
    list.replaceChildren();
    const rules = sheet.conditional || [];
    if (!rules.length) {
      list.append(el("div", { class: "palette-empty", text: "No rules for this sheet." }));
      return;
    }
    rules.forEach((rule, index) => {
      const item = el("div", { class: "list-item" }, [
        el("div", { text: (rule.range || "") + " — " + (rule.type === "colorScale" ? "Color Scale" : (rule.operator || "")) }),
        el("div", { class: "li-sub", text: `${rule.formula1 ?? ""}${rule.formula2 ? " … " + rule.formula2 : ""}` }),
      ]);
      const remove = el("button", { class: "btn ghost", type: "button", text: "Delete" });
      remove.style.marginLeft = "auto";
      remove.addEventListener("click", (event) => {
        event.stopPropagation();
        app.mutate("conditional", () => {
          sheet.conditional = rules.filter((entry) => entry !== rule);
        });
        draw();
      });
      item.style.display = "flex";
      item.style.alignItems = "center";
      item.prepend(el("span", { class: "muted", text: String(index + 1) }));
      item.append(remove);
      list.append(item);
    });
  };
  draw();
  return showDialog({
    title: "Manage Rules",
    size: "wide",
    body: list,
    buttons: [
      {
        label: "Delete All", left: true, onClick: () => {
          app.mutate("conditional", () => { sheet.conditional = []; });
          draw();
        },
      },
      { label: "Close", kind: "primary" },
    ],
  });
}

/* ------------------------------------------------------------------- sort */
export function sortDialog(app) {
  const grid = app.grid;
  const sheet = grid.sheet();
  const selected = grid.sel;
  const used = usedRange(sheet);
  const range = selected.r1 === selected.r2 && selected.c1 === selected.c2 && used ? { ...used } : { ...selected };
  const guessedHeader = (() => {
    const first = [];
    for (let col = range.c1; col <= Math.min(range.c2, range.c1 + 8); col += 1) {
      const cell = cellAt(sheet, range.r1, col);
      first.push(cell ? typeof cell.v : "empty");
    }
    return first.some((kind) => kind === "string");
  })();
  const state = { keys: [{ col: range.c1, order: "asc" }], hasHeader: guessedHeader };
  const rows = el("div");
  const header = el("input", { type: "checkbox" });
  header.checked = state.hasHeader;
  header.addEventListener("change", () => { state.hasHeader = header.checked; });

  const draw = () => {
    rows.replaceChildren();
    state.keys.forEach((key, index) => {
      const col = selectControl(columnOptions(app, range, state.hasHeader), key.col);
      col.addEventListener("change", () => { key.col = Number(col.value); });
      const order = selectControl([
        { value: "asc", label: "Smallest to Largest" },
        { value: "desc", label: "Largest to Smallest" },
      ], key.order);
      order.addEventListener("change", () => { key.order = order.value; });
      const remove = el("button", { class: "sort-remove", type: "button", html: icon("close"), title: "Remove" });
      remove.addEventListener("click", () => {
        if (state.keys.length === 1) return;
        state.keys.splice(index, 1);
        draw();
      });
      rows.append(el("div", { class: "sort-row" }, [
        el("span", { class: "unit", text: index === 0 ? "Sort by" : "Then by" }),
        col, order, remove,
      ]));
    });
  };
  draw();
  const add = el("button", { class: "btn ghost", type: "button", text: "Add Level" });
  add.addEventListener("click", () => {
    if (state.keys.length >= 4) return;
    state.keys.push({ col: range.c1, order: "asc" });
    draw();
  });
  const body = el("div", {}, [
    el("p", { class: "dialog-note", text: "Range: " + rangeLabel(range) }),
    rows,
    add,
    el("div", { style: "height:10px" }),
    el("label", { class: "check-row" }, [header, el("span", { text: "My data has headers" })]),
  ]);
  return showDialog({
    title: "Sort",
    size: "wide",
    body,
    buttons: [
      { label: "Cancel", kind: "ghost" },
      {
        label: "OK",
        kind: "primary",
        onClick: () => {
          grid.setSelection(range.r1, range.c1, range.r2, range.c2, { activeRow: range.r1, activeCol: range.c1 });
          if (!grid.applySort(state.keys, { hasHeader: state.hasHeader })) {
            toast("Select a range first.");
            return false;
          }
          toast("Data sorted.");
        },
      },
    ],
  });
}

/* ----------------------------------------------------------------- filter */
export function filterPopover(app, col, anchor) {
  const grid = app.grid;
  const sheet = grid.sheet();
  const range = parseRange(sheet.filters?.range || "") || (() => {
    const used = usedRange(sheet);
    return used ? { ...used } : { r1: 1, c1: 1, r2: 20, c2: 1 };
  })();
  if (col < range.c1 || col > range.c2) return;
  const values = grid.filterValues(col, range);
  const existing = (sheet.filters?.criteria || []).find((criterion) => Number(criterion.col) === col);
  const selected = new Set(existing?.values || values.map((item) => item.value));
  const panel = el("div", { class: "filter-panel" });
  const search = el("input", { type: "text", placeholder: "Search", class: "function-search" });
  const list = el("div", { class: "filter-list" });
  const all = el("input", { type: "checkbox" });
  all.checked = !existing || selected.size === values.length;

  const rebuild = () => {
    list.replaceChildren();
    const needle = search.value.trim().toLowerCase();
    let shown = 0;
    for (const item of values) {
      if (needle && !item.value.toLowerCase().includes(needle)) continue;
      if (shown > 500) break;
      shown += 1;
      const box = el("input", { type: "checkbox" });
      box.checked = selected.has(item.value);
      box.addEventListener("change", () => {
        if (box.checked) selected.add(item.value);
        else selected.delete(item.value);
        all.checked = selected.size === values.length;
      });
      list.append(el("label", { class: "filter-item" }, [
        box,
        el("span", { text: item.value || "(Blank)" }),
        el("span", { class: "muted", style: "margin-left:auto", text: String(item.count) }),
      ]));
    }
    if (!shown) list.append(el("div", { class: "palette-empty", text: "No results" }));
  };
  search.addEventListener("input", rebuild);
  all.addEventListener("change", () => {
    selected.clear();
    if (all.checked) for (const item of values) selected.add(item.value);
    rebuild();
  });
  rebuild();

  const apply = el("button", { class: "btn primary", type: "button", text: "OK" });
  apply.addEventListener("click", () => {
    const criteria = (sheet.filters?.criteria || []).filter((criterion) => Number(criterion.col) !== col);
    if (selected.size < values.length) {
      criteria.push({ col, values: [...selected] });
    }
    grid.applyFilter(range, criteria);
    toast(criteria.length ? "Filter applied." : "Filter cleared.");
    closeMenus();
  });
  const clear = el("button", { class: "btn", type: "button", text: "Clear Filter" });
  clear.addEventListener("click", () => {
    const criteria = (sheet.filters?.criteria || []).filter((criterion) => Number(criterion.col) !== col);
    if (criteria.length) grid.applyFilter(range, criteria);
    else grid.clearFilter();
    closeMenus();
  });
  panel.append(
    el("div", { class: "menu-head", text: "Filter by value" }),
    search,
    el("label", { class: "filter-item" }, [all, el("span", { text: "Select All" })]),
    list,
    el("div", { class: "input-row", style: "justify-content:flex-end;margin-top:8px" }, [clear, apply]),
  );
  openMenu(anchor, [{ type: "custom", node: panel }], { minWidth: 280 });
}

/* ------------------------------------------------------------- validation */
export function validationDialog(app) {
  const grid = app.grid;
  const sheet = grid.sheet();
  const range = rangeOf(grid);
  const current = (sheet.validations || []).find((rule) => rule.range === rangeLabel(range)) || {};
  const state = {
    type: current.type || "list",
    operator: current.operator || "between",
    formula1: current.formula1 || "",
    formula2: current.formula2 || "",
    allowBlank: current.allowBlank !== false,
    prompt: current.prompt || "",
    error: current.error || "The value does not match the data validation rule.",
  };
  const body = el("div");
  const draw = () => {
    body.replaceChildren();
    const type = selectControl([
      { value: "any", label: "Any value" },
      { value: "whole", label: "Whole number" },
      { value: "decimal", label: "Decimal" },
      { value: "list", label: "List" },
      { value: "date", label: "Date" },
      { value: "time", label: "Time" },
      { value: "textLength", label: "Text length" },
      { value: "custom", label: "Custom" },
    ], state.type);
    type.addEventListener("change", () => { state.type = type.value; draw(); });
    const operator = selectControl([
      { value: "between", label: "Between" }, { value: "notBetween", label: "Not Between" },
      { value: "equal", label: "Equal To" }, { value: "notEqual", label: "Not Equal To" },
      { value: "greaterThan", label: "Greater Than" }, { value: "lessThan", label: "Less Than" },
      { value: "greaterThanOrEqual", label: "Greater Than or Equal To" },
      { value: "lessThanOrEqual", label: "Less Than or Equal To" },
    ], state.operator);
    operator.addEventListener("change", () => { state.operator = operator.value; });
    const first = el("input", { type: "text", value: state.formula1, placeholder: state.type === "list" ? "Yes, No, Maybe" : "Value" });
    first.addEventListener("input", () => { state.formula1 = first.value; });
    const second = el("input", { type: "text", value: state.formula2, placeholder: "And" });
    second.addEventListener("input", () => { state.formula2 = second.value; });
    const allow = el("input", { type: "checkbox" });
    allow.checked = state.allowBlank;
    allow.addEventListener("change", () => { state.allowBlank = allow.checked; });
    const prompt = el("input", { type: "text", value: state.prompt });
    prompt.addEventListener("input", () => { state.prompt = prompt.value; });
    const error = el("input", { type: "text", value: state.error });
    error.addEventListener("input", () => { state.error = error.value; });
    body.append(
      el("div", { class: "dialog-grid" }, [
        field("Allow", type),
        field("Data", selectControl([{ value: "list", label: rangeLabel(range) }], "list", { disabled: true })),
        state.type === "list" || state.type === "custom" ? el("div") : field("Criteria", operator),
        field(state.type === "list" ? "Source" : "Value", first),
        ["between", "notBetween"].includes(state.operator) && !["list", "custom", "any"].includes(state.type)
          ? field("And", second) : el("div"),
      ]),
      el("div", { class: "dialog-section-title", text: "Input" }),
      el("label", { class: "check-row" }, [allow, el("span", { text: "Ignore blank" })]),
      field("Prompt", prompt),
      field("Error message", error),
    );
  };
  draw();
  return showDialog({
    title: "Data Validation",
    size: "wide",
    body,
    buttons: [
      {
        label: "Clear All", left: true, onClick: () => {
          app.mutate("validation", () => {
            sheet.validations = (sheet.validations || []).filter((rule) => rule.range !== rangeLabel(range));
          });
          toast("Validation saved.");
          return false;
        },
      },
      { label: "Cancel", kind: "ghost" },
      {
        label: "OK",
        kind: "primary",
        onClick: () => {
          const rule = {
            range: rangeLabel(range), type: state.type, operator: state.operator,
            formula1: state.formula1, formula2: state.formula2,
            allowBlank: state.allowBlank, prompt: state.prompt, error: state.error,
          };
          if (state.type !== "any" && !state.formula1) {
            toast("Enter a validation value.");
            return false;
          }
          app.mutate("validation", () => {
            sheet.validations = (sheet.validations || []).filter((entry) => entry.range !== rule.range);
            if (state.type !== "any") sheet.validations.push(rule);
          });
          toast("Validation saved.");
        },
      },
    ],
  });
}

/* ------------------------------------------------------------ names/functions */
export function nameManagerDialog(app) {
  const model = app.model;
  const table = el("table", { class: "name-table" });
  const body = el("div");
  const draw = () => {
    table.replaceChildren(el("thead", {}, [el("tr", {}, [
      el("th", { text: "Name" }), el("th", { text: "Value" }), el("th", { text: "Scope" }),
    ])]));
    const tbody = el("tbody");
    for (const [name, ref] of Object.entries(model.names || {})) {
      const nameInput = el("input", { type: "text", value: name });
      const refInput = el("input", { type: "text", value: ref });
      const remove = el("button", { class: "btn ghost", type: "button", text: "Delete" });
      const row = el("tr", {}, [
        el("td", {}, [nameInput]), el("td", {}, [refInput]), el("td", { text: "Workbook" }), el("td", {}, [remove]),
      ]);
      nameInput.addEventListener("change", () => {
        const next = nameInput.value.trim();
        if (!validName(next)) {
          toast("No spaces allowed");
          nameInput.value = name;
          return;
        }
        app.mutate("name", () => {
          const value = model.names[name];
          delete model.names[name];
          model.names[next] = value;
        });
        draw();
      });
      refInput.addEventListener("change", () => {
        app.mutate("name", () => { model.names[name] = refInput.value.trim(); });
      });
      remove.addEventListener("click", () => {
        app.mutate("name", () => { delete model.names[name]; });
        toast("Name deleted.");
        draw();
      });
      tbody.append(row);
    }
    if (!Object.keys(model.names || {}).length) {
      tbody.append(el("tr", {}, [el("td", { colspan: 4, class: "muted", text: "No names defined." })]));
    }
    table.append(tbody);
  };
  draw();
  body.append(table, el("p", { class: "dialog-note", text: "Names make formulas easier to read; use them anywhere a range is accepted." }));
  return showDialog({
    title: "Name Manager",
    size: "wide",
    body,
    buttons: [
      { label: "New…", left: true, onClick: () => { defineNameDialog(app).then(() => draw()); return false; } },
      { label: "Close", kind: "primary" },
    ],
  });
}

export function defineNameDialog(app, initialRef = null) {
  const grid = app.grid;
  const ref = initialRef || `${grid.sheet().name}!${rangeText(grid.sel)}`;
  const name = el("input", { type: "text", placeholder: "Name" });
  const refers = el("input", { type: "text", value: ref });
  const body = el("div", {}, [
    field("Name", name),
    field("Refers to", refers),
    el("p", { class: "dialog-note", text: "Names cannot contain spaces and must start with a letter, underscore or backslash." }),
  ]);
  return showDialog({
    title: "New Name",
    size: "narrow",
    body,
    buttons: [
      { label: "Cancel", kind: "ghost" },
      {
        label: "OK", kind: "primary", onClick: () => {
          const value = name.value.trim();
          if (!validName(value)) {
            toast("No spaces allowed");
            return false;
          }
          app.mutate("name", () => { app.model.names[value] = refers.value.trim(); });
          toast("Name saved.");
        },
      },
    ],
  });
}

const FUNCTION_CATEGORIES = [
  ["Math & Trig", ["SUM", "PRODUCT", "ABS", "SQRT", "EXP", "LN", "LOG", "LOG10", "POWER", "MOD", "INT", "ROUND", "ROUNDUP", "ROUNDDOWN", "RAND", "RANDBETWEEN"]],
  ["Statistical", ["AVERAGE", "COUNT", "COUNTA", "COUNTBLANK", "MIN", "MAX", "MEDIAN", "SUMIF", "SUMIFS", "AVERAGEIF", "COUNTIF", "COUNTIFS"]],
  ["Logical", ["IF", "IFS", "AND", "OR", "NOT", "IFERROR", "IFNA", "TRUE", "FALSE"]],
  ["Lookup & Reference", ["VLOOKUP", "HLOOKUP", "INDEX", "MATCH", "XLOOKUP", "CHOOSE", "ROW", "COLUMN", "ROWS", "COLUMNS"]],
  ["Text", ["CONCAT", "CONCATENATE", "TEXTJOIN", "LEFT", "RIGHT", "MID", "LEN", "LOWER", "UPPER", "PROPER", "TRIM", "SUBSTITUTE", "REPLACE", "FIND", "SEARCH", "TEXT", "VALUE"]],
  ["Date & Time", ["TODAY", "NOW", "DATE", "YEAR", "MONTH", "DAY", "HOUR", "MINUTE", "SECOND", "DAYS", "EDATE", "EOMONTH", "WEEKDAY"]],
  ["Financial", ["PMT", "FV", "PV", "NPV", "IRR"]],
];

export function functionsByCategory(category) {
  const entry = FUNCTION_CATEGORIES.find(([name]) => name === category);
  return entry ? entry[1] : [];
}

export function functionDialog(app) {
  const search = el("input", { class: "function-search", type: "text", placeholder: "Search functions", spellcheck: "false" });
  const categoryList = el("div", { class: "fmt-list" });
  const functionList = el("div", { class: "function-list" });
  const detail = el("div", { class: "function-detail" });
  const state = { category: "All", name: "" };

  const insert = (name) => {
    app.insertFunction(name);
    closeMenus();
  };
  const selectFunction = (name) => {
    state.name = name;
    for (const item of functionList.children) item.classList.toggle("active", item.dataset.name === name);
    const arity = app.functionArity(name);
    const usage = arity.max === 0 ? `=${name}()` : `=${name}(${Array.from({ length: Math.max(1, arity.min) }, (_, index) => "arg" + (index + 1)).join(", ")})`;
    detail.replaceChildren(
      el("h4", { text: name }),
      el("div", { html: `<code>${usage}</code>` }),
      el("p", { text: `Takes ${arity.min}${arity.max !== arity.min ? " to " + (arity.max === 255 ? "255" : arity.max) : ""} argument(s).` }),
      el("button", { class: "btn primary", type: "button", text: "Insert Function", onclick: () => insert(name) }),
    );
  };
  const fill = () => {
    functionList.replaceChildren();
    const needle = search.value.trim().toUpperCase();
    const names = state.category === "All"
      ? FUNCTION_NAMES
      : functionsByCategory(state.category);
    const filtered = names.filter((name) => !needle || name.includes(needle));
    for (const name of filtered) {
      const item = el("div", { class: "list-item", dataset: { name } }, [
        el("div", { text: name }),
        el("div", { class: "li-sub", text: "Function" }),
      ]);
      item.addEventListener("click", () => selectFunction(name));
      item.addEventListener("dblclick", () => insert(name));
      functionList.append(item);
    }
    if (!filtered.length) functionList.append(el("div", { class: "palette-empty", text: "No results" }));
    if (!state.name && filtered.length) selectFunction(filtered[0]);
  };
  for (const category of ["All", ...FUNCTION_CATEGORIES.map(([name]) => name)]) {
    const item = el("div", { class: "list-item" + (category === state.category ? " active" : ""), text: category });
    item.addEventListener("click", () => {
      state.category = category;
      for (const other of categoryList.children) other.classList.toggle("active", other === item);
      state.name = "";
      fill();
    });
    categoryList.append(item);
  }
  search.addEventListener("input", fill);
  fill();
  const body = el("div", {}, [
    search,
    el("div", { style: "display:grid;grid-template-columns:170px 220px 1fr;gap:12px" }, [
      el("div", {}, [el("div", { class: "dialog-section-title", text: "Category" }), categoryList]),
      el("div", {}, [el("div", { class: "dialog-section-title", text: "Function" }), functionList]),
      el("div", {}, [el("div", { class: "dialog-section-title", text: "About" }), detail]),
    ]),
  ]);
  return showDialog({
    title: "Insert Function",
    size: "xwide",
    body,
    buttons: [
      { label: "Cancel", kind: "ghost" },
      { label: "Insert Function", kind: "primary", onClick: () => { if (state.name) insert(state.name); } },
    ],
  });
}

/* ---------------------------------------------------------- row/column size */
export function rowHeightDialog(app) {
  const grid = app.grid;
  const input = numberInput(20, { min: 6, max: 400 });
  const body = el("div", {}, [field("Row height", input), el("p", { class: "dialog-note", text: "Applied to the selected rows." })]);
  return showDialog({
    title: "Row Height",
    size: "narrow",
    body,
    buttons: [
      { label: "AutoFit", left: true, onClick: () => grid.autoFitSelection() },
      { label: "Cancel", kind: "ghost" },
      { label: "OK", kind: "primary", onClick: () => grid.setSelectedRowHeight(Number(input.value) || 20) },
    ],
  });
}

export function columnWidthDialog(app) {
  const grid = app.grid;
  const input = numberInput(64, { min: 10, max: 480 });
  const body = el("div", {}, [field("Column width", input), el("p", { class: "dialog-note", text: "Applied to the selected columns." })]);
  return showDialog({
    title: "Column Width",
    size: "narrow",
    body,
    buttons: [
      { label: "AutoFit", left: true, onClick: () => grid.autoFitSelection() },
      { label: "Cancel", kind: "ghost" },
      { label: "OK", kind: "primary", onClick: () => grid.setSelectedColumnWidth(Number(input.value) || 64) },
    ],
  });
}

/* ------------------------------------------------------------- page setup */
export function pageSetupDialog(app, initialTab = "page") {
  const sheet = app.grid.sheet();
  const state = clone(sheet.page) || {};
  state.margins = { t: 72, r: 72, b: 72, l: 72, header: 30, footer: 30, ...(state.margins || {}) };
  state.orientation = state.orientation || "portrait";
  state.size = state.size || "A4";
  state.scale = state.scale || 100;
  const tabs = [["page", "Page"], ["margins", "Margins"], ["header", "Header/Footer"], ["sheet", "Sheet"]];
  const tabBar = el("div", { class: "dialog-tabs" });
  const host = el("div");
  let active = initialTab;
  const body = el("div", {}, [tabBar, host]);

  const marginField = (key, label) => {
    const input = numberInput(state.margins[key], { min: 0, max: 288, step: 1 });
    input.style.width = "80px";
    input.addEventListener("input", () => { state.margins[key] = Number(input.value) || 0; });
    return field(label, el("div", { class: "input-row" }, [input, el("span", { class: "unit", text: "pt" })]));
  };

  const draw = () => {
    for (const button of tabBar.children) button.classList.toggle("active", button.dataset.tab === active);
    host.replaceChildren();
    if (active === "page") {
      const orientation = segmented([
        { value: "portrait", label: "Portrait" }, { value: "landscape", label: "Landscape" },
      ], state.orientation, (value) => { state.orientation = value; });
      const size = selectControl([
        { value: "A4", label: "A4" }, { value: "Letter", label: "Letter" },
      ], state.size);
      size.addEventListener("change", () => { state.size = size.value; });
      const scale = numberInput(state.scale, { min: 10, max: 400 });
      scale.addEventListener("input", () => { state.scale = Number(scale.value) || 100; });
      host.append(el("div", { class: "dialog-grid" }, [
        el("div", { class: "full" }, [field("Orientation", orientation)]),
        field("Paper size", size),
        field("Scaling", el("div", { class: "input-row" }, [scale, el("span", { class: "unit", text: "% normal size" })])),
      ]));
    } else if (active === "margins") {
      host.append(el("div", { class: "dialog-grid three" }, [
        marginField("t", "Top"), marginField("b", "Bottom"), marginField("header", "Header"),
        marginField("l", "Left"), marginField("r", "Right"), marginField("footer", "Footer"),
      ]));
      host.append(el("p", { class: "dialog-note", text: "Margins are measured in points; 72 points make one inch." }));
    } else if (active === "header") {
      const header = el("input", { type: "text", value: state.header || "", placeholder: "e.g. &A" });
      const footer = el("input", { type: "text", value: state.footer || "&P" });
      header.addEventListener("input", () => { state.header = header.value; });
      footer.addEventListener("input", () => { state.footer = footer.value; });
      host.append(el("div", { class: "dialog-grid" }, [
        field("Header", header), field("Footer", footer),
      ]), el("p", { class: "dialog-note", text: "&P page number · &N page count · &A sheet name · &D date · &F file name" }));
    } else {
      const area = el("input", { type: "text", value: state.area || "", placeholder: "A1:H40" });
      area.addEventListener("input", () => { state.area = area.value; });
      const gridlines = el("input", { type: "checkbox" });
      gridlines.checked = !!state.gridlines;
      gridlines.addEventListener("change", () => { state.gridlines = gridlines.checked; });
      const headings = el("input", { type: "checkbox" });
      headings.checked = !!state.headings;
      headings.addEventListener("change", () => { state.headings = headings.checked; });
      host.append(el("div", { class: "dialog-grid" }, [
        field("Print area", area),
        el("div", {}, [
          el("label", { class: "check-row" }, [gridlines, el("span", { text: "Print gridlines" })]),
          el("label", { class: "check-row" }, [headings, el("span", { text: "Print headings" })]),
        ]),
      ]));
      const clear = el("button", { class: "btn ghost", type: "button", text: "Clear print area" });
      clear.addEventListener("click", () => { state.area = ""; area.value = ""; });
      host.append(clear);
    }
  };
  for (const [id, label] of tabs) {
    const button = el("button", { class: "dialog-tab", type: "button", text: label, dataset: { tab: id } });
    button.addEventListener("click", () => { active = id; draw(); });
    tabBar.append(button);
  }
  draw();
  return showDialog({
    title: "Page Setup",
    size: "wide",
    body,
    buttons: [
      { label: "Cancel", kind: "ghost" },
      {
        label: "OK", kind: "primary", onClick: () => {
          app.mutate("page", () => { sheet.page = { ...state }; }, { layout: false });
          toast("Page setup saved.");
        },
      },
    ],
  });
}

export function headerFooterDialog(app) {
  const sheet = app.grid.sheet();
  const header = el("input", { type: "text", value: sheet.page?.header || "", placeholder: "Header text" });
  const footer = el("input", { type: "text", value: sheet.page?.footer || "&P", placeholder: "Footer text" });
  const preview = el("div", { class: "fmt-preview", style: "flex-direction:column;align-items:stretch;gap:6px;font-size:11px" });
  const refresh = () => {
    preview.replaceChildren(
      el("div", { style: "border-bottom:1px solid var(--border);padding-bottom:4px", text: header.value || "\u00a0" }),
      el("div", { style: "height:26px;color:var(--muted);display:flex;align-items:center;justify-content:center", text: "Page content" }),
      el("div", { style: "border-top:1px solid var(--border);padding-top:4px;text-align:center", text: footer.value || "\u00a0" }),
    );
  };
  header.addEventListener("input", refresh);
  footer.addEventListener("input", refresh);
  refresh();
  return showDialog({
    title: "Header & Footer",
    size: "wide",
    body: el("div", {}, [
      el("div", { class: "dialog-grid" }, [field("Header", header), field("Footer", footer)]),
      el("div", { class: "dialog-section-title", text: "Preview" }),
      preview,
      el("p", { class: "dialog-note", text: "&P page number · &N page count · &A sheet name · &D date · &F file name" }),
    ]),
    buttons: [
      { label: "Cancel", kind: "ghost" },
      {
        label: "OK", kind: "primary", onClick: () => {
          app.mutate("page", () => {
            sheet.page = { ...(sheet.page || {}), header: header.value, footer: footer.value };
          }, { layout: false });
        },
      },
    ],
  });
}

/* --------------------------------------------------------------- hyperlink */
export function hyperlinkDialog(app) {
  const grid = app.grid;
  const sheet = grid.sheet();
  const address = grid.activeAddress();
  const cell = cellAt(sheet, grid.active.row, grid.active.col);
  const link = el("input", { type: "text", value: cell?.h || "", placeholder: "https://example.com or Sheet1!A1" });
  const display = el("input", { type: "text", value: cell?.v !== undefined ? String(cell.v) : "", placeholder: "Text to display" });
  const body = el("div", {}, [
    field("Address", link),
    field("Text to display", display),
    el("p", { class: "dialog-note", text: "Use a web address or an internal reference such as Sheet1!A1." }),
  ]);
  return showDialog({
    title: "Link",
    size: "narrow",
    body,
    buttons: [
      cell?.h ? {
        label: "Remove Link", left: true, onClick: () => {
          app.mutate("link", () => {
            const target = ensureCell(sheet, address);
            delete target.h;
          });
        },
      } : null,
      { label: "Cancel", kind: "ghost" },
      {
        label: "OK", kind: "primary", onClick: () => {
          const url = link.value.trim();
          if (!url) return false;
          app.mutate("link", () => {
            const target = ensureCell(sheet, address);
            target.h = url;
            if (display.value) target.v = display.value;
            target.s = { ...(target.s || {}), color: "#0563C1", u: "single" };
          });
        },
      },
    ].filter(Boolean),
  });
}

/* ------------------------------------------------------- symbols and icons */
const SYMBOLS = [
  "€", "£", "¥", "¢", "©", "®", "™", "±", "×", "÷", "≤", "≥", "≠", "≈", "∞", "π",
  "Ω", "µ", "°", "§", "¶", "†", "‡", "•", "…", "‰", "α", "β", "γ", "δ", "ε", "λ",
  "Δ", "Σ", "Π", "√", "∫", "∂", "→", "←", "↑", "↓", "↔", "⇒", "⇐", "∑", "∏", "√",
  "½", "⅓", "¼", "¾", "⅛", "⅜", "⅝", "⅞", "①", "②", "③", "④", "⑤", "✓", "✗", "★",
];

export function symbolDialog(app) {
  const grid = el("div", { class: "palette-grid", style: "grid-template-columns:repeat(8,34px)" });
  for (const symbol of SYMBOLS) {
    const button = el("button", { class: "swatch", type: "button", text: symbol, title: symbol });
    button.style.cssText = "width:34px;height:32px;font-size:16px;background:var(--surface);color:var(--text)";
    button.addEventListener("click", () => {
      app.insertIntoCell(symbol);
      closeMenus();
    });
    grid.append(button);
  }
  return showDialog({
    title: "Symbol",
    size: "narrow",
    body: el("div", {}, [grid, el("p", { class: "dialog-note", text: "The symbol is inserted into the active cell editor." })]),
    buttons: [{ label: "Close", kind: "primary" }],
  });
}

const PICKER_ICONS = [
  "chartColumn", "chartPie", "table", "picture", "shapes", "link", "comment", "function",
  "calendar", "clock", "check", "star", "warning", "info", "success", "error",
  "lock", "unlock", "eye", "eyeOff", "save", "print", "share", "account",
  "undo", "redo", "search", "trash", "edit", "settings", "filter", "sortAZ",
  "sum", "average", "max", "min", "percent", "palette", "keyboard", "help",
];

export function iconsDialog(app) {
  const grid = el("div", { class: "palette-grid", style: "grid-template-columns:repeat(8,34px)" });
  for (const name of PICKER_ICONS) {
    const button = el("button", { class: "swatch", type: "button", title: name, html: icon(name) });
    button.style.cssText = "width:34px;height:32px;display:flex;align-items:center;justify-content:center;background:var(--surface);color:var(--text)";
    button.addEventListener("click", () => {
      const svg = icon(name)
        .replace("<svg", '<svg xmlns="http://www.w3.org/2000/svg"')
        .replace(/currentColor/g, "#2F5597");
      app.addImage("data:image/svg+xml;charset=utf-8," + encodeURIComponent(svg), 32, 32);
      closeMenus();
    });
    grid.append(button);
  }
  return showDialog({
    title: "Icons",
    size: "narrow",
    body: el("div", {}, [grid, el("p", { class: "dialog-note", text: "The icon is placed on the sheet and can be moved or resized." })]),
    buttons: [{ label: "Close", kind: "primary" }],
  });
}

export function textBoxDialog(app, existing = null) {
  const area = el("textarea", { rows: 4, placeholder: "Text" });
  area.value = existing?.text || "";
  area.style.cssText = "width:100%;border:1px solid var(--border-strong);border-radius:4px;padding:7px 8px;font-size:12px;resize:vertical;outline:none";
  const width = numberInput(existing?.w || 220, { min: 60, max: 1200 });
  const height = numberInput(existing?.h || 90, { min: 30, max: 800 });
  return showDialog({
    title: existing ? "Text Box" : "Text Box",
    size: "narrow",
    body: el("div", {}, [
      field("Text", area),
      el("div", { class: "dialog-grid" }, [field("Width", width), field("Height", height)]),
    ]),
    buttons: [
      { label: "Cancel", kind: "ghost" },
      {
        label: "OK", kind: "primary", onClick: () => {
          const text = area.value.trim();
          if (!text) return false;
          if (existing) {
            app.mutate("textbox", () => {
              existing.text = text;
              existing.w = Number(width.value) || existing.w;
              existing.h = Number(height.value) || existing.h;
            });
          } else {
            app.addTextBox(text, { w: Number(width.value) || 220, h: Number(height.value) || 90 });
          }
        },
      },
    ],
  });
}

/* --------------------------------------------------------------- sparkline */
export function sparklineDialog(app) {
  const grid = app.grid;
  const used = usedRange(grid.sheet());
  const range = el("input", { type: "text", value: used ? rangeLabel(used) : "A1:A10" });
  const location = el("input", { type: "text", value: grid.activeAddress() });
  return showDialog({
    title: "Sparklines",
    size: "narrow",
    body: el("div", {}, [
      field("Data range", range),
      field("Location", location),
      el("p", { class: "dialog-note", text: "The sparkline is a small line chart anchored to the location cell." }),
    ]),
    buttons: [
      { label: "Cancel", kind: "ghost" },
      {
        label: "OK", kind: "primary", onClick: () => {
          const area = parseRange(range.value.trim());
          const target = parseAddr(location.value.trim());
          if (!area || !target) {
            toast("Select a data range first.");
            return false;
          }
          const sheet = grid.sheet();
          app.mutate("chart", () => {
            sheet.charts.push({
              id: "spark-" + Date.now().toString(36),
              type: "line",
              range: rangeLabel(area),
              title: "",
              headerRow: false,
              headerCol: false,
              x: Math.round((grid.colPrefix[target.col - 1] || 0) / grid.scale) + 2,
              y: Math.round((grid.rowPrefix[target.row - 1] || 0) / grid.scale) + 2,
              w: 160,
              h: 44,
            });
          });
        },
      },
    ],
  });
}

/* ---------------------------------------------------------------- tables */
export const TABLE_STYLES = [
  { label: "Blue, Table Style Medium 2", headerFill: "#4472C4", headerColor: "#FFFFFF", bandFill: "#D9E2F3", border: "#8EA9DB" },
  { label: "Orange, Table Style Medium 3", headerFill: "#ED7D31", headerColor: "#FFFFFF", bandFill: "#FBE5D6", border: "#F4B183" },
  { label: "Green, Table Style Medium 4", headerFill: "#70AD47", headerColor: "#FFFFFF", bandFill: "#E2EFDA", border: "#A9D08E" },
  { label: "Light Blue, Table Style Medium 5", headerFill: "#5B9BD5", headerColor: "#FFFFFF", bandFill: "#DDEBF7", border: "#9DC3E6" },
  { label: "Light Gray, Table Style Medium 1", headerFill: "#7F7F7F", headerColor: "#FFFFFF", bandFill: "#F2F2F2", border: "#BFBFBF" },
];

function paintStyle(sheet, address, patch) {
  const cell = ensureCell(sheet, address);
  cell.s = { ...(cell.s || {}), ...patch };
}

function nextTableName(sheet) {
  const used = new Set((sheet.tables || []).map((table) => String(table.name || "")));
  let index = 1;
  while (used.has("Table" + index)) index += 1;
  return "Table" + index;
}

/** Apply one of the built-in table styles and remember the table range. */
export function applyTableStyle(app, index, range, hasHeader = true) {
  const grid = app.grid;
  const sheet = grid.sheet();
  const area = normalizeRange(range.r1, range.c1, range.r2, range.c2);
  if (!area) return "";
  const spec = TABLE_STYLES[Math.max(0, Math.min(TABLE_STYLES.length - 1, Number(index) || 0))];
  const rangeString = rangeText(area);
  let name = "";
  app.mutate("table", () => {
    if (!Array.isArray(sheet.tables)) sheet.tables = [];
    const existing = sheet.tables.find((table) => table.range === rangeString);
    name = existing?.name || nextTableName(sheet);
    for (let row = area.r1; row <= area.r2; row += 1) {
      const header = hasHeader && row === area.r1;
      const band = !header && ((row - (hasHeader ? area.r1 + 1 : area.r1)) % 2 === 1);
      const fill = header ? spec.headerFill : band ? spec.bandFill : "#FFFFFF";
      for (let col = area.c1; col <= area.c2; col += 1) {
        const patch = {
          fill,
          border: {
            t: { style: "thin", color: spec.border },
            r: { style: "thin", color: spec.border },
            b: { style: "thin", color: spec.border },
            l: { style: "thin", color: spec.border },
          },
        };
        if (header) {
          patch.b = true;
          patch.color = spec.headerColor;
        }
        paintStyle(sheet, addr(row, col), patch);
      }
    }
    const record = {
      id: existing?.id || "tbl-" + Date.now().toString(36),
      name,
      range: rangeString,
      style: TABLE_STYLES.indexOf(spec),
      hasHeader,
    };
    if (existing) Object.assign(existing, record);
    else sheet.tables.push(record);
    sheet.filters = { range: rangeString, criteria: [], hidden: [] };
  }, { layout: true });
  return name;
}

export function tableDialog(app) {
  const grid = app.grid;
  const sheet = grid.sheet();
  const used = usedRange(sheet);
  const initial = grid.sel.r1 === grid.sel.r2 && grid.sel.c1 === grid.sel.c2 && used ? used : rangeOf(grid);
  const range = el("input", { type: "text", value: rangeLabel(initial) });
  const hasHeader = checkRow("My table has headers", true);
  const style = selectControl(
    TABLE_STYLES.map((entry, index) => ({ value: index, label: entry.label })),
    "0",
  );
  return showDialog({
    title: "Create Table",
    size: "narrow",
    body: el("div", {}, [
      field("Where is the data for your table?", range),
      field("Table style", style),
      hasHeader,
      el("p", { class: "dialog-note", text: "Tables add filter buttons and banded rows. The table range is saved with the workbook." }),
    ]),
    buttons: [
      { label: "Cancel", kind: "ghost" },
      {
        label: "OK", kind: "primary", onClick: () => {
          const area = parseRange(range.value.trim());
          if (!area) {
            toast("Select a range such as A1:D12.");
            return false;
          }
          const name = applyTableStyle(app, Number(style.value), area, hasHeader.querySelector("input").checked);
          toast('Table "' + name + '" created.');
        },
      },
    ],
  });
}

/* -------------------------------------------------------------- subtotals */
export function subtotalDialog(app) {
  const grid = app.grid;
  const sheet = grid.sheet();
  const used = usedRange(sheet);
  const range = grid.sel.r1 === grid.sel.r2 && grid.sel.c1 === grid.sel.c2 && used ? { ...used } : rangeOf(grid);
  const hasHeader = checkRow("My data has headers", true);
  const functions = [
    { value: 9, label: "Sum" },
    { value: 1, label: "Average" },
    { value: 2, label: "Count" },
    { value: 4, label: "Max" },
    { value: 5, label: "Min" },
  ];
  const func = selectControl(functions, "9");
  const replace = checkRow("Replace current subtotals", true);
  const options = columnOptions(app, range, true);
  const checks = [];
  const numeric = new Set();
  for (let col = range.c1 + 1; col <= range.c2; col += 1) {
    for (let row = range.r1 + 1; row <= Math.min(range.r2, range.r1 + 25); row += 1) {
      const value = cellAt(sheet, row, col)?.v;
      if (typeof value === "number") { numeric.add(col); break; }
    }
  }
  const list = el("div", { style: "max-height:160px;overflow:auto;border:1px solid var(--border);border-radius:4px;padding:7px 9px" });
  for (const option of options) {
    if (Number(option.value) === range.c1) continue;
    const row = checkRow(option.label, numeric.has(Number(option.value)) || !numeric.size);
    checks.push({ col: Number(option.value), input: row.querySelector("input") });
    list.append(row);
  }
  return showDialog({
    title: "Subtotal",
    size: "narrow",
    body: el("div", {}, [
      field("At each change in", el("input", { type: "text", value: indexToCol(range.c1), readonly: "readonly" })),
      field("Use function", func),
      el("div", { class: "dialog-section-title", text: "Add subtotal to" }),
      list,
      hasHeader,
      replace,
      el("p", { class: "dialog-note", text: "Subtotal rows are inserted below each group of equal values in the first column." }),
    ]),
    buttons: [
      { label: "Cancel", kind: "ghost" },
      {
        label: "OK", kind: "primary", onClick: () => {
          const columns = checks.filter((entry) => entry.input.checked).map((entry) => entry.col);
          if (!columns.length) {
            toast("Select at least one column for the subtotals.");
            return false;
          }
          const area = normalizeRange(range.r1, range.c1, range.r2, range.c2);
          if (area.r2 <= area.r1) {
            toast("The selected range is too small.");
            return false;
          }
          const withHeader = hasHeader.querySelector("input").checked;
          const fnNumber = Number(func.value) || 9;
          const doReplace = replace.querySelector("input").checked;
          app.mutate("subtotal", () => {
            let bottom = area.r2;
            if (doReplace) {
              for (let row = bottom; row >= area.r1; row -= 1) {
                let hasSubtotal = false;
                for (let col = area.c1; col <= area.c2; col += 1) {
                  const formula = cellAt(sheet, row, col)?.f;
                  if (formula && /^=SUBTOTAL\(/i.test(String(formula))) hasSubtotal = true;
                }
                if (hasSubtotal) {
                  deleteRows(sheet, row, 1);
                  bottom -= 1;
                }
              }
            }
            const start = area.r1 + (withHeader ? 1 : 0);
            if (bottom <= start) return;
            const groups = [];
            let groupStart = start;
            let value = String(cellAt(sheet, start, area.c1)?.v ?? "");
            for (let row = start + 1; row <= bottom + 1; row += 1) {
              const next = row <= bottom ? String(cellAt(sheet, row, area.c1)?.v ?? "") : "\u0000end";
              if (next !== value || row > bottom) {
                groups.push({ start: groupStart, end: row - 1, value });
                groupStart = row;
                value = next;
              }
            }
            for (let index = groups.length - 1; index >= 0; index -= 1) {
              const group = groups[index];
              insertRows(sheet, group.end + 1, 1);
              const totalRow = group.end + 1;
              const label = cellAt(sheet, totalRow, area.c1) || ensureCell(sheet, addr(totalRow, area.c1));
              label.v = group.value ? group.value + " Total" : "Total";
              label.s = { ...(label.s || {}), b: true, fill: "#F2F2F2" };
              for (const col of columns) {
                if (col === area.c1) continue;
                const target = ensureCell(sheet, addr(totalRow, col));
                target.f = "=SUBTOTAL(" + fnNumber + "," + indexToCol(col) + group.start + ":" + indexToCol(col) + group.end + ")";
                delete target.v;
                target.s = { ...(target.s || {}), b: true, fill: "#F2F2F2" };
              }
            }
          }, { layout: true });
          toast("Subtotal rows inserted.");
        },
      },
    ],
  });
}

/* ------------------------------------------------------- text to columns */
function splitDelimited(line, delimiter, qualifier, collapse) {
  const fields = [];
  let current = "";
  let quoted = false;
  for (let index = 0; index < line.length; index += 1) {
    const char = line[index];
    if (quoted) {
      if (char === qualifier) {
        if (line[index + 1] === qualifier) {
          current += qualifier;
          index += 1;
        } else {
          quoted = false;
        }
      } else {
        current += char;
      }
    } else if (qualifier && char === qualifier) {
      quoted = true;
    } else if (line.startsWith(delimiter, index)) {
      fields.push(current);
      current = "";
      index += delimiter.length - 1;
      if (collapse) {
        while (line.startsWith(delimiter, index + 1)) index += delimiter.length;
      }
    } else {
      current += char;
    }
  }
  fields.push(current);
  return fields;
}

export function textToColumnsDialog(app) {
  const grid = app.grid;
  const sheet = grid.sheet();
  const used = usedRange(sheet);
  const area = grid.sel.r1 === grid.sel.r2 && grid.sel.c1 === grid.sel.c2 && used ? { ...used } : rangeOf(grid);
  const delimiter = selectControl([
    { value: "tab", label: "Tab" },
    { value: ",", label: "Comma" },
    { value: ";", label: "Semicolon" },
    { value: " ", label: "Space" },
    { value: "custom", label: "Other" },
  ], "tab");
  const custom = el("input", { type: "text", value: "|", maxlength: "4" });
  const qualifier = selectControl([
    { value: "none", label: "None" },
    { value: '"', label: "Double quote" },
    { value: "'", label: "Single quote" },
  ], '"');
  const collapse = checkRow("Treat consecutive delimiters as one", true);
  const destination = el("input", { type: "text", value: addr(area.r1, area.c1) });
  return showDialog({
    title: "Text to Columns",
    size: "narrow",
    body: el("div", {}, [
      field("Source range", el("input", { type: "text", value: rangeLabel(area), readonly: "readonly" })),
      field("Delimiter", delimiter),
      field("Other delimiter", custom),
      field("Text qualifier", qualifier),
      collapse,
      field("Destination", destination),
      el("p", { class: "dialog-note", text: "Each row is split at the delimiter and written to the right of the destination cell." }),
    ]),
    buttons: [
      { label: "Cancel", kind: "ghost" },
      {
        label: "OK", kind: "primary", onClick: () => {
          const target = parseAddr(destination.value.trim());
          if (!target) {
            toast("Enter a destination cell such as A1.");
            return false;
          }
          const sep = delimiter.value === "custom" ? custom.value || "," : delimiter.value;
          const quote = qualifier.value === "none" ? "" : qualifier.value;
          const collapseValue = collapse.querySelector("input").checked;
          const lines = [];
          for (let row = area.r1; row <= area.r2; row += 1) {
            const parts = [];
            for (let col = area.c1; col <= area.c2; col += 1) parts.push(String(cellAt(sheet, row, col)?.v ?? ""));
            lines.push(parts.join(sep));
          }
          app.mutate("text-to-columns", () => {
            lines.forEach((line, rowOffset) => {
              splitDelimited(line, sep, quote, collapseValue).forEach((fieldText, colOffset) => {
                const address = addr(target.row + rowOffset, target.col + colOffset);
                const cell = ensureCell(sheet, address);
                const parsed = parseUserInput(fieldText);
                delete cell.f;
                if (parsed.kind === "number" || parsed.kind === "boolean") cell.v = parsed.value;
                else if (parsed.kind === "date") cell.v = parsed.value;
                else if (parsed.kind === "formula") cell.f = parsed.formula;
                else if (parsed.kind === "empty") delete cell.v;
                else cell.v = parsed.value;
                if (cell.v === undefined && !cell.f && !cell.s) delete sheet.cells[address];
              });
            });
          }, { layout: true });
          toast(lines.length + " rows converted.");
        },
      },
    ],
  });
}

/* ------------------------------------------------------ remove duplicates */
export function removeDuplicatesDialog(app) {
  const grid = app.grid;
  const sheet = grid.sheet();
  const used = usedRange(sheet);
  const area = grid.sel.r1 === grid.sel.r2 && grid.sel.c1 === grid.sel.c2 && used ? { ...used } : rangeOf(grid);
  const hasHeader = checkRow("My data has headers", true);
  const checks = [];
  const list = el("div", { class: "filter-list" });
  for (const option of columnOptions(app, area, true)) {
    const row = checkRow(option.label, true);
    checks.push({ col: Number(option.value), input: row.querySelector("input") });
    list.append(row);
  }
  return showDialog({
    title: "Remove Duplicates",
    size: "narrow",
    body: el("div", {}, [
      el("p", { class: "dialog-note", text: "Select the columns that contain duplicates." }),
      list,
      hasHeader,
    ]),
    buttons: [
      { label: "Cancel", kind: "ghost" },
      {
        label: "OK", kind: "primary", onClick: () => {
          const columns = checks.filter((entry) => entry.input.checked).map((entry) => entry.col);
          if (!columns.length) {
            toast("Select at least one column.");
            return false;
          }
          const removed = grid.removeDuplicates(columns);
          const total = area.r2 - area.r1 + 1 - (hasHeader.querySelector("input").checked ? 1 : 0);
          if (!removed) toast("No duplicate values were found.");
          else toast(removed + " duplicates removed, " + Math.max(0, total - removed) + " unique values remain.");
        },
      },
    ],
  });
}

/* ------------------------------------------------------------ protection */
const PERMISSIONS = [
  ["allowSelectLocked", "Select locked cells"],
  ["allowSelectUnlocked", "Select unlocked cells"],
  ["allowFormatCells", "Format cells"],
  ["allowFormatColumns", "Format columns"],
  ["allowFormatRows", "Format rows"],
  ["allowInsertRows", "Insert rows"],
  ["allowInsertColumns", "Insert columns"],
  ["allowDeleteRows", "Delete rows"],
  ["allowDeleteColumns", "Delete columns"],
  ["allowSort", "Sort"],
  ["allowFilter", "Use AutoFilter"],
];

export function protectionDialog(app, scope = "sheet") {
  const sheet = app.grid.sheet();
  const current = sheet.protection || {};
  const protect = checkRow(scope === "workbook" ? "Protect workbook" : "Protect sheet", !!current.protected);
  const checks = [];
  const list = el("div", { style: "max-height:210px;overflow:auto;padding-right:4px" });
  for (const [key, label] of PERMISSIONS) {
    const enabled = current.protected ? current[key] !== false : ["allowSelectLocked", "allowSelectUnlocked"].includes(key);
    const row = checkRow(label, enabled);
    checks.push({ key, input: row.querySelector("input") });
    list.append(row);
  }
  return showDialog({
    title: scope === "workbook" ? "Protect Workbook" : "Protect Sheet",
    size: "narrow",
    body: el("div", {}, [
      protect,
      el("p", { class: "dialog-note", text: "Prevent accidental changes to locked cells. Protection is not a security boundary and does not encrypt the workbook." }),
      el("div", { class: "dialog-section-title", text: "Allow users of this " + (scope === "workbook" ? "workbook" : "sheet") + " to" }),
      list,
    ]),
    buttons: [
      { label: "Cancel", kind: "ghost" },
      {
        label: "OK", kind: "primary", onClick: () => {
          const enabled = protect.querySelector("input").checked;
          const options = {};
          for (const entry of checks) options[entry.key] = entry.input.checked;
          app.mutate("protection", () => {
            const target = enabled ? { protected: true, ...options } : null;
            if (scope === "workbook") {
              for (const item of app.model.sheets) item.protection = target ? { ...target } : null;
              app.model.protection = target ? { ...target } : null;
            } else {
              sheet.protection = target;
            }
          });
          toast(scope === "workbook"
            ? (enabled ? "Workbook protected." : "Workbook unprotected.")
            : (enabled ? "Sheet protected." : "Sheet unprotected."));
        },
      },
    ],
  });
}

/* ----------------------------------------------------------- paste special */
const PASTE_MODES = [
  ["all", "All", "Paste values, formulas and formatting."],
  ["formulas", "Formulas", "Paste formulas without formatting."],
  ["values", "Values", "Paste only the calculated values."],
  ["formats", "Formats", "Paste only the cell formatting."],
  ["transpose", "Transpose", "Paste with rows and columns swapped."],
];

export function pasteSpecialDialog(app, onPick = null) {
  let mode = "all";
  const list = el("div", {}, []);
  const inputs = [];
  for (const [value, label, description] of PASTE_MODES) {
    const input = el("input", { type: "radio", name: "paste-mode", value, checked: value === mode });
    input.addEventListener("change", () => { if (input.checked) mode = value; });
    inputs.push(input);
    const row = el("label", { class: "check-row", style: "align-items:flex-start" }, [
      input,
      el("span", {}, [
        el("strong", { text: label, style: "display:block;color:var(--text)" }),
        el("span", { class: "dialog-note", text: description }),
      ]),
    ]);
    list.append(row);
  }
  return showDialog({
    title: "Paste Special",
    size: "narrow",
    body: el("div", {}, [
      list,
      el("p", { class: "dialog-note", text: "Copy a range first, then choose how its contents should be pasted." }),
    ]),
    buttons: [
      { label: "Cancel", kind: "ghost" },
      {
        label: "OK", kind: "primary", onClick: () => {
          if (onPick) onPick(mode);
          else app.pasteSpecial(mode);
        },
      },
    ],
  });
}

/* ------------------------------------------------------------------ go to */
const GOTO_SPECIAL = [
  ["comments", "Comments"],
  ["constants", "Constants"],
  ["formulas", "Formulas"],
  ["blanks", "Blanks"],
  ["currentRegion", "Current region"],
  ["lastCell", "Last cell"],
];

export function goToDialog(app) {
  const reference = el("input", { type: "text", value: app.grid.activeAddress() });
  const special = selectControl([{ value: "", label: "None" }, ...GOTO_SPECIAL.map(([value, label]) => ({ value, label }))], "");
  const names = Object.keys(app.model.names || {});
  const nameSelect = selectControl([{ value: "", label: "—" }, ...names.map((name) => ({ value: name, label: name }))], "");
  nameSelect.addEventListener("change", () => {
    if (nameSelect.value) reference.value = String(app.model.names[nameSelect.value] || "");
  });
  return showDialog({
    title: "Go To",
    size: "narrow",
    body: el("div", {}, [
      field("Reference", reference),
      field("Go to special", special),
      field("Defined names", nameSelect),
      el("p", { class: "dialog-note", text: "Type a cell, a range such as B4:F20, or pick a defined name." }),
    ]),
    buttons: [
      { label: "Cancel", kind: "ghost" },
      {
        label: "OK", kind: "primary", onClick: () => {
          if (special.value) {
            app.goToSpecial(special.value);
            return;
          }
          const text = reference.value.trim();
          if (!text) {
            toast("Enter a cell reference.");
            return false;
          }
          if (!app.selectReference(text)) {
            toast("Enter a cell reference such as B12 or A1:C10.");
            return false;
          }
        },
      },
    ],
  });
}

/* ------------------------------------------------------------- cell style */
export function newCellStyleDialog(app) {
  const name = el("input", { type: "text", value: "My Style" });
  const style = app.grid.currentStyle();
  const parts = [];
  if (style.b) parts.push("bold");
  if (style.i) parts.push("italic");
  if (style.fill) parts.push("fill " + style.fill);
  if (style.color) parts.push("text " + style.color);
  if (style.numFmt) parts.push(style.numFmt);
  return showDialog({
    title: "New Cell Style",
    size: "narrow",
    body: el("div", {}, [
      field("Style name", name),
      el("p", { class: "dialog-note", text: "The style is based on the active cell: " + (parts.join(", ") || "no formatting") + "." }),
    ]),
    buttons: [
      { label: "Cancel", kind: "ghost" },
      {
        label: "OK", kind: "primary", onClick: () => {
          const label = name.value.trim();
          if (!label) {
            toast("Enter a style name.");
            return false;
          }
          app.addCustomStyle(label, style);
          toast('Style "' + label + '" saved.');
        },
      },
    ],
  });
}

/* ------------------------------------------------------------ sheet tabs */
const SHEET_COLORS = [...new Set([...THEME_COLORS, ...STANDARD_COLORS])];

export function sheetColorDialog(app, index) {
  let close = null;
  const palette = el("div", { class: "palette-grid", style: "grid-template-columns:repeat(8,22px)" });
  const none = el("button", { class: "palette-none", type: "button", text: "No Color" });
  none.addEventListener("click", () => {
    app.setSheetColor(index, "");
    close?.();
  });
  for (const color of SHEET_COLORS) {
    const swatch = el("button", { class: "swatch", type: "button", title: color, style: "background:" + color });
    swatch.addEventListener("click", () => {
      app.setSheetColor(index, color);
      close?.();
    });
    palette.append(swatch);
  }
  showDialog({
    title: "Sheet Tab Color",
    size: "narrow",
    body: el("div", {}, [palette, el("div", { style: "margin-top:8px" }, [none])]),
    buttons: [{ label: "Cancel", kind: "ghost" }],
    onOpen: ({ close: closeDialog }) => { close = closeDialog; },
  });
}

export function insertSheetDialog(app) {
  const current = app.model.activeSheet;
  const name = el("input", { type: "text", value: "" , placeholder: "Sheet" });
  const position = selectControl([
    ...app.model.sheets.map((sheet, index) => ({ value: index, label: "Before " + sheet.name })),
    { value: app.model.sheets.length, label: "At the end" },
  ], String(current));
  return showDialog({
    title: "Insert Sheet",
    size: "narrow",
    body: el("div", {}, [
      field("Sheet name", name),
      field("Insert before", position),
      el("p", { class: "dialog-note", text: "Leave the name empty to use the next available sheet number." }),
    ]),
    buttons: [
      { label: "Cancel", kind: "ghost" },
      {
        label: "OK", kind: "primary", onClick: () => {
          app.insertSheet(name.value.trim(), Number(position.value));
          toast("Sheet added.");
        },
      },
    ],
  });
}

export async function renameSheetDialog(app, index = null) {
  const sheetIndex = index === null ? app.model.activeSheet : index;
  const sheet = app.model.sheets[sheetIndex];
  if (!sheet) return;
  const value = await promptDialog({ title: "Rename Sheet", label: "Sheet name", value: sheet.name, note: "Sheet names cannot contain : \\ / ? * [ ]" });
  if (value === null) return;
  const clean = String(value).trim();
  if (!clean) {
    toast("Sheet names cannot be empty.");
    return;
  }
  if (sheetByName(app.model, clean) && clean.toLowerCase() !== sheet.name.toLowerCase()) {
    toast("A sheet with that name already exists.");
    return;
  }
  const result = app.renameSheet(sheetIndex, clean);
  if (!result || result.ok === false) toast(result?.error || "The sheet could not be renamed.");
  else toast("Sheet renamed.");
}

export function moveSheetDialog(app, index = null) {
  const from = index === null ? app.model.activeSheet : index;
  const sheet = app.model.sheets[from];
  if (!sheet) return;
  const options = app.model.sheets.map((item, position) => ({ value: position, label: "Before " + item.name }));
  options.push({ value: app.model.sheets.length, label: "At the end" });
  const position = selectControl(options, String(Math.min(app.model.sheets.length, from + 1)));
  const copy = checkRow("Create a copy", false);
  return showDialog({
    title: "Move or Copy",
    size: "narrow",
    body: el("div", {}, [field("Move selected sheet", position), copy]),
    buttons: [
      { label: "Cancel", kind: "ghost" },
      {
        label: "OK", kind: "primary", onClick: () => {
          const target = Math.max(0, Math.min(app.model.sheets.length, Number(position.value)));
          app.moveSheet(from, target, copy.querySelector("input").checked);
          toast(copy.querySelector("input").checked ? "Sheet copied." : "Sheet moved.");
        },
      },
    ],
  });
}

export async function deleteSheetDialog(app, index = null) {
  const sheetIndex = index === null ? app.model.activeSheet : index;
  const sheet = app.model.sheets[sheetIndex];
  if (!sheet) return;
  if (app.model.sheets.length <= 1) {
    toast("Cannot delete the last sheet.");
    return;
  }
  const ok = await confirmDialog("Delete Sheet", "Are you sure you want to delete this sheet? This cannot be undone.", "Delete", "danger");
  if (!ok) return;
  app.removeSheet(sheetIndex);
  toast("Sheet deleted.");
}

export function unhideSheetDialog(app) {
  const hidden = app.model.sheets.map((sheet, index) => ({ sheet, index })).filter((entry) => entry.sheet.hidden);
  if (!hidden.length) {
    toast("No hidden sheets.");
    return;
  }
  const select = selectControl(hidden.map((entry) => ({ value: entry.index, label: entry.sheet.name })), String(hidden[0].index));
  return showDialog({
    title: "Unhide Sheet",
    size: "narrow",
    body: field("Hidden sheets", select),
    buttons: [
      { label: "Cancel", kind: "ghost" },
      {
        label: "OK", kind: "primary", onClick: () => {
          app.setSheetHidden(Number(select.value), false);
          toast("Sheet shown.");
        },
      },
    ],
  });
}

/* ------------------------------------------------------------- info dialogs */
export function aboutDialog(app) {
  const version = app?.info?.version || "1.0.0";
  return showDialog({
    title: "About mTails",
    size: "narrow",
    body: el("div", { style: "display:flex;gap:16px;align-items:flex-start" }, [
      el("img", {
        src: "../assets/mtails.png",
        alt: "mTails",
        style: "width:56px;height:56px;flex:0 0 auto;object-fit:contain",
      }),
      el("div", {}, [
        el("h3", { text: "mTails", style: "margin:0 0 2px;font-size:18px" }),
        el("div", { text: t("Version {version}", { version }), style: "color:var(--muted);font-size:12px;margin-bottom:10px" }),
        el("p", { class: "dialog-note", text: "A local-first spreadsheet with an Excel-compatible core, built with Qt, Chromium and open web standards." }),
        el("p", { class: "dialog-note", text: "mTails is an independent product. It is not affiliated with, endorsed by, or sponsored by Microsoft Corporation." }),
        el("p", { class: "dialog-note", text: "Excel is a trademark of Microsoft Corporation." }),
      ]),
    ]),
    buttons: [{ label: "Close", kind: "primary" }],
  });
}

const SHORTCUTS = [
  ["Files", [
    ["Ctrl+N", "New workbook"], ["Ctrl+O", "Open workbook"], ["Ctrl+S", "Save workbook"],
    ["Ctrl+Shift+S", "Save as"], ["Ctrl+P", "Print"],
  ]],
  ["Editing", [
    ["Ctrl+Z", "Undo"], ["Ctrl+Y", "Redo"], ["Ctrl+C", "Copy"], ["Ctrl+X", "Cut"], ["Ctrl+V", "Paste"],
    ["Ctrl+B", "Bold"], ["Ctrl+I", "Italic"], ["Ctrl+U", "Underline"], ["Ctrl+1", "Format Cells"],
    ["F2", "Edit cell"], ["Escape", "Cancel entry"], ["Delete", "Clear contents"], ["Ctrl+F", "Find"], ["Ctrl+H", "Replace"],
  ]],
  ["Navigation", [
    ["Ctrl+Home", "Go to A1"], ["Ctrl+End", "Go to last cell"],
    ["Ctrl+←/→/↑/↓", "Move to edge of data"],
    ["Ctrl+Page Down", "Next sheet"], ["Ctrl+Page Up", "Previous sheet"],
    ["Ctrl+Space", "Select column"], ["Shift+Space", "Select row"], ["Ctrl+A", "Select all"],
    ["Ctrl+G", "Go to"], ["Ctrl+Shift+L", "Filter"],
  ]],
  ["Formatting", [
    ["Ctrl+Shift+1", "Number format"], ["Ctrl+Shift+4", "Currency format"],
    ["Ctrl+Shift+5", "Percentage format"], ["Ctrl+Shift+#", "Date format"],
    ["Alt+Enter", "New line in cell"], ["Ctrl+`", "Show formulas"], ["F9", "Recalculate"],
  ]],
];

export function shortcutsDialog() {
  const columns = el("div", { style: "display:grid;grid-template-columns:1fr 1fr;gap:18px 26px" });
  for (const [title, entries] of SHORTCUTS) {
    const block = el("div", {}, [el("div", { class: "dialog-section-title", text: title })]);
    for (const [keys, action] of entries) {
      block.append(el("div", { style: "display:flex;justify-content:space-between;gap:10px;padding:3px 0;font-size:12px" }, [
        el("span", { text: t(action), style: "color:var(--text-soft)" }),
        el("kbd", { text: keys, style: "font-family:inherit;font-size:11px;border:1px solid var(--border);border-bottom-width:2px;border-radius:4px;padding:1px 5px;background:var(--surface-2);white-space:nowrap" }),
      ]));
    }
    columns.append(block);
  }
  return showDialog({
    title: "Keyboard shortcuts",
    size: "wide",
    body: columns,
    buttons: [{ label: "Close", kind: "primary" }],
  });
}

export function whatsNewDialog() {
  const items = [
    "A virtualized grid that scrolls smoothly through 100,000 rows and 1,000 columns.",
    "A full formula engine with dependency tracking, circular-reference detection and more than 150 functions.",
    "XLSX, XLSM, CSV and TSV import and export with number formats, merges, comments and conditional formatting.",
    "Charts, sparklines, pictures, text boxes, hyperlinks and cell comments on the sheet canvas.",
    "Eight ribbon tabs, a file backstage, page setup, print preview and print-ready PDF export.",
    "Autosave and crash recovery, recent files, workbook statistics and a Polish interface.",
    "Undo history, fill series, format painter, tables, filters, sorting and data validation.",
  ];
  return showDialog({
    title: "What's new in mTails 1.0",
    size: "narrow",
    body: el("div", {}, items.map((text) => el("div", { style: "display:flex;gap:8px;margin:0 0 9px;font-size:12.5px;line-height:1.5" }, [
      el("span", { text: "•", style: "color:var(--accent)" }),
      el("span", { text }),
    ]))),
    buttons: [{ label: "Close", kind: "primary" }],
  });
}

export function statsDialog(app) {
  const stats = workbookStats(app.model);
  const sheets = app.model.sheets;
  const charts = sheets.reduce((sum, sheet) => sum + (sheet.charts?.length || 0), 0);
  const comments = sheets.reduce((sum, sheet) => sum + (sheet.comments?.length || 0), 0);
  const merges = sheets.reduce((sum, sheet) => sum + (sheet.merges?.length || 0), 0);
  const rows = [
    ["Sheets: {count}", stats.sheets],
    ["Cells with data: {count}", stats.cells],
    ["Formulas: {count}", stats.formulas],
    ["Charts: {count}", charts],
    ["Comments: {count}", comments],
    ["Named ranges: {count}", stats.names],
    ["Merged ranges: {count}", merges],
  ];
  const table = el("table", { class: "name-table" }, [
    el("thead", {}, [el("tr", {}, [el("th", { text: "Sheet" }), el("th", { text: "Used range" }), el("th", { text: "Cells" }), el("th", { text: "Formulas" })])]),
  ]);
  const body = el("tbody", {});
  for (const sheet of sheets) {
    const range = usedRange(sheet);
    const formulas = Object.values(sheet.cells || {}).filter((cell) => cell.f).length;
    body.append(el("tr", {}, [
      el("td", { text: sheet.name }),
      el("td", { text: range ? rangeText(range) : "—", i18n: false }),
      el("td", { text: String(Object.keys(sheet.cells || {}).length) }),
      el("td", { text: String(formulas) }),
    ]));
  }
  table.append(body);
  return showDialog({
    title: "Workbook statistics",
    size: "wide",
    body: el("div", {}, [
      el("div", { style: "display:grid;grid-template-columns:1fr 1fr;gap:7px 22px;margin-bottom:16px" }, rows.map(([label, count]) =>
        el("div", { style: "display:flex;justify-content:space-between;font-size:12.5px;border-bottom:1px solid var(--border);padding:4px 0" }, [
          el("span", { text: t(label, { count }) }),
          el("strong", { text: String(count) }),
        ]))),
      table,
    ]),
    buttons: [{ label: "Close", kind: "primary" }],
  });
}

/* --------------------------------------------------------- accessibility */
export function accessibilityDialog(app) {
  const list = el("div", {});
  const summary = el("p", { class: "dialog-note", style: "margin:0 0 10px" });
  const render = () => {
    list.replaceChildren();
    const checks = [];
    for (const sheet of app.model.sheets) {
      const merges = sheet.merges?.length || 0;
      if (merges) checks.push(["warning", "Merged cells", "Merged cells can make navigation harder for screen readers.", merges]);
      if (/^Sheet\d+$/.test(sheet.name)) checks.push(["info", "Default sheet name", "Rename default sheet names to describe their content.", 1]);
      if ((sheet.charts || []).some((chart) => !chart.title)) checks.push(["warning", "Untitled chart", "Charts should have a title that describes the data.", (sheet.charts || []).filter((chart) => !chart.title).length]);
      if (sheet.hidden) checks.push(["info", "Hidden sheet", "Hidden sheets are invisible to keyboard navigation.", 1]);
    }
    summary.textContent = checks.length ? checks.length + " items need attention." : "No accessibility issues were found.";
    if (!checks.length) {
      list.append(el("div", { style: "padding:18px;text-align:center;color:var(--muted);font-size:12.5px" }, [
        el("div", { html: icon("success"), style: "width:28px;height:28px;margin:0 auto 8px;color:#217346" }),
        el("div", { text: "No accessibility issues were found." }),
      ]));
      return;
    }
    for (const [level, title, detail, count] of checks) {
      list.append(el("div", { class: "list-item", style: "align-items:flex-start" }, [
        el("span", {
          html: icon(level === "warning" ? "warning" : "info"),
          style: "width:18px;height:18px;flex:0 0 auto;color:" + (level === "warning" ? "#C55A11" : "#2F5597"),
        }),
        el("div", { style: "flex:1" }, [
          el("strong", { text: title, style: "display:block;font-size:12.5px" }),
          el("span", { class: "li-sub", text: detail }),
        ]),
        count > 1 ? el("span", { class: "pi-tab", text: String(count) }) : null,
      ]));
    }
  };
  render();
  return showDialog({
    title: "Accessibility Checker",
    size: "narrow",
    body: el("div", {}, [summary, list]),
    buttons: [
      { label: "Check Again", kind: "ghost", left: true, onClick: () => { render(); return false; } },
      { label: "Close", kind: "primary" },
    ],
  });
}

/* ------------------------------------------------------------- spellcheck */
export function spellcheckDialog(app) {
  const enabled = checkRow("Check spelling as you type", app.settings.spellcheck !== false);
  const language = selectControl([
    { value: "en-US", label: "English (US)" },
    { value: "pl-PL", label: "Polski (PL)" },
  ], app.settings.proofingLanguage || "en-US");
  return showDialog({
    title: "Spelling",
    size: "narrow",
    body: el("div", {}, [
      enabled,
      field("Proofing language", language),
      el("p", { class: "dialog-note", text: "The editor underlines words that the system dictionary does not recognise. Right-click a flagged word to see suggestions." }),
    ]),
    buttons: [
      { label: "Cancel", kind: "ghost" },
      {
        label: "OK", kind: "primary", onClick: () => {
          app.setSpelling(enabled.querySelector("input").checked, language.value);
          toast("Spelling preferences saved.");
        },
      },
    ],
  });
}

/* --------------------------------------------------------------------- help */
export function helpDialog(app) {
  const topics = [
    ["Formulas", "Start with = and use SUM, AVERAGE, IF, VLOOKUP, INDEX, MATCH and more. The function list documents every supported name.", () => functionDialog(app)],
    ["Charts", "Select a range and insert a column, bar, line, pie, area, scatter or doughnut chart. Double-click a chart to edit it.", () => chartDialog(app)],
    ["Keyboard shortcuts", "Every ribbon command has a keyboard route. Open the shortcut reference for the complete list.", () => shortcutsDialog()],
    ["About mTails", "Version information and the project disclaimer.", () => aboutDialog(app)],
  ];
  const list = el("div", {});
  for (const [title, text, run] of topics) {
    const item = el("button", { class: "list-item", type: "button", style: "width:100%;border:0;background:transparent;text-align:left;cursor:default" }, [
      el("div", { style: "flex:1" }, [
        el("strong", { text: title, style: "display:block;font-size:12.5px" }),
        el("span", { class: "li-sub", text }),
      ]),
      el("span", { html: icon("chevronRight"), style: "width:14px;height:14px;color:var(--muted)" }),
    ]);
    item.addEventListener("click", run);
    list.append(item);
  }
  return showDialog({
    title: "Help",
    size: "narrow",
    body: el("div", {}, [
      el("p", { class: "dialog-note", text: "Select cells and type to replace their content. Press F2 to edit. Drag the fill handle to continue a series. Formulas start with an equals sign, for example =SUM(B2:B10). Press Ctrl+S to save the workbook." }),
      el("div", { class: "dialog-section-title", text: "Topics" }),
      list,
    ]),
    buttons: [{ label: "Close", kind: "primary" }],
  });
}

/* -------------------------------------------------------------------- audit */
export function auditDialog(app, mode = "precedents") {
  if (mode === "remove") {
    app.clearAudit();
    toast("Arrows cleared.");
    return;
  }
  const grid = app.grid;
  const sheet = grid.sheet();
  const address = grid.activeAddress();
  const cell = cellAt(sheet, parseAddr(address)?.row || 1, parseAddr(address)?.col || 1);
  // Dependents can be traced from any cell; precedents only make sense for a
  // formula, since that is what defines them.
  if (mode !== "dependents" && (!cell || !cell.f)) {
    toast("The active cell does not contain a formula.");
    return;
  }
  const rows = [];
  if (mode === "dependents") {
    for (const entry of app.engine.dependentsOf(sheet, address)) rows.push({ sheet: entry.sheet || sheet.name, label: entry.address });
  } else {
    const info = app.engine.precedentsOf(sheet, address);
    for (const entry of info.cells) rows.push({ sheet: entry.sheet || sheet.name, label: entry.address });
    for (const entry of info.ranges) rows.push({ sheet: entry.sheet || sheet.name, label: entry.range, range: true });
  }
  app.traceArrows(rows, mode);
  const note = cell?.f
    ? address + " = " + cell.f
    : address + (cell && cell.v !== undefined && cell.v !== "" ? " = " + cell.v : "");
  const list = el("div", { class: "filter-list" });
  if (!rows.length) {
    list.append(el("div", { class: "palette-empty", text: mode === "dependents" ? "No dependents." : "No precedents." }));
  }
  for (const row of rows) {
    const item = el("button", { class: "list-item", type: "button", style: "width:100%;border:0;background:transparent;text-align:left" }, [
      el("span", { html: icon(row.range ? "crosshair" : "nameBox"), style: "width:15px;height:15px;color:var(--accent)" }),
      el("span", { text: row.sheet + "!" + row.label, i18n: false, style: "flex:1" }),
      el("span", { html: icon("chevronRight"), style: "width:13px;height:13px;color:var(--muted)" }),
    ]);
    item.addEventListener("click", () => {
      app.selectReference(row.sheet === sheet.name ? row.label : row.sheet + "!" + row.label);
    });
    list.append(item);
  }
  return showDialog({
    title: t(mode === "dependents" ? "Trace Dependents" : "Trace Precedents"),
    size: "narrow",
    body: el("div", {}, [
      el("p", { class: "dialog-note", text: note, i18n: false }),
      list,
    ]),
    buttons: [
      { label: "Remove Arrows", kind: "ghost", left: true, onClick: () => { app.clearAudit(); toast("Arrows cleared."); } },
      { label: "Close", kind: "primary" },
    ],
  });
}

/* --------------------------------------------------------------------- zoom */
export function zoomDialog(app) {
  const grid = app.grid;
  const input = numberInput(grid.zoom || 100, { min: "40", max: "200", step: "5" });
  const fit = () => {
    const rect = grid.sel;
    const width = grid.colPrefix[rect.c2] - grid.colPrefix[rect.c1 - 1];
    const height = grid.rowPrefix[rect.r2] - grid.rowPrefix[rect.r1 - 1];
    if (!(width > 0) || !(height > 0)) {
      toast("Select a range first.");
      return;
    }
    const value = Math.round(Math.min((grid.viewportWidth - 12) / width, (grid.viewportHeight - 12) / height) * 100);
    input.value = String(Math.max(40, Math.min(200, value)));
  };
  return showDialog({
    title: "Zoom",
    size: "narrow",
    body: el("div", {}, [
      field("Zoom level", el("div", { class: "input-row" }, [input, el("span", { class: "unit", text: "%" })])),
      el("p", { class: "dialog-note", text: "Use the buttons on the status bar to zoom quickly. Ctrl+wheel zooms the sheet." }),
    ]),
    buttons: [
      { label: "100%", kind: "ghost", left: true, onClick: () => { input.value = "100"; return false; } },
      { label: "Fit selection", kind: "ghost", left: true, onClick: () => { fit(); return false; } },
      { label: "Cancel", kind: "ghost" },
      {
        label: "OK", kind: "primary", onClick: () => {
          const value = Math.max(40, Math.min(200, Number(input.value) || 100));
          grid.setZoom(value);
          toast(t("Zoom set to {value}%.", { value }));
        },
      },
    ],
  });
}

/* --------------------------------------------------------------- fill series */
export function fillSeriesDialog(app) {
  const grid = app.grid;
  const range = rangeOf(grid);
  const across = range.r2 === range.r1;
  const orientation = selectControl([{ value: "down", label: "Columns" }, { value: "right", label: "Rows" }], across ? "right" : "down");
  const type = selectControl([
    { value: "linear", label: "Linear" },
    { value: "growth", label: "Growth" },
    { value: "date", label: "Date" },
    { value: "autofill", label: "AutoFill" },
  ], "linear");
  const step = numberInput(1, { step: "any" });
  const stop = numberInput("", { step: "any" });
  const trend = checkRow("Trend", false);
  return showDialog({
    title: "Series",
    size: "narrow",
    body: el("div", {}, [
      field("Series in", orientation),
      field("Type", type),
      el("div", { class: "dialog-grid" }, [field("Step value", step), field("Stop value", stop)]),
      trend,
      el("p", { class: "dialog-note", text: "The series starts from the first cell of the selection and fills the remaining cells." }),
    ]),
    buttons: [
      { label: "Cancel", kind: "ghost" },
      {
        label: "OK", kind: "primary", onClick: () => {
          const direction = orientation.value;
          const count = direction === "down" ? range.r2 - range.r1 + 1 : range.c2 - range.c1 + 1;
          if (count < 2) {
            toast("Select a range with at least two cells.");
            return false;
          }
          const sheet = grid.sheet();
          const first = cellAt(sheet, range.r1, range.c1);
          const startValue = Number(first?.v);
          const kind = type.value;
          const stepValue = Number(step.value);
          const stopValue = stop.value === "" ? null : Number(stop.value);
          if (kind !== "autofill" && !Number.isFinite(startValue)) {
            toast("The first cell must contain a number.");
            return false;
          }
          const useTrend = trend.querySelector("input").checked;
          let slope = Number.isFinite(stepValue) && stepValue !== 0 ? stepValue : 1;
          if (useTrend && direction === "down") {
            const values = [];
            for (let row = range.r1; row <= range.r2; row += 1) {
              const value = Number(cellAt(sheet, row, range.c1)?.v);
              if (Number.isFinite(value)) values.push(value);
            }
            if (values.length >= 2) slope = (values[values.length - 1] - values[0]) / (values.length - 1);
          }
          app.mutate("fill-series", () => {
            for (let offset = 1; offset < count; offset += 1) {
              let value;
              if (kind === "autofill") {
                const source = cellAt(sheet, range.r1 + (offset % count), range.c1);
                value = source?.v;
              } else if (kind === "growth") {
                value = startValue * Math.pow(slope, offset);
              } else if (kind === "date") {
                value = startValue + slope * offset;
              } else {
                value = startValue + slope * offset;
              }
              if (stopValue !== null && Number.isFinite(value)) {
                if (slope >= 0 && value > stopValue) break;
                if (slope < 0 && value < stopValue) break;
              }
              const target = direction === "down" ? { row: range.r1 + offset, col: range.c1 } : { row: range.r1, col: range.c1 + offset };
              const cell = ensureCell(sheet, addr(target.row, target.col));
              delete cell.f;
              cell.v = value;
            }
          }, { layout: true });
          toast("Fill applied.");
        },
      },
    ],
  });
}
