/* Print and PDF layout.
 *
 * The grid is virtualised, so Chromium cannot paginate it directly.  This
 * module turns the workbook into a stack of fixed-size paper pages (points),
 * complete with margins, headers, footers, repeatable headings and page
 * numbers.  file.js makes the page visible for a moment and asks Qt to print
 * or export it; the backstage and Page Layout view reuse the same renderer.
 */

import { el, clamp } from "./ui.js";
import { formatCell } from "./format.js";
import {
  cellAt,
  colDisplayWidth,
  rowDisplayHeight,
  usedRange,
} from "./model.js";
import { indexToCol, parseRange } from "./units.js";

/* ISO 216 and US paper sizes in PostScript points. */
export const PAPER_SIZES = {
  a3: { w: 841.89, h: 1190.55 },
  a4: { w: 595.28, h: 841.89 },
  a5: { w: 419.53, h: 595.28 },
  letter: { w: 612, h: 792 },
  legal: { w: 612, h: 1008 },
  tabloid: { w: 792, h: 1224 },
};

const PX_TO_PT = 0.75;

function numeric(value, fallback) {
  const number = Number(value);
  return Number.isFinite(number) ? number : fallback;
}

function paperFor(page) {
  const key = String(page?.paper || page?.size || "A4").toLowerCase();
  return PAPER_SIZES[key] ? key : "a4";
}

export function pageMetrics(page = {}) {
  const key = paperFor(page);
  const paper = PAPER_SIZES[key];
  const landscape = String(page.orientation || "portrait").toLowerCase() === "landscape";
  const width = landscape ? paper.h : paper.w;
  const height = landscape ? paper.w : paper.h;
  const margins = page.margins || {};
  const metrics = {
    paper: key,
    orientation: landscape ? "landscape" : "portrait",
    w: width,
    h: height,
    m: {
      t: Math.max(0, numeric(margins.t, 72)),
      r: Math.max(0, numeric(margins.r, 72)),
      b: Math.max(0, numeric(margins.b, 72)),
      l: Math.max(0, numeric(margins.l, 72)),
      header: Math.max(0, numeric(margins.header, 30)),
      footer: Math.max(0, numeric(margins.footer, 30)),
    },
    scale: clamp(numeric(page.scale, 100), 10, 400) / 100,
  };
  metrics.contentW = Math.max(24, width - metrics.m.l - metrics.m.r);
  metrics.contentH = Math.max(24, height - metrics.m.t - metrics.m.b);
  return metrics;
}

function printRange(sheet) {
  const area = parseRange(sheet?.page?.area || "");
  if (area) return area;
  const used = usedRange(sheet || {});
  if (!used) return { r1: 1, c1: 1, r2: 1, c2: 1 };
  return used;
}

function collectColumns(sheet, range, metrics) {
  const columns = [];
  const limit = 1000;
  for (let col = range.c1; col <= range.c2 && columns.length < limit; col += 1) {
    const width = colDisplayWidth(sheet, col);
    if (!width) continue;
    columns.push({ col, width: width * PX_TO_PT * metrics.scale });
  }
  return columns;
}

function collectRows(sheet, range, metrics) {
  const rows = [];
  const limit = 50000;
  for (let row = range.r1; row <= range.r2 && rows.length < limit; row += 1) {
    const height = rowDisplayHeight(sheet, row);
    if (!height) continue;
    rows.push({ row, height: height * PX_TO_PT * metrics.scale });
  }
  return rows;
}

function bandsOf(columns, available) {
  const bands = [];
  let current = [];
  let width = 0;
  for (const column of columns) {
    if (current.length && width + column.width > available) {
      bands.push(current);
      current = [];
      width = 0;
    }
    current.push(column);
    width += column.width;
  }
  if (current.length) bands.push(current);
  return bands.length ? bands : [[]];
}

function chunksOf(rows, available) {
  const chunks = [];
  let current = [];
  let height = 0;
  for (const row of rows) {
    if (current.length && height + row.height > available) {
      chunks.push(current);
      current = [];
      height = 0;
    }
    current.push(row);
    height += row.height;
  }
  if (current.length) chunks.push(current);
  return chunks.length ? chunks : [[]];
}

function normalizeHeaderFooter(text) {
  const sections = { left: "", center: "", right: "" };
  let target = "center";
  let buffer = "";
  const source = String(text ?? "");
  for (let index = 0; index < source.length; index += 1) {
    const char = source[index];
    if (char !== "&") {
      buffer += char;
      continue;
    }
    const next = source[index + 1];
    if (next === "&") {
      buffer += "&";
      index += 1;
      continue;
    }
    if (next === "L" || next === "C" || next === "R") {
      sections[target] += buffer;
      buffer = "";
      target = next === "L" ? "left" : next === "C" ? "center" : "right";
      index += 1;
      continue;
    }
    if (next === '"') {
      const end = source.indexOf('"', index + 2);
      index = end === -1 ? source.length : end;
      continue;
    }
    buffer += "&" + (next || "");
    index += 1;
  }
  sections[target] += buffer;
  return sections;
}

export function formatHeaderFooter(text, context = {}) {
  const sections = normalizeHeaderFooter(text);
  const date = context.date instanceof Date ? context.date : new Date();
  const pad = (value) => String(value).padStart(2, "0");
  const replacements = (value) => String(value || "")
    .replace(/&P/g, String(context.page ?? 1))
    .replace(/&N/g, String(context.total ?? 1))
    .replace(/&D/g, `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`)
    .replace(/&T/g, `${pad(date.getHours())}:${pad(date.getMinutes())}`)
    .replace(/&F/g, String(context.file || ""))
    .replace(/&A/g, String(context.sheet || ""));
  return {
    left: replacements(sections.left),
    center: replacements(sections.center),
    right: replacements(sections.right),
  };
}

function mergeMap(sheet) {
  const map = new Map();
  for (const entry of sheet.merges || []) {
    const range = typeof entry === "string" ? parseRange(entry) : entry;
    if (!range) continue;
    map.set(range.r1 + ":" + range.c1, range);
  }
  return map;
}

function borderFor(style, side, gridlines) {
  const spec = style?.border?.[side];
  if (spec && spec.style && spec.style !== "none") {
    const widths = { hair: 0.5, thin: 0.75, medium: 1.5, thick: 2.25, dashed: 0.75, dotted: 0.75, double: 1.5 };
    const type = spec.style === "dashed" ? "dashed" : spec.style === "dotted" ? "dotted" : spec.style === "double" ? "double" : "solid";
    return `${widths[spec.style] || 0.75}pt ${type} ${spec.color || "#000000"}`;
  }
  return gridlines ? "0.5pt solid #c9c9c9" : "none";
}

function makeCell(sheet, row, col, column, rowHeight, merges, metrics, gridlines) {
  const cell = cellAt(sheet, row, col);
  const style = cell?.s || {};
  const td = el("td", { i18n: false });
  td.style.width = column.width + "pt";
  td.style.height = rowHeight + "pt";
  td.style.padding = "1pt 2.5pt";
  td.style.boxSizing = "border-box";
  td.style.overflow = "hidden";
  td.style.verticalAlign = style.valign === "top" ? "top" : style.valign === "bottom" ? "bottom" : "middle";
  td.style.borderTop = borderFor(style, "t", gridlines);
  td.style.borderRight = borderFor(style, "r", gridlines);
  td.style.borderBottom = borderFor(style, "b", gridlines);
  td.style.borderLeft = borderFor(style, "l", gridlines);
  if (style.fill) td.style.background = style.fill;
  if (style.color) td.style.color = style.color;
  if (style.b) td.style.fontWeight = "700";
  if (style.i) td.style.fontStyle = "italic";
  if (style.u || style.strike) {
    td.style.textDecoration = [style.u ? "underline" : "", style.strike ? "line-through" : ""].filter(Boolean).join(" ");
  }
  if (style.rotate) td.style.transform = `rotate(${-Number(style.rotate)}deg)`;
  const points = Math.max(5, (Number(style.size) || 11) * metrics.scale);
  td.style.fontSize = points + "pt";
  if (style.font) td.style.fontFamily = `"${String(style.font).replace(/"/g, "")}", sans-serif`;
  td.style.whiteSpace = style.wrap ? "pre-wrap" : "nowrap";
  const align = style.align || "";
  if (align === "center") td.style.textAlign = "center";
  else if (align === "right") td.style.textAlign = "right";
  else if (align === "justify") td.style.textAlign = "justify";
  const formatted = formatCell(cell || {});
  let text = formatted.text;
  if (!text && cell?.f) text = "";
  if (!text && cell?.h) text = cell.h;
  if (!style.align && formatted.kind === "number") td.style.textAlign = "right";
  if (formatted.kind === "error") td.style.color = "#c62828";
  td.textContent = text;
  return td;
}

function mergeSpans(sheet, band, chunk) {
  const merges = mergeMap(sheet);
  const starts = new Map();
  const covered = new Set();
  for (const [key, merge] of merges) {
    const firstRow = chunk.findIndex((entry) => entry.row === merge.r1);
    const lastRow = chunk.findIndex((entry) => entry.row === merge.r2);
    if (firstRow === -1) continue;
    const firstCol = band.findIndex((entry) => entry.col === merge.c1);
    if (firstCol === -1) continue;
    const endRow = lastRow === -1 ? chunk.length - 1 : lastRow;
    let endCol = band.findIndex((entry) => entry.col === merge.c2);
    if (endCol === -1) endCol = band.length - 1;
    const rowSpan = Math.max(1, endRow - firstRow + 1);
    const colSpan = Math.max(1, endCol - firstCol + 1);
    starts.set(key, { rowSpan, colSpan });
    for (let r = firstRow; r <= endRow; r += 1) {
      for (let c = firstCol; c <= endCol; c += 1) {
        if (r === firstRow && c === firstCol) continue;
        covered.add(chunk[r].row + ":" + band[c].col);
      }
    }
  }
  return { starts, covered };
}

function renderPage(sheet, chunk, band, options) {
  const { metrics, pageNumber, totalPages, model, merges, gridlines, headings } = options;
  const page = el("div", { class: "print-page", i18n: false });
  page.style.width = metrics.w + "pt";
  page.style.height = metrics.h + "pt";
  page.style.position = "relative";

  const headingHeight = headings ? 14 * metrics.scale : 0;
  const headingWidth = headings ? 26 * metrics.scale : 0;
  const contentTop = metrics.m.t;
  const contentLeft = metrics.m.l;

  const header = formatHeaderFooter(sheet.page?.header || "", {
    page: pageNumber,
    total: totalPages,
    sheet: sheet.name,
    file: model.title || "",
  });
  const footer = formatHeaderFooter(sheet.page?.footer || "", {
    page: pageNumber,
    total: totalPages,
    sheet: sheet.name,
    file: model.title || "",
  });

  const addHeaderFooter = (text, kind, top) => {
    if (!text.left && !text.center && !text.right) return;
    const box = el("div", { class: "print-chrome print-" + kind, i18n: false });
    box.style.cssText = `position:absolute;left:${contentLeft}pt;width:${metrics.contentW}pt;top:${top}pt;`
      + `height:${Math.max(12, kind === "header" ? metrics.m.header : metrics.m.footer)}pt;display:flex;`
      + "align-items:" + (kind === "header" ? "flex-start" : "flex-end") + ";font-size:8pt;color:#555;";
    box.append(el("span", { text: text.left, style: "flex:1", i18n: false }));
    box.append(el("span", { text: text.center, style: "flex:1;text-align:center", i18n: false }));
    box.append(el("span", { text: text.right, style: "flex:1;text-align:right", i18n: false }));
    page.append(box);
  };
  addHeaderFooter(header, "header", Math.max(2, metrics.m.t - Math.max(12, metrics.m.header) - 2));
  addHeaderFooter(footer, "footer", metrics.h - metrics.m.b + 2);

  const body = el("div", { class: "print-body", i18n: false });
  body.style.cssText = `position:absolute;left:${contentLeft}pt;top:${contentTop}pt;`
    + `width:${metrics.contentW}pt;height:${metrics.contentH}pt;overflow:hidden;`;
  page.append(body);

  const { starts, covered } = mergeSpans(sheet, band, chunk);

  const table = el("table", { class: "print-table", i18n: false });
  table.style.cssText = "border-collapse:collapse;table-layout:fixed;width:" + band.reduce((sum, col) => sum + col.width, 0) + "pt;";
  const colgroup = el("colgroup", { i18n: false });
  if (headings) {
    const cornerCol = el("col", { i18n: false });
    cornerCol.style.width = headingWidth + "pt";
    colgroup.append(cornerCol);
  }
  for (const column of band) {
    const colNode = el("col", { i18n: false });
    colNode.style.width = column.width + "pt";
    colgroup.append(colNode);
  }
  table.append(colgroup);

  if (headings) {
    const headRow = el("tr", { i18n: false });
    const corner = el("th", { text: "", i18n: false });
    corner.style.height = headingHeight + "pt";
    corner.style.background = "#f1f1f1";
    corner.style.fontSize = Math.max(5, 8 * metrics.scale) + "pt";
    headRow.append(corner);
    for (const column of band) {
      const th = el("th", { text: indexToCol(column.col), i18n: false });
      th.style.height = headingHeight + "pt";
      th.style.background = "#f1f1f1";
      th.style.border = "0.5pt solid #bfbfbf";
      th.style.fontSize = Math.max(5, 8 * metrics.scale) + "pt";
      th.style.fontWeight = "600";
      headRow.append(th);
    }
    table.append(headRow);
  }

  for (const row of chunk) {
    const tr = el("tr", { i18n: false });
    tr.style.height = row.height + "pt";
    if (headings) {
      const th = el("th", { text: String(row.row), i18n: false });
      th.style.height = row.height + "pt";
      th.style.background = "#f1f1f1";
      th.style.border = "0.5pt solid #bfbfbf";
      th.style.fontSize = Math.max(5, 8 * metrics.scale) + "pt";
      th.style.fontWeight = "600";
      th.style.textAlign = "center";
      tr.append(th);
    }
    for (const column of band) {
      if (covered.has(row.row + ":" + column.col)) continue;
      const td = makeCell(sheet, row.row, column.col, column, row.height, merges, metrics, gridlines);
      const span = starts.get(row.row + ":" + column.col);
      if (span) {
        td.rowSpan = span.rowSpan;
        td.colSpan = span.colSpan;
      }
      tr.append(td);
    }
    table.append(tr);
  }

  body.append(table);
  page.dataset.pageNumber = String(pageNumber);
  page.dataset.sheet = sheet.name;
  page.dataset.rowStart = String(chunk[0]?.row || 1);
  page.dataset.rowEnd = String(chunk[chunk.length - 1]?.row || 1);
  page.dataset.colStart = String(band[0]?.col || 1);
  page.dataset.colEnd = String(band[band.length - 1]?.col || 1);
  return page;
}

/* Build the page plan without touching the DOM; used both by rendering and
   by the Page Break preview. */
export function planPages(sheet, model = { title: "" }) {
  const metrics = pageMetrics(sheet.page || {});
  const range = printRange(sheet);
  const columns = collectColumns(sheet, range, metrics);
  const rows = collectRows(sheet, range, metrics);
  const bands = bandsOf(columns, metrics.contentW);
  const chunks = chunksOf(rows, metrics.contentH);
  const pages = [];
  bands.forEach((band, bandIndex) => {
    chunks.forEach((chunk, chunkIndex) => {
      pages.push({ band, chunk, bandIndex, chunkIndex, metrics });
    });
  });
  return { metrics, range, bands, chunks, pages };
}

function ensurePageStyle(width, height) {
  let node = document.getElementById("print-page-style");
  if (!node) {
    node = document.createElement("style");
    node.id = "print-page-style";
    document.head.append(node);
  }
  node.textContent = `@page { size: ${width}pt ${height}pt; margin: 0; }`;
}

export function renderPrintPages(model, options = {}) {
  const root = options.root || document.getElementById("print-root");
  if (!root) return { ok: false, pages: 0 };
  const allSheets = options.allSheets !== false;
  const sheets = allSheets
    ? model.sheets.filter((sheet) => !sheet.hidden)
    : [model.sheets[clamp(Number(options.sheetIndex ?? model.activeSheet) || 0, 0, model.sheets.length - 1)]];
  root.replaceChildren();
  const plans = sheets.map((sheet) => planPages(sheet, model));
  const total = plans.reduce((sum, plan) => sum + plan.pages.length, 0);
  let pageNumber = 0;
  let firstMetrics = null;
  for (let index = 0; index < sheets.length; index += 1) {
    const sheet = sheets[index];
    const plan = plans[index];
    if (!firstMetrics) firstMetrics = plan.metrics;
    const gridlines = sheet.page?.gridlines !== false;
    const headings = !!sheet.page?.headings;
    const merges = mergeMap(sheet);
    for (const entry of plan.pages) {
      pageNumber += 1;
      const node = renderPage(sheet, entry.chunk, entry.band, {
        metrics: entry.metrics,
        pageNumber,
        totalPages: total,
        model,
        merges,
        gridlines,
        headings,
      });
      root.append(node);
    }
  }
  if (!total) {
    const metrics = pageMetrics(sheets[0]?.page || {});
    root.append(el("div", { class: "print-page", i18n: false }));
    firstMetrics = metrics;
  }
  const metrics = firstMetrics || pageMetrics(sheets[0]?.page || {});
  ensurePageStyle(metrics.w, metrics.h);
  root.dataset.pageCount = String(total || 1);
  return { ok: true, pages: total || 1, w: metrics.w, h: metrics.h, plan: plans[0] };
}

export function clearPrint(root = document.getElementById("print-root")) {
  root?.replaceChildren();
}

export function printSummary(model, sheetIndex = model.activeSheet) {
  const sheet = model.sheets[sheetIndex] || model.sheets[0];
  const metrics = pageMetrics(sheet?.page || {});
  const paper = metrics.paper.toUpperCase();
  const orientation = metrics.orientation.charAt(0).toUpperCase() + metrics.orientation.slice(1);
  return `${paper} \u00B7 ${orientation} \u00B7 ${Math.round(metrics.scale * 100)}%`;
}

/* Small, non-interactive thumbnail used by the backstage Print panel. */
export function renderPreview(container, model, { sheetIndex = model.activeSheet } = {}) {
  if (!container) return;
  container.replaceChildren();
  const holder = el("div", { class: "print-preview-pages", i18n: false });
  container.append(holder);
  const sheet = model.sheets[sheetIndex] || model.sheets[0];
  const metrics = pageMetrics(sheet.page || {});
  const plan = planPages(sheet, model);
  const merges = mergeMap(sheet);
  const page = renderPage(sheet, plan.pages[0]?.chunk || [], plan.pages[0]?.band || [], {
    metrics,
    pageNumber: 1,
    totalPages: plan.pages.length,
    model,
    merges,
    gridlines: sheet.page?.gridlines !== false,
    headings: !!sheet.page?.headings,
  });
  holder.append(page);
  const scale = Math.min(1, (container.clientWidth - 36) / metrics.w, (container.clientHeight - 36) / metrics.h);
  holder.style.transform = `scale(${Math.max(0.1, scale)})`;
  holder.style.width = metrics.w + "pt";
  holder.style.height = metrics.h + "pt";
  holder.style.transformOrigin = "top left";
  holder.style.position = "absolute";
  holder.style.left = "18px";
  holder.style.top = "18px";
  container.style.position = "relative";
}

/* Page boundaries for Page Break Preview: the first row/column of every
   continuation page. */
export function computePageBreaks(model, sheetIndex = model.activeSheet) {
  const sheet = model.sheets[sheetIndex] || model.sheets[0];
  if (!sheet) return { rows: [], cols: [] };
  const plan = planPages(sheet, model);
  const rows = [];
  const cols = [];
  for (let index = 1; index < plan.chunks.length; index += 1) {
    const chunk = plan.chunks[index];
    if (chunk.length) rows.push(chunk[0].row);
  }
  for (let index = 1; index < plan.bands.length; index += 1) {
    const band = plan.bands[index];
    if (band.length) cols.push(band[0].col);
  }
  return { rows, cols };
}

export function pageCountFor(model, sheetIndex = model.activeSheet) {
  const sheet = model.sheets[sheetIndex] || model.sheets[0];
  return planPages(sheet, model).pages.length || 1;
}
