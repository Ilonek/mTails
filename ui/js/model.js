/* Workbook model shared with the Python file layer.
 *
 * A workbook is a plain JSON-serialisable object:
 *
 *   {version, title, activeSheet, settings:{calcMode, r1c1}, names, sheets:[]}
 *
 * A sheet keeps sparse maps: cells keyed by "A1", column widths keyed by
 * "A", row heights keyed by "1".  Cell objects use {v, f, s, h} for value,
 * formula ("=SUM(A1:A3)"), style and hyperlink.  Everything else mirrors the
 * shapes produced by mtails/xlsx_reader.py so a model can round-trip through
 * openpyxl without translation tables.
 */

import {
  MAX_ROWS,
  MAX_COLS,
  addr,
  parseAddr,
  parseRange,
  rangeText,
  normalizeRange,
  rangeContains,
  colToIndex,
  indexToCol,
} from "./units.js";

export const MODEL_VERSION = 1;
export const DEFAULT_COL_W = 64;
export const DEFAULT_ROW_H = 20;
export const MIN_COL_W = 10;
export const MAX_COL_W = 480;
export const MIN_ROW_H = 6;
export const MAX_ROW_H = 400;
export const DEFAULT_ROW_HEADER_W = 46;
export const DEFAULT_COL_HEADER_H = 22;

export function defaultPage() {
  return {
    size: "A4",
    orientation: "portrait",
    margins: { t: 72, r: 72, b: 72, l: 72, header: 30, footer: 30 },
    scale: 100,
    gridlines: false,
    headings: false,
    header: "",
    footer: "&P",
    area: "",
  };
}

export function createSheet(name = "Sheet1") {
  return {
    name,
    cells: {},
    cols: {},
    rows: {},
    merges: [],
    freeze: { rows: 0, cols: 0 },
    charts: [],
    images: [],
    textBoxes: [],
    comments: [],
    filters: null,
    validations: [],
    conditional: [],
    tables: [],
    page: defaultPage(),
    protection: null,
    tabColor: "",
    hidden: false,
    showGridlines: true,
    showHeadings: true,
    zoom: 100,
  };
}

export function createWorkbook(options = {}) {
  const count = Math.max(1, Math.min(20, Number(options.sheets) || 1));
  const sheets = [];
  for (let index = 0; index < count; index += 1) {
    sheets.push(createSheet(count === 1 ? "Sheet1" : `Sheet${index + 1}`));
  }
  return {
    version: MODEL_VERSION,
    title: options.title || "Workbook1",
    activeSheet: 0,
    settings: { calcMode: "auto", r1c1: false, showFormulas: false },
    properties: { title: options.title || "Workbook1", author: "", subject: "", keywords: "" },
    names: {},
    sheets,
  };
}

export function cloneModel(model) {
  return JSON.parse(JSON.stringify(model));
}

/* ------------------------------------------------------------- sanitising */
export function normalizeModel(raw) {
  const input = raw && typeof raw === "object" ? raw : {};
  const model = {
    version: MODEL_VERSION,
    title: String(input.title || "Workbook1"),
    activeSheet: 0,
    settings: {
      calcMode: input.settings?.calcMode === "manual" ? "manual" : "auto",
      r1c1: !!input.settings?.r1c1,
      showFormulas: !!input.settings?.showFormulas,
    },
    properties: {
      title: String(input.properties?.title || input.title || "Workbook1"),
      author: String(input.properties?.author || ""),
      subject: String(input.properties?.subject || ""),
      keywords: String(input.properties?.keywords || ""),
    },
    names: {},
    sheets: [],
  };
  const used = new Set();
  for (const entry of Array.isArray(input.sheets) ? input.sheets : []) {
    if (!entry || typeof entry !== "object") continue;
    const sheet = createSheet();
    sheet.name = uniqueName(used, String(entry.name || "Sheet"));
    used.add(sheet.name.toLowerCase());
    sheet.cells = {};
    if (entry.cells && typeof entry.cells === "object") {
      for (const [address, cell] of Object.entries(entry.cells)) {
        const parsed = parseAddr(address);
        if (!parsed || !cell || typeof cell !== "object") continue;
        const clean = {};
        if (cell.v !== undefined && cell.v !== null) clean.v = cell.v;
        if (typeof cell.f === "string" && cell.f.trim()) {
          clean.f = cell.f.startsWith("=") ? cell.f : "=" + cell.f;
        }
        if (cell.s && typeof cell.s === "object" && Object.keys(cell.s).length) clean.s = cell.s;
        if (typeof cell.h === "string" && cell.h) clean.h = cell.h;
        if (Object.keys(clean).length) sheet.cells[address] = clean;
      }
    }
    if (entry.cols && typeof entry.cols === "object") {
      for (const [key, value] of Object.entries(entry.cols)) {
        if (!value || typeof value !== "object") continue;
        const clean = {};
        if (Number(value.width) > 0) clean.width = Math.round(Number(value.width));
        if (value.hidden) clean.hidden = true;
        if (Object.keys(clean).length) sheet.cols[key.toUpperCase()] = clean;
      }
    }
    if (entry.rows && typeof entry.rows === "object") {
      for (const [key, value] of Object.entries(entry.rows)) {
        if (!value || typeof value !== "object") continue;
        const index = Number(key);
        if (!Number.isFinite(index) || index < 1) continue;
        const clean = {};
        if (Number(value.height) > 0) clean.height = Math.round(Number(value.height));
        if (value.hidden) clean.hidden = true;
        if (Object.keys(clean).length) sheet.rows[String(index)] = clean;
      }
    }
    sheet.merges = (Array.isArray(entry.merges) ? entry.merges : [])
      .map((range) => parseRange(range))
      .filter(Boolean)
      .map(rangeText);
    const freeze = entry.freeze || {};
    sheet.freeze = {
      rows: Math.max(0, Math.min(MAX_ROWS - 1, Number(freeze.rows) || 0)),
      cols: Math.max(0, Math.min(MAX_COLS - 1, Number(freeze.cols) || 0)),
    };
    sheet.charts = Array.isArray(entry.charts) ? entry.charts.filter((item) => item && typeof item === "object") : [];
    sheet.images = Array.isArray(entry.images) ? entry.images.filter((item) => item && typeof item === "object") : [];
    sheet.textBoxes = Array.isArray(entry.textBoxes)
      ? entry.textBoxes.filter((item) => item && typeof item === "object").map((item, index) => ({
          id: String(item.id || "text-" + (index + 1)),
          text: String(item.text ?? ""),
          x: Number(item.x) || 0,
          y: Number(item.y) || 0,
          w: Math.max(40, Number(item.w) || 200),
          h: Math.max(30, Number(item.h) || 90),
          size: Math.max(6, Number(item.size) || 12),
          bold: !!item.bold,
          italic: !!item.italic,
          color: typeof item.color === "string" ? item.color : "",
        }))
      : [];
    sheet.comments = (Array.isArray(entry.comments) ? entry.comments : [])
      .filter((item) => item && typeof item === "object" && item.addr)
      .map((item, index) => ({
        id: Number(item.id) || index + 1,
        addr: String(item.addr),
        text: String(item.text ?? ""),
        author: String(item.author || "mTails"),
        created: Number(item.created) || 0,
      }));
    sheet.filters = entry.filters && typeof entry.filters === "object"
      ? {
          range: String(entry.filters.range || ""),
          criteria: Array.isArray(entry.filters.criteria) ? entry.filters.criteria : [],
          hidden: Array.isArray(entry.filters.hidden) ? entry.filters.hidden : [],
        }
      : null;
    sheet.validations = Array.isArray(entry.validations)
      ? entry.validations.filter((item) => item && typeof item === "object" && item.range)
      : [];
    sheet.conditional = Array.isArray(entry.conditional)
      ? entry.conditional.filter((item) => item && typeof item === "object" && item.range)
      : [];
    sheet.tables = Array.isArray(entry.tables) ? entry.tables.filter((item) => item && item.range) : [];
    sheet.page = { ...defaultPage(), ...(entry.page && typeof entry.page === "object" ? entry.page : {}) };
    sheet.page.margins = { ...defaultPage().margins, ...(entry.page?.margins || {}) };
    sheet.protection = entry.protection && entry.protection.protected
      ? { protected: true, imported: !!entry.protection.imported }
      : null;
    sheet.tabColor = typeof entry.tabColor === "string" ? entry.tabColor : "";
    sheet.hidden = !!entry.hidden;
    sheet.showGridlines = entry.showGridlines !== false;
    sheet.showHeadings = entry.showHeadings !== false;
    sheet.zoom = Math.max(40, Math.min(200, Number(entry.zoom) || 100));
    model.sheets.push(sheet);
  }
  if (!model.sheets.length) model.sheets.push(createSheet("Sheet1"));
  for (const [name, ref] of Object.entries(input.names && typeof input.names === "object" ? input.names : {})) {
    if (!name || typeof ref !== "string") continue;
    model.names[String(name)] = ref;
  }
  model.activeSheet = Math.max(0, Math.min(model.sheets.length - 1, Number(input.activeSheet) || 0));
  return model;
}

function uniqueName(used, base) {
  let name = String(base || "Sheet").replace(/[\\/*?:[\]]/g, " ").trim().slice(0, 31) || "Sheet";
  let candidate = name;
  let suffix = 1;
  while (used.has(candidate.toLowerCase())) {
    suffix += 1;
    const tail = ` (${suffix})`;
    candidate = name.slice(0, 31 - tail.length) + tail;
  }
  return candidate;
}

/* ----------------------------------------------------------------- sheets */
export function activeSheet(model) {
  return model.sheets[model.activeSheet] || model.sheets[0];
}

export function sheetAt(model, index) {
  return model.sheets[index] || null;
}

export function sheetByName(model, name) {
  const lower = String(name || "").toLowerCase();
  return model.sheets.find((sheet) => sheet.name.toLowerCase() === lower) || null;
}

export function sheetIndex(model, name) {
  const lower = String(name || "").toLowerCase();
  return model.sheets.findIndex((sheet) => sheet.name.toLowerCase() === lower);
}

export function uniqueSheetName(model, base = "Sheet") {
  const used = new Set(model.sheets.map((sheet) => sheet.name.toLowerCase()));
  return uniqueName(used, base);
}

export function addSheet(model, options = {}) {
  const sheet = createSheet(uniqueSheetName(model, options.name || "Sheet"));
  if (options.index !== undefined) {
    model.sheets.splice(Math.max(0, Math.min(model.sheets.length, options.index)), 0, sheet);
    model.activeSheet = options.index;
  } else {
    model.sheets.push(sheet);
    model.activeSheet = model.sheets.length - 1;
  }
  return sheet;
}

export function removeSheet(model, index) {
  if (model.sheets.length <= 1) return false;
  model.sheets.splice(index, 1);
  if (model.activeSheet >= model.sheets.length) model.activeSheet = model.sheets.length - 1;
  else if (model.activeSheet > index) model.activeSheet -= 1;
  return true;
}

export function moveSheet(model, from, to) {
  const target = Math.max(0, Math.min(model.sheets.length - 1, to));
  if (from === target) return;
  const [sheet] = model.sheets.splice(from, 1);
  model.sheets.splice(target, 0, sheet);
  model.activeSheet = target;
}

export function renameSheet(model, index, name) {
  const sheet = model.sheets[index];
  if (!sheet) return { ok: false, error: "The sheet does not exist." };
  const clean = String(name || "").replace(/[\\/*?:[\]]/g, " ").trim();
  if (!clean) return { ok: false, error: "A sheet name cannot be empty." };
  if (clean.length > 31) return { ok: false, error: "A sheet name cannot be longer than 31 characters." };
  const lower = clean.toLowerCase();
  if (model.sheets.some((other, position) => position !== index && other.name.toLowerCase() === lower)) {
    return { ok: false, error: "A sheet with that name already exists." };
  }
  sheet.name = clean;
  return { ok: true, name: clean };
}

/* ------------------------------------------------------------------ cells */
export function cellAt(sheet, row, col) {
  return sheet.cells[addr(row, col)] || null;
}

export function getCell(sheet, address) {
  return sheet.cells[address] || null;
}

export function ensureCell(sheet, address) {
  let cell = sheet.cells[address];
  if (!cell) {
    cell = {};
    sheet.cells[address] = cell;
  }
  return cell;
}

export function isCellEmpty(cell) {
  if (!cell) return true;
  return (cell.v === undefined || cell.v === null || cell.v === "")
    && !cell.f && !cell.h && (!cell.s || !Object.keys(cell.s).length);
}

export function pruneCell(sheet, address) {
  const cell = sheet.cells[address];
  if (cell && isCellEmpty(cell)) delete sheet.cells[address];
}

export function setCellValue(sheet, address, value) {
  if (value === undefined || value === null || value === "") {
    const cell = sheet.cells[address];
    if (!cell) return null;
    delete cell.v;
    delete cell.f;
    pruneCell(sheet, address);
    return sheet.cells[address] || null;
  }
  const cell = ensureCell(sheet, address);
  cell.v = value;
  delete cell.f;
  return cell;
}

export function setCellFormula(sheet, address, formula, cached = null) {
  const cell = ensureCell(sheet, address);
  if (!formula || !String(formula).trim()) {
    delete cell.f;
    pruneCell(sheet, address);
    return sheet.cells[address] || null;
  }
  cell.f = String(formula).startsWith("=") ? String(formula) : "=" + String(formula);
  if (cached !== null && cached !== undefined) cell.v = cached;
  return cell;
}

export function clearCell(sheet, address, what = "all") {
  const cell = sheet.cells[address];
  if (!cell) return;
  if (what === "contents" || what === "all") {
    delete cell.v;
    delete cell.f;
  }
  if (what === "formats" || what === "all") delete cell.s;
  if (what === "comments" || what === "all") deleteCellComments(sheet, address);
  pruneCell(sheet, address);
}

export function cellStyle(cell) {
  return cell && cell.s ? cell.s : {};
}

export function styleKey(cell) {
  return cell && cell.s ? JSON.stringify(cell.s) : "";
}

export function usedRange(sheet, includeCharts = false) {
  let minRow = Infinity;
  let minCol = Infinity;
  let maxRow = 0;
  let maxCol = 0;
  const visit = (range) => {
    minRow = Math.min(minRow, range.r1);
    minCol = Math.min(minCol, range.c1);
    maxRow = Math.max(maxRow, range.r2);
    maxCol = Math.max(maxCol, range.c2);
  };
  for (const address of Object.keys(sheet.cells || {})) {
    const parsed = parseAddr(address);
    if (!parsed) continue;
    visit({ r1: parsed.row, c1: parsed.col, r2: parsed.row, c2: parsed.col });
  }
  for (const merge of sheet.merges || []) {
    const range = parseRange(merge);
    if (range) visit(range);
  }
  if (!maxRow) return null;
  const range = { r1: minRow, c1: minCol, r2: maxRow, c2: maxCol };
  if (includeCharts) {
    for (const chart of sheet.charts || []) {
      const top = rowAtY(sheet, chart.y || 0);
      const left = colAtX(sheet, chart.x || 0);
      const bottom = rowAtY(sheet, (chart.y || 0) + (chart.h || 0));
      const right = colAtX(sheet, (chart.x || 0) + (chart.w || 0));
      visit({ r1: top, c1: left, r2: Math.max(top, bottom), c2: Math.max(left, right) });
    }
  }
  return range;
}

export function countCells(sheet) {
  let filled = 0;
  let formulas = 0;
  for (const cell of Object.values(sheet.cells || {})) {
    if (cell && (cell.v !== undefined || cell.f)) filled += 1;
    if (cell && cell.f) formulas += 1;
  }
  return { filled, formulas };
}

export function workbookStats(model) {
  let cells = 0;
  let formulas = 0;
  for (const sheet of model.sheets) {
    const stats = countCells(sheet);
    cells += stats.filled;
    formulas += stats.formulas;
  }
  return { sheets: model.sheets.length, cells, formulas, names: Object.keys(model.names || {}).length };
}

/* -------------------------------------------------------------- dimensions */
export function colWidth(sheet, col) {
  const entry = sheet.cols[indexToCol(col)];
  return entry && Number(entry.width) > 0 ? Number(entry.width) : DEFAULT_COL_W;
}

export function rowHeight(sheet, row) {
  const entry = sheet.rows[String(row)];
  return entry && Number(entry.height) > 0 ? Number(entry.height) : DEFAULT_ROW_H;
}

export function colDisplayWidth(sheet, col) {
  const entry = sheet.cols[indexToCol(col)];
  if (entry?.hidden) return 0;
  return entry && Number(entry.width) > 0 ? Number(entry.width) : DEFAULT_COL_W;
}

export function rowDisplayHeight(sheet, row) {
  const entry = sheet.rows[String(row)];
  if (entry?.hidden) return 0;
  return entry && Number(entry.height) > 0 ? Number(entry.height) : DEFAULT_ROW_H;
}

export function isColHidden(sheet, col) {
  return !!sheet.cols[indexToCol(col)]?.hidden;
}

export function isRowHidden(sheet, row) {
  return !!sheet.rows[String(row)]?.hidden;
}

export function setColWidth(sheet, col, width) {
  const key = indexToCol(col);
  const entry = sheet.cols[key] || {};
  const value = Math.round(Math.max(MIN_COL_W, Math.min(MAX_COL_W, Number(width) || DEFAULT_COL_W)));
  entry.width = value;
  sheet.cols[key] = entry;
  return value;
}

export function setRowHeight(sheet, row, height) {
  const key = String(row);
  const entry = sheet.rows[key] || {};
  const value = Math.round(Math.max(MIN_ROW_H, Math.min(MAX_ROW_H, Number(height) || DEFAULT_ROW_H)));
  entry.height = value;
  sheet.rows[key] = entry;
  return value;
}

export function setColHidden(sheet, col, hidden) {
  const key = indexToCol(col);
  if (hidden) {
    const entry = sheet.cols[key] || {};
    entry.hidden = true;
    sheet.cols[key] = entry;
  } else if (sheet.cols[key]) {
    delete sheet.cols[key].hidden;
    if (!Object.keys(sheet.cols[key]).length) delete sheet.cols[key];
  }
}

export function setRowHidden(sheet, row, hidden) {
  const key = String(row);
  if (hidden) {
    const entry = sheet.rows[key] || {};
    entry.hidden = true;
    sheet.rows[key] = entry;
  } else if (sheet.rows[key]) {
    delete sheet.rows[key].hidden;
    if (!Object.keys(sheet.rows[key]).length) delete sheet.rows[key];
  }
}

/* Pixel geometry.  Structural edits and drawing placement use these; the
   grid builds a faster prefix cache for scrolling. */
export function colLeft(sheet, col) {
  let left = 0;
  for (let index = 1; index < col; index += 1) left += colDisplayWidth(sheet, index);
  return left;
}

export function rowTop(sheet, row) {
  let top = 0;
  for (let index = 1; index < row; index += 1) top += rowDisplayHeight(sheet, index);
  return top;
}

export function rowsSpanHeight(sheet, at, count) {
  let height = 0;
  for (let index = 0; index < count; index += 1) height += rowDisplayHeight(sheet, at + index);
  return height;
}

export function colsSpanWidth(sheet, at, count) {
  let width = 0;
  for (let index = 0; index < count; index += 1) width += colDisplayWidth(sheet, at + index);
  return width;
}

export function rowAtY(sheet, y) {
  let top = 0;
  for (let row = 1; row <= MAX_ROWS; row += 1) {
    const height = rowDisplayHeight(sheet, row);
    if (y < top + height) return row;
    top += height;
  }
  return MAX_ROWS;
}

export function colAtX(sheet, x) {
  let left = 0;
  for (let col = 1; col <= MAX_COLS; col += 1) {
    const width = colDisplayWidth(sheet, col);
    if (x < left + width) return col;
    left += width;
  }
  return MAX_COLS;
}

/* ---------------------------------------------------------------- merging */
export function mergeContains(sheet, address) {
  const parsed = parseAddr(address);
  if (!parsed) return null;
  return mergeAt(sheet, parsed.row, parsed.col);
}

export function mergeAt(sheet, row, col) {
  for (const text of sheet.merges || []) {
    const range = parseRange(text);
    if (range && rangeContains(range, row, col)) return { text, ...range };
  }
  return null;
}

export function mergeAnchor(sheet, range) {
  const normalized = sanitizeRange(range);
  return normalized ? addr(normalized.r1, normalized.c1) : null;
}

export function addMerge(sheet, range) {
  const normalized = sanitizeRange(range);
  if (!normalized) return null;
  if (normalized.r1 === normalized.r2 && normalized.c1 === normalized.c2) return null;
  sheet.merges = (sheet.merges || []).filter((text) => {
    const existing = parseRange(text);
    if (!existing) return false;
    return !rangesOverlap(existing, normalized);
  });
  const text = rangeText(normalized);
  sheet.merges.push(text);
  return text;
}

export function removeMergeAt(sheet, address) {
  const merge = mergeContains(sheet, address);
  if (!merge) return false;
  sheet.merges = sheet.merges.filter((text) => text !== merge.text);
  return true;
}

export function removeMergesIn(sheet, range) {
  const normalized = sanitizeRange(range);
  if (!normalized) return [];
  const removed = [];
  sheet.merges = sheet.merges.filter((text) => {
    const existing = parseRange(text);
    if (existing && rangesOverlap(existing, normalized)) {
      removed.push(text);
      return false;
    }
    return true;
  });
  return removed;
}

function sanitizeRange(range) {
  const parsed = typeof range === "string" ? parseRange(range) : range;
  if (!parsed) return null;
  const { r1, c1, r2, c2 } = parsed;
  if (![r1, c1, r2, c2].every((value) => Number.isFinite(Number(value)))) return null;
  return normalizeRange(Number(r1), Number(c1), Number(r2), Number(c2));
}

function rangesOverlap(a, b) {
  return a.r1 <= b.r2 && a.r2 >= b.r1 && a.c1 <= b.c2 && a.c2 >= b.c1;
}

/* ---------------------------------------------------------------- comments */
export function commentsAt(sheet, address) {
  return (sheet.comments || []).filter((comment) => comment.addr === address);
}

export function addComment(sheet, address, text, author = "mTails") {
  if (!text || !String(text).trim()) return null;
  const existing = commentsAt(sheet, address)[0];
  if (existing) {
    existing.text = String(text);
    existing.author = author || existing.author;
    return existing;
  }
  const id = (sheet.comments || []).reduce((max, comment) => Math.max(max, Number(comment.id) || 0), 0) + 1;
  const comment = {
    id,
    addr: address,
    text: String(text),
    author: author || "mTails",
    created: Date.now(),
  };
  sheet.comments.push(comment);
  return comment;
}

export function deleteCellComments(sheet, address) {
  const before = sheet.comments.length;
  sheet.comments = sheet.comments.filter((comment) => comment.addr !== address);
  return before !== sheet.comments.length;
}

export function deleteComment(sheet, id) {
  const before = sheet.comments.length;
  sheet.comments = sheet.comments.filter((comment) => comment.id !== id);
  return before !== sheet.comments.length;
}

/* ------------------------------------------------------------------ names */
const NAME_PATTERN = /^[A-Za-z_\\][A-Za-z0-9_.\\]*$/;
const CELL_PATTERN = /^[A-Za-z]{1,3}[1-9][0-9]*$/;

export function validName(name) {
  const text = String(name || "").trim();
  if (!text || text.length > 255) return { ok: false, error: "Enter a valid name." };
  if (!NAME_PATTERN.test(text) || CELL_PATTERN.test(text) || /^[RrCc]$/.test(text) || /^[Rr]\d+[Cc]\d+$/.test(text)) {
    return { ok: false, error: "Names must start with a letter, a backslash or an underscore and cannot look like a cell reference." };
  }
  return { ok: true, name: text };
}

export function defineName(model, name, ref) {
  const check = validName(name);
  if (!check.ok) return check;
  if (model.names[check.name] && String(ref) === String(model.names[check.name])) return check;
  model.names[check.name] = String(ref);
  return { ok: true, name: check.name };
}

export function removeName(model, name) {
  if (!(name in model.names)) return false;
  delete model.names[name];
  return true;
}

export function renameName(model, from, to) {
  if (!(from in model.names)) return { ok: false, error: "The name does not exist." };
  const check = validName(to);
  if (!check.ok) return check;
  if (check.name !== from && from in model.names && check.name in model.names) {
    return { ok: false, error: "A name with that label already exists." };
  }
  const ref = model.names[from];
  delete model.names[from];
  model.names[check.name] = ref;
  return { ok: true, name: check.name };
}

/* -------------------------------------------------------- structural edits */
export function insertRows(sheet, at, count) {
  const row = Math.max(1, Math.min(MAX_ROWS, at));
  const amount = Math.max(1, Math.min(MAX_ROWS - row, count));
  if (!amount) return 0;
  const moved = {};
  for (const [address, cell] of Object.entries(sheet.cells)) {
    const parsed = parseAddr(address);
    if (!parsed) continue;
    if (parsed.row >= row) moved[addr(parsed.row + amount, parsed.col)] = cell;
    else moved[address] = cell;
  }
  sheet.cells = moved;

  const rows = {};
  for (const [key, value] of Object.entries(sheet.rows)) {
    const index = Number(key);
    rows[String(index >= row ? index + amount : index)] = value;
  }
  sheet.rows = rows;

  for (const comment of sheet.comments) {
    const parsed = parseAddr(comment.addr);
    if (parsed && parsed.row >= row) comment.addr = addr(parsed.row + amount, parsed.col);
  }
  sheet.merges = sheet.merges.map((text) => {
    const range = parseRange(text);
    if (!range) return text;
    if (row <= range.r1) return rangeText({ ...range, r1: range.r1 + amount, r2: range.r2 + amount });
    if (row <= range.r2) return rangeText({ ...range, r2: range.r2 + amount });
    return text;
  });
  if (sheet.freeze.rows >= row) sheet.freeze.rows += amount;
  retargetRanges(sheet, "row", row, amount, false);
  shiftDrawings(sheet, "row", row, amount, false);
  return amount;
}

export function deleteRows(sheet, at, count) {
  const row = Math.max(1, Math.min(MAX_ROWS, at));
  const amount = Math.max(1, Math.min(MAX_ROWS - row + 1, count));
  if (!amount) return 0;
  const last = row + amount - 1;
  const removedHeight = rowsSpanHeight(sheet, row, amount);
  const moved = {};
  for (const [address, cell] of Object.entries(sheet.cells)) {
    const parsed = parseAddr(address);
    if (!parsed) continue;
    if (parsed.row < row) moved[address] = cell;
    else if (parsed.row > last) moved[addr(parsed.row - amount, parsed.col)] = cell;
  }
  sheet.cells = moved;

  const rows = {};
  for (const [key, value] of Object.entries(sheet.rows)) {
    const index = Number(key);
    if (index < row) rows[String(index)] = value;
    else if (index > last) rows[String(index - amount)] = value;
  }
  sheet.rows = rows;

  sheet.comments = sheet.comments.filter((comment) => {
    const parsed = parseAddr(comment.addr);
    if (!parsed) return false;
    if (parsed.row >= row && parsed.row <= last) return false;
    if (parsed.row > last) comment.addr = addr(parsed.row - amount, parsed.col);
    return true;
  });
  sheet.merges = adjustRangesForDelete(sheet.merges, "row", row, amount);
  if (sheet.freeze.rows >= last) sheet.freeze.rows -= amount;
  else if (sheet.freeze.rows >= row) sheet.freeze.rows = row - 1;
  retargetRanges(sheet, "row", row, amount, true);
  shiftDrawings(sheet, "row", row, amount, true, removedHeight);
  return amount;
}

export function insertCols(sheet, at, count) {
  const col = Math.max(1, Math.min(MAX_COLS, at));
  const amount = Math.max(1, Math.min(MAX_COLS - col, count));
  if (!amount) return 0;
  const moved = {};
  for (const [address, cell] of Object.entries(sheet.cells)) {
    const parsed = parseAddr(address);
    if (!parsed) continue;
    if (parsed.col >= col) moved[addr(parsed.row, parsed.col + amount)] = cell;
    else moved[address] = cell;
  }
  sheet.cells = moved;

  const cols = {};
  for (const [key, value] of Object.entries(sheet.cols)) {
    const index = colToIndex(key);
    if (!index) continue;
    cols[indexToCol(index >= col ? index + amount : index)] = value;
  }
  sheet.cols = cols;

  for (const comment of sheet.comments) {
    const parsed = parseAddr(comment.addr);
    if (parsed && parsed.col >= col) comment.addr = addr(parsed.row, parsed.col + amount);
  }
  sheet.merges = sheet.merges.map((text) => {
    const range = parseRange(text);
    if (!range) return text;
    if (col <= range.c1) return rangeText({ ...range, c1: range.c1 + amount, c2: range.c2 + amount });
    if (col <= range.c2) return rangeText({ ...range, c2: range.c2 + amount });
    return text;
  });
  if (sheet.freeze.cols >= col) sheet.freeze.cols += amount;
  retargetRanges(sheet, "col", col, amount, false);
  shiftDrawings(sheet, "col", col, amount, false);
  return amount;
}

export function deleteCols(sheet, at, count) {
  const col = Math.max(1, Math.min(MAX_COLS, at));
  const amount = Math.max(1, Math.min(MAX_COLS - col + 1, count));
  if (!amount) return 0;
  const last = col + amount - 1;
  const removedWidth = colsSpanWidth(sheet, col, amount);
  const moved = {};
  for (const [address, cell] of Object.entries(sheet.cells)) {
    const parsed = parseAddr(address);
    if (!parsed) continue;
    if (parsed.col < col) moved[address] = cell;
    else if (parsed.col > last) moved[addr(parsed.row, parsed.col - amount)] = cell;
  }
  sheet.cells = moved;

  const cols = {};
  for (const [key, value] of Object.entries(sheet.cols)) {
    const index = colToIndex(key);
    if (!index) continue;
    if (index < col) cols[key] = value;
    else if (index > last) cols[indexToCol(index - amount)] = value;
  }
  sheet.cols = cols;

  sheet.comments = sheet.comments.filter((comment) => {
    const parsed = parseAddr(comment.addr);
    if (!parsed) return false;
    if (parsed.col >= col && parsed.col <= last) return false;
    if (parsed.col > last) comment.addr = addr(parsed.row, parsed.col - amount);
    return true;
  });
  sheet.merges = adjustRangesForDelete(sheet.merges, "col", col, amount);
  if (sheet.freeze.cols >= last) sheet.freeze.cols -= amount;
  else if (sheet.freeze.cols >= col) sheet.freeze.cols = col - 1;
  retargetRanges(sheet, "col", col, amount, true);
  shiftDrawings(sheet, "col", col, amount, true, removedWidth);
  return amount;
}

function adjustRangesForDelete(ranges, axis, at, count) {
  const last = at + count - 1;
  const output = [];
  for (const text of ranges) {
    const range = parseRange(text);
    if (!range) continue;
    const start = axis === "row" ? range.r1 : range.c1;
    const end = axis === "row" ? range.r2 : range.c2;
    let nextStart = start;
    let nextEnd = end;
    if (end < at) {
      nextStart = start;
      nextEnd = end;
    } else if (start > last) {
      nextStart = start - count;
      nextEnd = end - count;
    } else {
      const startInside = start >= at && start <= last;
      const endInside = end >= at && end <= last;
      if (startInside && endInside) continue;
      if (startInside) nextStart = at;
      if (endInside) nextEnd = at - 1;
      else nextEnd = end - count;
      if (nextEnd < nextStart) continue;
    }
    output.push(rangeText(axis === "row"
      ? { ...range, r1: nextStart, r2: nextEnd }
      : { ...range, c1: nextStart, c2: nextEnd }));
  }
  return output;
}

function retargetRanges(sheet, axis, at, count, deleting) {
  const shiftText = (text) => {
    const range = parseRange(text);
    if (!range) return text;
    if (deleting) {
      const adjusted = adjustRangesForDelete([text], axis, at, count);
      return adjusted[0] || "";
    }
    if (axis === "row") {
      if (at <= range.r1) return rangeText({ ...range, r1: range.r1 + count, r2: range.r2 + count });
      if (at <= range.r2) return rangeText({ ...range, r2: range.r2 + count });
    } else {
      if (at <= range.c1) return rangeText({ ...range, c1: range.c1 + count, c2: range.c2 + count });
      if (at <= range.c2) return rangeText({ ...range, c2: range.c2 + count });
    }
    return text;
  };
  if (sheet.filters?.range) {
    const next = shiftText(sheet.filters.range);
    if (next) sheet.filters.range = next;
  }
  for (const collection of [sheet.validations, sheet.conditional, sheet.tables]) {
    for (const entry of collection || []) {
      if (!entry.range) continue;
      const next = shiftText(entry.range);
      if (!next) entry.range = "";
      else entry.range = next;
    }
  }
  if (deleting) {
    sheet.validations = sheet.validations.filter((entry) => entry.range);
    sheet.conditional = sheet.conditional.filter((entry) => entry.range);
    sheet.tables = sheet.tables.filter((entry) => entry.range);
  }
}

function shiftDrawings(sheet, axis, at, count, deleting, removedSpan = 0) {
  if (deleting) {
    const removed = removedSpan;
    const start = axis === "row" ? rowTop(sheet, at) : colLeft(sheet, at);
    for (const item of [...(sheet.charts || []), ...(sheet.images || [])]) {
      const position = axis === "row" ? Number(item.y) || 0 : Number(item.x) || 0;
      if (position >= start) {
        const next = Math.max(start, position - removed);
        if (axis === "row") item.y = next;
        else item.x = next;
      }
    }
    return;
  }
  const added = axis === "row" ? rowsSpanHeight(sheet, at, count) : colsSpanWidth(sheet, at, count);
  const start = axis === "row" ? rowTop(sheet, at) : colLeft(sheet, at);
  for (const item of [...(sheet.charts || []), ...(sheet.images || [])]) {
    const position = axis === "row" ? Number(item.y) || 0 : Number(item.x) || 0;
    if (position >= start) {
      if (axis === "row") item.y = position + added;
      else item.x = position + added;
    }
  }
}

/* ---------------------------------------------------------------- sorting */
export function sortRange(sheet, range, keys) {
  const normalized = sanitizeRange(range);
  if (!normalized || !keys?.length) return false;
  const rows = [];
  for (let row = normalized.r1; row <= normalized.r2; row += 1) {
    rows.push({ row, cells: [] });
  }
  for (const entry of rows) {
    for (let col = normalized.c1; col <= normalized.c2; col += 1) {
      entry.cells.push(sheet.cells[addr(entry.row, col)] || null);
    }
  }
  const keyIndex = (key) => {
    const parsed = parseAddr(key?.address || "");
    return parsed ? Math.max(0, parsed.col - normalized.c1) : 0;
  };
  rows.sort((left, right) => {
    for (const key of keys) {
      const index = keyIndex(key);
      const factor = key.order === "desc" ? -1 : 1;
      const result = compareValues(cellSortValue(left.cells[index]), cellSortValue(right.cells[index]), key.caseSensitive);
      if (result) return result * factor;
    }
    return 0;
  });
  rows.forEach((entry, index) => {
    const target = normalized.r1 + index;
    for (let col = normalized.c1; col <= normalized.c2; col += 1) {
      const cell = entry.cells[col - normalized.c1];
      if (cell) sheet.cells[addr(target, col)] = cell;
      else delete sheet.cells[addr(target, col)];
    }
  });
  return true;
}

export function cellSortValue(cell) {
  if (!cell) return null;
  if (cell.v === undefined || cell.v === null || cell.v === "") return null;
  return cell.v;
}

export function compareValues(left, right, caseSensitive = false) {
  const leftEmpty = left === null || left === undefined || left === "";
  const rightEmpty = right === null || right === undefined || right === "";
  if (leftEmpty && rightEmpty) return 0;
  if (leftEmpty) return 1;
  if (rightEmpty) return -1;
  const leftNumber = typeof left === "number";
  const rightNumber = typeof right === "number";
  if (leftNumber && rightNumber) return left - right;
  if (leftNumber !== rightNumber) return leftNumber ? -1 : 1;
  if (typeof left === "boolean" || typeof right === "boolean") {
    return Number(left) - Number(right);
  }
  const a = String(left);
  const b = String(right);
  return a.localeCompare(b, undefined, { numeric: true, sensitivity: caseSensitive ? "case" : "base" });
}

export function removeDuplicateRows(sheet, range, columns = null) {
  const normalized = sanitizeRange(range);
  if (!normalized) return 0;
  const selectors = columns?.length ? columns : null;
  const seen = new Set();
  const keep = [];
  let removed = 0;
  for (let row = normalized.r1; row <= normalized.r2; row += 1) {
    const values = [];
    for (let col = normalized.c1; col <= normalized.c2; col += 1) {
      if (selectors && !selectors.includes(col)) continue;
      const cell = sheet.cells[addr(row, col)];
      values.push(cell ? String(cell.v ?? cell.f ?? "") : "");
    }
    const key = values.join("\u001f");
    if (seen.has(key)) {
      removed += 1;
      for (let col = normalized.c1; col <= normalized.c2; col += 1) delete sheet.cells[addr(row, col)];
    } else {
      seen.add(key);
      keep.push(row);
    }
  }
  if (!removed) return 0;
  const compact = {};
  const sortedKeep = keep.slice().sort((a, b) => a - b);
  const mapping = new Map();
  for (let index = 0; index < keep.length; index += 1) mapping.set(sortedKeep[index], normalized.r1 + index);
  for (const [address, cell] of Object.entries(sheet.cells)) {
    const parsed = parseAddr(address);
    if (!parsed) continue;
    if (parsed.row >= normalized.r1 && parsed.row <= normalized.r2) {
      const target = mapping.get(parsed.row);
      if (target === undefined) continue;
      compact[addr(target, parsed.col)] = cell;
    } else {
      compact[address] = cell;
    }
  }
  sheet.cells = compact;
  return removed;
}
