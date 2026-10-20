/* Virtualised worksheet grid - rendering core.
 *
 * The grid draws only the cells that intersect the viewport.  Row and column
 * offsets live in two prefix arrays (Float64Array) so scrolling and hit
 * testing are O(log n) even with 100,000 rows and 1,000 columns.  Freeze
 * panes split the drawing into four quadrants: fixed rows scroll horizontally
 * but not vertically, fixed columns the other way round.
 *
 * The scroller underneath the cell layer owns the native scrollbars.  Cells
 * themselves are pure visuals; pointer events are handled on the scroller so
 * one listener covers the whole sheet.
 */

import {
  $,
  el,
  escapeHtml,
  clamp,
} from "./ui.js";
import { icon } from "./icons.js";
import { formatCell, parseUserInput } from "./format.js";
import {
  DEFAULT_COL_W,
  DEFAULT_ROW_H,
  colDisplayWidth,
  rowDisplayHeight,
  cellAt,
  usedRange,
} from "./model.js";
import {
  MAX_ROWS,
  MAX_COLS,
  addr,
  parseRange,
  rangeText,
  indexToCol,
  rangeSize,
  rangeContains,
} from "./units.js";

const CELL_RENDER_LIMIT = 6000;
export const GRID_FONT_STACK = '"Segoe UI", "Selawik", "Liberation Sans", "Noto Sans", sans-serif';
const BORDER_STYLES = {
  hair: "1px solid",
  thin: "1px solid",
  medium: "2px solid",
  thick: "3px solid",
  dashed: "1px dashed",
  dotted: "1px dotted",
  double: "3px double",
};

function isNumeric(value) {
  return typeof value === "number" && Number.isFinite(value);
}

export function rawEditText(cell) {
  if (!cell) return "";
  if (cell.f) return cell.f;
  if (cell.v === undefined || cell.v === null) return "";
  if (typeof cell.v === "boolean") return cell.v ? "TRUE" : "FALSE";
  return String(cell.v);
}

function colorMix(from, to, amount) {
  const parse = (value) => {
    const text = String(value || "").replace("#", "");
    if (text.length !== 6) return null;
    return [parseInt(text.slice(0, 2), 16), parseInt(text.slice(2, 4), 16), parseInt(text.slice(4, 6), 16)];
  };
  const a = parse(from);
  const b = parse(to);
  if (!a || !b) return to;
  const mix = a.map((channel, index) => Math.round(channel + (b[index] - channel) * amount));
  return "#" + mix.map((channel) => channel.toString(16).padStart(2, "0")).join("");
}

/* Pick a readable font colour for a solid fill when the cell has no explicit
   font colour.  Without this a light fill inherits the dark-theme light text
   and disappears; a dark fill does the same in the light theme. */
function fillTextColor(fill) {
  const match = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(String(fill || "").trim());
  if (!match) return "";
  let hex = match[1];
  if (hex.length === 3) hex = hex.split("").map((ch) => ch + ch).join("");
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  return (0.299 * r + 0.587 * g + 0.114 * b) / 255 > 0.62 ? "#1f1f1f" : "#ffffff";
}

/* Extract chart series from the sheet.  The chart module can override the
   presentation, but the built-in renderer keeps drawings useful after an
   import even before a chart is edited. */
export function chartSeries(chart, sheet) {
  const range = parseRange(chart.range || "");
  if (!range) return null;
  const grid = [];
  for (let row = range.r1; row <= range.r2; row += 1) {
    const values = [];
    for (let col = range.c1; col <= range.c2; col += 1) {
      const cell = cellAt(sheet, row, col);
      values.push(cell && cell.v !== undefined ? cell.v : null);
    }
    grid.push(values);
  }
  if (!grid.length) return null;
  let headerRow = !!chart.headerRow;
  let headerCol = !!chart.headerCol;
  if (!headerRow && grid[0].every((value) => typeof value === "string" || value === null)) {
    headerRow = grid[0].some((value) => typeof value === "string" && value);
  }
  if (!headerCol && grid.length > 1 && grid.slice(1).every((row) => typeof row[0] === "string" || row[0] === null)) {
    headerCol = grid[0].length > 1;
  }
  const categories = [];
  const series = [];
  const firstDataRow = headerRow ? 1 : 0;
  const firstDataCol = headerCol ? 1 : 0;
  for (let row = firstDataRow; row < grid.length; row += 1) {
    categories.push(headerCol ? String(grid[row][0] ?? row + 1) : String(row - firstDataRow + 1));
  }
  for (let col = firstDataCol; col < grid[0].length; col += 1) {
    const name = headerRow ? String(grid[0][col] ?? `Series ${col + 1}`) : `Series ${col + 1}`;
    const values = [];
    for (let row = firstDataRow; row < grid.length; row += 1) {
      const value = grid[row][col];
      values.push(isNumeric(value) ? value : 0);
    }
    series.push({ name, values });
  }
  return { categories, series };
}

const CHART_COLORS = ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47", "#264478", "#9E480E", "#636363"];

export function chartSvg(chart, sheet) {
  const data = chartSeries(chart, sheet);
  const empty = `<svg viewBox="0 0 320 180" preserveAspectRatio="none"><text x="50%" y="50%" text-anchor="middle" fill="currentColor" font-size="10">No data</text></svg>`;
  if (!data || !data.series.length) return empty;
  const W = 320;
  const H = 180;
  const left = 34;
  const right = 8;
  const top = 10;
  const bottom = 26;
  const plotW = W - left - right;
  const plotH = H - top - bottom;
  const type = String(chart.type || "column").toLowerCase();
  const palette = CHART_COLORS;

  if (type === "pie" || type === "doughnut") {
    const values = data.series[0].values.map((value) => Math.max(0, value));
    const total = values.reduce((sum, value) => sum + value, 0);
    if (!total) return empty;
    const cx = W / 2;
    const cy = H / 2;
    const radius = Math.min(W, H) / 2 - 16;
    let angle = -Math.PI / 2;
    const slices = values.map((value, index) => {
      const sweep = (value / total) * Math.PI * 2;
      const x1 = cx + radius * Math.cos(angle);
      const y1 = cy + radius * Math.sin(angle);
      angle += sweep;
      const x2 = cx + radius * Math.cos(angle);
      const y2 = cy + radius * Math.sin(angle);
      const large = sweep > Math.PI ? 1 : 0;
      return `<path d="M${cx} ${cy} L${x1.toFixed(1)} ${y1.toFixed(1)} A${radius} ${radius} 0 ${large} 1 ${x2.toFixed(1)} ${y2.toFixed(1)} Z" fill="${palette[index % palette.length]}" stroke="#fff" stroke-width="1"/>`;
    }).join("");
    const hole = type === "doughnut"
      ? `<circle cx="${cx}" cy="${cy}" r="${(radius * 0.55).toFixed(1)}" fill="var(--surface)"/>`
      : "";
    return `<svg viewBox="0 0 ${W} ${H}" preserveAspectRatio="none">${slices}${hole}</svg>`;
  }

  const totals = data.series.flatMap((item) => item.values);
  const max = Math.max(0, ...totals);
  const min = Math.min(0, ...totals);
  const span = max - min || 1;
  const yOf = (value) => top + plotH - ((value - min) / span) * plotH;
  const zero = yOf(0);
  const count = Math.max(1, data.categories.length);
  const baseline = `<line x1="${left}" y1="${zero.toFixed(1)}" x2="${left + plotW}" y2="${zero.toFixed(1)}" stroke="currentColor" opacity=".45"/>`;
  const labels = data.categories.map((label, index) => {
    if (count > 8 && index % Math.ceil(count / 8) !== 0) return "";
    const x = left + (index + 0.5) * (plotW / count);
    return `<text x="${x.toFixed(1)}" y="${H - 8}" text-anchor="middle" font-size="7" fill="currentColor" opacity=".75">${escapeHtml(String(label).slice(0, 7))}</text>`;
  }).join("");
  const gridlines = [0, 0.5, 1].map((fraction) => {
    const y = top + plotH * fraction;
    const value = max - span * fraction;
    return `<line x1="${left}" y1="${y.toFixed(1)}" x2="${left + plotW}" y2="${y.toFixed(1)}" stroke="currentColor" opacity=".12"/><text x="${left - 3}" y="${(y + 2).toFixed(1)}" text-anchor="end" font-size="6.5" fill="currentColor" opacity=".65">${Math.abs(value) >= 1000 ? (value / 1000).toFixed(1) + "k" : Math.round(value)}</text>`;
  }).join("");

  let body = "";
  if (type === "line" || type === "area" || type === "scatter") {
    data.series.forEach((item, seriesIndex) => {
      const points = item.values.map((value, index) => {
        const x = left + (index + 0.5) * (plotW / count);
        return `${x.toFixed(1)},${yOf(value).toFixed(1)}`;
      });
      const color = palette[seriesIndex % palette.length];
      if (type === "area") {
        body += `<polygon points="${left + plotW / count / 2},${zero} ${points.join(" ")} ${left + plotW - plotW / count / 2},${zero}" fill="${color}" opacity=".55"/>`;
      }
      if (type === "scatter") {
        item.values.forEach((value, index) => {
          const x = left + (index + 0.5) * (plotW / count);
          body += `<circle cx="${x.toFixed(1)}" cy="${yOf(value).toFixed(1)}" r="2.2" fill="${color}"/>`;
        });
      } else {
        body += `<polyline points="${points.join(" ")}" fill="none" stroke="${color}" stroke-width="1.8"/>`;
      }
    });
  } else {
    const groups = Math.max(1, data.series.length);
    const groupWidth = plotW / count;
    const barWidth = Math.max(1, groupWidth * 0.72 / groups);
    data.series.forEach((item, seriesIndex) => {
      item.values.forEach((value, index) => {
        const x = left + index * groupWidth + groupWidth * 0.14 + seriesIndex * barWidth;
        const y = yOf(Math.max(value, 0));
        const height = Math.max(0.5, Math.abs(yOf(value) - zero));
        body += `<rect x="${x.toFixed(1)}" y="${(value >= 0 ? y : zero).toFixed(1)}" width="${barWidth.toFixed(1)}" height="${height.toFixed(1)}" fill="${palette[seriesIndex % palette.length]}"/>`;
      });
    });
  }
  return `<svg viewBox="0 0 ${W} ${H}" preserveAspectRatio="none">${gridlines}${baseline}${body}${labels}</svg>`;
}

export class GridBase {
  constructor(options = {}) {
    this.options = options;
    this.defaultFont = options.defaultFont || "";
    this.defaultSize = Number(options.defaultSize) || 0;
    this.model = options.model;
    this.engine = options.engine;
    this.sheetIndex = this.model?.activeSheet || 0;
    this.zoom = 100;
    this.sel = { r1: 1, c1: 1, r2: 1, c2: 1 };
    this.active = { row: 1, col: 1 };
    this.scrollLeft = 0;
    this.scrollTop = 0;
    this.freezeRows = 0;
    this.freezeCols = 0;
    this.colPrefix = new Float64Array(MAX_COLS + 2);
    this.rowPrefix = new Float64Array(MAX_ROWS + 2);
    this.totalWidth = 0;
    this.totalHeight = 0;
    this.freezeWidth = 0;
    this.freezeHeight = 0;
    this.scale = 1;
    this.mergeMap = new Map();
    this.mergeList = [];
    this.commentCells = new Set();
    this.filterHidden = new Set();
    this.cellNodes = new Map();
    this.colHeadNodes = new Map();
    this.rowHeadNodes = new Map();
    this.commentNodes = new Map();
    this.overlayNodes = new Map();
    this.epoch = 0;
    this.frame = 0;
    this.drag = null;
    this.editing = null;
    this.fillPreview = null;
    this.searchSet = new Set();
    this.searchActive = "";
    this.auditSet = new Set();
    this.auditLinks = [];
    this.pageBreaks = { rows: [], cols: [] };
    this.cfCache = new Map();
    this.clipboard = null;
    this.lastPointer = null;
    this.autoScroll = null;
    this.mounted = false;
    this.destroyed = false;
  }

  /* ---------------------------------------------------------------- mount */
  mount() {
    if (this.mounted) return;
    this.els = {
      viewport: $("#grid-viewport"),
      scroller: $("#grid-scroller"),
      spacer: $("#grid-spacer"),
      layer: $("#grid-layer"),
      cells: $("#grid-cells"),
      selection: $("#selection-layer"),
      pins: $("#comment-pins"),
      overlay: $("#overlay-layer"),
      editor: $("#editor-layer"),
      fill: $("#fill-handle"),
      corner: $("#grid-corner"),
      colHeaders: $("#col-headers"),
      rowHeaders: $("#row-headers"),
    };
    if (!this.els.scroller) return;
    this.els.scroller.tabIndex = 0;
    this.els.cells.classList.toggle("no-gridlines", !this.sheet().showGridlines);

    this._onScroll = () => {
      this.scrollLeft = this.els.scroller.scrollLeft;
      this.scrollTop = this.els.scroller.scrollTop;
      this.scheduleRender();
    };
    this._onResize = () => this.invalidateLayout();
    this._onMouseDown = (event) => this.handleMouseDown(event);
    this._onDoubleClick = (event) => this.handleDoubleClick(event);
    this._onContextMenu = (event) => this.handleContextMenu(event);
    this._onMouseMove = (event) => this.handleMouseMove(event);
    this._onMouseUp = (event) => this.handleMouseUp(event);
    this._onKeyDown = (event) => this.handleKeyDown(event);
    this._onCopy = (event) => this.handleCopy(event);
    this._onCut = (event) => this.handleCut(event);
    this._onPaste = (event) => this.handlePaste(event);

    this.els.scroller.addEventListener("scroll", this._onScroll, { passive: true });
    this.els.scroller.addEventListener("mousedown", this._onMouseDown);
    this.els.scroller.addEventListener("dblclick", this._onDoubleClick);
    this.els.scroller.addEventListener("contextmenu", this._onContextMenu);
    this.els.scroller.addEventListener("keydown", this._onKeyDown);
    this.els.scroller.addEventListener("copy", this._onCopy);
    this.els.scroller.addEventListener("cut", this._onCut);
    this.els.scroller.addEventListener("paste", this._onPaste);
    window.addEventListener("mousemove", this._onMouseMove);
    window.addEventListener("mouseup", this._onMouseUp);
    window.addEventListener("resize", this._onResize);

    this.els.fill.addEventListener("mousedown", (event) => {
      event.preventDefault();
      event.stopPropagation();
      this.beginFillDrag(event);
    });
    this.els.colHeaders.addEventListener("mousedown", (event) => this.handleHeaderMouseDown(event, "col"));
    this.els.rowHeaders.addEventListener("mousedown", (event) => this.handleHeaderMouseDown(event, "row"));
    this.els.colHeaders.addEventListener("dblclick", (event) => this.handleHeaderDoubleClick(event, "col"));
    this.els.rowHeaders.addEventListener("dblclick", (event) => this.handleHeaderDoubleClick(event, "row"));
    this.els.corner.addEventListener("mousedown", (event) => {
      event.preventDefault();
      this.selectAll();
      this.els.scroller.focus({ preventScroll: true });
    });

    if (window.ResizeObserver) {
      this.resizeObserver = new ResizeObserver(() => this.invalidateLayout());
      this.resizeObserver.observe(this.els.viewport);
    }
    this.mounted = true;
    this.setSheet(this.sheetIndex, { silent: true });
  }

  destroy() {
    this.destroyed = true;
    if (this.frame) cancelAnimationFrame(this.frame);
    if (this.autoScroll) cancelAnimationFrame(this.autoScroll);
    if (this.resizeObserver) this.resizeObserver.disconnect();
  }

  sheet() {
    return this.model.sheets[this.sheetIndex] || this.model.sheets[0];
  }

  setModel(model, engine = null) {
    this.model = model;
    if (engine) this.engine = engine;
    this.setSheet(model.activeSheet || 0, { silent: true });
  }

  setEngine(engine) {
    this.engine = engine;
    this.epoch += 1;
    this.cfCache.clear();
  }

  setSheet(index, { silent = false } = {}) {
    this.sheetIndex = clamp(Number(index) || 0, 0, this.model.sheets.length - 1);
    this.model.activeSheet = this.sheetIndex;
    this.scrollLeft = 0;
    this.scrollTop = 0;
    if (this.els?.scroller) {
      this.els.scroller.scrollLeft = 0;
      this.els.scroller.scrollTop = 0;
    }
    this.cellNodes.clear();
    this.colHeadNodes.clear();
    this.rowHeadNodes.clear();
    this.commentNodes.clear();
    this.overlayNodes.clear();
    this.zoom = clamp(Number(this.sheet().zoom) || 100, 40, 200);
    this.buildLayout();
    if (this.els) {
      this.els.cells.classList.toggle("no-gridlines", !this.sheet().showGridlines);
      this.els.colHeaders.style.display = this.sheet().showHeadings === false ? "none" : "";
      this.els.rowHeaders.style.display = this.sheet().showHeadings === false ? "none" : "";
      this.els.corner.style.display = this.sheet().showHeadings === false ? "none" : "";
    }
    this.setSelection(1, 1, 1, 1, { activeRow: 1, activeCol: 1, silent: true });
    this.epoch += 1;
    if (!silent) this.renderNow();
    this.notifySelection();
  }

  /* --------------------------------------------------------------- layout */
  buildLayout() {
    const sheet = this.sheet();
    this.scale = this.zoom / 100;
    this.filterHidden = new Set((sheet.filters?.hidden || []).map(Number));
    const colPrefix = this.colPrefix;
    colPrefix[0] = 0;
    for (let col = 1; col <= MAX_COLS; col += 1) {
      const raw = colDisplayWidth(sheet, col) * this.scale;
      const size = raw > 0 ? Math.max(1, Math.round(raw)) : 0;
      colPrefix[col] = colPrefix[col - 1] + size;
    }
    const rowPrefix = this.rowPrefix;
    rowPrefix[0] = 0;
    for (let row = 1; row <= MAX_ROWS; row += 1) {
      const hidden = this.filterHidden.has(row);
      const raw = hidden ? 0 : rowDisplayHeight(sheet, row) * this.scale;
      const size = raw > 0 ? Math.max(1, Math.round(raw)) : 0;
      rowPrefix[row] = rowPrefix[row - 1] + size;
    }
    this.totalWidth = colPrefix[MAX_COLS];
    this.totalHeight = rowPrefix[MAX_ROWS];
    this.freezeRows = clamp(Number(sheet.freeze?.rows) || 0, 0, MAX_ROWS - 1);
    this.freezeCols = clamp(Number(sheet.freeze?.cols) || 0, 0, MAX_COLS - 1);
    this.freezeWidth = colPrefix[this.freezeCols];
    this.freezeHeight = rowPrefix[this.freezeRows];

    this.commentCells = new Set((sheet.comments || []).map((comment) => comment.addr));
    this.mergeMap = new Map();
    this.mergeList = [];
    for (const text of sheet.merges || []) {
      const range = parseRange(text);
      if (!range) continue;
      const entry = { text, ...range };
      this.mergeList.push(entry);
      const area = rangeSize(range);
      if (area <= 60000) {
        for (let row = range.r1; row <= range.r2; row += 1) {
          for (let col = range.c1; col <= range.c2; col += 1) {
            this.mergeMap.set(row + ":" + col, entry);
          }
        }
      }
    }
    if (this.els) {
      this.els.spacer.style.width = Math.max(1, this.totalWidth) + "px";
      this.els.spacer.style.height = Math.max(1, this.totalHeight) + "px";
    }
    const maxLeft = Math.max(0, this.totalWidth - (this.els?.scroller.clientWidth || 0));
    const maxTop = Math.max(0, this.totalHeight - (this.els?.scroller.clientHeight || 0));
    if (this.scrollLeft > maxLeft) this.scrollLeft = maxLeft;
    if (this.scrollTop > maxTop) this.scrollTop = maxTop;
  }

  invalidateLayout() {
    if (!this.mounted) return;
    this.buildLayout();
    this.renderNow();
  }

  invalidateValues() {
    this.epoch += 1;
    this.cfCache.clear();
    this.scheduleRender();
  }

  scheduleRender() {
    if (this.frame || !this.mounted) return;
    this.frame = requestAnimationFrame(() => {
      this.frame = 0;
      if (!this.destroyed) this.renderNow();
    });
  }

  renderNow() {
    if (!this.mounted) return;
    const scroller = this.els.scroller;
    this.scrollLeft = scroller.scrollLeft;
    this.scrollTop = scroller.scrollTop;
    this.viewportWidth = scroller.clientWidth;
    this.viewportHeight = scroller.clientHeight;
    if (this.els.spacer.style.width !== Math.max(1, this.totalWidth) + "px") {
      this.els.spacer.style.width = Math.max(1, this.totalWidth) + "px";
    }
    if (this.els.spacer.style.height !== Math.max(1, this.totalHeight) + "px") {
      this.els.spacer.style.height = Math.max(1, this.totalHeight) + "px";
    }
    this.renderColumnHeaders();
    this.renderRowHeaders();
    this.renderCells();
    this.renderSelection();
    this.renderComments();
    this.renderOverlays();
    this.renderEditor();
  }

  /* ------------------------------------------------------------ geometry */
  colSize(col) {
    return this.colPrefix[col] - this.colPrefix[col - 1];
  }

  rowSize(row) {
    return this.rowPrefix[row] - this.rowPrefix[row - 1];
  }

  screenLeft(col) {
    return this.colPrefix[col - 1] - (col > this.freezeCols ? this.scrollLeft : 0);
  }

  screenTop(row) {
    return this.rowPrefix[row - 1] - (row > this.freezeRows ? this.scrollTop : 0);
  }

  findItem(prefix, position, max) {
    if (position <= 0) return 1;
    let low = 1;
    let high = max;
    while (low < high) {
      const mid = Math.ceil((low + high) / 2);
      if (prefix[mid] <= position) low = mid;
      else high = mid - 1;
    }
    return low;
  }

  findVisible(prefix, position, max) {
    let index = this.findItem(prefix, Math.max(0, position), max);
    while (index < max && prefix[index] === prefix[index - 1]) index += 1;
    if (index === max && prefix[index] === prefix[index - 1]) {
      while (index > 1 && prefix[index] === prefix[index - 1]) index -= 1;
    }
    return index;
  }

  hitTest(x, y) {
    x = clamp(x, 0, Math.max(0, this.viewportWidth - 1));
    y = clamp(y, 0, Math.max(0, this.viewportHeight - 1));
    const contentX = x < this.freezeWidth ? x : x + this.scrollLeft;
    const contentY = y < this.freezeHeight ? y : y + this.scrollTop;
    const col = this.findVisible(this.colPrefix, contentX, MAX_COLS);
    const row = this.findVisible(this.rowPrefix, contentY, MAX_ROWS);
    const merge = this.mergeFor(row, col);
    if (merge) return { row: merge.r1, col: merge.c1, merge, x, y };
    return { row, col, merge: null, x, y };
  }

  mergeFor(row, col) {
    return this.mergeMap.get(row + ":" + col) || null;
  }

  cellRect(row, col, merge = null) {
    const r2 = merge ? merge.r2 : row;
    const c2 = merge ? merge.c2 : col;
    return {
      left: this.screenLeft(col),
      top: this.screenTop(row),
      width: this.colPrefix[c2] - this.colPrefix[col - 1],
      height: this.rowPrefix[r2] - this.rowPrefix[row - 1],
    };
  }

  scrollToCell(row, col, { center = false } = {}) {
    if (!this.mounted) return;
    const scroller = this.els.scroller;
    const left = this.colPrefix[Math.max(0, col - 1)];
    const right = this.colPrefix[col];
    const top = this.rowPrefix[Math.max(0, row - 1)];
    const bottom = this.rowPrefix[row];
    const width = scroller.clientWidth;
    const height = scroller.clientHeight;
    const freezeW = this.freezeWidth;
    const freezeH = this.freezeHeight;
    if (col > this.freezeCols) {
      if (center) scroller.scrollLeft = Math.max(0, (left + right) / 2 - width / 2);
      else if (left - freezeW < scroller.scrollLeft) scroller.scrollLeft = Math.max(0, left - freezeW);
      else if (right > scroller.scrollLeft + width) scroller.scrollLeft = right - width;
    }
    if (row > this.freezeRows) {
      if (center) scroller.scrollTop = Math.max(0, (top + bottom) / 2 - height / 2);
      else if (top - freezeH < scroller.scrollTop) scroller.scrollTop = Math.max(0, top - freezeH);
      else if (bottom > scroller.scrollTop + height) scroller.scrollTop = bottom - height;
    }
    this.scrollLeft = scroller.scrollLeft;
    this.scrollTop = scroller.scrollTop;
    this.renderNow();
  }

  ensureVisible(row, col) {
    this.scrollToCell(row, col);
  }

  setZoom(percent) {
    const value = clamp(Math.round(Number(percent) || 100), 40, 200);
    if (value === this.zoom) return;
    this.zoom = value;
    this.sheet().zoom = value;
    const active = this.active;
    this.buildLayout();
    this.epoch += 1;
    this.scrollToCell(active.row, active.col);
    this.options.onZoomChange?.(value);
    this.options.onChange?.({ kind: "zoom", sheetIndex: this.sheetIndex });
  }

  /* -------------------------------------------------------------- headers */
  renderColumnHeaders() {
    if (!this.sheet().showHeadings) return;
    const width = this.viewportWidth;
    const used = new Set();
    const lastFrozen = Math.min(this.freezeCols, this.findVisible(this.colPrefix, this.freezeWidth, MAX_COLS) + 1);
    const ranges = [];
    if (this.freezeCols) ranges.push([1, lastFrozen]);
    let start = this.findItem(this.colPrefix, this.scrollLeft + this.freezeWidth, MAX_COLS) - 1;
    start = Math.max(this.freezeCols + 1, start, 1);
    const end = Math.min(MAX_COLS, this.findItem(this.colPrefix, this.scrollLeft + width, MAX_COLS) + 1);
    if (end >= start) ranges.push([start, end]);

    for (const [from, to] of ranges) {
      for (let col = from; col <= to; col += 1) {
        const size = this.colSize(col);
        if (!size) continue;
        const left = this.screenLeft(col);
        if (left >= width || left + size <= 0) continue;
        used.add(col);
        let node = this.colHeadNodes.get(col);
        if (!node) {
          node = el("div", { class: "colhead", dataset: { col: String(col) }, i18n: false });
          node.append(el("span", { i18n: false }));
          node.append(el("div", { class: "resize-handle", i18n: false }));
          this.els.colHeaders.append(node);
          this.colHeadNodes.set(col, node);
        }
        node.style.cssText = `left:${Math.round(left)}px;width:${size}px;`;
        const selected = col >= this.sel.c1 && col <= this.sel.c2;
        node.classList.toggle("sel", selected);
        const label = node.firstChild;
        if (label.textContent !== indexToCol(col)) label.textContent = indexToCol(col);
        this.syncFilterButton(node, col);
      }
    }
    for (const [col, node] of this.colHeadNodes) {
      if (!used.has(col)) {
        node.remove();
        this.colHeadNodes.delete(col);
      }
    }
    if (this.freezeCols && this.freezeWidth < width) {
      if (!this.els.frozenX) {
        this.els.frozenX = el("div", { class: "frozen-divider-x", i18n: false });
        this.els.layer.append(this.els.frozenX);
      }
      this.els.frozenX.style.left = Math.round(this.freezeWidth) + "px";
      this.els.frozenX.style.display = "";
    } else if (this.els.frozenX) {
      this.els.frozenX.style.display = "none";
    }
  }

  syncFilterButton(node, col) {
    const filters = this.sheet().filters;
    let wanted = false;
    let active = false;
    if (filters?.range) {
      const range = parseRange(filters.range);
      if (range && col >= range.c1 && col <= range.c2) {
        wanted = true;
        active = this.filterHidden.size > 0;
      }
    }
    let button = node.querySelector(".filter-btn");
    if (wanted && !button) {
      button = el("button", { class: "filter-btn", type: "button", title: "Filter", html: icon("chevronDown"), i18n: false });
      button.addEventListener("mousedown", (event) => event.stopPropagation());
      button.addEventListener("click", (event) => {
        event.stopPropagation();
        this.options.onFilterRequest?.(col, button);
      });
      node.append(button);
    } else if (!wanted && button) {
      button.remove();
    }
    if (button) {
      button.classList.toggle("active", active);
      button.title = active ? "Filter applied" : "Filter";
    }
  }

  renderRowHeaders() {
    if (!this.sheet().showHeadings) return;
    const height = this.viewportHeight;
    const used = new Set();
    const lastFrozen = Math.min(this.freezeRows, this.findVisible(this.rowPrefix, this.freezeHeight, MAX_ROWS) + 1);
    const ranges = [];
    if (this.freezeRows) ranges.push([1, lastFrozen]);
    let start = this.findItem(this.rowPrefix, this.scrollTop + this.freezeHeight, MAX_ROWS) - 1;
    start = Math.max(this.freezeRows + 1, start, 1);
    const end = Math.min(MAX_ROWS, this.findItem(this.rowPrefix, this.scrollTop + height, MAX_ROWS) + 1);
    if (end >= start) ranges.push([start, end]);

    for (const [from, to] of ranges) {
      for (let row = from; row <= to; row += 1) {
        const size = this.rowSize(row);
        if (!size) continue;
        const top = this.screenTop(row);
        if (top >= height || top + size <= 0) continue;
        used.add(row);
        let node = this.rowHeadNodes.get(row);
        if (!node) {
          node = el("div", { class: "rowhead", dataset: { row: String(row) }, i18n: false });
          node.append(el("span", { i18n: false }));
          node.append(el("div", { class: "resize-handle", i18n: false }));
          this.els.rowHeaders.append(node);
          this.rowHeadNodes.set(row, node);
        }
        node.style.cssText = `top:${Math.round(top)}px;height:${size}px;`;
        node.classList.toggle("sel", row >= this.sel.r1 && row <= this.sel.r2);
        const label = node.firstChild;
        const text = String(row);
        if (label.textContent !== text) label.textContent = text;
      }
    }
    for (const [row, node] of this.rowHeadNodes) {
      if (!used.has(row)) {
        node.remove();
        this.rowHeadNodes.delete(row);
      }
    }
    if (this.freezeRows && this.freezeHeight < height) {
      if (!this.els.frozenY) {
        this.els.frozenY = el("div", { class: "frozen-divider-y", i18n: false });
        this.els.layer.append(this.els.frozenY);
      }
      this.els.frozenY.style.top = Math.round(this.freezeHeight) + "px";
      this.els.frozenY.style.display = "";
    } else if (this.els.frozenY) {
      this.els.frozenY.style.display = "none";
    }
  }

  /* --------------------------------------------------------------- cells */
  visibleRows() {
    const height = this.viewportHeight;
    const rows = [];
    if (this.freezeRows) {
      const last = Math.min(this.freezeRows, this.findVisible(this.rowPrefix, this.freezeHeight, MAX_ROWS));
      for (let row = 1; row <= last; row += 1) {
        if (this.rowSize(row) && this.screenTop(row) < height) rows.push(row);
      }
    }
    let start = this.findItem(this.rowPrefix, this.scrollTop + this.freezeHeight, MAX_ROWS) - 1;
    start = Math.max(this.freezeRows + 1, start, 1);
    const end = Math.min(MAX_ROWS, this.findItem(this.rowPrefix, this.scrollTop + height, MAX_ROWS) + 1);
    for (let row = start; row <= end; row += 1) {
      const size = this.rowSize(row);
      if (!size) continue;
      const top = this.screenTop(row);
      if (top >= height || top + size <= 0) continue;
      if (row > this.freezeRows && top + size <= this.freezeHeight) continue;
      rows.push(row);
    }
    return rows;
  }

  visibleCols() {
    const width = this.viewportWidth;
    const cols = [];
    if (this.freezeCols) {
      const last = Math.min(this.freezeCols, this.findVisible(this.colPrefix, this.freezeWidth, MAX_COLS));
      for (let col = 1; col <= last; col += 1) {
        if (this.colSize(col) && this.screenLeft(col) < width) cols.push(col);
      }
    }
    let start = this.findItem(this.colPrefix, this.scrollLeft + this.freezeWidth, MAX_COLS) - 1;
    start = Math.max(this.freezeCols + 1, start, 1);
    const end = Math.min(MAX_COLS, this.findItem(this.colPrefix, this.scrollLeft + width, MAX_COLS) + 1);
    for (let col = start; col <= end; col += 1) {
      const size = this.colSize(col);
      if (!size) continue;
      const left = this.screenLeft(col);
      if (left >= width || left + size <= 0) continue;
      if (col > this.freezeCols && left + size <= this.freezeWidth) continue;
      cols.push(col);
    }
    return cols;
  }

  renderCells() {
    const rows = this.visibleRows();
    const cols = this.visibleCols();
    const frame = new Set();
    if (rows.length * cols.length > CELL_RENDER_LIMIT) {
      rows.length = Math.min(rows.length, 220);
      cols.length = Math.min(cols.length, 40);
    }
    for (const row of rows) {
      for (const col of cols) {
        const merge = this.mergeFor(row, col);
        if (merge && (merge.r1 !== row || merge.c1 !== col)) continue;
        this.paintCellNode(row, col, merge, frame);
      }
    }
    if (this.mergeList.length) {
      for (const merge of this.mergeList) {
        const rect = this.cellRect(merge.r1, merge.c1, merge);
        if (rect.left >= this.viewportWidth || rect.top >= this.viewportHeight) continue;
        if (rect.left + rect.width <= 0 || rect.top + rect.height <= 0) continue;
        const key = merge.r1 + ":" + merge.c1;
        if (!frame.has(key)) this.paintCellNode(merge.r1, merge.c1, merge, frame);
      }
    }
    for (const [key, node] of this.cellNodes) {
      if (!frame.has(key)) {
        node.remove();
        this.cellNodes.delete(key);
      }
    }
  }

  paintCellNode(row, col, merge, frame) {
    const sheet = this.sheet();
    const cell = cellAt(sheet, row, col);
    const key = row + ":" + col;
    frame.add(key);
    let node = this.cellNodes.get(key);
    if (!node) {
      node = el("div", { class: "cell", i18n: false });
      this.els.cells.append(node);
      this.cellNodes.set(key, node);
      node._epoch = -1;
      node._cellRef = undefined;
    }
    const rect = this.cellRect(row, col, merge);
    node.style.left = Math.round(rect.left) + "px";
    node.style.top = Math.round(rect.top) + "px";
    node.style.width = rect.width + "px";
    node.style.height = rect.height + "px";
    if (!cell && node._cellRef === null && node._epoch === this.epoch) return;
    this.paintCellContent(node, cell, row, col);
    node._cellRef = cell || null;
    node._epoch = this.epoch;
  }

  paintCellContent(node, cell, row, col) {
    const sheet = this.sheet();
    const style = cell?.s || {};
    const formatted = formatCell(cell || {});
    const cf = this.conditionalStyle(sheet, row, col, cell);
    const showFormulas = !!this.model?.settings?.showFormulas;
    let text = formatted.text;
    if (showFormulas && cell?.f) text = cell.f;
    else if (cell?.f && (cell.v === undefined || cell.v === null)) text = "";
    if (cell?.h && !text) text = cell.h;
    const rotate = Number(style.rotate) || 0;
    if (rotate) {
      let span = node.querySelector(".cell-rot");
      if (!span) {
        node.textContent = "";
        span = el("span", { class: "cell-rot", i18n: false });
        node.append(span);
      }
      if (span.textContent !== text) span.textContent = text;
      span.style.transform = `rotate(${-rotate}deg)`;
    } else {
      const span = node.querySelector(".cell-rot");
      if (span) span.remove();
      if (node.textContent !== text) node.textContent = text;
    }

    const classes = ["cell"];
    if (row <= this.freezeRows || col <= this.freezeCols) classes.push("frozen");
    const align = style.align || formatted.align;
    if (align === "center") classes.push("center");
    else if (align === "right" || (formatted.kind === "number" && !style.align)) classes.push("num");
    else if (align === "justify") classes.push("justify");
    if (style.wrap) classes.push("wrap");
    if (style.b) classes.push("bold");
    if (style.i) classes.push("italic");
    if (style.strike) classes.push("strike");
    if (style.u) classes.push("underline");
    if (cell?.h) classes.push("link");
    if (formatted.kind === "error") classes.push("error");
    if (this.searchSet.has(sheet.name.toLowerCase() + "!" + addr(row, col).toLowerCase())) classes.push("search-hit");
    if (this.auditSet.has(addr(row, col))) classes.push("audit");
    if (this.commentCells.has(addr(row, col))) classes.push("commented");
    node.className = classes.join(" ");

    const parts = [];
    if (style.fill) parts.push(`background:${style.fill}`);
    if (cf?.fill) parts.push(`background:${cf.fill}`);
    const color = cf?.color || style.color || formatted.color;
    if (color) parts.push(`color:${color}`);
    else {
      const fill = cf?.fill || style.fill;
      const fallback = fill ? fillTextColor(fill) : "";
      if (fallback) parts.push(`color:${fallback}`);
    }
    const fontName = style.font || this.defaultFont;
    if (fontName) parts.push(`font-family:"${String(fontName).replace(/"/g, "")}",${GRID_FONT_STACK}`);
    const points = Number(style.size) || Number(this.defaultSize) || 11;
    const pixelSize = Math.max(6, Math.round(points * (4 / 3) * this.scale * 10) / 10);
    parts.push(`font-size:${pixelSize}px`);
    if (style.b || cf?.bold) parts.push("font-weight:700");
    if (style.i) parts.push("font-style:italic");
    const decorations = [];
    if (style.u) decorations.push("underline");
    if (style.strike) decorations.push("line-through");
    if (decorations.length) parts.push(`text-decoration:${decorations.join(" ")}`);
    const valign = style.valign || "middle";
    parts.push(`align-items:${valign === "top" ? "flex-start" : valign === "bottom" ? "flex-end" : "center"}`);
    if (style.indent) parts.push(`padding-left:${4 + Number(style.indent) * 8 * this.scale}px`);
    if (style.border && typeof style.border === "object") {
      for (const side of ["t", "r", "b", "l"]) {
        const spec = style.border[side];
        if (!spec || !spec.style) continue;
        const width = BORDER_STYLES[spec.style] || "1px solid";
        const name = { t: "Top", r: "Right", b: "Bottom", l: "Left" }[side];
        parts.push(`border-${name.toLowerCase()}:${width} ${spec.color || "var(--grid-line-strong)"}`);
      }
    }
    const geometry = node.style.cssText.split(";").filter((part) => /^\s*(left|top|width|height):/.test(part)).join(";");
    node.style.cssText = geometry + ";" + parts.join(";");
  }

  conditionalStyle(sheet, row, col, cell) {
    const rules = sheet.conditional || [];
    if (!rules.length || !cell || cell.v === undefined) return null;
    let result = null;
    for (let index = 0; index < rules.length; index += 1) {
      const rule = rules[index];
      const range = parseRange(rule.range);
      if (!range || !rangeContains(range, row, col)) continue;
      if (rule.type === "cellIs") {
        if (this.matchesRule(rule, cell.v)) {
          result = {
            ...(result || {}),
            fill: rule.fill || result?.fill,
            color: rule.color || result?.color,
            bold: rule.bold || result?.bold,
          };
        }
      } else if (rule.type === "colorScale") {
        const color = this.scaleColor(rule, index, cell.v, range);
        if (color) result = { ...(result || {}), fill: color };
      }
    }
    return result;
  }

  matchesRule(rule, value) {
    const resolve = (text) => {
      if (text === undefined || text === null || text === "") return null;
      const raw = String(text);
      if (raw.startsWith("=")) {
        try {
          return this.engine?.evaluate(raw, this.sheet().name, "A1") ?? null;
        } catch {
          return null;
        }
      }
      const parsed = parseUserInput(raw);
      return parsed.kind === "number" ? parsed.value : parsed.kind === "boolean" ? parsed.value : raw;
    };
    const first = resolve(rule.formula1);
    const second = resolve(rule.formula2);
    const compare = (a, b) => {
      if (typeof a === "number" && typeof b === "number") return a - b;
      const left = String(a ?? "").toLowerCase();
      const right = String(b ?? "").toLowerCase();
      return left < right ? -1 : left > right ? 1 : 0;
    };
    switch (rule.operator) {
      case "lessThan": return compare(value, first) < 0;
      case "lessThanOrEqual": return compare(value, first) <= 0;
      case "greaterThan": return compare(value, first) > 0;
      case "greaterThanOrEqual": return compare(value, first) >= 0;
      case "equal": return compare(value, first) === 0;
      case "notEqual": return compare(value, first) !== 0;
      case "between": return compare(value, first) >= 0 && compare(value, second) <= 0;
      case "notBetween": return !(compare(value, first) >= 0 && compare(value, second) <= 0);
      case "containsText": return String(value).toLowerCase().includes(String(first ?? "").toLowerCase());
      case "notContains": return !String(value).toLowerCase().includes(String(first ?? "").toLowerCase());
      case "beginsWith": return String(value).toLowerCase().startsWith(String(first ?? "").toLowerCase());
      case "endsWith": return String(value).toLowerCase().endsWith(String(first ?? "").toLowerCase());
      default: return false;
    }
  }

  scaleColor(rule, index, value, range) {
    if (typeof value !== "number") return null;
    let cache = this.cfCache.get(index);
    if (!cache) {
      let min = Infinity;
      let max = -Infinity;
      const limit = Math.min(rangeSize(range), 200000);
      let count = 0;
      for (let row = range.r1; row <= range.r2 && count < limit; row += 1) {
        for (let col = range.c1; col <= range.c2 && count < limit; col += 1) {
          const cell = cellAt(this.sheet(), row, col);
          if (cell && typeof cell.v === "number") {
            min = Math.min(min, cell.v);
            max = Math.max(max, cell.v);
          }
          count += 1;
        }
      }
      cache = { min, max };
      this.cfCache.set(index, cache);
    }
    if (!Number.isFinite(cache.min) || cache.max <= cache.min) return null;
    const colors = (rule.colors || ["#F8696B", "#FFEB84", "#63BE7B"]).filter(Boolean);
    if (!colors.length) return null;
    const amount = clamp((value - cache.min) / (cache.max - cache.min), 0, 1);
    if (colors.length === 1) return colors[0];
    if (colors.length === 2) return colorMix(colors[0], colors[1], amount);
    if (amount < 0.5) return colorMix(colors[0], colors[1], amount * 2);
    return colorMix(colors[1], colors[2], (amount - 0.5) * 2);
  }
}
