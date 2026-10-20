/* Grid behaviour: selection, pointer and keyboard input, editing, clipboard,
 * fill series and the small structural commands used by the ribbon.  The
 * drawing core lives in grid-base.js. */

import { el, clamp, toast, openMenu, closeMenus, showDialog, confirmDialog } from "./ui.js";
import { icon } from "./icons.js";
import { formatCell, parseUserInput, resolveFormat, isDateFormat } from "./format.js";
import {
  DEFAULT_COL_W,
  DEFAULT_ROW_H,
  MIN_COL_W,
  MAX_COL_W,
  MIN_ROW_H,
  MAX_ROW_H,
  cellAt,
  ensureCell,
  clearCell,
  colDisplayWidth,
  rowDisplayHeight,
  setColWidth,
  setRowHeight,
  setColHidden,
  setRowHidden,
  isColHidden,
  isRowHidden,
  addMerge,
  removeMergeAt,
  usedRange,
  insertRows,
  deleteRows,
  insertCols,
  deleteCols,
  addComment,
  deleteCellComments,
  sortRange,
  removeDuplicateRows,
} from "./model.js";
import {
  MAX_ROWS,
  MAX_COLS,
  addr,
  parseAddr,
  parseRange,
  rangeText,
  normalizeRange,
  rangeContains,
  rangeSize,
  rangesIntersect,
  indexToCol,
  colToIndex,
} from "./units.js";
import { translateFormula, cellKey } from "./formula.js";
import { GridBase, GRID_FONT_STACK, rawEditText, chartSvg } from "./grid-base.js";

function escapeField(text) {
  const value = String(text ?? "");
  if (/[\t\n\r"]/.test(value)) return '"' + value.replace(/"/g, '""') + '"';
  return value;
}

function parseDelimited(text) {
  const rows = [];
  let row = [];
  let field = "";
  let quoted = false;
  const source = String(text ?? "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  for (let index = 0; index < source.length; index += 1) {
    const char = source[index];
    if (quoted) {
      if (char === '"') {
        if (source[index + 1] === '"') {
          field += '"';
          index += 1;
        } else {
          quoted = false;
        }
      } else {
        field += char;
      }
      continue;
    }
    if (char === '"') quoted = true;
    else if (char === "\t") {
      row.push(field);
      field = "";
    } else if (char === "\n") {
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else {
      field += char;
    }
  }
  if (field.length || row.length) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

function cloneStyle(style) {
  if (!style) return null;
  const clone = JSON.parse(JSON.stringify(style));
  return Object.keys(clone).length ? clone : null;
}

function cellHasData(cell) {
  return !!cell && ((cell.v !== undefined && cell.v !== null && cell.v !== "") || !!cell.f);
}

export class Grid extends GridBase {
  /* ------------------------------------------------------------ selection */
  activeAddress() {
    return addr(this.active.row, this.active.col);
  }

  effectiveRange() {
    const used = usedRange(this.sheet());
    return {
      r1: this.sel.r1,
      c1: this.sel.c1,
      r2: Math.min(this.sel.r2, used ? Math.max(used.r2, this.sel.r1) : this.sel.r1),
      c2: Math.min(this.sel.c2, used ? Math.max(used.c2, this.sel.c1) : this.sel.c1),
    };
  }

  setSelection(r1, c1, r2, c2, { activeRow = null, activeCol = null, silent = false, reveal = false, extend = false } = {}) {
    const range = normalizeRange(r1, c1, r2, c2);
    this.sel = range;
    if (activeRow === null || activeCol === null) {
      if (this.active.row < range.r1 || this.active.row > range.r2 || this.active.col < range.c1 || this.active.col > range.c2) {
        this.active = { row: range.r1, col: range.c1 };
      }
    } else {
      this.active = {
        row: clamp(activeRow, range.r1, range.r2),
        col: clamp(activeCol, range.c1, range.c2),
      };
    }
    if (!extend) this.selectionAnchor = { ...this.active };
    if (reveal) this.ensureVisible(this.active.row, this.active.col);
    if (!silent) {
      this.notifySelection();
      this.scheduleRender();
    }
  }

  selectRange(range, options = {}) {
    if (!range) return;
    this.setSelection(range.r1, range.c1, range.r2, range.c2, {
      activeRow: range.r1,
      activeCol: range.c1,
      ...options,
    });
  }

  selectCell(row, col, { extend = false } = {}) {
    row = clamp(row, 1, MAX_ROWS);
    col = clamp(col, 1, MAX_COLS);
    const merge = this.mergeFor(row, col);
    if (merge) {
      this.setSelection(merge.r1, merge.c1, merge.r2, merge.c2, {
        activeRow: merge.r1,
        activeCol: merge.c1,
        extend,
      });
      return;
    }
    if (extend) {
      const anchor = this.selectionAnchor || { ...this.active };
      this.setSelection(anchor.row, anchor.col, row, col, {
        activeRow: anchor.row,
        activeCol: anchor.col,
        extend: true,
      });
    } else {
      this.setSelection(row, col, row, col, { activeRow: row, activeCol: col });
    }
  }

  selectAll() {
    const used = usedRange(this.sheet());
    const range = used
      ? { r1: 1, c1: 1, r2: used.r2, c2: used.c2 }
      : { r1: 1, c1: 1, r2: 1, c2: 1 };
    this.selectRange(range);
    this.els?.scroller.focus({ preventScroll: true });
  }

  selectColumn(col, extend = false) {
    col = clamp(col, 1, MAX_COLS);
    if (extend && this.selectionAnchor) {
      const anchor = this.selectionAnchor;
      this.setSelection(1, Math.min(anchor.col, col), MAX_ROWS, Math.max(anchor.col, col), {
        activeRow: 1,
        activeCol: anchor.col,
        extend: true,
      });
    } else {
      this.setSelection(1, col, MAX_ROWS, col, { activeRow: 1, activeCol: col });
    }
  }

  selectRow(row, extend = false) {
    row = clamp(row, 1, MAX_ROWS);
    if (extend && this.selectionAnchor) {
      const anchor = this.selectionAnchor;
      this.setSelection(Math.min(anchor.row, row), 1, Math.max(anchor.row, row), MAX_COLS, {
        activeRow: anchor.row,
        activeCol: 1,
        extend: true,
      });
    } else {
      this.setSelection(row, 1, row, MAX_COLS, { activeRow: row, activeCol: 1 });
    }
  }

  isSelected(row, col) {
    return row >= this.sel.r1 && row <= this.sel.r2 && col >= this.sel.c1 && col <= this.sel.c2;
  }

  selectionInfo() {
    const sheet = this.sheet();
    const cell = cellAt(sheet, this.active.row, this.active.col);
    return {
      grid: this,
      sheet: sheet.name,
      sheetIndex: this.sheetIndex,
      range: { ...this.sel },
      text: rangeText(this.sel),
      address: this.activeAddress(),
      active: { ...this.active },
      cell: cell ? { ...cell } : null,
      style: cell?.s ? { ...cell.s } : {},
      editing: !!this.editing,
      stats: this.selectionStats(),
    };
  }

  notifySelection() {
    this.options.onSelection?.(this.selectionInfo());
  }

  selectionStats() {
    const sheet = this.sheet();
    const range = this.sel;
    let count = 0;
    let numeric = 0;
    let text = 0;
    let sum = 0;
    let min = Infinity;
    let max = -Infinity;
    for (const [address, cell] of Object.entries(sheet.cells || {})) {
      const parsed = parseAddr(address);
      if (!parsed || !rangeContains(range, parsed.row, parsed.col)) continue;
      if (cell.v === undefined || cell.v === null || cell.v === "") continue;
      count += 1;
      if (typeof cell.v === "number" && Number.isFinite(cell.v)) {
        numeric += 1;
        sum += cell.v;
        if (cell.v < min) min = cell.v;
        if (cell.v > max) max = cell.v;
      } else if (typeof cell.v === "string") {
        text += 1;
      }
    }
    return {
      count,
      numeric,
      text,
      sum,
      average: numeric ? sum / numeric : 0,
      min: numeric ? min : 0,
      max: numeric ? max : 0,
    };
  }

  move(dRow, dCol, { extend = false, jump = false } = {}) {
    let row = this.active.row;
    let col = this.active.col;
    if (jump) {
      const target = this.findEdge(row, col, Math.sign(dRow), Math.sign(dCol));
      row = target.row;
      col = target.col;
    } else {
      row = clamp(row + dRow, 1, MAX_ROWS);
      col = clamp(col + dCol, 1, MAX_COLS);
    }
    const target = this.mergeFor(row, col);
    if (target) {
      row = target.r1;
      col = target.c1;
    }
    if (extend) {
      const anchor = this.selectionAnchor || { ...this.active };
      this.setSelection(anchor.row, anchor.col, row, col, {
        activeRow: anchor.row,
        activeCol: anchor.col,
        extend: true,
      });
    } else {
      this.selectCell(row, col);
    }
    this.ensureVisible(row, col);
  }

  findEdge(row, col, dRow, dCol) {
    const sheet = this.sheet();
    const has = (r, c) => {
      if (r < 1 || r > MAX_ROWS || c < 1 || c > MAX_COLS) return false;
      const cell = cellAt(sheet, r, c);
      return !!cell && ((cell.v !== undefined && cell.v !== null && cell.v !== "") || !!cell.f || !!cell.h);
    };
    let current = { row, col };
    if (has(row, col)) {
      while (true) {
        const next = { row: current.row + dRow, col: current.col + dCol };
        if (next.row < 1 || next.row > MAX_ROWS || next.col < 1 || next.col > MAX_COLS) break;
        if (!has(next.row, next.col)) break;
        current = next;
      }
      return current;
    }
    while (true) {
      const next = { row: current.row + dRow, col: current.col + dCol };
      if (next.row < 1 || next.row > MAX_ROWS || next.col < 1 || next.col > MAX_COLS) break;
      current = next;
      if (has(current.row, current.col)) break;
    }
    return current;
  }

  /* ------------------------------------------------------------- selection UI */
  renderSelection() {
    const layer = this.els.selection;
    layer.replaceChildren();
    const segmentsX = [];
    if (this.sel.c1 <= this.freezeCols) {
      segmentsX.push([this.sel.c1, Math.min(this.sel.c2, this.freezeCols), true]);
    }
    if (this.sel.c2 > this.freezeCols) {
      segmentsX.push([Math.max(this.sel.c1, this.freezeCols + 1), this.sel.c2, false]);
    }
    const segmentsY = [];
    if (this.sel.r1 <= this.freezeRows) {
      segmentsY.push([this.sel.r1, Math.min(this.sel.r2, this.freezeRows), true]);
    }
    if (this.sel.r2 > this.freezeRows) {
      segmentsY.push([Math.max(this.sel.r1, this.freezeRows + 1), this.sel.r2, false]);
    }
    for (const [r1, r2, frozenRow] of segmentsY) {
      for (const [c1, c2, frozenCol] of segmentsX) {
        const left = this.colPrefix[c1 - 1] - (frozenCol ? 0 : this.scrollLeft);
        const top = this.rowPrefix[r1 - 1] - (frozenRow ? 0 : this.scrollTop);
        const width = this.colPrefix[c2] - this.colPrefix[c1 - 1];
        const height = this.rowPrefix[r2] - this.rowPrefix[r1 - 1];
        if (left >= this.viewportWidth || top >= this.viewportHeight) continue;
        if (left + width <= 0 || top + height <= 0) continue;
        const fill = el("div", { class: "sel-fill", i18n: false });
        fill.style.cssText = `left:${Math.round(left)}px;top:${Math.round(top)}px;width:${width}px;height:${height}px;`;
        layer.append(fill);
        const outline = el("div", { class: "sel-rect active", i18n: false });
        outline.style.cssText = `left:${Math.round(left)}px;top:${Math.round(top)}px;width:${width}px;height:${height}px;`
          + `border-top-width:${r1 === this.sel.r1 ? 2 : 0}px;border-bottom-width:${r2 === this.sel.r2 ? 2 : 0}px;`
          + `border-left-width:${c1 === this.sel.c1 ? 2 : 0}px;border-right-width:${c2 === this.sel.c2 ? 2 : 0}px;`;
        layer.append(outline);
      }
    }
    if (!this.editing) {
      const merge = this.mergeFor(this.active.row, this.active.col);
      const rect = this.cellRect(this.active.row, this.active.col, merge);
      if (rect.left < this.viewportWidth && rect.top < this.viewportHeight && rect.left + rect.width > 0 && rect.top + rect.height > 0) {
        const node = el("div", { class: "sel-active-cell", i18n: false });
        node.style.cssText = `left:${Math.round(rect.left)}px;top:${Math.round(rect.top)}px;`
          + `width:${rect.width}px;height:${rect.height}px;`;
        layer.append(node);
      }
    }
    this.updateFillHandle();
  }

  updateFillHandle() {
    const handle = this.els.fill;
    if (!handle) return;
    if (this.editing || this.sel.r2 >= MAX_ROWS || this.sel.c2 >= MAX_COLS) {
      handle.classList.add("hidden");
      return;
    }
    const right = this.colPrefix[this.sel.c2] - (this.sel.c2 > this.freezeCols ? this.scrollLeft : 0);
    const bottom = this.rowPrefix[this.sel.r2] - (this.sel.r2 > this.freezeRows ? this.scrollTop : 0);
    const left = right - 3;
    const top = bottom - 3;
    if (left < 0 || top < 0 || left > this.viewportWidth || top > this.viewportHeight) {
      handle.classList.add("hidden");
      return;
    }
    handle.classList.remove("hidden");
    handle.style.left = Math.round(left) + "px";
    handle.style.top = Math.round(top) + "px";
  }

  /* ------------------------------------------------------- comments and drawings */
  renderComments() {
    const layer = this.els.pins;
    if (!layer) return;
    const used = new Set();
    for (const comment of this.sheet().comments || []) {
      const parsed = parseAddr(comment.addr);
      if (!parsed) continue;
      const rect = this.cellRect(parsed.row, parsed.col, this.mergeFor(parsed.row, parsed.col));
      if (rect.left > this.viewportWidth || rect.top > this.viewportHeight) continue;
      if (rect.left + rect.width < 0 || rect.top + rect.height < 0) continue;
      used.add(comment.addr);
      let node = this.commentNodes.get(comment.addr);
      if (!node) {
        node = el("div", { class: "comment-pin", i18n: false });
        node.addEventListener("mousedown", (event) => event.stopPropagation());
        node.addEventListener("click", (event) => {
          event.stopPropagation();
          this.options.onCommentRequest?.(comment.addr);
        });
        node.addEventListener("contextmenu", (event) => {
          event.preventDefault();
          event.stopPropagation();
          this.options.onCommentRequest?.(comment.addr);
        });
        layer.append(node);
        this.commentNodes.set(comment.addr, node);
      }
      const author = (comment.author || "?").trim();
      node.textContent = author.charAt(0).toUpperCase() || "?";
      node.title = author ? `${author}: ${comment.text || ""}` : comment.text || "";
      node.style.left = Math.round(clamp(rect.left + rect.width - 18, 0, Math.max(0, this.viewportWidth - 18))) + "px";
      node.style.top = Math.round(clamp(rect.top + 1, 0, Math.max(0, this.viewportHeight - 18))) + "px";
    }
    for (const [address, node] of this.commentNodes) {
      if (!used.has(address)) {
        node.remove();
        this.commentNodes.delete(address);
      }
    }
  }

  renderOverlays() {
    const layer = this.els.overlay;
    if (!layer) return;
    const sheet = this.sheet();
    const used = new Set();
    const scale = this.scale;
    const place = (item, node) => {
      const left = (Number(item.x) || 0) * scale - this.scrollLeft;
      const top = (Number(item.y) || 0) * scale - this.scrollTop;
      const width = Math.max(24, (Number(item.w) || 160) * scale);
      const height = Math.max(24, (Number(item.h) || 120) * scale);
      if (left + width < 0 || top + height < 0 || left > this.viewportWidth || top > this.viewportHeight) return;
      node.style.left = Math.round(left) + "px";
      node.style.top = Math.round(top) + "px";
      node.style.width = Math.round(width) + "px";
      node.style.height = Math.round(height) + "px";
      node.classList.toggle("selected", this.selectedOverlay === item.id);
      layer.append(node);
      used.add(item.id);
    };
    for (const chart of sheet.charts || []) {
      let node = this.overlayNodes.get(chart.id);
      if (!node) {
        node = el("div", { class: "sheet-chart", i18n: false });
        node.append(el("div", { class: "chart-title", i18n: false }));
        node.append(el("div", { class: "chart-body", i18n: false }));
        node.append(el("div", { class: "chart-handle", i18n: false }));
        this.bindOverlay(node, chart, "chart");
        this.overlayNodes.set(chart.id, node);
      }
      const title = node.querySelector(".chart-title");
      title.textContent = chart.title || "";
      title.style.display = chart.title ? "" : "none";
      const body = node.querySelector(".chart-body");
      const signature = `${chart.range}|${chart.type}|${chart.title || ""}|${this.epoch}`;
      if (body.dataset.signature !== signature) {
        body.innerHTML = chartSvg(chart, sheet);
        body.dataset.signature = signature;
      }
      place(chart, node);
    }
    for (const image of sheet.images || []) {
      let node = this.overlayNodes.get(image.id);
      if (!node) {
        node = el("div", { class: "sheet-image", i18n: false });
        node.append(el("img", { i18n: false, alt: "" }));
        node.append(el("div", { class: "chart-handle", i18n: false }));
        this.bindOverlay(node, image, "image");
        this.overlayNodes.set(image.id, node);
      }
      const img = node.querySelector("img");
      if (img.getAttribute("src") !== image.data) img.setAttribute("src", image.data || "");
      place(image, node);
    }
    for (const box of sheet.textBoxes || []) {
      let node = this.overlayNodes.get(box.id);
      if (!node) {
        node = el("div", { class: "sheet-textbox", i18n: false });
        node.append(el("div", { class: "textbox-body", i18n: false }));
        node.append(el("div", { class: "chart-handle", i18n: false }));
        this.bindOverlay(node, box, "text");
        this.overlayNodes.set(box.id, node);
      }
      const body = node.querySelector(".textbox-body");
      const label = String(box.text ?? "");
      if (body.textContent !== label) body.textContent = label;
      node.style.fontSize = `${Math.max(8, Math.round((Number(box.size) || 12) * this.scale))}px`;
      node.style.fontWeight = box.bold ? "700" : "";
      node.style.fontStyle = box.italic ? "italic" : "";
      node.style.color = box.color || "";
      place(box, node);
    }
    for (const [id, node] of this.overlayNodes) {
      if (!used.has(id)) {
        node.remove();
        this.overlayNodes.delete(id);
      }
    }
    this.renderPageBreakLines(layer);
    this.renderAuditArrows(layer);
  }

  renderPageBreakLines(layer) {
    for (const node of [...layer.querySelectorAll(".page-break-line")]) node.remove();
    const rows = this.pageBreaks?.rows || [];
    const cols = this.pageBreaks?.cols || [];
    for (const row of rows) {
      if (!Number.isFinite(row) || row < 1) continue;
      const top = this.screenTop(row);
      if (top < 0 || top > this.viewportHeight) continue;
      const line = el("div", { class: "page-break-line horizontal", i18n: false });
      line.style.top = Math.round(top) + "px";
      line.style.left = "0px";
      line.style.width = this.viewportWidth + "px";
      layer.append(line);
    }
    for (const col of cols) {
      if (!Number.isFinite(col) || col < 1) continue;
      const left = this.screenLeft(col);
      if (left < 0 || left > this.viewportWidth) continue;
      const line = el("div", { class: "page-break-line vertical", i18n: false });
      line.style.left = Math.round(left) + "px";
      line.style.top = "0px";
      line.style.height = this.viewportHeight + "px";
      layer.append(line);
    }
  }

  renderAuditArrows(layer) {
    for (const node of [...layer.querySelectorAll(".audit-arrows")]) node.remove();
    const links = this.auditLinks || [];
    if (!links.length) return;
    const sheetName = this.sheet().name.toLowerCase();
    const viewportWidth = this.viewportWidth;
    const viewportHeight = this.viewportHeight;
    const parts = [];
    for (const link of links) {
      if (link.sheet && String(link.sheet).toLowerCase() !== sheetName) continue;
      const from = this.cellRect(link.row1, link.col1);
      const to = this.cellRect(link.row2, link.col2);
      const x1 = from.left + from.width / 2;
      const y1 = from.top + from.height / 2;
      const x2 = to.left + to.width / 2;
      const y2 = to.top + to.height / 2;
      if (Math.max(x1, x2) < 0 || Math.min(x1, x2) > viewportWidth) continue;
      if (Math.max(y1, y2) < 0 || Math.min(y1, y2) > viewportHeight) continue;
      parts.push(`<line x1="${x1.toFixed(1)}" y1="${y1.toFixed(1)}" x2="${x2.toFixed(1)}" y2="${y2.toFixed(1)}" marker-end="url(#audit-arrow-head)"/>`);
    }
    if (!parts.length) return;
    const svg = el("svg", { class: "audit-arrows", i18n: false });
    svg.setAttribute("width", String(viewportWidth));
    svg.setAttribute("height", String(viewportHeight));
    svg.innerHTML = '<defs><marker id="audit-arrow-head" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="6" markerHeight="6" orient="auto-start-reverse"><path d="M 0 0 L 10 5 L 0 10 z" fill="#1a73e8"/></marker></defs>'
      + '<g fill="none" stroke="#1a73e8" stroke-width="1.4">' + parts.join("") + "</g>";
    layer.append(svg);
  }

  bindOverlay(node, item, kind) {
    node.addEventListener("mousedown", (event) => {
      if (event.button !== 0) return;
      event.preventDefault();
      event.stopPropagation();
      this.selectedOverlay = item.id;
      if (event.target.classList.contains("chart-handle")) {
        this.drag = {
          type: kind + "resize", item,
          startX: event.clientX, startY: event.clientY,
          startW: Number(item.w) || 160, startH: Number(item.h) || 120,
          before: this.options.captureState?.() ?? null,
        };
      } else {
        this.drag = {
          type: kind + "move", item,
          startX: event.clientX, startY: event.clientY,
          startL: Number(item.x) || 0, startT: Number(item.y) || 0,
          before: this.options.captureState?.() ?? null,
        };
      }
      this.renderOverlays();
      this.els.scroller.focus({ preventScroll: true });
    });
    node.addEventListener("dblclick", (event) => {
      event.preventDefault();
      event.stopPropagation();
      this.options.onOverlayActivate?.(item, kind);
    });
    node.addEventListener("contextmenu", (event) => {
      event.preventDefault();
      event.stopPropagation();
      this.selectedOverlay = item.id;
      this.renderOverlays();
      this.options.onOverlayContextMenu?.(event.clientX, event.clientY, item, kind);
    });
  }

  /* ---------------------------------------------------------------- pointer */
  pointerCell(x, y) {
    const rect = this.els.viewport.getBoundingClientRect();
    return this.hitTest(x - rect.left, y - rect.top);
  }

  handleMouseDown(event) {
    if (event.button === 2) return;
    if (event.button !== 0) return;
    if (this.editing) this.commitEdit("none");
    this.selectedOverlay = null;
    const scroller = this.els.scroller;
    const rect = this.els.viewport.getBoundingClientRect();
    const x = event.clientX - rect.left;
    const y = event.clientY - rect.top;
    if (x < 0 || y < 0 || x > scroller.clientWidth || y > scroller.clientHeight) return;
    closeMenus();
    event.preventDefault();
    scroller.focus({ preventScroll: true });
    const hit = this.hitTest(x, y);
    if (event.shiftKey) {
      const anchor = this.selectionAnchor || { ...this.active };
      this.setSelection(anchor.row, anchor.col, hit.row, hit.col, {
        activeRow: anchor.row,
        activeCol: anchor.col,
        extend: true,
      });
    } else {
      this.selectCell(hit.row, hit.col);
    }
    this.drag = {
      type: "cells",
      anchorRow: this.active.row,
      anchorCol: this.active.col,
      lastRow: hit.row,
      lastCol: hit.col,
    };
    this.lastPointer = { x: event.clientX, y: event.clientY };
    this.startAutoScroll();
  }

  handleMouseMove(event) {
    if (!this.drag) return;
    this.lastPointer = { x: event.clientX, y: event.clientY };
    const drag = this.drag;
    if (drag.type === "colresize") {
      const delta = (event.clientX - drag.startX) / this.scale;
      const width = clamp(drag.startWidth + delta, MIN_COL_W, MAX_COL_W);
      setColWidth(this.sheet(), drag.col, width);
      this.buildLayout();
      this.renderNow();
      this.options.onStatus?.(`Width: ${Math.round(width)} px`);
      return;
    }
    if (drag.type === "rowresize") {
      const delta = (event.clientY - drag.startY) / this.scale;
      const height = clamp(drag.startHeight + delta, MIN_ROW_H, MAX_ROW_H);
      setRowHeight(this.sheet(), drag.row, height);
      this.buildLayout();
      this.renderNow();
      this.options.onStatus?.(`Height: ${Math.round(height)} px`);
      return;
    }
    if (drag.type === "cols") {
      const rect = this.els.colHeaders.getBoundingClientRect();
      const x = clamp(event.clientX - rect.left, 0, Math.max(0, rect.width - 1));
      const col = this.findVisible(this.colPrefix, x + (x < this.freezeWidth ? 0 : this.scrollLeft), MAX_COLS);
      this.selectColumn(col, true);
      return;
    }
    if (drag.type === "rows") {
      const rect = this.els.rowHeaders.getBoundingClientRect();
      const y = clamp(event.clientY - rect.top, 0, Math.max(0, rect.height - 1));
      const row = this.findVisible(this.rowPrefix, y + (y < this.freezeHeight ? 0 : this.scrollTop), MAX_ROWS);
      this.selectRow(row, true);
      return;
    }
    if (drag.type === "fill") {
      this.updateFillPreview(event);
      return;
    }
    if (drag.type === "chartmove" || drag.type === "imagemove" || drag.type === "textmove") {
      drag.item.x = Math.max(0, Math.round(drag.startL + (event.clientX - drag.startX) / this.scale));
      drag.item.y = Math.max(0, Math.round(drag.startT + (event.clientY - drag.startY) / this.scale));
      this.renderOverlays();
      return;
    }
    if (drag.type === "chartresize" || drag.type === "imageresize" || drag.type === "textresize") {
      drag.item.w = Math.max(40, Math.round(drag.startW + (event.clientX - drag.startX) / this.scale));
      drag.item.h = Math.max(30, Math.round(drag.startH + (event.clientY - drag.startY) / this.scale));
      this.renderOverlays();
      return;
    }
    if (drag.type === "cells") {
      const hit = this.pointerCell(event.clientX, event.clientY);
      const row = clamp(hit.row, 1, MAX_ROWS);
      const col = clamp(hit.col, 1, MAX_COLS);
      if (row === drag.lastRow && col === drag.lastCol) return;
      drag.lastRow = row;
      drag.lastCol = col;
      const anchor = this.selectionAnchor || { row: drag.anchorRow, col: drag.anchorCol };
      this.setSelection(anchor.row, anchor.col, row, col, {
        activeRow: anchor.row,
        activeCol: anchor.col,
        extend: true,
      });
    }
  }

  handleMouseUp() {
    if (!this.drag) return;
    const drag = this.drag;
    this.drag = null;
    this.stopAutoScroll();
    document.body.style.cursor = "";
    if (drag.type === "fill") {
      this.finishFillDrag();
      return;
    }
    if (drag.type === "colresize" || drag.type === "rowresize") {
      this.options.onStatus?.("");
      this.afterMutation("resize", drag.before, { layout: false });
      return;
    }
    if (drag.type === "chartmove" || drag.type === "imagemove" || drag.type === "chartresize"
      || drag.type === "imageresize" || drag.type === "textmove" || drag.type === "textresize") {
      this.afterMutation("drawing", drag.before, { layout: false });
    }
  }

  startAutoScroll() {
    if (this.autoScroll) return;
    const tick = () => {
      this.autoScroll = null;
      if (!this.drag || !this.lastPointer) return;
      const rect = this.els.viewport.getBoundingClientRect();
      const x = this.lastPointer.x - rect.left;
      const y = this.lastPointer.y - rect.top;
      const margin = 26;
      let dx = 0;
      let dy = 0;
      if (x < margin) dx = -1;
      else if (x > this.els.scroller.clientWidth - margin) dx = 1;
      if (y < margin) dy = -1;
      else if (y > this.els.scroller.clientHeight - margin) dy = 1;
      if (!dx && !dy) return;
      this.els.scroller.scrollLeft = Math.max(0, this.els.scroller.scrollLeft + dx * 18);
      this.els.scroller.scrollTop = Math.max(0, this.els.scroller.scrollTop + dy * 18);
      this.scrollLeft = this.els.scroller.scrollLeft;
      this.scrollTop = this.els.scroller.scrollTop;
      this.scheduleRender();
      if (this.drag.type === "cells" || this.drag.type === "fill") {
        const hit = this.pointerCell(this.lastPointer.x, this.lastPointer.y);
        if (this.drag.type === "cells") {
          const anchor = this.selectionAnchor || { row: this.drag.anchorRow, col: this.drag.anchorCol };
          this.setSelection(anchor.row, anchor.col, hit.row, hit.col, {
            activeRow: anchor.row,
            activeCol: anchor.col,
            extend: true,
            silent: false,
          });
        } else {
          this.updateFillPreviewFromCell(hit.row, hit.col);
        }
      }
      this.autoScroll = requestAnimationFrame(tick);
    };
    this.autoScroll = requestAnimationFrame(tick);
  }

  stopAutoScroll() {
    if (this.autoScroll) cancelAnimationFrame(this.autoScroll);
    this.autoScroll = null;
  }

  handleDoubleClick(event) {
    const scroller = this.els.scroller;
    const rect = this.els.viewport.getBoundingClientRect();
    const x = event.clientX - rect.left;
    const y = event.clientY - rect.top;
    if (x < 0 || y < 0 || x > scroller.clientWidth || y > scroller.clientHeight) return;
    const hit = this.hitTest(x, y);
    const cell = cellAt(this.sheet(), hit.row, hit.col);
    if (cell?.h) {
      this.options.onHyperlink?.(cell.h, cell);
      return;
    }
    this.beginEdit(null, { selectAll: false });
  }

  handleContextMenu(event) {
    const scroller = this.els.scroller;
    const rect = this.els.viewport.getBoundingClientRect();
    const x = event.clientX - rect.left;
    const y = event.clientY - rect.top;
    if (x < 0 || y < 0 || x > scroller.clientWidth || y > scroller.clientHeight) return;
    event.preventDefault();
    const hit = this.hitTest(x, y);
    if (!this.isSelected(hit.row, hit.col)) this.selectCell(hit.row, hit.col);
    this.options.onContextMenu?.(event.clientX, event.clientY, this);
  }

  /* ---------------------------------------------------------------- headers */
  handleHeaderMouseDown(event, axis) {
    if (event.button !== 0) return;
    const target = event.target;
    if (target.closest(".resize-handle")) {
      event.preventDefault();
      const holder = target.closest(".colhead, .rowhead");
      if (!holder) return;
      if (axis === "col") this.beginColResize(Number(holder.dataset.col), event);
      else this.beginRowResize(Number(holder.dataset.row), event);
      return;
    }
    const holder = target.closest(axis === "col" ? ".colhead" : ".rowhead");
    if (!holder) return;
    event.preventDefault();
    this.els.scroller.focus({ preventScroll: true });
    if (axis === "col") {
      const col = Number(holder.dataset.col);
      this.selectColumn(col, event.shiftKey);
      this.drag = { type: "cols" };
    } else {
      const row = Number(holder.dataset.row);
      this.selectRow(row, event.shiftKey);
      this.drag = { type: "rows" };
    }
  }

  handleHeaderDoubleClick(event, axis) {
    const target = event.target;
    if (!target.closest(".resize-handle")) return;
    const holder = target.closest(".colhead, .rowhead");
    if (!holder) return;
    event.preventDefault();
    event.stopPropagation();
    if (axis === "col") this.autoFitColumn(Number(holder.dataset.col));
    else this.autoFitRow(Number(holder.dataset.row));
  }

  beginColResize(col, event) {
    if (!col) return;
    const before = this.options.captureState?.() ?? null;
    const width = colDisplayWidth(this.sheet(), col) || DEFAULT_COL_W;
    this.drag = { type: "colresize", col, startX: event.clientX, startWidth: width, before };
    document.body.style.cursor = "col-resize";
  }

  beginRowResize(row, event) {
    if (!row) return;
    const before = this.options.captureState?.() ?? null;
    const height = rowDisplayHeight(this.sheet(), row) || DEFAULT_ROW_H;
    this.drag = { type: "rowresize", row, startY: event.clientY, startHeight: height, before };
    document.body.style.cursor = "row-resize";
  }

  measureText(text, style, pixelSize) {
    if (!this.measureCanvas) {
      this.measureCanvas = document.createElement("canvas");
      this.measureContext = this.measureCanvas.getContext("2d");
    }
    const context = this.measureContext;
    context.font = `${style.i ? "italic " : ""}${style.b ? "700 " : "400 "}${pixelSize}px ${style.font ? '"' + style.font + '",' : ""}${GRID_FONT_STACK}`;
    return context.measureText(String(text ?? "")).width;
  }

  autoFitColumn(col) {
    const sheet = this.sheet();
    const used = usedRange(sheet);
    if (!used) return;
    const before = this.options.captureState?.() ?? null;
    let width = 25;
    const limit = Math.min(used.r2, 4000);
    for (let row = used.r1; row <= limit; row += 1) {
      const cell = cellAt(sheet, row, col);
      if (!cell || cell.v === undefined) continue;
      const style = cell.s || {};
      const points = Number(style.size) || 11;
      const pixelSize = points * (4 / 3);
      const formatted = formatCell(cell);
      width = Math.max(width, this.measureText(formatted.text, style, pixelSize) + 12 + (Number(style.indent) || 0) * 8);
    }
    const value = clamp(Math.round(width / this.scale), MIN_COL_W, MAX_COL_W);
    setColWidth(sheet, col, value);
    this.buildLayout();
    this.renderNow();
    this.afterMutation("auto-fit", before, { layout: false });
  }

  autoFitRow(row) {
    const sheet = this.sheet();
    const used = usedRange(sheet);
    if (!used) return;
    const before = this.options.captureState?.() ?? null;
    let height = 20;
    for (let col = used.c1; col <= Math.min(used.c2, 60); col += 1) {
      const cell = cellAt(sheet, row, col);
      if (!cell || cell.v === undefined) continue;
      const style = cell.s || {};
      const points = Number(style.size) || 11;
      const pixelSize = points * (4 / 3);
      const formatted = formatCell(cell);
      if (!style.wrap) {
        height = Math.max(height, DEFAULT_ROW_H);
        continue;
      }
      const available = Math.max(20, colDisplayWidth(sheet, col) - 10);
      const words = formatted.text.split(/\s+/);
      let lines = 1;
      let line = "";
      for (const word of words) {
        const candidate = line ? line + " " + word : word;
        if (this.measureText(candidate, style, pixelSize) > available && line) {
          lines += 1;
          line = word;
        } else {
          line = candidate;
        }
      }
      height = Math.max(height, lines * pixelSize * 1.32 + 4);
    }
    const value = clamp(Math.round(height / this.scale), MIN_ROW_H, MAX_ROW_H);
    setRowHeight(sheet, row, value);
    this.buildLayout();
    this.renderNow();
    this.afterMutation("auto-fit", before, { layout: false });
  }

  /* -------------------------------------------------------------- keyboard */
  handleKeyDown(event) {
    const tag = event.target?.tagName;
    if (tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT" || event.target?.isContentEditable) return;
    if (this.editing) return;
    const ctrl = event.ctrlKey || event.metaKey;
    const shift = event.shiftKey;
    const key = event.key;
    let handled = true;
    switch (key) {
      case "ArrowUp":
        this.move(-1, 0, { extend: shift, jump: ctrl });
        break;
      case "ArrowDown":
        this.move(1, 0, { extend: shift, jump: ctrl });
        break;
      case "ArrowLeft":
        this.move(0, -1, { extend: shift, jump: ctrl });
        break;
      case "ArrowRight":
        this.move(0, 1, { extend: shift, jump: ctrl });
        break;
      case "Tab":
        this.move(0, shift ? -1 : 1, { extend: false });
        break;
      case "Enter":
        this.move(shift ? -1 : 1, 0, { extend: false });
        break;
      case "Home":
        if (ctrl) this.selectCell(1, 1);
        else this.move(0, -this.active.col + 1, { extend: shift });
        break;
      case "End": {
        const used = usedRange(this.sheet());
        if (ctrl) this.selectCell(used ? used.r2 : 1, used ? used.c2 : 1);
        else this.move(0, (used ? used.c2 : 1) - this.active.col, { extend: shift });
        break;
      }
      case "PageUp":
        this.move(-Math.max(1, Math.floor((this.viewportHeight || 400) / 22)), 0, { extend: shift });
        break;
      case "PageDown":
        this.move(Math.max(1, Math.floor((this.viewportHeight || 400) / 22)), 0, { extend: shift });
        break;
      case "Delete":
        this.clearSelection("contents");
        break;
      case "Backspace":
        this.clearSelection("contents");
        this.beginEdit("", { selectAll: false });
        break;
      case "F2":
        this.beginEdit(null, { selectAll: false });
        break;
      case "Escape":
        this.options.onEscape?.();
        break;
      case " ": {
        if (ctrl || event.altKey) {
          handled = false;
          break;
        }
        event.preventDefault();
        this.beginEdit(" ", { selectAll: false });
        return;
      }
      default:
        if (ctrl && !event.altKey) {
          const lower = key.toLowerCase();
          if (lower === "d") this.fillSelection("down");
          else if (lower === "r") this.fillSelection("right");
          else if (lower === "z") this.options.onUndo?.(shift);
          else if (lower === "y") this.options.onRedo?.();
          else if (lower === "f") this.options.onFind?.();
          else if (lower === "g") this.options.onGoTo?.();
          else if (lower === "1") this.options.onShortcut?.("format-cells");
          else if (lower === "b" || lower === "i" || lower === "u") this.options.onShortcut?.("format-" + lower);
          else if (lower === "k") this.options.onShortcut?.("hyperlink");
          else if (lower === ";") this.insertDate(shift);
          else if (key === "Home") this.selectCell(1, 1);
          else handled = false;
          break;
        }
        if (event.altKey && (key === "=" || key === "+")) {
          this.options.onShortcut?.("autosum");
          break;
        }
        if (event.altKey) {
          handled = false;
          break;
        }
        if (key.length === 1) {
          event.preventDefault();
          this.beginEdit(key, { selectAll: false });
          return;
        }
        handled = false;
    }
    if (handled) event.preventDefault();
  }

  insertDate(withTime) {
    const now = new Date();
    const pad = (value) => String(value).padStart(2, "0");
    const text = withTime
      ? `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())} ${pad(now.getHours())}:${pad(now.getMinutes())}`
      : `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}`;
    this.setActiveValue(text);
  }

  /* ---------------------------------------------------------------- editor */
  beginEdit(initial = null, { selectAll = true } = {}) {
    if (this.editing) this.commitEdit("none");
    const sheet = this.sheet();
    const merge = this.mergeFor(this.active.row, this.active.col);
    const row = merge ? merge.r1 : this.active.row;
    const col = merge ? merge.c1 : this.active.col;
    const address = addr(row, col);
    const cell = cellAt(sheet, row, col);
    const before = this.options.captureState?.() ?? null;
    const originalText = rawEditText(cell);
    const text = initial === null || initial === undefined ? originalText : String(initial);
    this.ensureVisible(row, col);
    const rect = this.cellRect(row, col, merge);
    const box = el("div", { class: "cell-editor", i18n: false });
    box.style.cssText = `left:${Math.round(rect.left)}px;top:${Math.round(rect.top)}px;`
      + `width:${Math.max(rect.width, 40)}px;height:${Math.max(rect.height, 18)}px;pointer-events:auto;`;
    const input = el("input", { type: "text", spellcheck: this.spellcheck ? "true" : "false", autocomplete: "off", i18n: false });
    input.value = text;
    const style = cell?.s || {};
    const points = Number(style.size) || Number(this.defaultSize) || 11;
    input.style.fontSize = `${Math.max(6, Math.round(points * (4 / 3) * this.scale * 10) / 10)}px`;
    input.style.fontFamily = style.font || this.defaultFont || "";
    input.style.fontWeight = style.b ? "700" : "";
    input.style.fontStyle = style.i ? "italic" : "";
    if (style.align === "right" || (!style.align && formatCell(cell || {}).kind === "number")) input.style.textAlign = "right";
    else if (style.align === "center") input.style.textAlign = "center";
    box.append(input);
    const actions = el("div", { class: "editor-actions", i18n: false });
    const accept = el("button", { type: "button", title: "Enter", html: icon("check"), i18n: false });
    accept.addEventListener("mousedown", (event) => event.preventDefault());
    accept.addEventListener("click", () => this.commitEdit("down"));
    const cancel = el("button", { type: "button", title: "Cancel", html: icon("close"), i18n: false });
    cancel.addEventListener("mousedown", (event) => event.preventDefault());
    cancel.addEventListener("click", () => this.cancelEdit());
    actions.append(accept, cancel);
    box.append(actions);
    this.els.editor.append(box);
    this.editing = { row, col, address, merge, box, input, before, originalText };
    input.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        this.commitEdit(event.shiftKey ? "up" : "down");
      } else if (event.key === "Tab") {
        event.preventDefault();
        this.commitEdit(event.shiftKey ? "left" : "right");
      } else if (event.key === "Escape") {
        event.preventDefault();
        this.cancelEdit();
      }
    });
    input.addEventListener("input", () => this.options.onEditInput?.(input.value, address));
    input.focus();
    if (selectAll) input.select();
    else {
      const length = input.value.length;
      input.setSelectionRange(length, length);
    }
    this.updateFillHandle();
    this.options.onEditStart?.({ address, text, row, col });
  }

  commitEdit(move = "down") {
    if (!this.editing) return;
    const editing = this.editing;
    const text = editing.input.value;
    this.finishEditing();
    if (text === editing.originalText) {
      this.moveAfterEdit(move);
      return;
    }
    const changed = this.applyText(editing.address, text);
    this.recalculate();
    this.afterMutation("edit", editing.before, { layout: false, changed });
    this.moveAfterEdit(move);
  }

  cancelEdit() {
    if (!this.editing) return;
    this.finishEditing();
    this.els.scroller.focus({ preventScroll: true });
    this.notifySelection();
  }

  finishEditing() {
    const editing = this.editing;
    this.editing = null;
    editing?.box.remove();
    this.options.onEditEnd?.();
    this.updateFillHandle();
  }

  moveAfterEdit(move) {
    if (move === "down") this.move(1, 0);
    else if (move === "up") this.move(-1, 0);
    else if (move === "right") this.move(0, 1);
    else if (move === "left") this.move(0, -1);
    else this.notifySelection();
    this.els.scroller.focus({ preventScroll: true });
    this.renderNow();
  }

  applyText(address, text, { silent = false, validate = true } = {}) {
    const sheet = this.sheet();
    const parsed = parseUserInput(text);
    if (validate && parsed.kind !== "formula" && parsed.kind !== "empty") {
      const rule = this.validationFor(address);
      if (rule && !this.validateValue(rule, parsed.value)) {
        toast(rule.error || "This value does not match the data validation rule.", "error");
        return false;
      }
    }
    const existing = this.sheet().cells[address];
    const style = existing?.s && Object.keys(existing.s).length ? existing.s : null;
    if (parsed.kind === "empty") {
      clearCell(sheet, address, "contents");
      return true;
    }
    const cell = ensureCell(sheet, address);
    if (style) cell.s = style;
    if (parsed.kind === "formula") {
      cell.f = parsed.formula;
      delete cell.v;
    } else {
      delete cell.f;
      cell.v = parsed.value;
      if (parsed.numFmt && (!cell.s?.numFmt || cell.s.numFmt === "General")) {
        cell.s = { ...(cell.s || {}), numFmt: parsed.numFmt };
      }
    }
    return true;
  }

  setActiveValue(text, { move = "none" } = {}) {
    if (this.editing) this.commitEdit("none");
    const before = this.options.captureState?.() ?? null;
    const address = this.activeAddress();
    this.applyText(address, text);
    this.recalculate();
    this.afterMutation("edit", before, { layout: false });
    this.moveAfterEdit(move);
  }

  editActive(text = null) {
    this.beginEdit(text, { selectAll: true });
  }

  renderEditor() {
    if (!this.editing) return;
    const { row, col, merge, box } = this.editing;
    if (!box.isConnected) return;
    const rect = this.cellRect(row, col, merge);
    box.style.left = Math.round(rect.left) + "px";
    box.style.top = Math.round(rect.top) + "px";
    box.style.width = Math.max(rect.width, 40) + "px";
    box.style.height = Math.max(rect.height, 18) + "px";
  }

  recalculate(force = false) {
    if (!this.engine) return;
    if (!force && this.model?.settings?.calcMode === "manual") return;
    this.engine.rebuild();
  }

  afterMutation(kind, before, { layout = false, changed = true } = {}) {
    this.epoch += 1;
    this.cfCache.clear();
    if (layout) this.buildLayout();
    if (!this.destroyed) this.renderNow();
    this.notifySelection();
    this.options.onChange?.({ kind, before, sheetIndex: this.sheetIndex, changed });
  }

  /* -------------------------------------------------------------- clipboard */
  serializeSelection() {
    const sheet = this.sheet();
    const range = this.effectiveRange();
    const rows = [];
    const cells = [];
    const height = range.r2 - range.r1 + 1;
    const width = range.c2 - range.c1 + 1;
    for (let row = range.r1; row <= range.r2; row += 1) {
      const fields = [];
      for (let col = range.c1; col <= range.c2; col += 1) {
        const cell = cellAt(sheet, row, col);
        fields.push(cell && cell.v !== undefined ? formatCell(cell).text : "");
        if (cell && (cell.v !== undefined || cell.f || cell.s || cell.h)) {
          cells.push({
            row: row - range.r1,
            col: col - range.c1,
            v: cell.v,
            f: cell.f,
            s: cell.s ? JSON.parse(JSON.stringify(cell.s)) : undefined,
            h: cell.h,
          });
        }
      }
      rows.push(fields.map(escapeField).join("\t"));
    }
    return { text: rows.join("\n"), cells, range: { ...range }, height, width };
  }

  copySelection() {
    const data = this.serializeSelection();
    this.clipboard = { ...data, cut: false };
    return data;
  }

  cutSelection() {
    const data = this.copySelection();
    this.clipboard.cut = true;
    this.clearSelection("contents", { kind: "cut" });
    return data;
  }

  handleCopy(event) {
    if (this.editing) return;
    const data = this.serializeSelection();
    this.clipboard = { ...data, cut: false };
    event.preventDefault();
    event.clipboardData.setData("text/plain", data.text);
    event.clipboardData.setData("text/html", this.selectionHtml(data));
    this.options.onStatus?.("Copied");
  }

  handleCut(event) {
    if (this.editing) return;
    const data = this.serializeSelection();
    this.clipboard = { ...data, cut: true };
    event.preventDefault();
    event.clipboardData.setData("text/plain", data.text);
    this.clearSelection("contents", { kind: "cut" });
    this.options.onStatus?.("Cut");
  }

  handlePaste(event) {
    if (this.editing) return;
    const text = event.clipboardData?.getData("text/plain") || "";
    if (!text) return;
    event.preventDefault();
    this.pasteText(text);
  }

  selectionHtml(data) {
    const rows = data.text.split("\n").map((line) => {
      const fields = line.split("\t").map((field) => `<td>${escapeField(field).replace(/"/g, "")}</td>`).join("");
      return `<tr>${fields}</tr>`;
    }).join("");
    return `<table>${rows}</table>`;
  }

  /* mode: all | values | formulas | formats | transpose.  valuesOnly is kept
     for the classic plain-text paste path used by the context menu. */
  pasteText(text, { valuesOnly = false, mode = null, skipBlanks = false } = {}) {
    if (!text) return 0;
    const pasteMode = mode || (valuesOnly ? "values" : "all");
    const internal = this.clipboard && this.clipboard.text === text ? this.clipboard : null;
    let rows;
    let sourceHeight;
    let sourceWidth;
    if (internal) {
      sourceHeight = internal.height;
      sourceWidth = internal.width;
    } else {
      rows = parseDelimited(text);
      if (!rows.length) return 0;
      sourceHeight = rows.length;
      sourceWidth = Math.max(1, ...rows.map((row) => row.length));
      if (pasteMode === "formats") return 0;
    }
    const transpose = pasteMode === "transpose";
    const height = transpose ? sourceWidth : sourceHeight;
    const width = transpose ? sourceHeight : sourceWidth;
    const startRow = clamp(this.active.row, 1, MAX_ROWS);
    const startCol = clamp(this.active.col, 1, MAX_COLS);
    const endRow = Math.min(MAX_ROWS, startRow + height - 1);
    const endCol = Math.min(MAX_COLS, startCol + width - 1);
    const before = this.options.captureState?.() ?? null;
    const sheet = this.sheet();

    if (internal) {
      if (pasteMode !== "formats") {
        for (let row = startRow; row <= endRow; row += 1) {
          for (let col = startCol; col <= endCol; col += 1) {
            const cell = cellAt(sheet, row, col);
            if (!cell) continue;
            if (pasteMode === "values" || pasteMode === "formulas") {
              delete cell.v;
              delete cell.f;
              if (!cell.s && !cell.h) delete sheet.cells[addr(row, col)];
            } else {
              delete cell.v;
              delete cell.f;
              delete cell.h;
              delete cell.s;
              if (!Object.keys(cell).length) delete sheet.cells[addr(row, col)];
            }
          }
        }
      }
      for (const entry of internal.cells) {
        const sourceRow = transpose ? entry.col : entry.row;
        const sourceCol = transpose ? entry.row : entry.col;
        if (sourceRow < 0 || sourceCol < 0 || sourceRow >= height || sourceCol >= width) continue;
        const row = startRow + sourceRow;
        const col = startCol + sourceCol;
        if (row > endRow || col > endCol) continue;
        const hasValue = entry.v !== undefined;
        if (skipBlanks && !hasValue && !entry.f) continue;
        if (pasteMode === "formats" && !entry.s) continue;
        if (pasteMode !== "formats" && !hasValue && !entry.f) continue;
        const cell = ensureCell(sheet, addr(row, col));
        if (pasteMode === "all" || pasteMode === "transpose" || pasteMode === "formats") {
          if (entry.s) cell.s = JSON.parse(JSON.stringify(entry.s));
          if (entry.h) cell.h = entry.h;
        }
        if (pasteMode === "formats") continue;
        if (entry.f && pasteMode !== "values") {
          const sourceRowAbs = internal.range.r1 + entry.row;
          const sourceColAbs = internal.range.c1 + entry.col;
          cell.f = translateFormula(entry.f, row - sourceRowAbs, col - sourceColAbs);
          if (entry.v !== undefined) cell.v = entry.v;
        } else if (entry.v !== undefined) {
          cell.v = entry.v;
        }
        if (!cell.s && cell.v === undefined && !cell.f && !cell.h) delete sheet.cells[addr(row, col)];
      }
    } else {
      let matrix = rows;
      if (transpose) {
        const flipped = [];
        for (let col = 0; col < width; col += 1) {
          const line = [];
          for (let row = 0; row < height; row += 1) line.push((rows[row] || [])[col] ?? "");
          flipped.push(line);
        }
        matrix = flipped;
      }
      for (let row = startRow; row <= endRow; row += 1) {
        const source = matrix[row - startRow] || [];
        for (let col = startCol; col <= endCol; col += 1) {
          const field = source[col - startCol];
          if (field === undefined || field === "") {
            if (!skipBlanks) clearCell(sheet, addr(row, col), "contents");
            continue;
          }
          if (pasteMode === "values" && field.startsWith("=")) {
            ensureCell(sheet, addr(row, col)).v = field;
            continue;
          }
          if (pasteMode === "values") {
            const parsed = parseUserInput(field);
            const cell = ensureCell(sheet, addr(row, col));
            delete cell.f;
            if (parsed.kind === "empty") clearCell(sheet, addr(row, col), "contents");
            else cell.v = parsed.kind === "formula" ? field : parsed.value;
            continue;
          }
          this.applyText(addr(row, col), field);
        }
      }
    }
    this.setSelection(startRow, startCol, endRow, endCol, { activeRow: startRow, activeCol: startCol });
    this.recalculate();
    this.afterMutation("paste", before, { layout: false });
    this.ensureVisible(startRow, startCol);
    this.options.onStatus?.("Pasted");
    return height * width;
  }

  clearSelection(what = "contents", { kind = "clear" } = {}) {
    const sheet = this.sheet();
    const before = this.options.captureState?.() ?? null;
    let changed = false;
    for (const [address, cell] of Object.entries(sheet.cells)) {
      const parsed = parseAddr(address);
      if (!parsed || !rangeContains(this.sel, parsed.row, parsed.col)) continue;
      if (what === "contents" && (cell.v === undefined || cell.v === null || cell.v === "") && !cell.f && cell.h === undefined) continue;
      clearCell(sheet, address, what);
      changed = true;
    }
    if (!changed) return;
    this.recalculate();
    this.afterMutation(kind, before, { layout: false });
  }

  deleteSelection() {
    this.clearSelection("contents", { kind: "delete" });
  }

  /* ------------------------------------------------------------ formatting */
  activeCell() {
    return cellAt(this.sheet(), this.active.row, this.active.col);
  }

  currentStyle() {
    return cloneStyle(this.activeCell()?.s) || {};
  }

  selectedAddresses(range = this.sel, limit = 4000) {
    const sheet = this.sheet();
    const addresses = new Set();
    if (rangeSize(range) <= limit) {
      for (let row = range.r1; row <= range.r2; row += 1) {
        for (let col = range.c1; col <= range.c2; col += 1) addresses.add(addr(row, col));
      }
    } else {
      for (const address of Object.keys(sheet.cells)) {
        const parsed = parseAddr(address);
        if (parsed && rangeContains(range, parsed.row, parsed.col)) addresses.add(address);
      }
    }
    return addresses;
  }

  /* Merge a style patch into every cell of the selection.  A value of null,
     false or an empty string removes the property, which is what the ribbon
     toggles send when they are switched off. */
  setStyle(patch, { scope = "selection" } = {}) {
    const sheet = this.sheet();
    const range = scope === "cell"
      ? { r1: this.active.row, c1: this.active.col, r2: this.active.row, c2: this.active.col }
      : this.sel;
    const before = this.options.captureState?.() ?? null;
    const addresses = this.selectedAddresses(range);
    for (const address of addresses) {
      const parsed = parseAddr(address);
      if (!parsed) continue;
      const cell = ensureCell(sheet, address);
      const style = { ...(cell.s || {}) };
      for (const [key, value] of Object.entries(patch)) {
        if (key === "border" && value && typeof value === "object") {
          const border = { ...(style.border || {}) };
          for (const [side, spec] of Object.entries(value)) {
            if (!spec) delete border[side];
            else border[side] = { ...spec };
          }
          if (Object.keys(border).length) style.border = border;
          else delete style.border;
          continue;
        }
        const clear = value === null || value === undefined || value === "" || value === "general"
          || (value === false && key !== "locked");
        if (clear) delete style[key];
        else style[key] = value;
      }
      if (Object.keys(style).length) cell.s = style;
      else delete cell.s;
      if (!cell.s && cell.v === undefined && !cell.f && !cell.h) delete sheet.cells[address];
    }
    this.afterMutation("format", before, { layout: true });
  }

  getStyle() {
    return this.currentStyle();
  }

  toggleStyle(key, value) {
    const enabled = this.currentStyle()[key] === value;
    this.setStyle({ [key]: enabled ? null : value });
    return !enabled;
  }

  setNumberFormat(format) {
    this.setStyle({ numFmt: !format || format === "General" ? null : String(format) });
  }

  setAlignment(align) {
    this.setStyle({ align: !align || align === "general" ? null : align });
  }

  setVerticalAlignment(valign) {
    this.setStyle({ valign: !valign || valign === "middle" ? null : valign });
  }

  applyBorders(preset, spec = { style: "thin", color: "#000000" }) {
    if (preset === "none") {
      this.setStyle({ border: { t: null, r: null, b: null, l: null } });
      return;
    }
    const sheet = this.sheet();
    const range = this.sel;
    const before = this.options.captureState?.() ?? null;
    const addresses = this.selectedAddresses(range, 20000);
    const limit = addresses.size <= 20000;
    for (const address of addresses) {
      let row = 0;
      let col = 0;
      if (limit) {
        const parsed = parseAddr(address);
        row = parsed.row;
        col = parsed.col;
      }
      const cell = ensureCell(sheet, address);
      const style = { ...(cell.s || {}) };
      const border = { ...(style.border || {}) };
      const sides = [];
      const outside = preset === "outside";
      if (preset === "all" || (outside && row === range.r1) || preset === "top") sides.push("t");
      if (preset === "all" || (outside && row === range.r2) || preset === "bottom") sides.push("b");
      if (preset === "all" || (outside && col === range.c1) || preset === "left") sides.push("l");
      if (preset === "all" || (outside && col === range.c2) || preset === "right") sides.push("r");
      if (preset === "inside" || preset === "all") {
        if (limit && row < range.r2) sides.push("b");
        if (limit && col < range.c2) sides.push("r");
      }
      if (!sides.length) continue;
      for (const side of sides) border[side] = { ...spec };
      style.border = border;
      cell.s = style;
    }
    this.afterMutation("borders", before, { layout: true });
  }

  /* ---------------------------------------------------------------- merges */
  mergeCells() {
    const range = { ...this.sel };
    if (range.r1 === range.r2 && range.c1 === range.c2) {
      toast("Select two or more cells to merge.");
      return false;
    }
    const sheet = this.sheet();
    const before = this.options.captureState?.() ?? null;
    addMerge(sheet, range);
    for (let row = range.r1; row <= range.r2; row += 1) {
      for (let col = range.c1; col <= range.c2; col += 1) {
        if (row === range.r1 && col === range.c1) continue;
        const address = addr(row, col);
        if (sheet.cells[address]) clearCell(sheet, address, "contents");
      }
    }
    this.recalculate();
    this.afterMutation("merge", before, { layout: true });
    this.setSelection(range.r1, range.c1, range.r2, range.c2, { activeRow: range.r1, activeCol: range.c1 });
    return true;
  }

  unmergeCells() {
    const sheet = this.sheet();
    const before = this.options.captureState?.() ?? null;
    const targets = (sheet.merges || []).filter((text) => {
      const parsed = parseRange(text);
      return parsed && rangesIntersect(parsed, this.sel);
    });
    if (!targets.length) return false;
    for (const text of targets) {
      const parsed = parseRange(text);
      removeMergeAt(sheet, addr(parsed.r1, parsed.c1));
    }
    this.afterMutation("merge", before, { layout: true });
    return true;
  }

  toggleMerge() {
    const sheet = this.sheet();
    const inside = (sheet.merges || []).some((text) => {
      const parsed = parseRange(text);
      return parsed && rangesIntersect(parsed, this.sel);
    });
    if (inside) return this.unmergeCells();
    return this.mergeCells();
  }

  /* ---------------------------------------------------------------- freeze */
  setFreeze(rows, cols) {
    const sheet = this.sheet();
    const before = this.options.captureState?.() ?? null;
    const rowCount = clamp(Number(rows) || 0, 0, MAX_ROWS - 1);
    const colCount = clamp(Number(cols) || 0, 0, MAX_COLS - 1);
    if (rowCount || colCount) sheet.freeze = { rows: rowCount, cols: colCount };
    else delete sheet.freeze;
    this.buildLayout();
    this.afterMutation("freeze", before, { layout: false });
    this.options.onStatus?.(rowCount || colCount ? `Frozen: ${rowCount} rows, ${colCount} columns` : "Freeze panes off");
    return { rows: rowCount, cols: colCount };
  }

  freezePanes(rows = null, cols = null) {
    const rowCount = rows === null ? this.active.row - 1 : rows;
    const colCount = cols === null ? this.active.col - 1 : cols;
    return this.setFreeze(rowCount, colCount);
  }

  unfreezePanes() {
    return this.setFreeze(0, 0);
  }

  toggleFreezeAtActive() {
    const freeze = this.sheet().freeze || {};
    if (freeze.rows || freeze.cols) return this.unfreezePanes();
    return this.freezePanes();
  }

  /* ------------------------------------------------------------- appearance */
  setGridlines(visible) {
    const sheet = this.sheet();
    const before = this.options.captureState?.() ?? null;
    sheet.showGridlines = !!visible;
    this.els.cells.classList.toggle("no-gridlines", !visible);
    this.afterMutation("view", before, { layout: false });
  }

  toggleGridlines() {
    this.setGridlines(this.sheet().showGridlines === false);
  }

  setHeadings(visible) {
    const sheet = this.sheet();
    const before = this.options.captureState?.() ?? null;
    sheet.showHeadings = !!visible;
    const area = this.els.viewport?.closest("#sheet-area");
    area?.classList.toggle("no-headings", !visible);
    this.els.colHeaders.style.display = visible ? "" : "none";
    this.els.rowHeaders.style.display = visible ? "" : "none";
    this.els.corner.style.display = visible ? "" : "none";
    this.afterMutation("view", before, { layout: true });
  }

  toggleHeadings() {
    this.setHeadings(this.sheet().showHeadings === false);
  }

  /* -------------------------------------------------------------- comments */
  commentAt(address = this.activeAddress()) {
    return (this.sheet().comments || []).find((comment) => comment.addr === address) || null;
  }

  setComment(text, author = "") {
    const sheet = this.sheet();
    const address = this.activeAddress();
    const before = this.options.captureState?.() ?? null;
    deleteCellComments(sheet, address);
    if (String(text || "").trim()) {
      addComment(sheet, address, String(text), author || "mTails");
    }
    this.afterMutation("comment", before, { layout: true });
    return this.commentAt(address);
  }

  deleteCommentAt(address = this.activeAddress()) {
    const sheet = this.sheet();
    const before = this.options.captureState?.() ?? null;
    if (!deleteCellComments(sheet, address)) return false;
    this.afterMutation("comment", before, { layout: true });
    return true;
  }

  /* ---------------------------------------------------------------- search */
  search(text, { select = true, forward = true, caseSensitive = false, entire = false, wrap = true } = {}) {
    const term = String(text ?? "");
    const sheet = this.sheet();
    this.searchActive = term;
    this.searchSet = new Set();
    if (!term) {
      this.invalidateValues();
      return { count: 0, address: null, matches: [] };
    }
    const needle = caseSensitive ? term : term.toLowerCase();
    const matches = [];
    for (const [address, cell] of Object.entries(sheet.cells)) {
      const parsed = parseAddr(address);
      if (!parsed) continue;
      const text = formatCell(cell).text;
      const hay = caseSensitive ? String(text ?? "") : String(text ?? "").toLowerCase();
      const hit = entire ? hay === needle : hay.includes(needle);
      if (!hit) continue;
      matches.push({ address, row: parsed.row, col: parsed.col });
      this.searchSet.add(sheet.name.toLowerCase() + "!" + address.toLowerCase());
    }
    matches.sort((left, right) => left.row - right.row || left.col - right.col);
    this.invalidateValues();
    if (!select || !matches.length) {
      return { count: matches.length, address: null, matches };
    }
    const startRow = this.active.row;
    const startCol = this.active.col;
    let chosen = null;
    if (forward) {
      chosen = matches.find((match) => match.row > startRow || (match.row === startRow && match.col > startCol)) || null;
      if (!chosen && wrap) chosen = matches[0];
    } else {
      chosen = [...matches].reverse().find((match) => match.row < startRow || (match.row === startRow && match.col < startCol)) || null;
      if (!chosen && wrap) chosen = matches[matches.length - 1];
    }
    if (chosen) {
      this.selectCell(chosen.row, chosen.col);
      this.scrollToCell(chosen.row, chosen.col, { center: true });
    }
    return { count: matches.length, address: chosen?.address || null, matches };
  }

  findNext(text, options = {}) {
    return this.search(text, { ...options, select: true, forward: options.forward !== false });
  }

  clearSearch() {
    this.searchActive = "";
    this.searchSet = new Set();
    this.invalidateValues();
  }

  /* ---------------------------------------------------------------- filters */
  filterValues(col, range = null) {
    const sheet = this.sheet();
    const area = range || parseRange(sheet.filters?.range || "") || {
      r1: 1, c1: 1, r2: Math.max(1, usedRange(sheet)?.r2 || 1), c2: Math.max(1, usedRange(sheet)?.c2 || 1),
    };
    const counts = new Map();
    for (let row = area.r1 + 1; row <= area.r2; row += 1) {
      const cell = cellAt(sheet, row, col);
      const label = cell ? formatCell(cell).text : "";
      const key = label === undefined || label === null ? "" : String(label);
      counts.set(key, (counts.get(key) || 0) + 1);
    }
    return [...counts.entries()]
      .map(([value, count]) => ({ value, count }))
      .sort((left, right) => String(left.value).localeCompare(String(right.value), undefined, { numeric: true, sensitivity: "base" }));
  }

  matchesFilter(sheet, row, criterion, range) {
    const col = Number(criterion.col);
    if (!col || col < range.c1 || col > range.c2) return true;
    const cell = cellAt(sheet, row, col);
    const raw = cell?.v;
    const display = cell ? String(formatCell(cell).text ?? "") : "";
    const blank = raw === undefined || raw === null || raw === "";
    if (Array.isArray(criterion.values)) {
      if (criterion.values.includes("\u0000blank") && blank) return true;
      return criterion.values.includes(display);
    }
    const operator = criterion.operator || "equals";
    const value = criterion.value;
    const text = String(display).toLowerCase();
    const needle = String(value ?? "").toLowerCase();
    const number = Number(raw);
    const target = Number(value);
    switch (operator) {
      case "equals": return text === needle;
      case "notEquals": return text !== needle;
      case "contains": return text.includes(needle);
      case "notContains": return !text.includes(needle);
      case "begins": return text.startsWith(needle);
      case "ends": return text.endsWith(needle);
      case "blank": return blank;
      case "nonBlank": return !blank;
      case "greater": return Number.isFinite(number) && Number.isFinite(target) && number > target;
      case "less": return Number.isFinite(number) && Number.isFinite(target) && number < target;
      case "greaterOrEqual": return Number.isFinite(number) && Number.isFinite(target) && number >= target;
      case "lessOrEqual": return Number.isFinite(number) && Number.isFinite(target) && number <= target;
      case "between": return Number.isFinite(number) && number >= target && number <= Number(criterion.value2);
      case "top10": return Number.isFinite(number);
      default: return true;
    }
  }

  applyFilter(range, criteria) {
    const sheet = this.sheet();
    const normalized = normalizeRange(range.r1, range.c1, range.r2, range.c2);
    const before = this.options.captureState?.() ?? null;
    const hidden = [];
    for (let row = normalized.r1 + 1; row <= normalized.r2; row += 1) {
      const visible = (criteria || []).every((criterion) => this.matchesFilter(sheet, row, criterion, normalized));
      if (!visible) hidden.push(row);
    }
    sheet.filters = {
      range: rangeText(normalized),
      criteria: criteria || [],
      hidden,
    };
    this.buildLayout();
    this.afterMutation("filter", before, { layout: false });
    this.setSelection(normalized.r1, normalized.c1, normalized.r2, normalized.c2, { activeRow: normalized.r1, activeCol: normalized.c1 });
    return hidden.length;
  }

  clearFilter() {
    const sheet = this.sheet();
    if (!sheet.filters) return false;
    const before = this.options.captureState?.() ?? null;
    delete sheet.filters;
    this.buildLayout();
    this.afterMutation("filter", before, { layout: false });
    return true;
  }

  /* ------------------------------------------------------------------ sort */
  applySort(keys, { hasHeader = true } = {}) {
    const sheet = this.sheet();
    let range = { ...this.sel };
    if (range.r1 === range.r2 && range.c1 === range.c2) {
      const used = usedRange(sheet);
      if (used) range = { ...used };
    }
    if (range.r1 === range.r2 && range.c1 === range.c2) return false;
    const before = this.options.captureState?.() ?? null;
    const dataRange = hasHeader && range.r1 < range.r2
      ? { r1: range.r1 + 1, c1: range.c1, r2: range.r2, c2: range.c2 }
      : range;
    const list = (keys || []).map((key) => ({
      address: addr(range.r1, Number(key.col) || range.c1),
      order: key.order === "desc" ? "desc" : "asc",
      caseSensitive: !!key.caseSensitive,
    }));
    if (!list.length) return false;
    if (!sortRange(sheet, dataRange, list)) return false;
    this.recalculate();
    this.afterMutation("sort", before, { layout: false });
    return true;
  }

  removeDuplicates(columns = null) {
    const sheet = this.sheet();
    let range = { ...this.sel };
    if (range.r1 === range.r2 && range.c1 === range.c2) {
      const used = usedRange(sheet);
      if (used) range = { ...used };
    }
    const before = this.options.captureState?.() ?? null;
    const removed = removeDuplicateRows(sheet, range, columns);
    if (!removed) return 0;
    this.recalculate();
    this.afterMutation("dedupe", before, { layout: false });
    return removed;
  }

  /* ------------------------------------------------------------ structure */
  insertRow(count = 1, where = "above") {
    const sheet = this.sheet();
    const at = where === "below" ? this.sel.r2 + 1 : this.sel.r1;
    const before = this.options.captureState?.() ?? null;
    const amount = insertRows(sheet, at, clamp(Number(count) || 1, 1, MAX_ROWS - 1));
    if (!amount) return;
    this.recalculate();
    this.afterMutation("insert-rows", before, { layout: true });
    this.setSelection(at, this.sel.c1, at + amount - 1, this.sel.c2, { activeRow: at, activeCol: this.sel.c1 });
  }

  insertColumn(count = 1, where = "left") {
    const sheet = this.sheet();
    const at = where === "right" ? this.sel.c2 + 1 : this.sel.c1;
    const before = this.options.captureState?.() ?? null;
    const amount = insertCols(sheet, at, clamp(Number(count) || 1, 1, MAX_COLS - 1));
    if (!amount) return;
    this.recalculate();
    this.afterMutation("insert-cols", before, { layout: true });
    this.setSelection(this.sel.r1, at, this.sel.r2, at + amount - 1, { activeRow: this.sel.r1, activeCol: at });
  }

  deleteRowSelection() {
    const sheet = this.sheet();
    const count = this.sel.r2 - this.sel.r1 + 1;
    const before = this.options.captureState?.() ?? null;
    const amount = deleteRows(sheet, this.sel.r1, count);
    if (!amount) return;
    this.recalculate();
    this.afterMutation("delete-rows", before, { layout: true });
    this.setSelection(this.sel.r1, 1, this.sel.r1, MAX_COLS, { activeRow: this.sel.r1, activeCol: 1 });
  }

  deleteColumnSelection() {
    const sheet = this.sheet();
    const count = this.sel.c2 - this.sel.c1 + 1;
    const before = this.options.captureState?.() ?? null;
    const amount = deleteCols(sheet, this.sel.c1, count);
    if (!amount) return;
    this.recalculate();
    this.afterMutation("delete-cols", before, { layout: true });
    this.setSelection(1, this.sel.c1, MAX_ROWS, this.sel.c1, { activeRow: 1, activeCol: this.sel.c1 });
  }

  hideSelectedRows() {
    const sheet = this.sheet();
    const before = this.options.captureState?.() ?? null;
    for (let row = this.sel.r1; row <= this.sel.r2; row += 1) setRowHidden(sheet, row, true);
    this.afterMutation("hide-rows", before, { layout: true });
  }

  hideSelectedColumns() {
    const sheet = this.sheet();
    const before = this.options.captureState?.() ?? null;
    for (let col = this.sel.c1; col <= this.sel.c2; col += 1) setColHidden(sheet, col, true);
    this.afterMutation("hide-cols", before, { layout: true });
  }

  showAllRows() {
    const sheet = this.sheet();
    const before = this.options.captureState?.() ?? null;
    for (const row of Object.keys(sheet.rows || {})) setRowHidden(sheet, Number(row), false);
    this.afterMutation("show-rows", before, { layout: true });
  }

  showAllColumns() {
    const sheet = this.sheet();
    const before = this.options.captureState?.() ?? null;
    for (const col of Object.keys(sheet.cols || {})) setColHidden(sheet, colToIndex(col), false);
    this.afterMutation("show-cols", before, { layout: true });
  }

  setSelectedRowHeight(height) {
    const sheet = this.sheet();
    const before = this.options.captureState?.() ?? null;
    for (let row = this.sel.r1; row <= this.sel.r2; row += 1) setRowHeight(sheet, row, height);
    this.afterMutation("row-height", before, { layout: true });
  }

  setSelectedColumnWidth(width) {
    const sheet = this.sheet();
    const before = this.options.captureState?.() ?? null;
    for (let col = this.sel.c1; col <= this.sel.c2; col += 1) setColWidth(sheet, col, width);
    this.afterMutation("col-width", before, { layout: true });
  }

  autoFitSelection() {
    for (let col = this.sel.c1; col <= this.sel.c2; col += 1) this.autoFitColumn(col);
    for (let row = this.sel.r1; row <= this.sel.r2; row += 1) this.autoFitRow(row);
  }

  /* ------------------------------------------------------------ validation */
  validationFor(address) {
    const parsed = parseAddr(address);
    if (!parsed) return null;
    for (const rule of this.sheet().validations || []) {
      const range = parseRange(rule.range);
      if (range && rangeContains(range, parsed.row, parsed.col)) return rule;
    }
    return null;
  }

  validateValue(rule, value) {
    const blank = value === undefined || value === null || value === "";
    if (blank) return !!rule.allowBlank;
    const type = rule.type || "any";
    const stringValue = String(value);
    if (type === "list") {
      const source = String(rule.formula1 || "");
      let allowed = [];
      const reference = parseRange(source.replace(/^=/, ""));
      if (reference) {
        for (let row = reference.r1; row <= reference.r2; row += 1) {
          for (let col = reference.c1; col <= reference.c2; col += 1) {
            const cell = cellAt(this.sheet(), row, col);
            if (cell && cell.v !== undefined && cell.v !== null && cell.v !== "") allowed.push(String(cell.v));
          }
        }
      } else {
        allowed = source.split(",").map((item) => item.trim().replace(/^"|"$/g, ""));
      }
      return allowed.some((item) => item === stringValue);
    }
    let number = Number(value);
    if (type === "date" || type === "time") {
      if (!Number.isFinite(number)) {
        const parsed = parseUserInput(stringValue);
        number = parsed.kind === "number" ? parsed.value : NaN;
      }
    } else {
      if (!Number.isFinite(number)) return false;
    }
    let first = Number(rule.formula1);
    let second = Number(rule.formula2);
    if (type === "date" || type === "time") {
      if (!Number.isFinite(first)) first = Number(parseUserInput(String(rule.formula1 || "")).value);
      if (!Number.isFinite(second)) second = Number(parseUserInput(String(rule.formula2 || "")).value);
    }
    const length = stringValue.length;
    const left = type === "textLength" ? length : number;
    const right = type === "textLength" ? Number(rule.formula2) : second;
    const limit = type === "textLength" ? Number(rule.formula1) : first;
    const operator = rule.operator || "between";
    if (type === "textLength") {
      if (operator === "between") return length >= limit && length <= right;
      if (operator === "equal") return length === limit;
      if (operator === "notEqual") return length !== limit;
      if (operator === "greaterThan") return length > limit;
      if (operator === "lessThan") return length < limit;
      if (operator === "greaterThanOrEqual") return length >= limit;
      if (operator === "lessThanOrEqual") return length <= limit;
    }
    if (!Number.isFinite(left) || !Number.isFinite(limit)) return true;
    switch (operator) {
      case "between": return left >= limit && left <= right;
      case "notBetween": return left < limit || left > right;
      case "equal": return left === limit;
      case "notEqual": return left !== limit;
      case "greaterThan": return left > limit;
      case "lessThan": return left < limit;
      case "greaterThanOrEqual": return left >= limit;
      case "lessThanOrEqual": return left <= limit;
      default: return true;
    }
  }

  /* -------------------------------------------------------------- fill */
  beginFillDrag(event) {
    if (this.editing) return;
    const rect = this.els.viewport.getBoundingClientRect();
    const hit = this.hitTest(event.clientX - rect.left, event.clientY - rect.top);
    this.drag = { type: "fill", source: { ...this.sel }, target: { ...this.sel } };
    if (!this.fillPreview) {
      this.fillPreview = el("div", { class: "fill-preview hidden", i18n: false });
      this.fillTip = el("div", { class: "fill-tip hidden", i18n: false });
      this.els.layer.append(this.fillPreview, this.fillTip);
    }
    this.updateFillPreviewFromCell(hit.row, hit.col);
    this.els.scroller.focus({ preventScroll: true });
  }

  updateFillPreview(event) {
    const rect = this.els.viewport.getBoundingClientRect();
    const hit = this.hitTest(event.clientX - rect.left, event.clientY - rect.top);
    this.updateFillPreviewFromCell(hit.row, hit.col);
  }

  updateFillPreviewFromCell(row, col) {
    if (!this.drag || this.drag.type !== "fill") return;
    const source = this.drag.source;
    const target = { ...source };
    const dy = row - source.r2;
    const dx = col - source.c2;
    if (Math.abs(dy) >= Math.abs(dx)) {
      if (row > source.r2) target.r2 = row;
      else if (row < source.r1) target.r1 = row;
    } else if (col > source.c2) {
      target.c2 = col;
    } else if (col < source.c1) {
      target.c1 = col;
    }
    this.drag.target = target;
    let extension = null;
    if (target.r2 > source.r2) extension = { r1: source.r2 + 1, r2: target.r2, c1: source.c1, c2: source.c2 };
    else if (target.r1 < source.r1) extension = { r1: target.r1, r2: source.r1 - 1, c1: source.c1, c2: source.c2 };
    else if (target.c2 > source.c2) extension = { r1: source.r1, r2: source.r2, c1: source.c2 + 1, c2: target.c2 };
    else if (target.c1 < source.c1) extension = { r1: source.r1, r2: source.r2, c1: target.c1, c2: source.c1 - 1 };
    if (!extension) {
      this.fillPreview.classList.add("hidden");
      this.fillTip.classList.add("hidden");
      return;
    }
    const left = this.colPrefix[extension.c1 - 1] - (extension.c1 > this.freezeCols ? this.scrollLeft : 0);
    const top = this.rowPrefix[extension.r1 - 1] - (extension.r1 > this.freezeRows ? this.scrollTop : 0);
    const right = this.colPrefix[extension.c2] - (extension.c2 > this.freezeCols ? this.scrollLeft : 0);
    const bottom = this.rowPrefix[extension.r2] - (extension.r2 > this.freezeRows ? this.scrollTop : 0);
    const width = Math.max(0, right - left);
    const height = Math.max(0, bottom - top);
    this.fillPreview.style.cssText = `left:${Math.round(Math.max(0, left))}px;top:${Math.round(Math.max(0, top))}px;width:${Math.round(Math.min(this.viewportWidth, width) - Math.max(0, -left))}px;height:${Math.round(Math.min(this.viewportHeight, height + top) - Math.max(0, -top))}px;`;
    this.fillPreview.classList.remove("hidden");
    this.fillTip.textContent = `${target.r2 - target.r1 + 1} × ${target.c2 - target.c1 + 1}`;
    this.fillTip.style.left = Math.round(Math.max(0, Math.min(this.viewportWidth - 50, right - 24))) + "px";
    this.fillTip.style.top = Math.round(Math.max(0, Math.min(this.viewportHeight - 20, bottom + 2))) + "px";
    this.fillTip.classList.remove("hidden");
  }

  finishFillDrag() {
    const drag = this.drag;
    if (this.fillPreview) {
      this.fillPreview.remove();
      this.fillPreview = null;
    }
    if (this.fillTip) {
      this.fillTip.remove();
      this.fillTip = null;
    }
    if (!drag || drag.type !== "fill") return;
    const { source, target } = drag;
    if (target.r1 === source.r1 && target.r2 === source.r2 && target.c1 === source.c1 && target.c2 === source.c2) return;
    this.fillTo(source, target);
  }

  fillTo(source, target) {
    const direction = target.r2 > source.r2 ? "down"
      : target.r1 < source.r1 ? "up"
        : target.c2 > source.c2 ? "right"
          : target.c1 < source.c1 ? "left" : "";
    if (!direction) return;
    const sheet = this.sheet();
    const before = this.options.captureState?.() ?? null;
    if (direction === "down" || direction === "up") {
      for (let col = source.c1; col <= source.c2; col += 1) {
        const entries = [];
        for (let row = source.r1; row <= source.r2; row += 1) {
          entries.push({ cell: cellAt(sheet, row, col), row, col });
        }
        const targets = [];
        if (direction === "down") {
          for (let row = source.r2 + 1; row <= target.r2; row += 1) targets.push({ row, col });
        } else {
          for (let row = source.r1 - 1; row >= target.r1; row -= 1) targets.push({ row, col });
        }
        this.fillLine(entries, targets, direction === "up");
      }
    } else {
      for (let row = source.r1; row <= source.r2; row += 1) {
        const entries = [];
        for (let col = source.c1; col <= source.c2; col += 1) {
          entries.push({ cell: cellAt(sheet, row, col), row, col });
        }
        const targets = [];
        if (direction === "right") {
          for (let col = source.c2 + 1; col <= target.c2; col += 1) targets.push({ row, col });
        } else {
          for (let col = source.c1 - 1; col >= target.c1; col -= 1) targets.push({ row, col });
        }
        this.fillLine(entries, targets, direction === "left");
      }
    }
    this.recalculate();
    this.afterMutation("fill", before, { layout: false });
  }

  fillLine(entries, targets, backwards) {
    if (!entries.length || !targets.length) return;
    const sheet = this.sheet();
    const n = entries.length;
    const values = entries.map((entry) => entry.cell?.v);
    const formulas = entries.map((entry) => entry.cell?.f);
    const allNumeric = values.every((value) => typeof value === "number" && Number.isFinite(value));
    const allFormulas = formulas.every((formula) => !!formula);
    const strings = values.map((value) => typeof value === "string" ? value : null);
    const matches = strings.map((value) => value === null ? null : /^(.*?)(\d+)$/.exec(value));
    const prefix = matches.every((match) => match && match[1] === matches[0][1]) ? matches[0]?.[1] : null;
    const textSeries = prefix !== null;
    const numbers = matches.map((match) => match ? parseInt(match[2], 10) : NaN);
    const textWidth = matches[0] ? matches[0][2].length : 0;
    let step = 0;
    if (allNumeric) {
      step = n >= 2 ? (values[n - 1] - values[0]) / (n - 1) : 0;
    } else if (textSeries) {
      step = n >= 2 ? numbers[n - 1] - numbers[0] : 1;
      if (!Number.isFinite(step)) step = 1;
    }
    targets.forEach((target, index) => {
      const sourceIndex = index % n;
      const entry = entries[sourceIndex];
      const sourceCell = entry.cell;
      const address = addr(target.row, target.col);
      const cell = ensureCell(sheet, address);
      delete cell.v;
      delete cell.f;
      delete cell.h;
      if (sourceCell?.s) cell.s = cloneStyle(sourceCell.s);
      else delete cell.s;
      if (sourceCell?.h) cell.h = sourceCell.h;
      if (allFormulas) {
        cell.f = translateFormula(sourceCell.f, target.row - entry.row, target.col - entry.col);
      } else if (allNumeric) {
        cell.v = backwards
          ? values[0] - step * (index + 1)
          : values[n - 1] + step * (index + 1);
      } else if (textSeries) {
        const base = backwards ? numbers[0] - step * (index + 1) : numbers[n - 1] + step * (index + 1);
        cell.v = prefix + String(Math.max(0, Math.round(base))).padStart(textWidth, "0");
      } else if (sourceCell) {
        if (sourceCell.f) cell.f = translateFormula(sourceCell.f, target.row - entry.row, target.col - entry.col);
        else if (sourceCell.v !== undefined) cell.v = sourceCell.v;
      }
      if (!Object.keys(cell).length) delete sheet.cells[address];
    });
  }

  fillSelection(direction) {
    const sel = this.sel;
    let source;
    let target;
    if (direction === "down" && sel.r2 > sel.r1) {
      source = { r1: sel.r1, c1: sel.c1, r2: sel.r1, c2: sel.c2 };
      target = { ...sel };
    } else if (direction === "up" && sel.r2 > sel.r1) {
      source = { r1: sel.r2, c1: sel.c1, r2: sel.r2, c2: sel.c2 };
      target = { ...sel };
    } else if (direction === "right" && sel.c2 > sel.c1) {
      source = { r1: sel.r1, c1: sel.c1, r2: sel.r2, c2: sel.c1 };
      target = { ...sel };
    } else if (direction === "left" && sel.c2 > sel.c1) {
      source = { r1: sel.r1, c1: sel.c2, r2: sel.r2, c2: sel.c2 };
      target = { ...sel };
    } else {
      return;
    }
    this.fillTo(source, target);
  }
}

