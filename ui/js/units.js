/* Address arithmetic shared by the model, formula engine and grid. */

export const MAX_ROWS = 100000;
export const MAX_COLS = 1000;

export function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

export function colToIndex(letters) {
  let index = 0;
  const text = String(letters || "").toUpperCase().replace(/\$/g, "");
  for (const char of text) {
    const code = char.charCodeAt(0);
    if (code < 65 || code > 90) return 0;
    index = index * 26 + (code - 64);
  }
  return index;
}

export function indexToCol(index) {
  let value = Math.max(1, Math.trunc(index) || 1);
  let name = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    value = Math.floor((value - 1) / 26);
  }
  return name;
}

export function addr(row, col) {
  return indexToCol(col) + row;
}

const ADDRESS_RE = /^\$?([A-Za-z]{1,3})\$?(\d+)$/;

export function parseAddr(text) {
  const match = ADDRESS_RE.exec(String(text || "").trim());
  if (!match) return null;
  const col = colToIndex(match[1]);
  const row = parseInt(match[2], 10);
  if (col < 1 || col > MAX_COLS || row < 1 || row > MAX_ROWS) return null;
  return { row, col };
}

export function parseRange(text) {
  const parts = String(text || "").split(":");
  if (parts.length === 1) {
    const single = parseAddr(parts[0]);
    return single ? { r1: single.row, c1: single.col, r2: single.row, c2: single.col } : null;
  }
  if (parts.length !== 2) return null;
  const first = parseAddr(parts[0]);
  const last = parseAddr(parts[1]);
  if (!first || !last) return null;
  return {
    r1: Math.min(first.row, last.row),
    c1: Math.min(first.col, last.col),
    r2: Math.max(first.row, last.row),
    c2: Math.max(first.col, last.col),
  };
}

export function rangeText(range) {
  if (!range) return "";
  return addr(range.r1, range.c1) + (range.r1 === range.r2 && range.c1 === range.c2 ? "" : ":" + addr(range.r2, range.c2));
}

export function rangeContains(range, row, col) {
  return !!range && row >= range.r1 && row <= range.r2 && col >= range.c1 && col <= range.c2;
}

export function rangeSize(range) {
  if (!range) return 0;
  return (range.r2 - range.r1 + 1) * (range.c2 - range.c1 + 1);
}

export function rangesIntersect(a, b) {
  return !!a && !!b && a.r1 <= b.r2 && a.r2 >= b.r1 && a.c1 <= b.c2 && a.c2 >= b.c1;
}

export function rangeUnion(a, b) {
  if (!a) return b ? { ...b } : null;
  if (!b) return { ...a };
  return {
    r1: Math.min(a.r1, b.r1),
    c1: Math.min(a.c1, b.c1),
    r2: Math.max(a.r2, b.r2),
    c2: Math.max(a.c2, b.c2),
  };
}

export function normalizeRange(r1, c1, r2, c2) {
  return {
    r1: clamp(Math.min(r1, r2), 1, MAX_ROWS),
    c1: clamp(Math.min(c1, c2), 1, MAX_COLS),
    r2: clamp(Math.max(r1, r2), 1, MAX_ROWS),
    c2: clamp(Math.max(c1, c2), 1, MAX_COLS),
  };
}

export function * iterateRange(range, limit = Infinity) {
  let count = 0;
  for (let row = range.r1; row <= range.r2; row += 1) {
    for (let col = range.c1; col <= range.c2; col += 1) {
      yield { row, col };
      count += 1;
      if (count >= limit) return;
    }
  }
}

export function splitSheetRef(text) {
  const value = String(text || "");
  let sheet = "";
  let rest = value;
  const bang = value.lastIndexOf("!");
  if (bang >= 0) {
    sheet = value.slice(0, bang);
    rest = value.slice(bang + 1);
    if (sheet.startsWith("'") && sheet.endsWith("'")) sheet = sheet.slice(1, -1).replace(/''/g, "'");
  }
  return { sheet, ref: rest };
}

export function quoteSheet(name) {
  const text = String(name || "");
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(text) ? text : "'" + text.replace(/'/g, "''") + "'";
}

/* Column letters for a range, used by charts and labels. */
export function colRange(c1, c2) {
  return indexToCol(c1) + ":" + indexToCol(c2);
}

export function rowRange(r1, r2) {
  return r1 + ":" + r2;
}

/* Parse a sheet-qualified range such as Sheet2!A1:B3. */
export function parseSheetRange(text) {
  const { sheet, ref } = splitSheetRef(text);
  const range = parseRange(ref);
  return range ? { sheet, range } : null;
}
