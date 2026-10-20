/* Formula parsing and evaluation.
 *
 * Formulas are parsed into a small AST, then evaluated against the workbook
 * model.  Nothing here uses eval() on user input: the tokenizer, the
 * precedence climbing parser and the function table are all explicit.
 *
 * A formula cell keeps its source text ("=SUM(A1:A3)") in cell.f and the
 * cached result in cell.v.  The engine owns the dependency graph, recalculates
 * dependents after an edit and detects circular references before evaluation
 * so a cycle can never hang the UI.
 *
 * Whole-column and whole-row references (A:A, 3:3) are accepted, but they are
 * bounded to the used extent of the sheet before iteration.  A reference such
 * as A:A therefore sums the cells that exist rather than 100,000 blanks, and
 * an oversized explicit range is rejected with #NUM! instead of freezing the
 * renderer.
 */

import {
  MAX_ROWS,
  MAX_COLS,
  addr,
  parseAddr,
  colToIndex,
  indexToCol,
  splitSheetRef,
  rangeContains,
  rangeText,
} from "./units.js";
import {
  generalFormat,
  formatValue,
  parseUserInput,
  serialToParts,
  serialFromParts,
  dateToSerial,
} from "./format.js";

/* A bounded range keeps whole-column references usable on the 100,000-row
   model without materialising millions of empty cells. */
const MAX_RANGE_CELLS = 400000;
const VOLATILE_FUNCTIONS = new Set(["RAND", "RANDBETWEEN", "TODAY", "NOW"]);

export class FormulaError {
  constructor(code, message = "") {
    this.code = code;
    this.message = message || code;
  }

  toString() {
    return this.code;
  }
}

export function isError(value) {
  return value instanceof FormulaError;
}

function error(code, message = "") {
  return new FormulaError(code, message);
}

export function errorText(value) {
  if (value instanceof FormulaError) return value.code;
  if (typeof value === "string" && ERROR_CODES.includes(value.toUpperCase())) return value.toUpperCase();
  return "";
}

/* ------------------------------------------------------------------ tokens */
const ERROR_CODES = [
  "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#NULL!", "#CIRC!", "#SPILL!", "#CALC!",
];

const REF_PATTERN = /^\$?([A-Za-z]{1,3})\$?(\d+)/;
const IDENT_START = /[A-Za-z_\\]/;
const IDENT_BODY = /^[A-Za-z_\\][A-Za-z0-9_.\\]*/;

function syntax(message) {
  return new FormulaError("#VALUE!", message || "The formula contains a syntax error.");
}

function readQuotedSheet(text, index) {
  let name = "";
  let cursor = index + 1;
  while (cursor < text.length) {
    if (text[cursor] === "'") {
      if (text[cursor + 1] === "'") {
        name += "'";
        cursor += 2;
        continue;
      }
      return { name, end: cursor + 1 };
    }
    name += text[cursor];
    cursor += 1;
  }
  throw syntax("Unterminated sheet name.");
}

function tokenBoundary(text, end) {
  return end >= text.length || !/[A-Za-z0-9_.]/.test(text[end]);
}

function readAbsoluteColumn(text, index) {
  const match = /^\$([A-Za-z]{1,3})(?![A-Za-z0-9_.])/.exec(text.slice(index));
  if (!match) return null;
  return {
    token: { type: "name", value: match[1], sheet: "" },
    end: index + match[0].length,
  };
}

function readRefToken(text, index, sheet = "") {
  const match = REF_PATTERN.exec(text.slice(index));
  if (!match) return null;
  const end = index + match[0].length;
  if (!tokenBoundary(text, end)) return null;
  const col = colToIndex(match[1]);
  const row = parseInt(match[2], 10);
  if (!col || !row) return null;
  return {
    token: {
      type: "ref",
      sheet,
      row,
      col,
      rowAbs: match[0].includes("$", 1),
      colAbs: match[0].startsWith("$"),
    },
    end,
  };
}

export function tokenize(text) {
  const tokens = [];
  const source = String(text ?? "");
  let index = 0;

  while (index < source.length) {
    const char = source[index];

    if (char === " " || char === "\t" || char === "\r" || char === "\n") {
      index += 1;
      continue;
    }

    if (char === '"') {
      let value = "";
      index += 1;
      let closed = false;
      while (index < source.length) {
        if (source[index] === '"') {
          if (source[index + 1] === '"') {
            value += '"';
            index += 2;
            continue;
          }
          index += 1;
          closed = true;
          break;
        }
        value += source[index];
        index += 1;
      }
      if (!closed) throw syntax("Unterminated text value.");
      tokens.push({ type: "str", value });
      continue;
    }

    if (char === "#") {
      const upper = source.slice(index).toUpperCase();
      let found = "";
      for (const code of ERROR_CODES) {
        if (upper.startsWith(code) && code.length > found.length) found = code;
      }
      if (!found) throw syntax("Unknown error value.");
      tokens.push({ type: "err", value: found });
      index += found.length;
      continue;
    }

    if (char === "'") {
      const quoted = readQuotedSheet(source, index);
      index = quoted.end;
      if (source[index] !== "!") throw syntax("Expected ! after a sheet name.");
      index += 1;
      const ref = readRefToken(source, index, quoted.name);
      if (ref) {
        tokens.push(ref.token);
        index = ref.end;
        continue;
      }
      const absolute = readAbsoluteColumn(source, index);
      if (absolute) {
        tokens.push({ ...absolute.token, sheet: quoted.name });
        index = absolute.end;
        continue;
      }
      const ident = IDENT_BODY.exec(source.slice(index));
      if (!ident) throw syntax("Expected a reference after the sheet name.");
      tokens.push({ type: "name", value: ident[0], sheet: quoted.name });
      index += ident[0].length;
      continue;
    }

    if (char === "$" || IDENT_START.test(char)) {
      const ref = readRefToken(source, index);
      if (ref && source[ref.end] !== "!") {
        tokens.push(ref.token);
        index = ref.end;
        continue;
      }
      const absolute = readAbsoluteColumn(source, index);
      if (absolute) {
        tokens.push(absolute.token);
        index = absolute.end;
        continue;
      }
      const ident = IDENT_BODY.exec(source.slice(index));
      if (!ident) throw syntax("Unexpected character '$'.");
      const word = ident[0];
      let cursor = index + word.length;
      if (source[cursor] === "!") {
        cursor += 1;
        const inner = readRefToken(source, cursor, word);
        if (inner) {
          tokens.push(inner.token);
          index = inner.end;
          continue;
        }
        const absolute = readAbsoluteColumn(source, cursor);
        if (absolute) {
          tokens.push({ ...absolute.token, sheet: word });
          index = absolute.end;
          continue;
        }
        const name = IDENT_BODY.exec(source.slice(cursor));
        if (!name) throw syntax("Expected a reference after the sheet name.");
        tokens.push({ type: "name", value: name[0], sheet: word });
        index = cursor + name[0].length;
        continue;
      }
      let lookahead = cursor;
      while (lookahead < source.length && /\s/.test(source[lookahead])) lookahead += 1;
      if (source[lookahead] === "(") {
        tokens.push({ type: "func", value: word });
        index = lookahead;
        continue;
      }
      if (/^(TRUE|FALSE)$/i.test(word)) {
        tokens.push({ type: "bool", value: /^TRUE$/i.test(word) });
      } else {
        tokens.push({ type: "name", value: word, sheet: "" });
      }
      index = cursor;
      continue;
    }

    if (/[0-9.]/.test(char)) {
      const match = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/.exec(source.slice(index));
      if (!match) throw syntax("Invalid number.");
      const value = Number(match[0]);
      if (!Number.isFinite(value)) throw syntax("Invalid number.");
      tokens.push({ type: "num", value });
      index += match[0].length;
      continue;
    }

    if (char === "(") {
      tokens.push({ type: "lparen" });
      index += 1;
      continue;
    }
    if (char === ")") {
      tokens.push({ type: "rparen" });
      index += 1;
      continue;
    }
    if (char === "," || char === ";") {
      tokens.push({ type: "comma" });
      index += 1;
      continue;
    }
    if (char === "<" || char === ">" || char === "=") {
      const pair = source.slice(index, index + 2);
      if (pair === "<=" || pair === ">=" || pair === "<>") {
        tokens.push({ type: "op", value: pair });
        index += 2;
      } else {
        tokens.push({ type: "op", value: char });
        index += 1;
      }
      continue;
    }
    if ("+-*/^&%:".includes(char)) {
      tokens.push({ type: "op", value: char });
      index += 1;
      continue;
    }
    throw syntax(`Unexpected character '${char}'.`);
  }

  tokens.push({ type: "eof" });
  return tokens;
}

/* Rewrite the relative references of a formula.  Copied and filled formulas
   keep their structure while their A1 references move with the cell.  The
   scanner skips string literals, quoted sheet names and error values, and it
   leaves $-anchored row or column parts alone. */
export function translateFormula(formula, dRow, dCol) {
  const text = String(formula ?? "");
  if ((!dRow && !dCol) || !text) return text;
  let output = "";
  let index = 0;
  while (index < text.length) {
    const char = text[index];
    if (char === '"') {
      let end = index + 1;
      while (end < text.length) {
        if (text[end] === '"') {
          if (text[end + 1] === '"') {
            end += 2;
            continue;
          }
          end += 1;
          break;
        }
        end += 1;
      }
      output += text.slice(index, end);
      index = end;
      continue;
    }
    if (char === "'") {
      let end = index + 1;
      while (end < text.length) {
        if (text[end] === "'") {
          if (text[end + 1] === "'") {
            end += 2;
            continue;
          }
          end += 1;
          break;
        }
        end += 1;
      }
      output += text.slice(index, end);
      index = end;
      continue;
    }
    if (char === "#") {
      const match = /^#[A-Za-z0-9/!?]+/.exec(text.slice(index));
      if (match) {
        output += match[0].toUpperCase();
        index += match[0].length;
        continue;
      }
    }
    const match = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)/.exec(text.slice(index));
    if (match) {
      const end = index + match[0].length;
      const previous = index > 0 ? text[index - 1] : "";
      const next = text[end] || "";
      const boundaryBefore = !previous || !/[A-Za-z0-9_.$]/.test(previous);
      const boundaryAfter = !next || !/[A-Za-z0-9_.(!]/.test(next);
      if (boundaryBefore && boundaryAfter) {
        const row = parseInt(match[4], 10) + (match[1] ? 0 : dRow);
        const col = colToIndex(match[2]) + (match[3] ? 0 : dCol);
        if (row >= 1 && row <= MAX_ROWS && col >= 1 && col <= MAX_COLS) {
          output += match[1] + indexToCol(col) + match[3] + row;
          index = end;
          continue;
        }
      }
    }
    output += char;
    index += 1;
  }
  return output;
}

/* ------------------------------------------------------------------ parser */
class Parser {
  constructor(tokens) {
    this.tokens = tokens;
    this.pos = 0;
  }

  peek(offset = 0) {
    return this.tokens[this.pos + offset] || this.tokens[this.tokens.length - 1];
  }

  next() {
    const token = this.peek();
    this.pos += 1;
    return token;
  }

  checkType(type) {
    return this.peek().type === type;
  }

  matchType(type) {
    if (this.checkType(type)) {
      this.pos += 1;
      return true;
    }
    return false;
  }

  expectType(type, message = "") {
    if (!this.matchType(type)) throw syntax(message || "Missing closing parenthesis.");
  }

  checkOp(value) {
    const token = this.peek();
    return token.type === "op" && token.value === value;
  }

  matchOp(value) {
    if (this.checkOp(value)) {
      this.pos += 1;
      return true;
    }
    return false;
  }

  parse() {
    if (this.checkType("eof")) throw syntax("Empty formula.");
    const node = this.parseComparison();
    if (!this.checkType("eof")) throw syntax("Unexpected text after the formula.");
    return node;
  }

  parseComparison() {
    let node = this.parseConcat();
    while (this.checkOp("=") || this.checkOp("<>") || this.checkOp("<") || this.checkOp(">")
      || this.checkOp("<=") || this.checkOp(">=")) {
      const operator = this.next().value;
      node = { type: "bin", op: operator, left: node, right: this.parseConcat() };
    }
    return node;
  }

  parseConcat() {
    let node = this.parseAdditive();
    while (this.checkOp("&")) {
      this.next();
      node = { type: "bin", op: "&", left: node, right: this.parseAdditive() };
    }
    return node;
  }

  parseAdditive() {
    let node = this.parseMultiplicative();
    while (this.checkOp("+") || this.checkOp("-")) {
      const operator = this.next().value;
      node = { type: "bin", op: operator, left: node, right: this.parseMultiplicative() };
    }
    return node;
  }

  parseMultiplicative() {
    let node = this.parsePower();
    while (this.checkOp("*") || this.checkOp("/")) {
      const operator = this.next().value;
      node = { type: "bin", op: operator, left: node, right: this.parsePower() };
    }
    return node;
  }

  parsePower() {
    const left = this.parseUnary();
    if (this.checkOp("^")) {
      this.next();
      return { type: "bin", op: "^", left, right: this.parsePower() };
    }
    return left;
  }

  parseUnary() {
    if (this.checkOp("-") || this.checkOp("+")) {
      const operator = this.next().value;
      return { type: "un", op: operator, operand: this.parseUnary() };
    }
    return this.parsePostfix();
  }

  parsePostfix() {
    let node = this.parsePrimary();
    while (this.checkOp("%")) {
      this.next();
      node = { type: "pct", operand: node };
    }
    return node;
  }

  parsePrimary() {
    const token = this.peek();

    if (token.type === "num") {
      this.next();
      if (this.checkOp(":")) {
        this.next();
        const last = this.next();
        if (last.type !== "num" || !Number.isInteger(token.value) || !Number.isInteger(last.value)) {
          throw syntax("A row reference must use whole numbers.");
        }
        const r1 = Math.min(token.value, last.value);
        const r2 = Math.max(token.value, last.value);
        if (r1 < 1 || r2 > MAX_ROWS) throw syntax("The row reference is outside the sheet.");
        return { type: "range", sheet: "", r1, c1: 1, r2, c2: MAX_COLS, whole: "row" };
      }
      return { type: "num", value: token.value };
    }

    if (token.type === "str") {
      this.next();
      return { type: "str", value: token.value };
    }

    if (token.type === "bool") {
      this.next();
      return { type: "bool", value: token.value };
    }

    if (token.type === "err") {
      this.next();
      return { type: "err", value: token.value };
    }

    if (token.type === "ref") {
      this.next();
      if (this.checkOp(":")) {
        this.next();
        const last = this.next();
        if (last.type !== "ref") throw syntax("A range must end in a cell reference.");
        if (token.sheet && last.sheet && token.sheet.toLowerCase() !== last.sheet.toLowerCase()) {
          throw syntax("A range cannot span two sheets.");
        }
        return {
          type: "range",
          sheet: token.sheet || last.sheet,
          r1: Math.min(token.row, last.row),
          c1: Math.min(token.col, last.col),
          r2: Math.max(token.row, last.row),
          c2: Math.max(token.col, last.col),
          whole: "",
        };
      }
      return { type: "ref", sheet: token.sheet, row: token.row, col: token.col };
    }

    if (token.type === "name") {
      this.next();
      if (this.checkOp(":")) {
        this.next();
        const last = this.next();
        if (last.type !== "name") throw syntax("A column range must end in a column letter.");
        if (token.sheet && last.sheet && token.sheet.toLowerCase() !== last.sheet.toLowerCase()) {
          throw syntax("A range cannot span two sheets.");
        }
        const firstCol = columnIndex(token.value);
        const lastCol = columnIndex(last.value);
        if (!firstCol || !lastCol) throw syntax("Invalid column reference.");
        return {
          type: "range",
          sheet: token.sheet || last.sheet || "",
          r1: 1,
          c1: Math.min(firstCol, lastCol),
          r2: MAX_ROWS,
          c2: Math.max(firstCol, lastCol),
          whole: "col",
        };
      }
      return { type: "name", value: token.value, sheet: token.sheet || "" };
    }

    if (token.type === "func") {
      this.next();
      this.expectType("lparen", "Expected an opening parenthesis after a function name.");
      return { type: "call", name: token.value, args: this.parseArgs() };
    }

    if (token.type === "lparen") {
      this.next();
      const node = this.parseComparison();
      this.expectType("rparen", "Missing closing parenthesis.");
      return node;
    }

    throw syntax("Unexpected token in the formula.");
  }

  parseArgs() {
    const args = [];
    if (this.matchType("rparen")) return args;
    while (true) {
      if (this.checkType("rparen")) {
        args.push({ type: "empty" });
        this.next();
        break;
      }
      if (this.checkType("comma")) {
        this.next();
        args.push({ type: "empty" });
        continue;
      }
      args.push(this.parseComparison());
      if (this.matchType("comma")) continue;
      this.expectType("rparen", "Missing closing parenthesis.");
      break;
    }
    return args;
  }
}

function columnIndex(text) {
  const clean = String(text || "").replace(/\$/g, "");
  if (!/^[A-Za-z]{1,3}$/.test(clean)) return 0;
  const index = colToIndex(clean);
  return index >= 1 && index <= MAX_COLS ? index : 0;
}

export function parseFormula(text) {
  const source = String(text ?? "").trim();
  const body = source.startsWith("=") ? source.slice(1) : source;
  if (!body.trim()) throw syntax("The formula is empty.");
  const parser = new Parser(tokenize(body));
  return parser.parse();
}

/* Collect the references a formula depends on.  Sheets are resolved against
   the model so the graph can use canonical keys. */
export function collectPrecedents(ast, model, sheet) {
  const result = { refs: new Set(), ranges: [], volatile: false };
  const visit = (node) => {
    if (!node) return;
    switch (node.type) {
      case "ref": {
        const target = node.sheet ? findSheet(model, node.sheet) : sheet;
        if (target) result.refs.add(cellKey(target.name, addr(node.row, node.col)));
        break;
      }
      case "range": {
        const target = node.sheet ? findSheet(model, node.sheet) : sheet;
        if (target) result.ranges.push({ sheet: target.name, range: boundedRange(target, node) });
        break;
      }
      case "name": {
        const resolved = resolveNameRange(model, node.value, node.sheet || sheet?.name, sheet);
        if (resolved) {
          if (resolved.range.r1 === resolved.range.r2 && resolved.range.c1 === resolved.range.c2) {
            result.refs.add(cellKey(resolved.sheet.name, addr(resolved.range.r1, resolved.range.c1)));
          } else {
            result.ranges.push({ sheet: resolved.sheet.name, range: resolved.range });
          }
        }
        break;
      }
      case "call": {
        const name = normalizeFunction(node.name);
        if (VOLATILE_FUNCTIONS.has(name)) result.volatile = true;
        for (const arg of node.args) visit(arg);
        break;
      }
      case "bin":
        visit(node.left);
        visit(node.right);
        break;
      case "un":
      case "pct":
        visit(node.operand);
        break;
      default:
        break;
    }
  };
  visit(ast);
  return result;
}

/* --------------------------------------------------------------- utilities */
export function cellKey(sheetName, address) {
  return String(sheetName).toLowerCase() + "!" + String(address).toUpperCase();
}

export function keySheet(key) {
  const bang = key.lastIndexOf("!");
  return bang < 0 ? "" : key.slice(0, bang);
}

export function keyAddress(key) {
  const bang = key.lastIndexOf("!");
  return bang < 0 ? key : key.slice(bang + 1);
}

function normalizeFunction(name) {
  return String(name || "").toUpperCase().replace(/^_XLFN\./, "").replace(/^_XLWS\./, "");
}

export function findSheet(model, name) {
  const lower = String(name || "").toLowerCase();
  return (model.sheets || []).find((sheet) => sheet.name.toLowerCase() === lower) || null;
}

export function usedExtent(sheet) {
  let maxRow = 0;
  let maxCol = 0;
  for (const address of Object.keys(sheet.cells || {})) {
    const parsed = parseAddr(address);
    if (!parsed) continue;
    if (parsed.row > maxRow) maxRow = parsed.row;
    if (parsed.col > maxCol) maxCol = parsed.col;
  }
  return { maxRow, maxCol };
}

export function boundedRange(sheet, node) {
  if (node.whole === "col") {
    const extent = usedExtent(sheet);
    return {
      r1: 1,
      c1: node.c1,
      r2: Math.max(1, extent.maxRow),
      c2: node.c2,
    };
  }
  if (node.whole === "row") {
    const extent = usedExtent(sheet);
    return {
      r1: node.r1,
      c1: 1,
      r2: node.r2,
      c2: Math.max(1, extent.maxCol),
    };
  }
  return { r1: node.r1, c1: node.c1, r2: node.r2, c2: node.c2 };
}

/* A defined name can point at a range, a single cell, a whole column or a
   whole row.  Returning null lets the caller surface #NAME?. */
export function resolveNameRange(model, name, sheetName, fallbackSheet) {
  const wanted = String(name || "").toLowerCase();
  let definition = "";
  for (const [key, value] of Object.entries(model.names || {})) {
    if (key.toLowerCase() === wanted) {
      definition = String(value || "");
      break;
    }
  }
  if (!definition) return null;
  const text = definition.replace(/^=/, "").trim();
  const { sheet: prefix, ref } = splitSheetRef(text);
  const sheet = (prefix && findSheet(model, prefix))
    || (sheetName && findSheet(model, sheetName))
    || fallbackSheet;
  if (!sheet) return null;
  const clean = ref.replace(/\$/g, "");
  let match = /^([A-Za-z]{1,3}):([A-Za-z]{1,3})$/.exec(clean);
  if (match) {
    const firstCol = columnIndex(match[1]);
    const lastCol = columnIndex(match[2]);
    if (!firstCol || !lastCol) return null;
    const extent = usedExtent(sheet);
    return {
      sheet,
      range: {
        r1: 1,
        c1: Math.min(firstCol, lastCol),
        r2: Math.max(1, extent.maxRow),
        c2: Math.max(firstCol, lastCol),
      },
    };
  }
  match = /^(\d+):(\d+)$/.exec(clean);
  if (match) {
    const firstRow = parseInt(match[1], 10);
    const lastRow = parseInt(match[2], 10);
    if (firstRow < 1 || lastRow > MAX_ROWS) return null;
    const extent = usedExtent(sheet);
    return {
      sheet,
      range: {
        r1: Math.min(firstRow, lastRow),
        c1: 1,
        r2: Math.max(firstRow, lastRow),
        c2: Math.max(1, extent.maxCol),
      },
    };
  }
  const parts = clean.split(":");
  if (parts.length > 2) return null;
  const first = parseAddr(parts[0]);
  const last = parts.length === 2 ? parseAddr(parts[1]) : first;
  if (!first || !last) return null;
  return {
    sheet,
    range: {
      r1: Math.min(first.row, last.row),
      c1: Math.min(first.col, last.col),
      r2: Math.max(first.row, last.row),
      c2: Math.max(first.col, last.col),
    },
  };
}

/* --------------------------------------------------------- value coercion */
function boundsOf(value) {
  if (value && value.__range) return { sheet: value.sheet, range: value.range };
  if (value && value.__ref) {
    return {
      sheet: value.sheet,
      range: { r1: value.row, c1: value.col, r2: value.row, c2: value.col },
    };
  }
  return null;
}

function deref(value, ctx) {
  if (value instanceof FormulaError) throw value;
  if (value && value.__ref) return ctx.cellValue(value.sheet, value.row, value.col);
  if (value && value.__range) {
    const range = value.range;
    if (range.r1 === range.r2 && range.c1 === range.c2) {
      return ctx.cellValue(value.sheet, range.r1, range.c1);
    }
    throw error("#VALUE!", "A range cannot be used as a single value.");
  }
  return value;
}

function numberValue(value, ctx) {
  const item = deref(value, ctx);
  if (item instanceof FormulaError) throw item;
  if (item === null || item === undefined || item === "") return 0;
  if (typeof item === "number") return item;
  if (typeof item === "boolean") return item ? 1 : 0;
  if (typeof item === "string") {
    const text = item.trim();
    if (!text) return 0;
    const number = Number(text);
    if (Number.isFinite(number)) return number;
    throw error("#VALUE!");
  }
  throw error("#VALUE!");
}

function textValue(value, ctx) {
  const item = deref(value, ctx);
  if (item instanceof FormulaError) throw item;
  if (item === null || item === undefined) return "";
  if (typeof item === "boolean") return item ? "TRUE" : "FALSE";
  if (typeof item === "number") return generalFormat(item);
  return String(item);
}

function boolValue(value, ctx) {
  const item = deref(value, ctx);
  if (item instanceof FormulaError) throw item;
  if (item === null || item === undefined || item === "") return false;
  if (typeof item === "boolean") return item;
  if (typeof item === "number") return item !== 0;
  if (typeof item === "string") {
    if (/^true$/i.test(item)) return true;
    if (/^false$/i.test(item)) return false;
    throw error("#VALUE!");
  }
  throw error("#VALUE!");
}

function* cellsOf(value, ctx) {
  if (value instanceof FormulaError) throw value;
  const bounds = boundsOf(value);
  if (!bounds) {
    yield { row: 0, col: 0, value };
    return;
  }
  const range = bounds.range;
  const size = (range.r2 - range.r1 + 1) * (range.c2 - range.c1 + 1);
  if (size > MAX_RANGE_CELLS) throw error("#NUM!", "The range is too large to evaluate.");
  for (let row = range.r1; row <= range.r2; row += 1) {
    for (let col = range.c1; col <= range.c2; col += 1) {
      yield { row, col, value: ctx.cellValue(bounds.sheet, row, col) };
    }
  }
}

function isDirect(value) {
  return !(value && (value.__range || value.__ref));
}

function collectNumbers(args, ctx) {
  const numbers = [];
  for (const arg of args) {
    const direct = isDirect(arg);
    for (const cell of cellsOf(arg, ctx)) {
      const value = cell.value;
      if (value instanceof FormulaError) throw value;
      if (typeof value === "number") {
        numbers.push(value);
      } else if (direct && typeof value === "boolean") {
        numbers.push(value ? 1 : 0);
      } else if (direct && typeof value === "string" && value.trim()) {
        if (!Number.isFinite(Number(value))) throw error("#VALUE!");
        numbers.push(Number(value));
      }
    }
  }
  return numbers;
}

function collectText(args, ctx) {
  const parts = [];
  for (const arg of args) {
    for (const cell of cellsOf(arg, ctx)) {
      const value = cell.value;
      if (value instanceof FormulaError) throw value;
      if (value === null || value === undefined) continue;
      parts.push(typeof value === "boolean" ? (value ? "TRUE" : "FALSE")
        : typeof value === "number" ? generalFormat(value) : String(value));
    }
  }
  return parts.join("");
}

function compareScalars(a, b) {
  const emptyA = a === null || a === undefined || a === "";
  const emptyB = b === null || b === undefined || b === "";
  if (emptyA && emptyB) return 0;
  if (emptyA) {
    if (typeof b === "number") a = 0;
    else if (typeof b === "boolean") a = false;
    else a = "";
  }
  if (emptyB) {
    if (typeof a === "number") b = 0;
    else if (typeof a === "boolean") b = false;
    else b = "";
  }
  const rank = (value) => (typeof value === "number" ? 0 : typeof value === "string" ? 1 : 2);
  let left = a;
  let right = b;
  if (typeof left === "boolean" && typeof right === "number") left = left ? 1 : 0;
  if (typeof right === "boolean" && typeof left === "number") right = right ? 1 : 0;
  const leftRank = rank(left);
  const rightRank = rank(right);
  if (leftRank !== rightRank) return leftRank < rightRank ? -1 : 1;
  if (leftRank === 0) return left === right ? 0 : (left < right ? -1 : 1);
  if (leftRank === 2) return left === right ? 0 : (left ? 1 : -1);
  const leftText = String(left).toLowerCase();
  const rightText = String(right).toLowerCase();
  if (leftText === rightText) return 0;
  return leftText < rightText ? -1 : 1;
}

function compareValues(left, right, ctx) {
  const a = deref(left, ctx);
  const b = deref(right, ctx);
  if (a instanceof FormulaError) throw a;
  if (b instanceof FormulaError) throw b;
  return compareScalars(a, b);
}

function binaryResult(op, left, right, ctx) {
  if (op === "&") return textValue(left, ctx) + textValue(right, ctx);
  if (op === "=") return compareValues(left, right, ctx) === 0;
  if (op === "<>") return compareValues(left, right, ctx) !== 0;
  if (op === "<") return compareValues(left, right, ctx) < 0;
  if (op === ">") return compareValues(left, right, ctx) > 0;
  if (op === "<=") return compareValues(left, right, ctx) <= 0;
  if (op === ">=") return compareValues(left, right, ctx) >= 0;
  const a = numberValue(left, ctx);
  const b = numberValue(right, ctx);
  switch (op) {
    case "+": return a + b;
    case "-": return a - b;
    case "*": return a * b;
    case "/":
      if (b === 0) throw error("#DIV/0!");
      return a / b;
    case "^": {
      if (a === 0 && b === 0) throw error("#NUM!");
      const value = Math.pow(a, b);
      if (!Number.isFinite(value)) throw error("#NUM!");
      return value;
    }
    default:
      throw error("#VALUE!");
  }
}

/* ------------------------------------------------------------ criteria */
function wildcardToRegExp(pattern) {
  let source = "";
  for (let index = 0; index < pattern.length; index += 1) {
    const char = pattern[index];
    if (char === "~" && index + 1 < pattern.length) {
      source += pattern[index + 1].replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      index += 1;
    } else if (char === "*") {
      source += ".*";
    } else if (char === "?") {
      source += ".";
    } else {
      source += char.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    }
  }
  return new RegExp("^" + source + "$", "i");
}

function textMatches(value, pattern) {
  if (pattern.includes("*") || pattern.includes("?") || pattern.includes("~")) {
    return wildcardToRegExp(pattern).test(value);
  }
  return value.toLowerCase() === pattern.toLowerCase();
}

function criteriaMatcher(criteria, ctx) {
  let raw = deref(criteria, ctx);
  if (raw instanceof FormulaError) throw raw;
  let operator = "=";
  let target = raw;
  if (typeof raw === "string") {
    const match = /^(<=|>=|<>|<|>|=)\s*([\s\S]*)$/.exec(raw);
    if (match) {
      operator = match[1];
      target = match[2];
    }
  }
  const targetText = typeof target === "string" ? target : "";
  const numericTarget = typeof target === "number"
    || (typeof target === "string" && target.trim() !== "" && Number.isFinite(Number(target)));
  const number = numericTarget ? Number(target) : 0;
  const boolTarget = typeof target === "boolean" ? target : /^(true|false)$/i.test(targetText) ? /^true$/i.test(targetText) : null;

  return (value) => {
    if (value instanceof FormulaError) return false;
    if (operator === "=" && targetText === "") {
      return value === null || value === undefined || value === "";
    }
    if (operator === "<>" && targetText === "") {
      return !(value === null || value === undefined || value === "");
    }
    if (typeof value === "number" && numericTarget) {
      switch (operator) {
        case "=": return value === number;
        case "<>": return value !== number;
        case "<": return value < number;
        case ">": return value > number;
        case "<=": return value <= number;
        case ">=": return value >= number;
        default: return false;
      }
    }
    if (typeof value === "boolean" && boolTarget !== null) {
      const equal = value === boolTarget;
      return operator === "<>" ? !equal : equal;
    }
    if (numericTarget && operator === "<>") {
      return typeof value === "number" ? value !== number : value !== null && value !== undefined && value !== "";
    }
    if (typeof value === "string" && !numericTarget && boolTarget === null) {
      const equal = textMatches(value, targetText);
      if (operator === "=") return equal;
      if (operator === "<>") return !equal;
      const order = compareScalars(value, targetText);
      switch (operator) {
        case "<": return order < 0;
        case ">": return order > 0;
        case "<=": return order <= 0;
        case ">=": return order >= 0;
        default: return false;
      }
    }
    if ((value === null || value === undefined) && operator === "<>") return targetText !== "";
    return false;
  };
}

/* -------------------------------------------------------------- functions */
function fn(spec) {
  return spec;
}

const FUNCTIONS = {
  SUM: fn({ min: 0, max: 255, run: (args, ctx) => {
    let total = 0;
    for (const arg of args) {
      const direct = isDirect(arg);
      for (const cell of cellsOf(arg, ctx)) {
        const value = cell.value;
        if (value instanceof FormulaError) throw value;
        if (typeof value === "number") total += value;
        else if (direct && typeof value === "boolean") total += value ? 1 : 0;
        else if (direct && typeof value === "string" && value.trim()) {
          if (!Number.isFinite(Number(value))) throw error("#VALUE!");
          total += Number(value);
        }
      }
    }
    return total;
  } }),

  PRODUCT: fn({ min: 0, max: 255, run: (args, ctx) => {
    const numbers = collectNumbers(args, ctx);
    if (!numbers.length) return 0;
    return numbers.reduce((total, value) => total * value, 1);
  } }),

  ABS: fn({ min: 1, max: 1, run: (args, ctx) => Math.abs(numberValue(args[0], ctx)) }),
  SQRT: fn({ min: 1, max: 1, run: (args, ctx) => {
    const value = numberValue(args[0], ctx);
    if (value < 0) throw error("#NUM!");
    return Math.sqrt(value);
  } }),
  EXP: fn({ min: 1, max: 1, run: (args, ctx) => Math.exp(numberValue(args[0], ctx)) }),
  LN: fn({ min: 1, max: 1, run: (args, ctx) => {
    const value = numberValue(args[0], ctx);
    if (value <= 0) throw error("#NUM!");
    return Math.log(value);
  } }),
  LOG: fn({ min: 1, max: 2, run: (args, ctx) => {
    const value = numberValue(args[0], ctx);
    const base = args.length > 1 ? numberValue(args[1], ctx) : 10;
    if (value <= 0 || base <= 0 || base === 1) throw error("#NUM!");
    return Math.log(value) / Math.log(base);
  } }),
  LOG10: fn({ min: 1, max: 1, run: (args, ctx) => {
    const value = numberValue(args[0], ctx);
    if (value <= 0) throw error("#NUM!");
    return Math.log10(value);
  } }),
  POWER: fn({ min: 2, max: 2, run: (args, ctx) => {
    const base = numberValue(args[0], ctx);
    const exponent = numberValue(args[1], ctx);
    if (base === 0 && exponent === 0) throw error("#NUM!");
    const value = Math.pow(base, exponent);
    if (!Number.isFinite(value)) throw error("#NUM!");
    return value;
  } }),
  MOD: fn({ min: 2, max: 2, run: (args, ctx) => {
    const number = numberValue(args[0], ctx);
    const divisor = numberValue(args[1], ctx);
    if (divisor === 0) throw error("#DIV/0!");
    return number - divisor * Math.floor(number / divisor);
  } }),
  INT: fn({ min: 1, max: 1, run: (args, ctx) => Math.floor(numberValue(args[0], ctx)) }),
  ROUND: fn({ min: 2, max: 2, run: (args, ctx) => {
    const value = numberValue(args[0], ctx);
    const digits = Math.trunc(numberValue(args[1], ctx));
    return roundTo(value, digits);
  } }),
  ROUNDUP: fn({ min: 2, max: 2, run: (args, ctx) => {
    const value = numberValue(args[0], ctx);
    const digits = Math.trunc(numberValue(args[1], ctx));
    return roundDirection(value, digits, "up");
  } }),
  ROUNDDOWN: fn({ min: 2, max: 2, run: (args, ctx) => {
    const value = numberValue(args[0], ctx);
    const digits = Math.trunc(numberValue(args[1], ctx));
    return roundDirection(value, digits, "down");
  } }),
  RAND: fn({ min: 0, max: 0, run: () => Math.random() }),
  RANDBETWEEN: fn({ min: 2, max: 2, run: (args, ctx) => {
    let low = Math.ceil(numberValue(args[0], ctx));
    let high = Math.floor(numberValue(args[1], ctx));
    if (low > high) [low, high] = [high, low];
    return low + Math.floor(Math.random() * (high - low + 1));
  } }),

  SUMIF: fn({ min: 2, max: 3, run: (args, ctx) => {
    const criteriaRange = boundsOf(args[0]);
    if (!criteriaRange) throw error("#VALUE!");
    const matcher = criteriaMatcher(args[1], ctx);
    const sumRange = args.length > 2 ? boundsOf(args[2]) : criteriaRange;
    if (!sumRange) throw error("#VALUE!");
    const range = criteriaRange.range;
    const size = (range.r2 - range.r1 + 1) * (range.c2 - range.c1 + 1);
    if (size > MAX_RANGE_CELLS) throw error("#NUM!");
    let total = 0;
    for (let row = range.r1; row <= range.r2; row += 1) {
      for (let col = range.c1; col <= range.c2; col += 1) {
        const value = ctx.cellValue(criteriaRange.sheet, row, col);
        if (!matcher(value)) continue;
        const target = ctx.cellValue(
          sumRange.sheet,
          sumRange.range.r1 + (row - range.r1),
          sumRange.range.c1 + (col - range.c1),
        );
        if (target instanceof FormulaError) throw target;
        if (typeof target === "number") total += target;
      }
    }
    return total;
  } }),

  SUMIFS: fn({ min: 3, max: 255, run: (args, ctx) => {
    const sumRange = boundsOf(args[0]);
    if (!sumRange) throw error("#VALUE!");
    return sumIfsCore(sumRange, args.slice(1), ctx);
  } }),

  AVERAGEIF: fn({ min: 2, max: 3, run: (args, ctx) => {
    const bounds = boundsOf(args[0]);
    if (!bounds) throw error("#VALUE!");
    const matcher = criteriaMatcher(args[1], ctx);
    const averageRange = args.length > 2 ? boundsOf(args[2]) : bounds;
    if (!averageRange) throw error("#VALUE!");
    const range = bounds.range;
    const size = (range.r2 - range.r1 + 1) * (range.c2 - range.c1 + 1);
    if (size > MAX_RANGE_CELLS) throw error("#NUM!");
    let total = 0;
    let count = 0;
    for (let row = range.r1; row <= range.r2; row += 1) {
      for (let col = range.c1; col <= range.c2; col += 1) {
        if (!matcher(ctx.cellValue(bounds.sheet, row, col))) continue;
        const value = ctx.cellValue(
          averageRange.sheet,
          averageRange.range.r1 + (row - range.r1),
          averageRange.range.c1 + (col - range.c1),
        );
        if (value instanceof FormulaError) throw value;
        if (typeof value === "number") {
          total += value;
          count += 1;
        }
      }
    }
    if (!count) throw error("#DIV/0!");
    return total / count;
  } }),

  COUNTIF: fn({ min: 2, max: 2, run: (args, ctx) => {
    if (!boundsOf(args[0])) throw error("#VALUE!");
    const matcher = criteriaMatcher(args[1], ctx);
    let count = 0;
    for (const cell of cellsOf(args[0], ctx)) {
      if (matcher(cell.value)) count += 1;
    }
    return count;
  } }),

  COUNTIFS: fn({ min: 2, max: 255, run: (args, ctx) => {
    const pairs = [];
    for (let index = 0; index + 1 < args.length; index += 2) {
      const bounds = boundsOf(args[index]);
      if (!bounds) throw error("#VALUE!");
      pairs.push({ bounds, matcher: criteriaMatcher(args[index + 1], ctx) });
    }
    if (!pairs.length) throw error("#VALUE!");
    const first = pairs[0].bounds.range;
    const size = (first.r2 - first.r1 + 1) * (first.c2 - first.c1 + 1);
    if (size > MAX_RANGE_CELLS) throw error("#NUM!");
    let count = 0;
    for (let row = first.r1; row <= first.r2; row += 1) {
      for (let col = first.c1; col <= first.c2; col += 1) {
        let matched = true;
        for (const pair of pairs) {
          const value = ctx.cellValue(
            pair.bounds.sheet,
            pair.bounds.range.r1 + (row - first.r1),
            pair.bounds.range.c1 + (col - first.c1),
          );
          if (!pair.matcher(value)) {
            matched = false;
            break;
          }
        }
        if (matched) count += 1;
      }
    }
    return count;
  } }),

  AVERAGE: fn({ min: 1, max: 255, run: (args, ctx) => {
    const numbers = collectNumbers(args, ctx);
    if (!numbers.length) throw error("#DIV/0!");
    return numbers.reduce((total, value) => total + value, 0) / numbers.length;
  } }),

  COUNT: fn({ min: 0, max: 255, run: (args, ctx) => collectNumbers(args, ctx).length }),

  COUNTA: fn({ min: 0, max: 255, run: (args, ctx) => {
    let count = 0;
    for (const arg of args) {
      for (const cell of cellsOf(arg, ctx)) {
        const value = cell.value;
        if (value instanceof FormulaError || (value !== null && value !== undefined)) count += 1;
      }
    }
    return count;
  } }),

  COUNTBLANK: fn({ min: 1, max: 1, run: (args, ctx) => {
    let count = 0;
    for (const cell of cellsOf(args[0], ctx)) {
      const value = cell.value;
      if (value === null || value === undefined || value === "") count += 1;
    }
    return count;
  } }),

  MIN: fn({ min: 1, max: 255, run: (args, ctx) => {
    const numbers = collectNumbers(args, ctx);
    return numbers.length ? Math.min(...numbers) : 0;
  } }),
  MAX: fn({ min: 1, max: 255, run: (args, ctx) => {
    const numbers = collectNumbers(args, ctx);
    return numbers.length ? Math.max(...numbers) : 0;
  } }),
  MEDIAN: fn({ min: 1, max: 255, run: (args, ctx) => {
    const numbers = collectNumbers(args, ctx).sort((a, b) => a - b);
    if (!numbers.length) throw error("#NUM!");
    const middle = Math.floor(numbers.length / 2);
    return numbers.length % 2 ? numbers[middle] : (numbers[middle - 1] + numbers[middle]) / 2;
  } }),
  "STDEV.S": fn({ min: 1, max: 255, run: (args, ctx) => {
    const numbers = collectNumbers(args, ctx);
    if (numbers.length < 2) throw error("#DIV/0!");
    const mean = numbers.reduce((total, value) => total + value, 0) / numbers.length;
    const variance = numbers.reduce((total, value) => total + (value - mean) ** 2, 0) / (numbers.length - 1);
    return Math.sqrt(variance);
  } }),
  "VAR.S": fn({ min: 1, max: 255, run: (args, ctx) => {
    const numbers = collectNumbers(args, ctx);
    if (numbers.length < 2) throw error("#DIV/0!");
    const mean = numbers.reduce((total, value) => total + value, 0) / numbers.length;
    return numbers.reduce((total, value) => total + (value - mean) ** 2, 0) / (numbers.length - 1);
  } }),

  IF: fn({ min: 2, max: 3, lazy: true, run: (nodes, ctx) => {
    const condition = ctx.bool(ctx.eval(nodes[0]));
    if (condition) return chosenValue(nodes[1], ctx);
    if (nodes.length > 2) return chosenValue(nodes[2], ctx);
    return false;
  } }),

  IFS: fn({ min: 2, max: 255, lazy: true, run: (nodes, ctx) => {
    for (let index = 0; index + 1 < nodes.length; index += 2) {
      if (ctx.bool(ctx.eval(nodes[index]))) return chosenValue(nodes[index + 1], ctx);
    }
    throw error("#N/A");
  } }),

  AND: fn({ min: 1, max: 255, run: (args, ctx) => {
    let seen = false;
    let result = true;
    for (const arg of args) {
      for (const cell of cellsOf(arg, ctx)) {
        const value = cell.value;
        if (value instanceof FormulaError) throw value;
        if (typeof value === "boolean") {
          seen = true;
          result = result && value;
        } else if (typeof value === "number") {
          seen = true;
          result = result && value !== 0;
        }
      }
    }
    if (!seen) throw error("#VALUE!");
    return result;
  } }),

  OR: fn({ min: 1, max: 255, run: (args, ctx) => {
    let seen = false;
    let result = false;
    for (const arg of args) {
      for (const cell of cellsOf(arg, ctx)) {
        const value = cell.value;
        if (value instanceof FormulaError) throw value;
        if (typeof value === "boolean") {
          seen = true;
          result = result || value;
        } else if (typeof value === "number") {
          seen = true;
          result = result || value !== 0;
        }
      }
    }
    if (!seen) throw error("#VALUE!");
    return result;
  } }),

  NOT: fn({ min: 1, max: 1, run: (args, ctx) => !boolValue(args[0], ctx) }),

  IFERROR: fn({ min: 2, max: 2, lazy: true, run: (nodes, ctx) => {
    const value = ctx.tryValue(nodes[0]);
    if (value instanceof FormulaError) return chosenValue(nodes[1], ctx);
    return value;
  } }),

  IFNA: fn({ min: 2, max: 2, lazy: true, run: (nodes, ctx) => {
    const value = ctx.tryValue(nodes[0]);
    if (value instanceof FormulaError && value.code === "#N/A") return chosenValue(nodes[1], ctx);
    return value;
  } }),

  TRUE: fn({ min: 0, max: 0, run: () => true }),
  FALSE: fn({ min: 0, max: 0, run: () => false }),

  VLOOKUP: fn({ min: 3, max: 4, run: (args, ctx) => {
    const lookup = deref(args[0], ctx);
    if (lookup instanceof FormulaError) throw lookup;
    const table = boundsOf(args[1]);
    if (!table) throw error("#VALUE!");
    const column = Math.trunc(numberValue(args[2], ctx));
    const approximate = args.length > 3 ? boolValue(args[3], ctx) : true;
    const range = table.range;
    if (column < 1 || range.c1 + column - 1 > range.c2) throw error("#REF!");
    let foundRow = 0;
    let best = null;
    for (let row = range.r1; row <= range.r2; row += 1) {
      const value = ctx.cellValue(table.sheet, row, range.c1);
      if (value instanceof FormulaError) throw value;
      if (!approximate) {
        if (compareScalars(value, lookup) === 0) {
          foundRow = row;
          break;
        }
      } else if (compareScalars(value, lookup) <= 0) {
        if (best === null || compareScalars(value, best) > 0) {
          best = value;
          foundRow = row;
        }
      }
    }
    if (!foundRow) throw error("#N/A");
    return ctx.cellValue(table.sheet, foundRow, range.c1 + column - 1);
  } }),

  HLOOKUP: fn({ min: 3, max: 4, run: (args, ctx) => {
    const lookup = deref(args[0], ctx);
    if (lookup instanceof FormulaError) throw lookup;
    const table = boundsOf(args[1]);
    if (!table) throw error("#VALUE!");
    const rowOffset = Math.trunc(numberValue(args[2], ctx));
    const approximate = args.length > 3 ? boolValue(args[3], ctx) : true;
    const range = table.range;
    if (rowOffset < 1 || range.r1 + rowOffset - 1 > range.r2) throw error("#REF!");
    let foundCol = 0;
    let best = null;
    for (let col = range.c1; col <= range.c2; col += 1) {
      const value = ctx.cellValue(table.sheet, range.r1, col);
      if (value instanceof FormulaError) throw value;
      if (!approximate) {
        if (compareScalars(value, lookup) === 0) {
          foundCol = col;
          break;
        }
      } else if (compareScalars(value, lookup) <= 0) {
        if (best === null || compareScalars(value, best) > 0) {
          best = value;
          foundCol = col;
        }
      }
    }
    if (!foundCol) throw error("#N/A");
    return ctx.cellValue(table.sheet, range.r1 + rowOffset - 1, foundCol);
  } }),

  INDEX: fn({ min: 2, max: 3, run: (args, ctx) => {
    const bounds = boundsOf(args[0]);
    if (!bounds) throw error("#VALUE!");
    const range = bounds.range;
    const rowNumber = Math.trunc(numberValue(args[1], ctx));
    const columnNumber = args.length > 2 ? Math.trunc(numberValue(args[2], ctx)) : 0;
    if (rowNumber < 0 || columnNumber < 0) throw error("#VALUE!");
    if (rowNumber > range.r2 - range.r1 + 1 || columnNumber > range.c2 - range.c1 + 1) throw error("#REF!");
    if (rowNumber === 0 && columnNumber === 0) {
      return { __range: true, sheet: bounds.sheet, range: { ...range }, whole: "" };
    }
    if (rowNumber === 0) {
      return {
        __range: true,
        sheet: bounds.sheet,
        range: { r1: range.r1, c1: range.c1 + columnNumber - 1, r2: range.r2, c2: range.c1 + columnNumber - 1 },
        whole: "",
      };
    }
    if (columnNumber === 0) {
      if (range.c1 !== range.c2) {
        return {
          __range: true,
          sheet: bounds.sheet,
          range: { r1: range.r1 + rowNumber - 1, c1: range.c1, r2: range.r1 + rowNumber - 1, c2: range.c2 },
          whole: "",
        };
      }
      return ctx.cellValue(bounds.sheet, range.r1 + rowNumber - 1, range.c1);
    }
    return ctx.cellValue(bounds.sheet, range.r1 + rowNumber - 1, range.c1 + columnNumber - 1);
  } }),

  MATCH: fn({ min: 2, max: 3, run: (args, ctx) => {
    const lookup = deref(args[0], ctx);
    if (lookup instanceof FormulaError) throw lookup;
    const bounds = boundsOf(args[1]);
    if (!bounds) throw error("#N/A");
    const range = bounds.range;
    if (range.r1 !== range.r2 && range.c1 !== range.c2) throw error("#N/A");
    const matchType = args.length > 2 ? Math.trunc(numberValue(args[2], ctx)) : 1;
    let index = 0;
    let bestIndex = 0;
    let best = null;
    for (let row = range.r1; row <= range.r2; row += 1) {
      for (let col = range.c1; col <= range.c2; col += 1) {
        index += 1;
        const value = ctx.cellValue(bounds.sheet, row, col);
        if (value instanceof FormulaError) throw value;
        if (matchType === 0) {
          if (typeof lookup === "string" && typeof value === "string" && /[*?~]/.test(lookup)) {
            if (textMatches(value, lookup)) return index;
          } else if (compareScalars(value, lookup) === 0) {
            return index;
          }
        } else if (matchType > 0) {
          if (compareScalars(value, lookup) <= 0 && (best === null || compareScalars(value, best) > 0)) {
            best = value;
            bestIndex = index;
          }
        } else if (compareScalars(value, lookup) >= 0 && (best === null || compareScalars(value, best) < 0)) {
          best = value;
          bestIndex = index;
        }
      }
    }
    if (!bestIndex) throw error("#N/A");
    return bestIndex;
  } }),

  XLOOKUP: fn({ min: 3, max: 6, lazy: true, run: (nodes, ctx) => {
    const lookup = ctx.eval(nodes[0]);
    const lookupValue = deref(lookup, ctx);
    if (lookupValue instanceof FormulaError) throw lookupValue;
    const lookupBounds = boundsOf(ctx.eval(nodes[1]));
    const returnBounds = boundsOf(ctx.eval(nodes[2]));
    if (!lookupBounds || !returnBounds) throw error("#VALUE!");
    const lookupRange = lookupBounds.range;
    const returnRange = returnBounds.range;
    const byRow = lookupRange.c1 === lookupRange.c2;
    const lookupSize = byRow ? lookupRange.r2 - lookupRange.r1 + 1 : lookupRange.c2 - lookupRange.c1 + 1;
    const returnSize = byRow ? returnRange.r2 - returnRange.r1 + 1 : returnRange.c2 - returnRange.c1 + 1;
    if (lookupSize !== returnSize) throw error("#VALUE!");
    const matchMode = nodes.length > 4 && nodes[4].type !== "empty" ? Math.trunc(numberValue(ctx.eval(nodes[4]), ctx)) : 0;
    const searchMode = nodes.length > 5 && nodes[5].type !== "empty" ? Math.trunc(numberValue(ctx.eval(nodes[5]), ctx)) : 1;
    const reverse = searchMode === -1;
    let bestIndex = -1;
    let best = null;
    for (let offset = 0; offset < lookupSize; offset += 1) {
      const position = reverse ? lookupSize - 1 - offset : offset;
      const value = byRow
        ? ctx.cellValue(lookupBounds.sheet, lookupRange.r1 + position, lookupRange.c1)
        : ctx.cellValue(lookupBounds.sheet, lookupRange.r1, lookupRange.c1 + position);
      if (value instanceof FormulaError) throw value;
      const order = compareScalars(value, lookupValue);
      if (order === 0) {
        bestIndex = position;
        best = value;
        break;
      }
      if (matchMode === -1 && order < 0) {
        if (best === null || compareScalars(value, best) > 0) {
          best = value;
          bestIndex = position;
        }
      } else if (matchMode === 1 && order > 0) {
        if (best === null || compareScalars(value, best) < 0) {
          best = value;
          bestIndex = position;
        }
      } else if (matchMode === 2 && typeof value === "string" && typeof lookupValue === "string"
        && textMatches(value, lookupValue)) {
        bestIndex = position;
        break;
      }
    }
    if (bestIndex < 0) {
      if (nodes.length > 3 && nodes[3].type !== "empty") return ctx.eval(nodes[3]);
      throw error("#N/A");
    }
    if (byRow && returnRange.c1 !== returnRange.c2) {
      return {
        __range: true,
        sheet: returnBounds.sheet,
        range: {
          r1: returnRange.r1 + bestIndex,
          c1: returnRange.c1,
          r2: returnRange.r1 + bestIndex,
          c2: returnRange.c2,
        },
        whole: "",
      };
    }
    if (!byRow && returnRange.r1 !== returnRange.r2) {
      return {
        __range: true,
        sheet: returnBounds.sheet,
        range: {
          r1: returnRange.r1,
          c1: returnRange.c1 + bestIndex,
          r2: returnRange.r2,
          c2: returnRange.c1 + bestIndex,
        },
        whole: "",
      };
    }
    return byRow
      ? ctx.cellValue(returnBounds.sheet, returnRange.r1 + bestIndex, returnRange.c1)
      : ctx.cellValue(returnBounds.sheet, returnRange.r1, returnRange.c1 + bestIndex);
  } }),

  CHOOSE: fn({ min: 2, max: 255, lazy: true, run: (nodes, ctx) => {
    const index = Math.trunc(numberValue(ctx.eval(nodes[0]), ctx));
    if (index < 1 || index > nodes.length - 1) throw error("#VALUE!");
    return chosenValue(nodes[index], ctx);
  } }),

  CONCAT: fn({ min: 1, max: 255, run: (args, ctx) => collectText(args, ctx) }),
  CONCATENATE: fn({ min: 1, max: 255, run: (args, ctx) => collectText(args, ctx) }),

  TEXTJOIN: fn({ min: 3, max: 255, run: (args, ctx) => {
    const delimiter = textValue(args[0], ctx);
    const ignoreEmpty = boolValue(args[1], ctx);
    const parts = [];
    for (const arg of args.slice(2)) {
      for (const cell of cellsOf(arg, ctx)) {
        const value = cell.value;
        if (value instanceof FormulaError) throw value;
        const text = typeof value === "boolean" ? (value ? "TRUE" : "FALSE")
          : typeof value === "number" ? generalFormat(value)
            : value === null || value === undefined ? "" : String(value);
        if (ignoreEmpty && text === "") continue;
        parts.push(text);
      }
    }
    return parts.join(delimiter);
  } }),

  LEFT: fn({ min: 1, max: 2, run: (args, ctx) => {
    const text = textValue(args[0], ctx);
    const count = args.length > 1 ? Math.trunc(numberValue(args[1], ctx)) : 1;
    if (count < 0) throw error("#VALUE!");
    return text.slice(0, count);
  } }),
  RIGHT: fn({ min: 1, max: 2, run: (args, ctx) => {
    const text = textValue(args[0], ctx);
    const count = args.length > 1 ? Math.trunc(numberValue(args[1], ctx)) : 1;
    if (count < 0) throw error("#VALUE!");
    return count === 0 ? "" : text.slice(-count);
  } }),
  MID: fn({ min: 3, max: 3, run: (args, ctx) => {
    const text = textValue(args[0], ctx);
    const start = Math.trunc(numberValue(args[1], ctx));
    const count = Math.trunc(numberValue(args[2], ctx));
    if (start < 1 || count < 0) throw error("#VALUE!");
    return text.slice(start - 1, start - 1 + count);
  } }),
  LEN: fn({ min: 1, max: 1, run: (args, ctx) => textValue(args[0], ctx).length }),
  LOWER: fn({ min: 1, max: 1, run: (args, ctx) => textValue(args[0], ctx).toLowerCase() }),
  UPPER: fn({ min: 1, max: 1, run: (args, ctx) => textValue(args[0], ctx).toUpperCase() }),
  PROPER: fn({ min: 1, max: 1, run: (args, ctx) => textValue(args[0], ctx).toLowerCase().replace(/(^|[^A-Za-z])([A-Za-z])/g, (match, lead, letter) => lead + letter.toUpperCase()) }),
  TRIM: fn({ min: 1, max: 1, run: (args, ctx) => textValue(args[0], ctx).replace(/\s+/g, " ").trim() }),

  SUBSTITUTE: fn({ min: 3, max: 4, run: (args, ctx) => {
    const text = textValue(args[0], ctx);
    const oldText = textValue(args[1], ctx);
    const newText = textValue(args[2], ctx);
    if (oldText === "") return text;
    if (args.length > 3) {
      const instance = Math.trunc(numberValue(args[3], ctx));
      if (instance < 1) throw error("#VALUE!");
      let seen = 0;
      return text.replace(new RegExp(oldText.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "g"), (match) => {
        seen += 1;
        return seen === instance ? newText : match;
      });
    }
    return text.split(oldText).join(newText);
  } }),

  REPLACE: fn({ min: 4, max: 4, run: (args, ctx) => {
    const text = textValue(args[0], ctx);
    const start = Math.trunc(numberValue(args[1], ctx));
    const count = Math.trunc(numberValue(args[2], ctx));
    const replacement = textValue(args[3], ctx);
    if (start < 1 || count < 0) throw error("#VALUE!");
    return text.slice(0, start - 1) + replacement + text.slice(start - 1 + count);
  } }),

  FIND: fn({ min: 2, max: 3, run: (args, ctx) => {
    const needle = textValue(args[0], ctx);
    const haystack = textValue(args[1], ctx);
    const start = args.length > 2 ? Math.trunc(numberValue(args[2], ctx)) : 1;
    if (start < 1 || start > haystack.length + 1) throw error("#VALUE!");
    const found = haystack.indexOf(needle, start - 1);
    if (found < 0) throw error("#VALUE!");
    return found + 1;
  } }),

  SEARCH: fn({ min: 2, max: 3, run: (args, ctx) => {
    const needle = textValue(args[0], ctx);
    const haystack = textValue(args[1], ctx);
    const start = args.length > 2 ? Math.trunc(numberValue(args[2], ctx)) : 1;
    if (start < 1 || start > haystack.length + 1) throw error("#VALUE!");
    const source = needle.split("").map((char) => {
      if (char === "*") return ".*";
      if (char === "?") return ".";
      return char.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    }).join("");
    const match = new RegExp(source, "i").exec(haystack.slice(start - 1));
    if (!match) throw error("#VALUE!");
    return start + match.index;
  } }),

  TEXT: fn({ min: 2, max: 2, run: (args, ctx) => {
    const value = deref(args[0], ctx);
    if (value instanceof FormulaError) throw value;
    const code = textValue(args[1], ctx);
    if (typeof value === "number") return formatValue(value, code);
    if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
    return value === null || value === undefined ? "" : String(value);
  } }),

  VALUE: fn({ min: 1, max: 1, run: (args, ctx) => {
    const text = textValue(args[0], ctx).trim();
    const parsed = parseUserInput(text);
    if (parsed.kind === "number") return parsed.value;
    if (parsed.kind === "boolean") return parsed.value ? 1 : 0;
    throw error("#VALUE!");
  } }),

  TODAY: fn({ min: 0, max: 0, run: () => {
    const now = new Date();
    return serialFromParts(now.getFullYear(), now.getMonth() + 1, now.getDate());
  } }),
  NOW: fn({ min: 0, max: 0, run: () => dateToSerial(new Date()) }),

  DATE: fn({ min: 3, max: 3, run: (args, ctx) => {
    const year = Math.trunc(numberValue(args[0], ctx));
    const month = Math.trunc(numberValue(args[1], ctx));
    const day = Math.trunc(numberValue(args[2], ctx));
    if (year < 0 || year > 9999) throw error("#NUM!");
    return serialFromParts(year, month, day);
  } }),

  YEAR: fn({ min: 1, max: 1, run: (args, ctx) => serialToParts(numberValue(args[0], ctx)).year }),
  MONTH: fn({ min: 1, max: 1, run: (args, ctx) => serialToParts(numberValue(args[0], ctx)).month }),
  DAY: fn({ min: 1, max: 1, run: (args, ctx) => serialToParts(numberValue(args[0], ctx)).day }),
  HOUR: fn({ min: 1, max: 1, run: (args, ctx) => serialToParts(numberValue(args[0], ctx)).hours }),
  MINUTE: fn({ min: 1, max: 1, run: (args, ctx) => serialToParts(numberValue(args[0], ctx)).minutes }),
  SECOND: fn({ min: 1, max: 1, run: (args, ctx) => serialToParts(numberValue(args[0], ctx)).seconds }),

  DAYS: fn({ min: 2, max: 2, run: (args, ctx) => {
    const end = numberValue(args[0], ctx);
    const start = numberValue(args[1], ctx);
    return Math.trunc(end) - Math.trunc(start);
  } }),

  EDATE: fn({ min: 2, max: 2, run: (args, ctx) => {
    const serial = numberValue(args[0], ctx);
    const months = Math.trunc(numberValue(args[1], ctx));
    const parts = serialToParts(serial);
    const target = new Date(Date.UTC(parts.year, parts.month - 1 + months, 1));
    const year = target.getUTCFullYear();
    const month = target.getUTCMonth() + 1;
    const lastDay = new Date(Date.UTC(year, month, 0)).getUTCDate();
    return serialFromParts(year, month, Math.min(parts.day, lastDay));
  } }),

  EOMONTH: fn({ min: 2, max: 2, run: (args, ctx) => {
    const serial = numberValue(args[0], ctx);
    const months = Math.trunc(numberValue(args[1], ctx));
    const parts = serialToParts(serial);
    const target = new Date(Date.UTC(parts.year, parts.month - 1 + months + 1, 0));
    return serialFromParts(target.getUTCFullYear(), target.getUTCMonth() + 1, target.getUTCDate());
  } }),

  WEEKDAY: fn({ min: 1, max: 2, run: (args, ctx) => {
    const serial = numberValue(args[0], ctx);
    const type = args.length > 1 ? Math.trunc(numberValue(args[1], ctx)) : 1;
    const weekday = serialToParts(serial).weekday;
    if (type >= 11 && type <= 17) {
      const start = type - 10;
      const day = weekday === 0 ? 7 : weekday;
      return ((day - start + 7) % 7) + 1;
    }
    switch (type) {
      case 1: return weekday + 1;
      case 2: return weekday === 0 ? 7 : weekday;
      case 3: return weekday === 0 ? 6 : weekday - 1;
      default: throw error("#NUM!");
    }
  } }),

  ROW: fn({ min: 0, max: 1, run: (args, ctx) => {
    if (!args.length) {
      const parsed = parseAddr(keyAddress(ctx.key));
      return parsed ? parsed.row : 0;
    }
    const bounds = boundsOf(args[0]);
    return bounds ? bounds.range.r1 : 0;
  } }),
  COLUMN: fn({ min: 0, max: 1, run: (args, ctx) => {
    if (!args.length) {
      const parsed = parseAddr(keyAddress(ctx.key));
      return parsed ? parsed.col : 0;
    }
    const bounds = boundsOf(args[0]);
    return bounds ? bounds.range.c1 : 0;
  } }),
  ROWS: fn({ min: 1, max: 1, run: (args) => {
    const bounds = boundsOf(args[0]);
    if (!bounds) throw error("#VALUE!");
    return bounds.range.r2 - bounds.range.r1 + 1;
  } }),
  COLUMNS: fn({ min: 1, max: 1, run: (args) => {
    const bounds = boundsOf(args[0]);
    if (!bounds) throw error("#VALUE!");
    return bounds.range.c2 - bounds.range.c1 + 1;
  } }),

  /* ------------------------------------------------------------ financial */
  PMT: fn({ min: 3, max: 5, run: (args, ctx) => {
    const rate = numberValue(args[0], ctx);
    const nper = numberValue(args[1], ctx);
    const pv = numberValue(args[2], ctx);
    const fv = args.length > 3 ? numberValue(args[3], ctx) : 0;
    const type = args.length > 4 && numberValue(args[4], ctx) ? 1 : 0;
    if (nper === 0) throw error("#NUM!");
    if (rate === 0) return -(pv + fv) / nper;
    const growth = Math.pow(1 + rate, nper);
    if (!Number.isFinite(growth)) throw error("#NUM!");
    return -(pv * growth + fv) * rate / ((growth - 1) * (1 + rate * type));
  } }),
  FV: fn({ min: 3, max: 5, run: (args, ctx) => {
    const rate = numberValue(args[0], ctx);
    const nper = numberValue(args[1], ctx);
    const payment = numberValue(args[2], ctx);
    const pv = args.length > 3 ? numberValue(args[3], ctx) : 0;
    const type = args.length > 4 && numberValue(args[4], ctx) ? 1 : 0;
    if (rate === 0) return -(pv + payment * nper);
    const growth = Math.pow(1 + rate, nper);
    if (!Number.isFinite(growth)) throw error("#NUM!");
    return -(pv * growth + payment * (1 + rate * type) * (growth - 1) / rate);
  } }),
  PV: fn({ min: 3, max: 5, run: (args, ctx) => {
    const rate = numberValue(args[0], ctx);
    const nper = numberValue(args[1], ctx);
    const payment = numberValue(args[2], ctx);
    const fv = args.length > 3 ? numberValue(args[3], ctx) : 0;
    const type = args.length > 4 && numberValue(args[4], ctx) ? 1 : 0;
    if (rate === 0) return -(fv + payment * nper);
    const growth = Math.pow(1 + rate, nper);
    if (!Number.isFinite(growth) || growth === 0) throw error("#NUM!");
    return -(fv + payment * (1 + rate * type) * (growth - 1) / rate) / growth;
  } }),
  NPV: fn({ min: 2, max: 255, run: (args, ctx) => {
    const rate = numberValue(args[0], ctx);
    if (rate <= -1) throw error("#NUM!");
    let total = 0;
    collectNumbers(args.slice(1), ctx).forEach((value, index) => {
      total += value / Math.pow(1 + rate, index + 1);
    });
    return total;
  } }),
  IRR: fn({ min: 1, max: 2, run: (args, ctx) => {
    const flows = collectNumbers([args[0]], ctx);
    if (flows.length < 2) throw error("#NUM!");
    if (!flows.some((value) => value > 0) || !flows.some((value) => value < 0)) throw error("#NUM!");
    const npvAt = (rate) => flows.reduce(
      (sum, value, index) => sum + value / Math.pow(1 + rate, index),
      0,
    );
    let rate = args.length > 1 ? numberValue(args[1], ctx) : 0.1;
    if (!Number.isFinite(rate)) rate = 0.1;
    let converged = false;
    for (let iteration = 0; iteration < 100; iteration += 1) {
      const value = npvAt(rate);
      const delta = 1e-6;
      const derivative = (npvAt(rate + delta) - npvAt(rate - delta)) / (2 * delta);
      if (!Number.isFinite(derivative) || Math.abs(derivative) < 1e-12) break;
      const next = rate - value / derivative;
      if (!Number.isFinite(next)) break;
      if (Math.abs(next - rate) < 1e-9) {
        rate = next;
        converged = true;
        break;
      }
      rate = next;
    }
    if (!converged || !Number.isFinite(rate) || Math.abs(npvAt(rate)) > 1e-5) throw error("#NUM!");
    return rate;
  } }),
};

export const FUNCTION_NAMES = Object.keys(FUNCTIONS);

export function functionArity(name) {
  const spec = FUNCTIONS[String(name || "").trim().toUpperCase()];
  if (!spec) return { min: 0, max: 255 };
  return { min: spec.min, max: spec.max };
}

function chosenValue(node, ctx) {
  if (!node || node.type === "empty") return 0;
  return ctx.eval(node);
}

function roundTo(value, digits) {
  const places = Math.max(-12, Math.min(12, digits));
  const factor = 10 ** places;
  const scaled = value * factor;
  return Math.round(scaled + (scaled >= 0 ? 1e-9 : -1e-9)) / factor;
}

function roundDirection(value, digits, direction) {
  const factor = 10 ** Math.max(-12, Math.min(12, digits));
  const scaled = value * factor;
  const rounded = direction === "up"
    ? (scaled >= 0 ? Math.ceil(scaled - 1e-9) : Math.floor(scaled + 1e-9))
    : (scaled >= 0 ? Math.floor(scaled + 1e-9) : Math.ceil(scaled - 1e-9));
  return rounded / factor;
}

function sumIfsCore(sumRange, pairs, ctx) {
  const parsed = [];
  for (let index = 0; index + 1 < pairs.length; index += 2) {
    const bounds = boundsOf(pairs[index]);
    if (!bounds) throw error("#VALUE!");
    parsed.push({ bounds, matcher: criteriaMatcher(pairs[index + 1], ctx) });
  }
  if (!parsed.length) throw error("#VALUE!");
  const first = parsed[0].bounds.range;
  const size = (first.r2 - first.r1 + 1) * (first.c2 - first.c1 + 1);
  if (size > MAX_RANGE_CELLS) throw error("#NUM!");
  let total = 0;
  for (let row = first.r1; row <= first.r2; row += 1) {
    for (let col = first.c1; col <= first.c2; col += 1) {
      let matched = true;
      for (const pair of parsed) {
        const value = ctx.cellValue(
          pair.bounds.sheet,
          pair.bounds.range.r1 + (row - first.r1),
          pair.bounds.range.c1 + (col - first.c1),
        );
        if (!pair.matcher(value)) {
          matched = false;
          break;
        }
      }
      if (!matched) continue;
      const value = ctx.cellValue(
        sumRange.sheet,
        sumRange.range.r1 + (row - first.r1),
        sumRange.range.c1 + (col - first.c1),
      );
      if (value instanceof FormulaError) throw value;
      if (typeof value === "number") total += value;
    }
  }
  return total;
}

/* ------------------------------------------------------------- the engine */
export class FormulaEngine {
  constructor(model) {
    this.model = model;
    this.astCache = new Map();
    this.parsed = new Map();
    this.precise = new Map();
    this.dependents = new Map();
    this.rangeDeps = new Map();
    this.volatile = new Set();
    this.cycle = new Set();
    this.values = new Map();
    this.states = new Map();
    this.formulaBySheet = new Map();
    this.extents = new Map();
    this.rebuild();
  }

  setModel(model) {
    this.model = model;
    this.rebuild();
  }

  /* Parse every formula, rebuild the dependency graph and recalculate the
     workbook.  Called on load, after a formula edit and after structural
     edits that move cells. */
  rebuild() {
    this.astCache.clear();
    this.parsed.clear();
    this.precise.clear();
    this.dependents.clear();
    this.rangeDeps.clear();
    this.volatile.clear();
    this.cycle.clear();
    this.values.clear();
    this.states.clear();
    this.formulaBySheet.clear();
    this.extents.clear();
    if (!this.model || !Array.isArray(this.model.sheets)) return 0;

    for (const sheet of this.model.sheets) {
      this.extents.set(sheet.name.toLowerCase(), usedExtent(sheet));
    }

    for (const sheet of this.model.sheets) {
      const formulaCells = [];
      for (const [address, cell] of Object.entries(sheet.cells || {})) {
        if (!cell || !cell.f) continue;
        const position = parseAddr(address);
        if (!position) continue;
        const key = cellKey(sheet.name, address);
        formulaCells.push({ key, row: position.row, col: position.col });
        const parsed = this.parseCellFormula(cell.f);
        this.parsed.set(key, parsed);
        if (!parsed.ast) continue;
        const deps = collectPrecedents(parsed.ast, this.model, sheet);
        this.precise.set(key, deps);
        for (const ref of deps.refs) this.addDependent(ref, key);
        for (const entry of deps.ranges) {
          const list = this.rangeDeps.get(entry.sheet.toLowerCase()) || [];
          list.push({ key, sheet: entry.sheet, range: entry.range });
          this.rangeDeps.set(entry.sheet.toLowerCase(), list);
        }
        if (deps.volatile) this.volatile.add(key);
      }
      this.formulaBySheet.set(sheet.name.toLowerCase(), formulaCells);
    }

    this.detectCycles();
    return this.recalculateAll();
  }

  addDependent(sourceKey, dependentKey) {
    const set = this.dependents.get(sourceKey) || new Set();
    set.add(dependentKey);
    this.dependents.set(sourceKey, set);
  }

  parseCellFormula(text) {
    const source = String(text || "");
    if (this.astCache.has(source)) return this.astCache.get(source);
    let result;
    try {
      result = { ast: parseFormula(source), error: "", code: "" };
    } catch (failure) {
      result = {
        ast: null,
        error: failure?.message || "The formula contains a syntax error.",
        code: failure instanceof FormulaError ? failure.code : "#VALUE!",
      };
    }
    this.astCache.set(source, result);
    return result;
  }

  /* Static cycle detection: an edge exists between two formula cells when the
     first reads the second, including reads through ranges.  Marking cycles
     before evaluation keeps IFERROR from hiding a self-reference and keeps a
     cycle from recursing forever. */
  detectCycles() {
    this.cycle.clear();
    const color = new Map();
    const stack = [];
    const visit = (key) => {
      color.set(key, 1);
      stack.push(key);
      const deps = this.precise.get(key);
      const edges = new Set();
      if (deps) {
        for (const ref of deps.refs) {
          if (this.parsed.has(ref)) edges.add(ref);
        }
        for (const entry of deps.ranges) {
          const cells = this.formulaBySheet.get(entry.sheet.toLowerCase()) || [];
          for (const item of cells) {
            if (rangeContains(entry.range, item.row, item.col)) edges.add(item.key);
          }
        }
      }
      for (const next of edges) {
        const state = color.get(next) || 0;
        if (state === 1) {
          const start = stack.indexOf(next);
          for (let index = Math.max(0, start); index < stack.length; index += 1) {
            this.cycle.add(stack[index]);
          }
        } else if (state === 0) {
          visit(next);
        }
      }
      stack.pop();
      color.set(key, 2);
    };
    for (const key of this.parsed.keys()) {
      if (!color.has(key)) visit(key);
    }
  }

  /* Recalculate a set of cells.  With no argument every formula is
     recalculated; with a list only the transitive dependents are touched. */
  recalculate(changedKeys = null) {
    if (!this.model || !Array.isArray(this.model.sheets)) return 0;
    const targets = new Set();
    if (changedKeys === null || changedKeys === undefined) {
      for (const key of this.parsed.keys()) targets.add(key);
    } else {
      const seeded = new Set();
      const queue = [];
      const seed = (key) => {
        if (!key || seeded.has(key)) return;
        seeded.add(key);
        queue.push(key);
      };
      for (const key of changedKeys) {
        const text = String(key);
        seed(text);
        if (this.parsed.has(text)) targets.add(text);
      }
      for (const key of this.volatile) {
        seed(key);
        targets.add(key);
      }
      while (queue.length) {
        const key = queue.shift();
        for (const dependent of this.dependents.get(key) || []) {
          if (seeded.has(dependent)) continue;
          seeded.add(dependent);
          queue.push(dependent);
          targets.add(dependent);
        }
        const address = parseAddr(keyAddress(key));
        if (!address) continue;
        for (const entry of this.rangeDeps.get(keySheet(key)) || []) {
          if (!rangeContains(entry.range, address.row, address.col)) continue;
          if (seeded.has(entry.key)) continue;
          seeded.add(entry.key);
          queue.push(entry.key);
          targets.add(entry.key);
        }
      }
    }
    for (const key of targets) {
      this.states.delete(key);
      this.values.delete(key);
    }
    for (const key of targets) this.computeKey(key);
    return targets.size;
  }

  recalculateAll() {
    return this.recalculate(null);
  }

  recalculateFrom(sheet, addresses) {
    const keys = [];
    for (const address of addresses || []) {
      const parsed = parseAddr(address);
      if (parsed) keys.push(cellKey(sheet.name, addr(parsed.row, parsed.col)));
    }
    return this.recalculate(keys);
  }

  /* --------------------------------------------------------- computation */
  locate(key) {
    const sheet = findSheet(this.model, keySheet(key));
    if (!sheet) return null;
    const position = parseAddr(keyAddress(key));
    if (!position) return null;
    return { sheet, position, cell: sheet.cells[keyAddress(key)] || null };
  }

  computeKey(key) {
    const state = this.states.get(key);
    if (state === "done") return this.values.get(key);
    if (state === "visiting") return error("#CIRC!");

    const located = this.locate(key);
    if (!located) return null;
    const { sheet, position, cell } = located;
    if (!cell || !cell.f) {
      const plain = cell && cell.v !== undefined ? cell.v : null;
      this.values.set(key, plain);
      this.states.set(key, "done");
      return plain;
    }

    if (this.cycle.has(key)) {
      const circular = error("#CIRC!", "Circular reference");
      this.values.set(key, circular);
      this.states.set(key, "done");
      this.writeCached(sheet, addr(position.row, position.col), circular);
      return circular;
    }

    let parsed = this.parsed.get(key);
    if (!parsed) {
      parsed = this.parseCellFormula(cell.f);
      this.parsed.set(key, parsed);
    }

    this.states.set(key, "visiting");
    let value;
    try {
      if (!parsed.ast) {
        value = error(parsed.code || "#VALUE!", parsed.error || "The formula could not be parsed.");
      } else {
        value = this.evaluateAst(parsed.ast, sheet, key);
      }
      value = this.finalizeValue(value);
    } catch (failure) {
      if (failure instanceof FormulaError) {
        value = failure;
      } else if (failure instanceof RangeError) {
        value = error("#NUM!", "The formula is too deeply nested.");
      } else {
        value = error("#VALUE!", failure?.message ? String(failure.message) : "The formula could not be evaluated.");
      }
    }

    if (typeof value === "number" && !Number.isFinite(value)) value = error("#NUM!");

    this.states.set(key, "done");
    this.values.set(key, value);
    this.writeCached(sheet, addr(position.row, position.col), value);
    return value;
  }

  finalizeValue(value) {
    if (value instanceof FormulaError) return value;
    if (value && value.__ref) {
      const result = this.cellValue(value.sheet, value.row, value.col);
      return result === undefined ? null : result;
    }
    if (value && value.__range) {
      const range = value.range;
      if (range.r1 === range.r2 && range.c1 === range.c2) {
        const result = this.cellValue(value.sheet, range.r1, range.c1);
        return result === undefined ? null : result;
      }
      return error("#VALUE!", "A formula cannot return a multi-cell range without dynamic arrays.");
    }
    return value === undefined ? null : value;
  }

  writeCached(sheet, address, value) {
    const cell = sheet.cells[address];
    if (!cell) return;
    if (value instanceof FormulaError) cell.v = value.code;
    else if (value === null || value === undefined) delete cell.v;
    else cell.v = value;
  }

  cellValue(sheet, row, col) {
    if (!sheet) return error("#REF!");
    if (row < 1 || col < 1 || row > MAX_ROWS || col > MAX_COLS) return error("#REF!");
    const address = addr(row, col);
    const cell = sheet.cells[address];
    if (!cell) return null;
    if (cell.f) return this.computeKey(cellKey(sheet.name, address));
    return cell.v === undefined ? null : cell.v;
  }

  getValue(sheet, address) {
    const position = parseAddr(address);
    if (!position) return null;
    return this.cellValue(sheet, position.row, position.col);
  }

  /* ------------------------------------------------------------ evaluation */
  context(sheet, key) {
    const ctx = {
      engine: this,
      sheet,
      key,
      eval: (node) => this.evalNode(node, sheet, key, ctx),
      tryEval: (node) => {
        try {
          return ctx.eval(node);
        } catch (failure) {
          return failure instanceof FormulaError ? failure : error("#VALUE!");
        }
      },
      tryValue: (node) => {
        try {
          return ctx.engine.finalizeValue(ctx.eval(node));
        } catch (failure) {
          return failure instanceof FormulaError ? failure : error("#VALUE!");
        }
      },
      number: (value) => numberValue(value, ctx),
      text: (value) => textValue(value, ctx),
      bool: (value) => boolValue(value, ctx),
      cellValue: (targetSheet, row, col) => this.cellValue(targetSheet, row, col),
    };
    return ctx;
  }

  evaluateAst(ast, sheet, key) {
    const ctx = this.context(sheet, key);
    return this.evalNode(ast, sheet, key, ctx);
  }

  evalNode(node, sheet, key, ctx) {
    if (!node) return null;
    switch (node.type) {
      case "num":
      case "str":
      case "bool":
        return node.value;
      case "err":
        return error(node.value);
      case "empty":
        return null;
      case "ref":
        return this.makeRef(node, sheet);
      case "range":
        return this.makeRange(node, sheet);
      case "name":
        return this.evalName(node, sheet);
      case "call":
        return this.evalCall(node, sheet, key, ctx);
      case "bin": {
        const left = this.evalNode(node.left, sheet, key, ctx);
        const right = this.evalNode(node.right, sheet, key, ctx);
        return binaryResult(node.op, left, right, ctx);
      }
      case "un": {
        const raw = this.evalNode(node.operand, sheet, key, ctx);
        const value = numberValue(raw, ctx);
        return node.op === "-" ? -value : value;
      }
      case "pct": {
        const raw = this.evalNode(node.operand, sheet, key, ctx);
        return numberValue(raw, ctx) / 100;
      }
      default:
        return error("#VALUE!");
    }
  }

  makeRef(node, sheet) {
    const target = node.sheet ? findSheet(this.model, node.sheet) : sheet;
    if (!target) return error("#REF!");
    if (node.row < 1 || node.col < 1 || node.row > MAX_ROWS || node.col > MAX_COLS) return error("#REF!");
    return { __ref: true, sheet: target, row: node.row, col: node.col };
  }

  makeRange(node, sheet) {
    const target = node.sheet ? findSheet(this.model, node.sheet) : sheet;
    if (!target) return error("#REF!");
    const extent = this.extents.get(target.name.toLowerCase()) || { maxRow: 1, maxCol: 1 };
    let { r1, c1, r2, c2 } = node;
    if (node.whole === "col") {
      r1 = 1;
      r2 = Math.max(1, extent.maxRow);
    } else if (node.whole === "row") {
      c1 = 1;
      c2 = Math.max(1, extent.maxCol);
    }
    if (r1 < 1 || c1 < 1 || r2 > MAX_ROWS || c2 > MAX_COLS) return error("#REF!");
    return { __range: true, sheet: target, range: { r1, c1, r2, c2 }, whole: node.whole || "" };
  }

  evalName(node, sheet) {
    const resolved = resolveNameRange(this.model, node.value, node.sheet || sheet?.name, sheet);
    if (!resolved) return error("#NAME?");
    const range = resolved.range;
    if (range.r1 === range.r2 && range.c1 === range.c2) {
      return { __ref: true, sheet: resolved.sheet, row: range.r1, col: range.c1 };
    }
    return { __range: true, sheet: resolved.sheet, range, whole: "" };
  }

  evalCall(node, sheet, key, ctx) {
    const name = normalizeFunction(node.name);
    const spec = FUNCTIONS[name];
    if (!spec) return error("#NAME?");
    const count = node.args.length;
    const minimum = spec.min ?? 0;
    const maximum = spec.max ?? 255;
    if (count < minimum || count > maximum) return error("#VALUE!");
    try {
      if (spec.lazy) return spec.run(node.args, ctx);
      const args = node.args.map((arg) => this.evalNode(arg, sheet, key, ctx));
      return spec.run(args, ctx);
    } catch (failure) {
      if (failure instanceof FormulaError) return failure;
      if (failure instanceof RangeError) return error("#NUM!", "The formula is too complex.");
      return error("#VALUE!", `${name}: ${failure?.message || "invalid arguments"}`);
    }
  }

  /* ------------------------------------------------------------ inspection */
  precedentsOf(sheet, address) {
    const key = cellKey(sheet.name, address);
    const deps = this.precise.get(key);
    if (!deps) return { cells: [], ranges: [] };
    return {
      cells: [...deps.refs].map((ref) => ({ sheet: keySheet(ref), address: keyAddress(ref) })),
      ranges: deps.ranges.map((entry) => ({ sheet: entry.sheet, range: rangeText(entry.range) })),
    };
  }

  dependentsOf(sheet, address) {
    const position = parseAddr(address);
    if (!position) return [];
    const key = cellKey(sheet.name, address);
    const result = new Set(this.dependents.get(key) || []);
    for (const entry of this.rangeDeps.get(sheet.name.toLowerCase()) || []) {
      if (rangeContains(entry.range, position.row, position.col)) result.add(entry.key);
    }
    return [...result].map((ref) => ({ sheet: keySheet(ref), address: keyAddress(ref) }));
  }

  isVolatile(sheet, address) {
    return this.volatile.has(cellKey(sheet.name, address));
  }

  /* Evaluate a formula without storing it; used by the function dialog. */
  evaluate(text, sheet, address = "A1") {
    try {
      const ast = parseFormula(text);
      const key = cellKey(sheet.name, address);
      const ctx = this.context(sheet, key);
      return this.finalizeValue(this.evalNode(ast, sheet, key, ctx));
    } catch (failure) {
      return failure instanceof FormulaError ? failure : error("#VALUE!");
    }
  }
}

export function formulaDisplay(value) {
  if (value instanceof FormulaError) return value.code;
  if (value === null || value === undefined) return "";
  if (typeof value === "number") return generalFormat(value);
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value);
}
