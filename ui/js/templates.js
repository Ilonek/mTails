/* Ready-made workbooks offered by the File > New backstage panel.
 *
 * Every template is a real workbook model: cells may contain formulas so the
 * engine recalculates them on load, and the sheets carry the same page setup,
 * widths and freeze panes a user would configure by hand.  build() runs on
 * demand and returns a fresh model, so opening a template twice never shares
 * cell objects. */

import { createSheet, createWorkbook } from "./model.js";
import { serialFromParts } from "./format.js";

const INK = "#1F3864";
const SOFT = "#44546A";
const HEAD_FILL = "#1F3864";
const BAND_FILL = "#D9E2F3";
const MONEY = "$#,##0.00";
const MONEY_RED = "$#,##0.00;[Red]-$#,##0.00";
const PLAIN_MONEY = "#,##0";
const DATE = "m/d/yyyy";
const PERCENT = "0.0%";

function style(fill, color, extra = {}) {
  const spec = { ...extra };
  if (fill) spec.fill = fill;
  if (color) spec.color = color;
  return spec;
}

const TITLE = style("", INK, { b: true, size: 20 });
const SUBTITLE = style("", SOFT, { i: true, size: 11 });
const HEAD = style(HEAD_FILL, "#FFFFFF", { b: true, size: 11 });
const SECTION = style(BAND_FILL, INK, { b: true });
const LABEL = style("", SOFT, { b: true });
const TOTAL = style(BAND_FILL, INK, {
  b: true,
  border: { t: { style: "thin", color: INK }, b: { style: "double", color: INK } },
});
const MONEY_CELL = style("", "", { numFmt: MONEY });
const MONEY_TOTAL = style(BAND_FILL, INK, {
  b: true,
  numFmt: MONEY,
  border: { t: { style: "thin", color: INK }, b: { style: "double", color: INK } },
});
const DATE_CELL = style("", "", { numFmt: DATE, align: "left" });

function put(cells, address, value, spec = null, formula = "") {
  const cell = {};
  if (formula) cell.f = formula;
  else if (value !== "" && value !== null && value !== undefined) cell.v = value;
  if (spec && Object.keys(spec).length) cell.s = { ...spec };
  cells[address] = cell;
}

function titleBlock(cells, title, subtitle) {
  put(cells, "A1", title, TITLE);
  if (subtitle) put(cells, "A2", subtitle, SUBTITLE);
}

function headerRow(cells, row, labels, startCol = 1) {
  labels.forEach((label, index) => {
    put(cells, columnLetter(startCol + index) + row, label, HEAD);
  });
}

function columnLetter(index) {
  let value = "";
  let number = index;
  while (number > 0) {
    const remainder = (number - 1) % 26;
    value = String.fromCharCode(65 + remainder) + value;
    number = Math.floor((number - 1) / 26);
  }
  return value;
}

function widths(sheet, entries) {
  for (const [col, width] of Object.entries(entries)) {
    sheet.cols[col] = { width };
  }
}

function setup(sheet, page = {}) {
  sheet.freeze = { rows: page.freezeRows || 0, cols: page.freezeCols || 0 };
  if (page.landscape) sheet.page.orientation = "landscape";
  sheet.page.fitToWidth = true;
  if (page.area) sheet.page.area = page.area;
}

function makeBook(title, sheets) {
  const book = createWorkbook({ title, sheets: 1 });
  book.sheets = sheets;
  book.activeSheet = 0;
  book.title = title;
  book.properties.title = title;
  book.properties.author = "mTails";
  book.properties.subject = "mTails template";
  return book;
}

/* ------------------------------------------------------------ blank workbook */
function blankWorkbook() {
  return makeBook("Workbook1", [createSheet("Sheet1")]);
}

/* --------------------------------------------------------------- monthly budget */
function monthlyBudget() {
  const sheet = createSheet("Budget");
  const cells = sheet.cells;
  titleBlock(cells, "Monthly budget", "Plan income and expenses for the month.");
  headerRow(cells, 4, ["Category", "Planned", "Actual", "Difference"]);
  const rows = [
    ["Housing", 1400, 1400],
    ["Groceries", 480, 512],
    ["Transport", 220, 196],
    ["Utilities", 180, 174],
    ["Dining out", 160, 208],
    ["Savings", 400, 400],
    ["Other", 120, 86],
  ];
  rows.forEach(([label, planned, actual], index) => {
    const row = 5 + index;
    put(cells, "A" + row, label, style("", "", { b: index === 0 }));
    put(cells, "B" + row, planned, MONEY_CELL);
    put(cells, "C" + row, actual, MONEY_CELL);
    put(cells, "D" + row, "", style("", "", { numFmt: MONEY_RED }), "=C" + row + "-B" + row);
  });
  put(cells, "A12", "Total", TOTAL);
  put(cells, "B12", "", MONEY_TOTAL, "=SUM(B5:B11)");
  put(cells, "C12", "", MONEY_TOTAL, "=SUM(C5:C11)");
  put(cells, "D12", "", MONEY_TOTAL, "=C12-B12");
  put(cells, "A14", "Planned vs. actual", LABEL);
  put(cells, "B14", "", style("", INK, { b: true, numFmt: MONEY_RED }), "=B12-C12");
  widths(sheet, { A: 22, B: 14, C: 14, D: 14 });
  setup(sheet, { freezeRows: 4 });
  sheet.charts.push({
    id: "chart-budget",
    type: "pie",
    range: "A4:B11",
    title: "Planned spending",
    headerRow: true,
    headerCol: true,
    x: 384,
    y: 16,
    w: 380,
    h: 264,
  });
  return makeBook("Monthly budget", [sheet]);
}

/* ---------------------------------------------------------------- sales tracker */
function salesTracker() {
  const sheet = createSheet("Sales");
  const cells = sheet.cells;
  titleBlock(cells, "Sales tracker", "Quarterly revenue by region, in thousands.");
  headerRow(cells, 4, ["Region", "Q1", "Q2", "Q3", "Q4", "Total"]);
  const rows = [
    ["North", 128, 142, 156, 171],
    ["South", 96, 108, 101, 124],
    ["East", 143, 138, 162, 179],
    ["West", 88, 94, 112, 118],
    ["Online", 64, 79, 95, 132],
  ];
  rows.forEach(([region, ...values], index) => {
    const row = 5 + index;
    put(cells, "A" + row, region, style(BAND_FILL, INK, { b: true }));
    values.forEach((value, column) => {
      put(cells, columnLetter(2 + column) + row, value, style("", "", { numFmt: PLAIN_MONEY }));
    });
    put(cells, "F" + row, "", style("", "", { numFmt: PLAIN_MONEY }), "=SUM(B" + row + ":E" + row + ")");
  });
  put(cells, "A10", "Total", TOTAL);
  for (const col of ["B", "C", "D", "E", "F"]) {
    put(cells, col + "10", "", MONEY_TOTAL, "=SUM(" + col + "5:" + col + "9)");
  }
  widths(sheet, { A: 16, B: 11, C: 11, D: 11, E: 11, F: 13 });
  setup(sheet, { freezeRows: 4 });
  sheet.charts.push({
    id: "chart-sales",
    type: "column",
    range: "A4:E9",
    title: "Revenue by region",
    headerRow: true,
    headerCol: true,
    x: 448,
    y: 16,
    w: 440,
    h: 280,
  });
  return makeBook("Sales tracker", [sheet]);
}

/* -------------------------------------------------------------------- invoice */
function invoice() {
  const sheet = createSheet("Invoice");
  const cells = sheet.cells;
  put(cells, "A1", "INVOICE", style("", INK, { b: true, size: 26 }));
  put(cells, "A2", "mTails Studio", style("", INK, { b: true, size: 12 }));
  put(cells, "A3", "12 Example Street, Warsaw", style("", SOFT, { size: 10 }));
  put(cells, "A5", "Bill to", LABEL);
  put(cells, "A6", "Acme Trading Ltd.", style("", "", { b: true }));
  put(cells, "A7", "42 Market Square, Kraków", style("", SOFT, { size: 10 }));
  put(cells, "E1", "Invoice #", LABEL);
  put(cells, "F1", "INV-1001");
  put(cells, "E2", "Issued", LABEL);
  put(cells, "F2", "", DATE_CELL, "=TODAY()");
  put(cells, "E3", "Due", LABEL);
  put(cells, "F3", "", DATE_CELL, "=TODAY()+30");
  headerRow(cells, 9, ["Description", "Qty", "Rate", "Amount"]);
  const items = [
    ["Spreadsheet design", 2, 1200],
    ["Template development", 5, 320],
    ["Training session (hours)", 3, 180],
  ];
  items.forEach(([label, qty, rate], index) => {
    const row = 10 + index;
    put(cells, "A" + row, label);
    put(cells, "B" + row, qty, style("", "", { numFmt: "0" }));
    put(cells, "C" + row, rate, style("", "", { numFmt: MONEY }));
    put(cells, "D" + row, "", style("", "", { numFmt: MONEY }), "=B" + row + "*C" + row);
  });
  put(cells, "A15", "Subtotal", LABEL);
  put(cells, "D15", "", style("", "", { numFmt: MONEY }), "=SUM(D10:D14)");
  put(cells, "A16", "Tax", LABEL);
  put(cells, "C16", 0.23, style("", "", { numFmt: PERCENT }));
  put(cells, "D16", "", style("", "", { numFmt: MONEY }), "=D15*C16");
  put(cells, "A17", "Total due", TOTAL);
  put(cells, "D17", "", MONEY_TOTAL, "=D15+D16");
  put(cells, "A19", "Payment is due within 30 days. Thank you for your business.", SUBTITLE);
  sheet.merges.push("A1:D1", "A5:D5", "A19:D19");
  widths(sheet, { A: 34, B: 10, C: 14, D: 16, E: 12, F: 18 });
  setup(sheet);
  return makeBook("Invoice", [sheet]);
}

/* ------------------------------------------------------------------ to-do list */
function todoList() {
  const sheet = createSheet("To-do");
  const cells = sheet.cells;
  titleBlock(cells, "To-do list", "Track assignments and their due dates.");
  headerRow(cells, 4, ["Done", "Task", "Owner", "Due", "Days left"]);
  const tasks = [
    ["Yes", "Draft quarterly report", "Anna", 3],
    ["No", "Review marketing budget", "Piotr", 6],
    ["No", "Update price list", "Marta", 9],
    ["No", "Prepare client presentation", "Jan", 12],
    ["Yes", "Archive invoices", "Anna", 1],
    ["No", "Plan team offsite", "Marta", 18],
    ["No", "Refresh website copy", "Piotr", 21],
    ["No", "Order office supplies", "Jan", 5],
  ];
  const today = new Date();
  const todaySerial = serialFromParts(today.getFullYear(), today.getMonth() + 1, today.getDate());
  tasks.forEach(([done, task, owner, offset], index) => {
    const row = 5 + index;
    put(cells, "A" + row, done, style("", done === "Yes" ? "#006100" : SOFT, { align: "center", b: done === "Yes" }));
    put(cells, "B" + row, task);
    put(cells, "C" + row, owner, style("", SOFT, {}));
    put(cells, "D" + row, todaySerial + offset, DATE_CELL);
    put(cells, "E" + row, "", style("", "", { numFmt: "0", align: "center" }), "=D" + row + "-TODAY()");
  });
  sheet.conditional.push({
    range: "A5:A12",
    type: "cellIs",
    operator: "equal",
    formula1: "Yes",
    fill: "#C6EFCE",
    color: "#006100",
    bold: true,
  });
  widths(sheet, { A: 9, B: 32, C: 12, D: 13, E: 11 });
  setup(sheet, { freezeRows: 4 });
  return makeBook("To-do list", [sheet]);
}

/* ------------------------------------------------------------------- timesheet */
function timesheet() {
  const sheet = createSheet("Timesheet");
  const cells = sheet.cells;
  titleBlock(cells, "Timesheet", "Weekly hours and billing summary.");
  put(cells, "A3", "Employee", LABEL);
  put(cells, "B3", "Your name");
  put(cells, "D3", "Rate", LABEL);
  put(cells, "E3", 95, style("", "", { numFmt: MONEY }));
  headerRow(cells, 5, ["Date", "Project", "Hours", "Rate", "Amount", "Notes"]);
  const base = serialFromParts(2026, 10, 19);
  const entries = [
    ["Discovery workshop", 6, "Kick-off"],
    ["Spreadsheet audit", 7.5, ""],
    ["Template build", 8, ""],
    ["Template build", 6.5, "Review with Anna"],
    ["Client training", 4, ""],
  ];
  entries.forEach(([project, hours, note], index) => {
    const row = 6 + index;
    put(cells, "A" + row, base + index, DATE_CELL);
    put(cells, "B" + row, project);
    put(cells, "C" + row, hours, style("", "", { numFmt: "0.0" }));
    put(cells, "D" + row, "", style("", "", { numFmt: MONEY }), "=IF($E$3=\"\",0,$E$3)");
    put(cells, "E" + row, "", style("", "", { numFmt: MONEY }), "=C" + row + "*D" + row);
    put(cells, "F" + row, note, style("", SOFT, { size: 10 }));
  });
  put(cells, "A12", "Total", TOTAL);
  put(cells, "C12", "", MONEY_TOTAL, "=SUM(C6:C11)");
  put(cells, "E12", "", MONEY_TOTAL, "=SUM(E6:E11)");
  widths(sheet, { A: 13, B: 24, C: 10, D: 12, E: 14, F: 24 });
  setup(sheet, { freezeRows: 5 });
  return makeBook("Timesheet", [sheet]);
}

/* -------------------------------------------------------- loan amortization */
function loanAmortization() {
  const sheet = createSheet("Loan");
  const cells = sheet.cells;
  titleBlock(cells, "Loan amortization", "Monthly payment and a twelve-month schedule.");
  put(cells, "A4", "Loan amount", LABEL);
  put(cells, "B4", 250000, style("", "", { numFmt: MONEY }));
  put(cells, "A5", "Annual rate", LABEL);
  put(cells, "B5", 0.065, style("", "", { numFmt: PERCENT }));
  put(cells, "A6", "Years", LABEL);
  put(cells, "B6", 30, style("", "", { numFmt: "0" }));
  put(cells, "A7", "Monthly payment", LABEL);
  put(cells, "B7", "", style(BAND_FILL, INK, { b: true, size: 12, numFmt: MONEY }),
    "=(B4*B5/12)/(1-POWER(1+B5/12,-B6*12))");
  headerRow(cells, 9, ["Period", "Payment", "Interest", "Principal", "Balance"]);
  for (let period = 1; period <= 12; period += 1) {
    const row = 9 + period;
    put(cells, "A" + row, period, style("", "", { numFmt: "0", align: "center" }));
    put(cells, "B" + row, "", style("", "", { numFmt: MONEY }), "=$B$7");
    if (period === 1) {
      put(cells, "C" + row, "", style("", "", { numFmt: MONEY }), "=$B$4*$B$5/12");
      put(cells, "D" + row, "", style("", "", { numFmt: MONEY }), "=B10-C10");
      put(cells, "E" + row, "", style("", "", { numFmt: MONEY }), "=$B$4-D10");
    } else {
      put(cells, "C" + row, "", style("", "", { numFmt: MONEY }), "=E" + (row - 1) + "*$B$5/12");
      put(cells, "D" + row, "", style("", "", { numFmt: MONEY }), "=B" + row + "-C" + row);
      put(cells, "E" + row, "", style("", "", { numFmt: MONEY }), "=E" + (row - 1) + "-D" + row);
    }
  }
  put(cells, "A22", "After 12 payments", LABEL);
  put(cells, "E22", "", style(BAND_FILL, INK, { b: true, numFmt: MONEY }), "=E21");
  widths(sheet, { A: 18, B: 14, C: 14, D: 14, E: 16 });
  setup(sheet, { freezeRows: 9 });
  return makeBook("Loan amortization", [sheet]);
}

/* ---------------------------------------------------------------- grade tracker */
function gradeTracker() {
  const sheet = createSheet("Grades");
  const cells = sheet.cells;
  titleBlock(cells, "Grade tracker", "Assignment scores and final letter grades.");
  headerRow(cells, 4, ["Student", "Quiz 1", "Quiz 2", "Midterm", "Final", "Average", "Grade"]);
  const students = [
    ["Anna Kowalska", 92, 88, 90, 94],
    ["Piotr Nowak", 78, 84, 81, 86],
    ["Marta Zielińska", 88, 91, 86, 89],
    ["Jan Wiśniewski", 65, 72, 68, 74],
    ["Ewa Lewandowska", 96, 94, 98, 97],
    ["Michał Szymański", 74, 69, 77, 72],
    ["Katarzyna Wójcik", 83, 87, 85, 88],
  ];
  students.forEach(([name, ...scores], index) => {
    const row = 5 + index;
    put(cells, "A" + row, name, style("", "", { b: true }));
    scores.forEach((score, column) => {
      put(cells, columnLetter(2 + column) + row, score, style("", "", { numFmt: "0", align: "center" }));
    });
    put(cells, "F" + row, "", style("", INK, { b: true, numFmt: "0.0", align: "center" }),
      "=AVERAGE(B" + row + ":E" + row + ")");
    put(cells, "G" + row, "",
      style("", "", { align: "center", b: true }),
      "=IF(F" + row + ">=90,\"A\",IF(F" + row + ">=80,\"B\",IF(F" + row + ">=70,\"C\",IF(F" + row + ">=60,\"D\",\"F\"))))");
  });
  put(cells, "A13", "Class average", TOTAL);
  for (const col of ["B", "C", "D", "E", "F"]) {
    put(cells, col + "13", "", style(BAND_FILL, INK, {
      b: true,
      numFmt: col === "F" ? "0.0" : "0.0",
      border: { t: { style: "thin", color: INK }, b: { style: "double", color: INK } },
    }), "=AVERAGE(" + col + "5:" + col + "12)");
  }
  widths(sheet, { A: 24, B: 9, C: 9, D: 10, E: 9, F: 10, G: 8 });
  setup(sheet, { freezeRows: 4 });
  sheet.charts.push({
    id: "chart-grades",
    type: "column",
    range: "A4:F12",
    title: "Average by student",
    headerRow: true,
    headerCol: true,
    x: 520,
    y: 16,
    w: 420,
    h: 280,
  });
  return makeBook("Grade tracker", [sheet]);
}

/* ------------------------------------------------------------------ catalogue */
export const TEMPLATES = [
  {
    id: "blank",
    label: "Blank workbook",
    description: "An empty grid, ready for anything.",
    accent: "accent",
    thumb: { title: "Blank workbook", head: ["A", "B", "C"], rows: [["", "", ""], ["", "", ""], ["", "", ""], ["", "", ""]] },
    build: blankWorkbook,
  },
  {
    id: "budget",
    label: "Monthly budget",
    description: "Planned versus actual spending.",
    accent: "green",
    thumb: { title: "Monthly budget", head: ["Category", "Plan", "Actual"], rows: [["Housing", "1,400", "1,400"], ["Food", "480", "512"], ["Transport", "220", "196"], ["Total", "2,960", "2,976"]] },
    build: monthlyBudget,
  },
  {
    id: "sales",
    label: "Sales tracker",
    description: "Quarterly revenue by region.",
    accent: "accent",
    thumb: { title: "Sales tracker", head: ["Region", "Q1", "Q2"], rows: [["North", "128", "142"], ["South", "96", "108"], ["East", "143", "138"], ["Total", "519", "561"]] },
    build: salesTracker,
  },
  {
    id: "invoice",
    label: "Invoice",
    description: "Billable items with tax and total.",
    accent: "warm",
    thumb: { title: "INVOICE", head: ["Item", "Qty", "Amount"], rows: [["Design", "2", "2,400"], ["Template", "5", "1,600"], ["Tax", "23%", "920"], ["Total", "—", "4,920"]] },
    build: invoice,
  },
  {
    id: "todo",
    label: "To-do list",
    description: "Tasks, owners and due dates.",
    accent: "purple",
    thumb: { title: "To-do list", head: ["Done", "Task", "Due"], rows: [["Yes", "Report", "Oct 22"], ["No", "Budget", "Oct 25"], ["No", "Prices", "Oct 28"], ["No", "Deck", "Oct 31"]] },
    build: todoList,
  },
  {
    id: "timesheet",
    label: "Timesheet",
    description: "Weekly hours and billing.",
    accent: "green",
    thumb: { title: "Timesheet", head: ["Date", "Hours", "Amount"], rows: [["Mon", "6.0", "570"], ["Tue", "7.5", "713"], ["Wed", "8.0", "760"], ["Total", "21.5", "2,043"]] },
    build: timesheet,
  },
  {
    id: "loan",
    label: "Loan amortization",
    description: "Payment formula and schedule.",
    accent: "accent",
    thumb: { title: "Loan", head: ["Period", "Interest", "Balance"], rows: [["1", "1,354", "249,273"], ["2", "1,350", "248,541"], ["3", "1,346", "247,806"], ["…", "…", "…"]] },
    build: loanAmortization,
  },
  {
    id: "grades",
    label: "Grade tracker",
    description: "Scores, averages and letters.",
    accent: "warm",
    thumb: { title: "Grades", head: ["Student", "Avg", "Grade"], rows: [["Anna", "91.0", "A"], ["Piotr", "82.3", "B"], ["Jan", "69.8", "D"], ["Ewa", "96.3", "A"]] },
    build: gradeTracker,
  },
];

export function templateById(id) {
  return TEMPLATES.find((template) => template.id === id) || null;
}
