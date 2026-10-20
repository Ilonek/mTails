"""Capture the marketing screenshots used in the project report.

The script boots a real mTails window, builds a small quarterly report with a
chart, and saves PNG captures for the English UI, the Polish UI, the filled
sheet, the chart and an open dialog. Run it with the project virtualenv:

    .venv/bin/python tests/screenshots.py
"""

from __future__ import annotations

import os
import sys

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE_DIR)

OUT_DIR = os.path.join(BASE_DIR, "screenshots")
os.makedirs(OUT_DIR, exist_ok=True)
os.environ.setdefault("XDG_DATA_HOME", os.path.join(OUT_DIR, ".settings"))

from PySide6.QtCore import QCoreApplication, QTimer  # noqa: E402
from PySide6.QtWidgets import QApplication  # noqa: E402

from mtails.app import MainWindow  # noqa: E402

SETUP = r"""
(async () => {
  const A = window.app;
  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
  while (!A || !A.booted) await sleep(100);
  A.setLanguage("en");
  await sleep(200);
  await A.newWorkbook();
  A.state.title = "Quarterly Report";
  A.model.title = "Quarterly Report";
  A.syncShell();
  const grid = A.grid;
  const split = (address) => {
    const match = /^([A-Z]+)(\d+)$/.exec(address);
    let col = 0;
    for (const char of match[1]) col = col * 26 + (char.charCodeAt(0) - 64);
    return { row: Number(match[2]), col };
  };
  const put = (address, text) => {
    const parsed = split(address);
    grid.selectCell(parsed.row, parsed.col);
    grid.setActiveValue(String(text), { move: "none" });
  };

  const rows = [
    ["Region", "Q1", "Q2", "Q3", "Q4", "Total"],
    ["North", "128000", "142000", "151000", "163000", "=SUM(B2:E2)"],
    ["South", "96000", "108000", "117000", "121000", "=SUM(B3:E3)"],
    ["East", "143000", "138000", "149000", "158000", "=SUM(B4:E4)"],
    ["West", "88000", "97000", "104000", "112000", "=SUM(B5:E5)"],
    ["Total", "=SUM(B2:B5)", "=SUM(C2:C5)", "=SUM(D2:D5)", "=SUM(E2:E5)", "=SUM(F2:F5)"],
  ];
  rows.forEach((row, index) => row.forEach((value, column) => {
    const address = String.fromCharCode(65 + column) + (index + 1);
    put(address, value);
    if (index > 0 && column > 0) grid.setNumberFormat("#,##0");
  }));

  grid.selectRange({ r1: 1, c1: 1, r2: 1, c2: 6 });
  grid.setStyle({ b: true, color: "#FFFFFF", fill: "#1F3864" });
  grid.setAlignment("center");
  grid.applyBorders("all", { style: "thin", color: "#8EA9DB" });
  grid.selectRange({ r1: 6, c1: 1, r2: 6, c2: 6 });
  grid.setStyle({ b: true, fill: "#D6DCE4" });
  grid.applyBorders("top", { style: "double", color: "#1F3864" });
  grid.selectColumn(1);
  grid.setSelectedColumnWidth(120);
  for (let col = 2; col <= 6; col += 1) {
    grid.selectColumn(col);
    grid.setSelectedColumnWidth(92);
  }

  grid.selectRange({ r1: 1, c1: 1, r2: 6, c2: 5 });
  A.execute("insert.chart:column");
  const chart = A.grid.sheet().charts[0];
  if (chart) {
    chart.title = "Quarterly revenue by region";
    chart.x = 660;
    chart.y = 20;
    chart.w = 470;
    chart.h = 290;
    grid.renderOverlays();
  }

  grid.selectCell(2, 1);
  grid.freezePanes();
  grid.selectCell(1, 1);
  A.setDirty(false);
  window.__shotReady = true;
})();
"""

CAPTURE_DIALOG = r"""
(async () => {
  const A = window.app;
  A.applyTheme("light");
  A.grid.selectedOverlay = null;
  A.grid.renderOverlays();
  A.execute("formatCells");
})();
"""

CLOSE_DIALOGS = r"""
(function () {
  window.app.closeAllDialogs ? window.app.closeAllDialogs() : null;
  document.querySelectorAll("#dialog-layer .dialog-close").forEach((node) => node.click());
})();
"""


def main() -> int:
    app = QApplication(sys.argv[:1])
    QCoreApplication.setApplicationName("mTailsShot")
    QCoreApplication.setOrganizationName("mTailsShot")

    window = MainWindow(startup_file="", smoke_path="", language="en")
    window.resize(1440, 900)
    window.show()

    shots: list[tuple[str, str]] = [
        ("01-home-en.png", ""),
        ("02-home-pl.png",
         "(async () => { const A = window.app; A.setLanguage('pl'); "
         "const chart = A.model.sheets[0].charts[0]; "
         "if (chart) { chart.title = 'Przychody kwartalne według regionu'; A.grid.renderOverlays(); } })();"),
        ("03-sheet-filled.png",
         "(async () => { const A = window.app; A.setLanguage('en'); "
         "const chart = A.model.sheets[0].charts[0]; "
         "if (chart) { chart.title = 'Quarterly revenue by region'; A.grid.renderOverlays(); } })();"),
        ("04-chart.png",
         "(function () { const A = window.app; const chart = A.model.sheets[0].charts[0]; "
         "if (chart) { A.grid.selectedOverlay = chart.id; A.grid.renderOverlays(); } })();"),
        ("05-dark.png",
         "(function () { const A = window.app; A.applyTheme('dark'); "
         "A.grid.selectedOverlay = null; A.grid.renderOverlays(); })();"),
        ("06-dialog-format-cells.png", CAPTURE_DIALOG),
    ]
    state = {"index": 0, "started": False, "keep": []}

    def grab(name: str) -> None:
        path = os.path.join(OUT_DIR, name)
        window.grab().save(path, "PNG")
        print("SHOT:" + path, flush=True)

    def step() -> None:
        index = state["index"]
        if index >= len(shots):
            print("SHOTS:" + str(index), flush=True)
            QTimer.singleShot(200, app.quit)
            QTimer.singleShot(1500, hard_exit)
            return
        name, script = shots[index]
        if script:
            window.view.page().runJavaScript(script)
        window.view.page().runJavaScript(
            "document.querySelectorAll('#toast-layer .toast').forEach((node) => node.remove())"
        )
        state["index"] = index + 1

        def capture() -> None:
            if name == "04-chart.png":
                window.view.page().runJavaScript("window.app.grid.renderOverlays && window.app.grid.renderOverlays()")
            grab(name)
            if name == "06-dialog-format-cells.png":
                window.view.page().runJavaScript(CLOSE_DIALOGS)
            QTimer.singleShot(700, step)

        QTimer.singleShot(650, capture)
        # PySide drops callbacks that are only referenced from C++, so keep a
        # strong reference until the timer fires.
        state["keep"].append(capture)

    def poll() -> None:
        def receive(value):
            if value and not state["started"]:
                state["started"] = True
                QTimer.singleShot(700, step)

        window.view.page().runJavaScript("window.__shotReady ? 'ready' : ''", 0, receive)

    def run_setup() -> None:
        window.view.page().runJavaScript(SETUP)

    def hard_exit() -> None:
        os._exit(0)

    ticker = QTimer()
    ticker.timeout.connect(poll)
    ticker.start(500)
    state["keep"].append(run_setup)
    state["keep"].append(step)
    QTimer.singleShot(2200, run_setup)
    QTimer.singleShot(60000, app.quit)
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
