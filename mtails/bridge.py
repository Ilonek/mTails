"""Qt <-> JavaScript bridge and spreadsheet file handling."""

from __future__ import annotations

import json
import os
import sys
import time
from typing import Optional

from PySide6.QtCore import QMarginsF, QObject, QSizeF, QUrl, Slot
from PySide6.QtGui import QDesktopServices, QFontDatabase, QPageLayout, QPageSize
from PySide6.QtPrintSupport import QPrintDialog, QPrinter
from PySide6.QtWidgets import QFileDialog, QMessageBox

from . import APP_VERSION, i18n
from .config import Settings
from .formats import default_page, sheet_to_delimited, text_to_sheet
from .xlsx_reader import read_xlsx
from .xlsx_writer import write_xlsx


def _filter(name: str, patterns: str) -> str:
    return "%s (%s)" % (i18n.tr(name), patterns)


def filters() -> dict:
    """File dialog filters in the active UI language, keyed by format."""
    return {
        "xlsx": _filter("Excel Workbook", "*.xlsx"),
        "xlsm": _filter("Excel Macro-Enabled Workbook", "*.xlsm"),
        "csv": _filter("CSV (Comma delimited)", "*.csv"),
        "tsv": _filter("Text (Tab delimited)", "*.tsv"),
    }


def open_filters() -> str:
    combined = _filter("All supported spreadsheets", "*.xlsx *.xlsm *.csv *.tsv")
    return ";;".join([combined] + list(filters().values()))


def save_filters() -> str:
    return ";;".join(filters().values())


EXT_TO_FORMAT = {
    ".xlsx": "xlsx",
    ".xlsm": "xlsx",
    ".csv": "csv",
    ".tsv": "tsv",
}

FORMAT_TO_EXT = {
    "xlsx": ".xlsx",
    "xlsm": ".xlsx",
    "csv": ".csv",
    "tsv": ".tsv",
}


def _decode_text(data: bytes) -> str:
    for encoding in ("utf-8-sig", "utf-8", "cp1250", "latin-1"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")


class Bridge(QObject):
    def __init__(self, window, settings: Settings) -> None:
        super().__init__(window)
        self.window = window
        self.settings = settings
        self.document_path: Optional[str] = None
        self.document_name: str = "Workbook1"
        self.modified = False
        self._pdf_callback = None
        self._print_callback = None
        self._pdf_result = ""
        self._print_result = ""

    # ----------------------------------------------------------- bootstrap
    @Slot(result=str)
    def bootstrap(self) -> str:
        fonts = QFontDatabase.families()
        fonts = sorted({family for family in fonts if family and not family.startswith("@")}, key=str.lower)
        recents = []
        for path in self.settings.get("recent", []):
            recents.append({"path": path, "name": os.path.basename(path), "exists": os.path.exists(path)})
        autosaves = self._list_autosaves()
        return json.dumps(
            {
                "version": APP_VERSION,
                "platform": sys.platform,
                "fonts": fonts,
                "recent": recents,
                "settings": {
                    "theme": self.settings.get("theme", "light"),
                    "autosave": self.settings.get("autosave", True),
                    "autosaveInterval": self.settings.get("autosaveInterval", 120),
                    "autorecover": self.settings.get("autorecover", True),
                    "zoom": self.settings.get("zoom", 100),
                    "uiLanguage": self.settings.get("uiLanguage", "system"),
                    "defaultFont": self.settings.get("defaultFont", "Calibri"),
                    "defaultSize": self.settings.get("defaultSize", 11),
                    "showGridlines": self.settings.get("showGridlines", True),
                    "showHeadings": self.settings.get("showHeadings", True),
                    "showFormulaBar": self.settings.get("showFormulaBar", True),
                    "defaultSheets": self.settings.get("defaultSheets", 1),
                },
                "defaultFont": self.settings.get("defaultFont", "Calibri"),
                "defaultSize": self.settings.get("defaultSize", 11),
                "startupFile": self.window.startup_file or "",
                "autosaves": autosaves,
                "home": os.path.expanduser("~"),
                "separator": os.sep,
            }
        )

    # -------------------------------------------------------------- opening
    @Slot(str, result=str)
    def openDocument(self, path: str) -> str:
        if not path:
            start_dir = self.settings.get("lastFolder") or os.path.expanduser("~")
            path, _ = QFileDialog.getOpenFileName(self.window, i18n.tr("Open"), start_dir, open_filters())
            if not path:
                return json.dumps({"ok": False, "canceled": True})
        return self._open_path(path)

    @Slot(str, result=str)
    def openPath(self, path: str) -> str:
        return self._open_path(path)

    def _open_path(self, path: str) -> str:
        path = os.path.abspath(path)
        if not os.path.exists(path):
            return json.dumps({"ok": False, "error": i18n.tr("File not found: {path}", path=path)})
        ext = os.path.splitext(path)[1].lower()
        try:
            if ext in (".xlsx", ".xlsm", ".xltx", ".xltm"):
                model = read_xlsx(path)
                fmt = "xlsx"
            elif ext in (".csv", ".tsv", ".txt"):
                with open(path, "rb") as handle:
                    text = _decode_text(handle.read())
                delimiter = "\t" if ext in (".tsv", ".txt") else None
                sheet = text_to_sheet(text, os.path.splitext(os.path.basename(path))[0], delimiter)
                sheet["name"] = "Sheet1"
                model = {
                    "version": 1,
                    "title": os.path.splitext(os.path.basename(path))[0],
                    "activeSheet": 0,
                    "settings": {"calcMode": "auto", "r1c1": False},
                    "names": {},
                    "sheets": [sheet],
                }
                fmt = "csv"
            elif ext == ".xls":
                return json.dumps(
                    {
                        "ok": False,
                        "error": i18n.tr(
                            "Legacy .xls files are not supported. Please open the file in Excel/LibreOffice and save it as .xlsx first."
                        ),
                    }
                )
            else:
                return json.dumps(
                    {"ok": False, "error": i18n.tr("Unsupported file type: {ext}", ext=ext or i18n.tr("unknown"))}
                )
        except Exception as exc:  # noqa: BLE001 - surfaced to the UI
            return json.dumps({"ok": False, "error": i18n.tr("Could not open the workbook:\n{error}", error=exc)})

        self.document_path = path
        self.document_name = os.path.splitext(os.path.basename(path))[0]
        self.settings.add_recent(path)
        self.settings.set("lastFolder", os.path.dirname(path))
        model.setdefault("title", self.document_name)
        return json.dumps(
            {
                "ok": True,
                "kind": "model",
                "model": model,
                "path": path,
                "name": os.path.basename(path),
                "format": fmt,
            }
        )

    # -------------------------------------------------------------- saving
    @Slot(str, str, str, str, bool, result=str)
    def saveDocument(self, model_json: str, fmt: str, suggested: str, current_path: str, save_as: bool) -> str:
        try:
            model = json.loads(model_json)
        except json.JSONDecodeError as exc:
            return json.dumps({"ok": False, "error": i18n.tr("Invalid workbook data: {error}", error=exc)})

        path = current_path
        if save_as or not path:
            start_dir = self.settings.get("lastFolder") or os.path.expanduser("~")
            suggested = suggested or (model.get("title") or "Workbook")
            if not os.path.splitext(suggested)[1]:
                suggested += FORMAT_TO_EXT.get(fmt, ".xlsx")
            initial = os.path.join(start_dir, suggested)
            dialog = QFileDialog(self.window, i18n.tr("Save As"), initial, save_filters())
            dialog.setAcceptMode(QFileDialog.AcceptSave)
            dialog.setOption(QFileDialog.DontUseNativeDialog, False)
            current = filters()
            wanted = current.get(fmt, current["xlsx"])
            dialog.selectNameFilter(wanted)
            if dialog.exec() != QFileDialog.Accepted:
                return json.dumps({"ok": False, "canceled": True})
            selected = dialog.selectedFiles()
            if not selected:
                return json.dumps({"ok": False, "canceled": True})
            path = selected[0]
            chosen_filter = dialog.selectedNameFilter()
            if not os.path.splitext(path)[1]:
                filter_ext = {current[key]: value for key, value in FORMAT_TO_EXT.items()}.get(
                    chosen_filter, FORMAT_TO_EXT.get(fmt, ".xlsx")
                )
                path += filter_ext

        ext = os.path.splitext(path)[1].lower()
        actual_fmt = EXT_TO_FORMAT.get(ext, fmt)
        try:
            self._write(path, actual_fmt, model)
        except Exception as exc:  # noqa: BLE001
            message = i18n.tr("The workbook could not be saved:\n{error}", error=exc)
            QMessageBox.critical(self.window, i18n.tr("Save failed"), message)
            return json.dumps({"ok": False, "error": str(exc)})

        self.document_path = path
        self.document_name = os.path.splitext(os.path.basename(path))[0]
        self.settings.add_recent(path)
        self.settings.set("lastFolder", os.path.dirname(path))
        self.modified = False
        self.window.set_document_title(self.document_name, False)
        return json.dumps(
            {
                "ok": True,
                "path": path,
                "name": os.path.basename(path),
                "title": self.document_name,
                "format": actual_fmt,
            }
        )

    @Slot(str, str, str, result=str)
    def exportTo(self, model_json: str, fmt: str, path: str) -> str:
        """Write an export to an explicit path chosen by the caller."""
        try:
            model = json.loads(model_json)
        except json.JSONDecodeError as exc:
            return json.dumps({"ok": False, "error": i18n.tr("Invalid workbook data: {error}", error=exc)})
        path = os.path.abspath(path)
        try:
            self._write(path, fmt, model)
        except Exception as exc:  # noqa: BLE001
            return json.dumps({"ok": False, "error": str(exc)})
        self.settings.add_recent(path)
        return json.dumps({"ok": True, "path": path, "name": os.path.basename(path)})

    @Slot(str, str, str, result=str)
    def exportDocument(self, model_json: str, fmt: str, suggested: str) -> str:
        try:
            model = json.loads(model_json)
        except json.JSONDecodeError as exc:
            return json.dumps({"ok": False, "error": i18n.tr("Invalid workbook data: {error}", error=exc)})
        current = filters()
        selected_filter = current.get(fmt, current["xlsx"])
        start_dir = self.settings.get("lastFolder") or os.path.expanduser("~")
        base = suggested or (model.get("title") or "Workbook")
        base = os.path.splitext(base)[0] + FORMAT_TO_EXT.get(fmt, ".xlsx")
        path, _ = QFileDialog.getSaveFileName(
            self.window, i18n.tr("Export"), os.path.join(start_dir, base), selected_filter
        )
        if not path:
            return json.dumps({"ok": False, "canceled": True})
        if not os.path.splitext(path)[1]:
            path += FORMAT_TO_EXT.get(fmt, ".xlsx")
        try:
            self._write(path, fmt, model)
        except Exception as exc:  # noqa: BLE001
            return json.dumps({"ok": False, "error": str(exc)})
        self.settings.add_recent(path)
        return json.dumps({"ok": True, "path": path, "name": os.path.basename(path)})

    def _write(self, path: str, fmt: str, model: dict) -> None:
        model = dict(model)
        model["title"] = os.path.splitext(os.path.basename(path))[0]
        if fmt == "xlsx":
            data = write_xlsx(model)
            with open(path, "wb") as handle:
                handle.write(data)
        elif fmt in ("csv", "tsv"):
            delimiter = "\t" if fmt == "tsv" else ","
            sheets = model.get("sheets") or []
            index = int(model.get("activeSheet") or 0)
            text = sheet_to_delimited(model, index if sheets else 0, delimiter)
            with open(path, "w", encoding="utf-8", newline="") as handle:
                handle.write(text)
        else:
            raise ValueError(i18n.tr("Unsupported format: {format}", format=fmt))

    # ------------------------------------------------------------------ pdf
    @Slot(str, result=str)
    def choosePdfPath(self, suggested: str) -> str:
        start_dir = self.settings.get("lastFolder") or os.path.expanduser("~")
        base = os.path.splitext(suggested or "Workbook")[0] + ".pdf"
        path, _ = QFileDialog.getSaveFileName(
            self.window,
            i18n.tr("Export as PDF"),
            os.path.join(start_dir, base),
            _filter("PDF Document", "*.pdf"),
        )
        if not path:
            return json.dumps({"ok": False, "canceled": True})
        if not path.lower().endswith(".pdf"):
            path += ".pdf"
        return json.dumps({"ok": True, "path": path})

    @Slot(str, str, result=str)
    def renderPdf(self, path: str, meta_json: str) -> str:
        try:
            meta = json.loads(meta_json or "{}")
        except json.JSONDecodeError:
            meta = {}
        width = float(meta.get("w") or 595.3)
        height = float(meta.get("h") or 841.9)
        page_size = QPageSize(QSizeF(width, height), QPageSize.Point, "Document", QPageSize.ExactMatch)
        layout = QPageLayout(page_size, QPageLayout.Portrait, QMarginsF(0, 0, 0, 0), QPageLayout.Millimeter)
        page = self.window.view.page()
        self._pdf_result = ""

        if self._pdf_callback is not None:
            try:
                page.pdfPrintingFinished.disconnect(self._pdf_callback)
            except (RuntimeError, TypeError):
                pass
        self._pdf_callback = self._on_pdf_finished
        page.pdfPrintingFinished.connect(self._pdf_callback)
        page.printToPdf(path, layout)
        return json.dumps({"ok": True, "pending": True, "path": path})

    @Slot(result=str)
    def pdfStatus(self) -> str:
        return self._pdf_result

    def _on_pdf_finished(self, path: str, success: bool) -> None:
        page = self.window.view.page()
        if self._pdf_callback is not None:
            try:
                page.pdfPrintingFinished.disconnect(self._pdf_callback)
            except (RuntimeError, TypeError):
                pass
            self._pdf_callback = None
        payload = {"ok": bool(success), "path": path}
        if success:
            self.settings.add_recent(path)
        else:
            payload["error"] = i18n.tr("Chromium could not render the PDF")
        self._pdf_result = json.dumps(payload)

    @Slot(str, result=str)
    def printDocument(self, meta_json: str) -> str:
        try:
            meta = json.loads(meta_json or "{}")
        except json.JSONDecodeError:
            meta = {}
        width = float(meta.get("w") or 595.3)
        height = float(meta.get("h") or 841.9)
        printer = QPrinter(QPrinter.HighResolution)
        page_size = QPageSize(QSizeF(width, height), QPageSize.Point, "Document", QPageSize.ExactMatch)
        printer.setPageSize(page_size)
        printer.setPageMargins(QMarginsF(0, 0, 0, 0), QPageLayout.Millimeter)
        printer.setDocName(self.document_name)
        dialog = QPrintDialog(printer, self.window)
        dialog.setWindowTitle(i18n.tr("Print"))
        if dialog.exec() != QPrintDialog.Accepted:
            return json.dumps({"ok": False, "canceled": True})
        page = self.window.view.page()
        self._print_result = ""
        if self._print_callback is not None:
            try:
                page.printingFinished.disconnect(self._print_callback)
            except (RuntimeError, TypeError):
                pass
        self._print_callback = self._on_print_finished
        page.printingFinished.connect(self._print_callback)
        page.print(printer)
        return json.dumps({"ok": True, "pending": True})

    @Slot(result=str)
    def printStatus(self) -> str:
        return self._print_result

    def _on_print_finished(self, success: bool) -> None:
        page = self.window.view.page()
        if self._print_callback is not None:
            try:
                page.printingFinished.disconnect(self._print_callback)
            except (RuntimeError, TypeError):
                pass
            self._print_callback = None
        payload = {"ok": bool(success)}
        if not success:
            payload["error"] = i18n.tr("The printer did not accept the job")
        self._print_result = json.dumps(payload)

    # ------------------------------------------------------------- autosave
    @Slot(str, str, str, str, result=str)
    def autosave(self, model_json: str, doc_id: str, name: str, source: str) -> str:
        if not self.settings.get("autosave", True):
            return json.dumps({"ok": False})
        try:
            payload = {
                "id": doc_id,
                "name": name,
                "source": source,
                "savedAt": time.time(),
                "model": json.loads(model_json),
            }
            target = self.settings.autosave_dir() / (doc_id + ".json")
            tmp = target.with_suffix(".tmp")
            with open(tmp, "w", encoding="utf-8") as handle:
                json.dump(payload, handle, ensure_ascii=False)
            os.replace(tmp, target)
            return json.dumps({"ok": True, "path": str(target)})
        except Exception as exc:  # noqa: BLE001
            return json.dumps({"ok": False, "error": str(exc)})

    def _list_autosaves(self) -> list:
        items = []
        folder = self.settings.autosave_dir()
        for path in sorted(folder.glob("*.json")):
            try:
                with open(path, "r", encoding="utf-8") as handle:
                    data = json.load(handle)
                items.append(
                    {
                        "id": data.get("id") or path.stem,
                        "name": data.get("name") or path.stem,
                        "source": data.get("source") or "",
                        "savedAt": data.get("savedAt") or 0,
                    }
                )
            except Exception:
                continue
        items.sort(key=lambda item: item.get("savedAt", 0), reverse=True)
        return items[:10]

    @Slot(str, result=str)
    def loadAutosave(self, doc_id: str) -> str:
        target = self.settings.autosave_dir() / (doc_id + ".json")
        try:
            with open(target, "r", encoding="utf-8") as handle:
                data = json.load(handle)
            return json.dumps({"ok": True, **data})
        except Exception as exc:  # noqa: BLE001
            return json.dumps({"ok": False, "error": str(exc)})

    @Slot(str, result=str)
    def discardAutosave(self, doc_id: str) -> str:
        target = self.settings.autosave_dir() / (doc_id + ".json")
        try:
            if target.exists():
                target.unlink()
            return json.dumps({"ok": True})
        except Exception as exc:  # noqa: BLE001
            return json.dumps({"ok": False, "error": str(exc)})

    @Slot(str, result=str)
    def clearAutosave(self, doc_id: str) -> str:
        return self.discardAutosave(doc_id)

    # -------------------------------------------------------------- window
    @Slot(str)
    def setTitle(self, title: str) -> None:
        self.document_name = title or "Workbook1"
        self.window.set_document_title(self.document_name, self.modified)

    @Slot(bool)
    def setModified(self, modified: bool) -> None:
        self.modified = bool(modified)
        self.window.set_document_title(self.document_name, self.modified)

    @Slot(str)
    def setDocumentPath(self, path: str) -> None:
        self.document_path = path or None
        if path:
            self.document_name = os.path.splitext(os.path.basename(path))[0]
            self.window.set_document_title(self.document_name, self.modified)

    @Slot()
    def requestClose(self) -> None:
        self.window.close()

    @Slot()
    def closeNow(self) -> None:
        self.window.force_close()

    @Slot(str)
    def openExternal(self, url: str) -> None:
        if url:
            QDesktopServices.openUrl(QUrl(url))

    @Slot(str, str)
    def setSetting(self, key: str, value: str) -> None:
        try:
            parsed = json.loads(value)
        except json.JSONDecodeError:
            parsed = value
        self.settings.set(key, parsed)
        if key == "uiLanguage":
            self.window.apply_ui_language(parsed)

    @Slot(result=str)
    def getRecent(self) -> str:
        recents = []
        for path in self.settings.get("recent", []):
            recents.append({"path": path, "name": os.path.basename(path), "exists": os.path.exists(path)})
        return json.dumps(recents)
