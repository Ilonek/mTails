"""Translation of native Qt strings.

The web UI has its own dictionary (``ui/js/i18n.js``); this module covers the
strings that Qt itself draws: the grid context menu, the close prompt, file
dialog titles and native error boxes.  Qt's own standard widgets (file and
print dialogs) are localised through the ``qtbase_pl`` catalog loaded in
``app.py``.
"""

from __future__ import annotations

from PySide6.QtCore import QLocale

_PL = {
    # context menu
    "Undo": "Cofnij",
    "Redo": "Ponów",
    "Cut": "Wytnij",
    "Copy": "Kopiuj",
    "Paste": "Wklej",
    "Delete": "Usuń",
    "Select All": "Zaznacz wszystko",
    "Open Link": "Otwórz łącze",
    # close prompt
    'Do you want to save changes to "{name}"?': "Czy chcesz zapisać zmiany w skoroszycie „{name}”?",
    "Your changes will be lost if you don't save them.": "Zmiany zostaną utracone, jeśli ich nie zapiszesz.",
    "Save": "Zapisz",
    "Don't Save": "Nie zapisuj",
    "Cancel": "Anuluj",
    # native dialogs
    "Open": "Otwórz",
    "Save As": "Zapisz jako",
    "Export": "Eksportuj",
    "Export as PDF": "Eksportuj jako PDF",
    "Print": "Drukuj",
    "Save failed": "Nie można zapisać",
    "The workbook could not be saved:\n{error}": "Nie można zapisać skoroszytu:\n{error}",
    # file filters
    "All supported spreadsheets": "Wszystkie obsługiwane arkusze",
    "Excel Workbook": "Skoroszyt programu Excel",
    "Excel Macro-Enabled Workbook": "Skoroszyt programu Excel z obsługą makr",
    "CSV (Comma delimited)": "CSV (rozdzielany przecinkami)",
    "Text (Tab delimited)": "Tekst (rozdzielany tabulatorami)",
    "PDF Document": "Dokument PDF",
    # errors returned to the UI
    "File not found: {path}": "Nie znaleziono pliku: {path}",
    "Unsupported file type: {ext}": "Nieobsługiwany typ pliku: {ext}",
    "unknown": "nieznany",
    "Could not open the workbook:\n{error}": "Nie można otworzyć skoroszytu:\n{error}",
    "Invalid workbook data: {error}": "Nieprawidłowe dane skoroszytu: {error}",
    "Unsupported format: {format}": "Nieobsługiwany format: {format}",
    "Chromium could not render the PDF": "Nie udało się wygenerować pliku PDF",
    "The printer did not accept the job": "Drukarka nie przyjęła zadania drukowania",
    "Legacy .xls files are not supported. Please open the file in Excel/LibreOffice and save it as .xlsx first.": "Pliki .xls w starym formacie nie są obsługiwane. Otwórz plik w programie Excel/LibreOffice i zapisz go najpierw jako .xlsx.",
    "The workbook is read-only. Use Save As to store a copy.": "Skoroszyt jest tylko do odczytu. Użyj opcji Zapisz jako, aby zapisać kopię.",
}


_current = "system"


def set_language(value: str) -> None:
    """Remember the ``uiLanguage`` setting (``system``, ``en`` or ``pl``)."""
    global _current
    _current = value if value in ("system", "en", "pl") else "system"


def language(value: str | None = None) -> str:
    """Resolve a setting to the concrete UI language."""
    setting = _current if value is None else value
    if setting in ("en", "pl"):
        return setting
    name = QLocale.system().name()
    return "pl" if name.startswith("pl") else "en"


def tr(text: str, **values: str) -> str:
    """Translate a native string, filling ``{name}`` placeholders."""
    translated = _PL.get(text, text) if language() == "pl" else text
    for key, value in values.items():
        translated = translated.replace("{" + key + "}", str(value))
    return translated
