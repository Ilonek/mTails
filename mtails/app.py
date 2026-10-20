"""Qt application shell hosting the mTails user interface."""

from __future__ import annotations

import os
import sys

from PySide6.QtCore import (
    QCoreApplication,
    QLibraryInfo,
    QObject,
    Qt,
    QTimer,
    QTranslator,
    QUrl,
)
from PySide6.QtGui import QColor, QDesktopServices, QIcon, QKeySequence, QPalette
from PySide6.QtWebChannel import QWebChannel
from PySide6.QtWebEngineCore import QWebEnginePage, QWebEngineProfile, QWebEngineSettings
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtWidgets import QApplication, QMainWindow, QMenu, QMessageBox

from . import APP_ID, APP_NAME, APP_TITLE, APP_VERSION, i18n
from .bridge import Bridge
from .config import Settings


def resource_dir() -> str:
    """Return the folder that holds the ``ui`` and ``assets`` trees.

    PyInstaller unpacks bundled data into ``sys._MEIPASS`` while a source
    checkout keeps everything next to the ``mtails`` package.
    """
    if getattr(sys, "frozen", False):
        return getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


BASE_DIR = resource_dir()
UI_DIR = os.path.join(BASE_DIR, "ui")


_qt_translator = None


def install_qt_translations(language: str) -> None:
    """Load Qt's own catalog so standard dialogs follow the UI language."""
    global _qt_translator
    app = QApplication.instance()
    if app is None:
        return
    if _qt_translator is not None:
        app.removeTranslator(_qt_translator)
        _qt_translator = None
    if language == "en":
        return
    translator = QTranslator(app)
    folders = [
        os.path.join(BASE_DIR, "translations"),
        QLibraryInfo.path(QLibraryInfo.TranslationsPath),
        os.path.join(BASE_DIR, "PySide6", "Qt", "translations"),
        os.path.join(BASE_DIR, "PySide6", "translations"),
        os.path.join(BASE_DIR, "_internal", "PySide6", "Qt", "translations"),
        os.path.join(BASE_DIR, "_internal", "PySide6", "translations"),
    ]
    for folder in folders:
        if folder and translator.load("qtbase_" + language, folder):
            app.installTranslator(translator)
            _qt_translator = translator
            return


class WebView(QWebEngineView):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setContextMenuPolicy(Qt.DefaultContextMenu)

    def contextMenuEvent(self, event) -> None:  # noqa: N802
        page = self.page()
        request = page.lastContextMenuRequest()
        # The sheet, its tabs, charts and comments draw their own HTML menus.
        # Qt only supplies the native edit menu for text fields.
        if request is None or not request.isContentEditable():
            event.accept()
            return
        menu = QMenu(self)
        action_specs = [
            ("Undo", QWebEnginePage.Undo, QKeySequence.Undo),
            ("Redo", QWebEnginePage.Redo, QKeySequence.Redo),
            (None, None, None),
            ("Cut", QWebEnginePage.Cut, QKeySequence.Cut),
            ("Copy", QWebEnginePage.Copy, QKeySequence.Copy),
            ("Paste", QWebEnginePage.Paste, QKeySequence.Paste),
            ("Delete", QWebEnginePage.Delete, QKeySequence.Delete),
            (None, None, None),
            ("Select All", QWebEnginePage.SelectAll, QKeySequence.SelectAll),
        ]
        for label, page_action, shortcut in action_specs:
            if label is None:
                menu.addSeparator()
                continue
            action = menu.addAction(i18n.tr(label))
            action.setShortcut(shortcut or QKeySequence())
            action.setEnabled(page.action(page_action).isEnabled())
            action.triggered.connect(page.action(page_action).trigger)

        if request.linkUrl().isValid() and not request.linkUrl().isEmpty():
            menu.addSeparator()
            open_link = menu.addAction(i18n.tr("Open Link"))
            open_link.triggered.connect(lambda: QDesktopServices.openUrl(request.linkUrl()))

        menu.exec(event.globalPos())


class MainWindow(QMainWindow):
    def __init__(self, startup_file: str = "", smoke_path: str = "", language: str = "") -> None:
        super().__init__()
        self.settings = Settings()
        if language:
            self.settings.override("uiLanguage", language)
        self.apply_ui_language(self.settings.get("uiLanguage", "system"))
        self.startup_file = startup_file or ""
        self.smoke_path = smoke_path
        self._allow_close = False
        self._page_ready = False

        self.setWindowTitle(APP_TITLE)
        self.resize(1440, 900)
        self.setMinimumSize(1024, 640)

        icon_path = os.path.join(BASE_DIR, "assets", "mtails.png")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        profile = QWebEngineProfile.defaultProfile()
        profile.setHttpCacheType(QWebEngineProfile.NoCache)
        profile.setSpellCheckEnabled(False)

        self.view = WebView(self)
        self.setCentralWidget(self.view)

        settings = self.view.settings()
        settings.setAttribute(QWebEngineSettings.LocalContentCanAccessFileUrls, True)
        settings.setAttribute(QWebEngineSettings.LocalContentCanAccessRemoteUrls, False)
        settings.setAttribute(QWebEngineSettings.JavascriptCanAccessClipboard, True)
        settings.setAttribute(QWebEngineSettings.JavascriptCanOpenWindows, False)
        settings.setAttribute(QWebEngineSettings.FocusOnNavigationEnabled, True)
        settings.setAttribute(QWebEngineSettings.ScrollAnimatorEnabled, True)
        settings.setAttribute(QWebEngineSettings.PlaybackRequiresUserGesture, True)

        self.channel = QWebChannel(self.view.page())
        self.bridge = Bridge(self, self.settings)
        self.channel.registerObject("bridge", self.bridge)
        self.view.page().setWebChannel(self.channel)

        self.view.loadFinished.connect(self._on_load_finished)
        index_path = os.path.join(UI_DIR, "index.html")
        self.view.load(QUrl.fromLocalFile(index_path))

        QTimer.singleShot(300, self._apply_palette)

    # ------------------------------------------------------------- lifecycle
    def apply_ui_language(self, setting: str) -> None:
        """Switch the native Qt strings to the chosen UI language."""
        i18n.set_language(setting)
        install_qt_translations(i18n.language())

    def _apply_palette(self) -> None:
        palette = QPalette()
        palette.setColor(QPalette.Window, QColor("#202020"))
        self.setPalette(palette)

    def _on_load_finished(self, ok: bool) -> None:
        self._page_ready = bool(ok)
        if not ok:
            print("mTails: failed to load the user interface", file=sys.stderr)
            return
        if self.smoke_path:
            QTimer.singleShot(3500, self._run_smoke)

    def _run_smoke(self) -> None:
        script = """
        (function(){
          try {
            return JSON.stringify(window.__smoke ? window.__smoke() : {ok:false, reason:'no hook'});
          } catch (err) { return JSON.stringify({ok:false, error:String(err)}); }
        })();
        """
        self.view.page().runJavaScript(script, 0, self._on_smoke_result)

    def _on_smoke_result(self, result) -> None:
        print("SMOKE:" + str(result), flush=True)
        path = self.smoke_path
        try:
            if path and path not in ("1", "true"):
                pixmap = self.grab()
                pixmap.save(path, "PNG")
                print("SCREENSHOT:" + path, flush=True)
        except Exception as exc:  # noqa: BLE001
            print("SCREENSHOT_ERROR:" + str(exc), flush=True)
        # QtWebEngine sometimes keeps the process alive after quit().
        QTimer.singleShot(2000, lambda: os._exit(0))
        QTimer.singleShot(250, QApplication.instance().quit)

    def set_document_title(self, name: str, modified: bool) -> None:
        marker = " *" if modified else ""
        self.setWindowTitle("%s%s - %s" % (name or "Workbook1", marker, APP_TITLE))
        self.setWindowFilePath(self.bridge.document_path or "")

    def force_close(self) -> None:
        self._allow_close = True
        self.close()

    def closeEvent(self, event) -> None:  # noqa: N802
        if self._allow_close:
            event.accept()
            return
        if not self.bridge.modified:
            event.accept()
            return

        event.ignore()
        box = QMessageBox(self)
        box.setWindowTitle(APP_TITLE)
        box.setIcon(QMessageBox.Warning)
        box.setText(i18n.tr('Do you want to save changes to "{name}"?', name=self.bridge.document_name))
        box.setInformativeText(i18n.tr("Your changes will be lost if you don't save them."))
        box.setStandardButtons(QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel)
        box.button(QMessageBox.Save).setText(i18n.tr("Save"))
        box.button(QMessageBox.Discard).setText(i18n.tr("Don't Save"))
        box.button(QMessageBox.Cancel).setText(i18n.tr("Cancel"))
        box.setDefaultButton(QMessageBox.Save)
        choice = box.exec()
        if choice == QMessageBox.Save:
            self.view.page().runJavaScript("window.__saveBeforeClose && window.__saveBeforeClose();")
        elif choice == QMessageBox.Discard:
            self.force_close()


def main() -> int:
    args = list(sys.argv[1:])
    smoke_path = ""
    language = ""
    if "--smoke" in args:
        index = args.index("--smoke")
        args.pop(index)
        if index < len(args) and not args[index].startswith("-"):
            smoke_path = args.pop(index)
        else:
            smoke_path = os.path.join("/tmp", "mtails_smoke.png")
    for flag in ("--lang", "--language"):
        if flag in args:
            index = args.index(flag)
            args.pop(index)
            if index < len(args) and not args[index].startswith("-"):
                language = args.pop(index).lower()
    if not language:
        for entry in list(args):
            if entry.startswith("--lang="):
                language = entry.split("=", 1)[1].lower()
                args.remove(entry)
    if language not in ("en", "pl"):
        language = ""
    startup_file = args[0] if args else ""

    QCoreApplication.setApplicationName(APP_NAME)
    QCoreApplication.setOrganizationName(APP_NAME)
    QCoreApplication.setApplicationVersion(APP_VERSION)
    os.environ.setdefault("QTWEBENGINE_CHROMIUM_FLAGS", "--disable-features=TranslateUI")

    app = QApplication(sys.argv)
    app.setApplicationDisplayName(APP_TITLE)
    app.setStyle("Fusion")
    app.setDesktopFileName(APP_ID)
    icon_path = os.path.join(BASE_DIR, "assets", "mtails.png")
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))

    window = MainWindow(startup_file=startup_file, smoke_path=smoke_path, language=language)
    window.show()
    return app.exec()
