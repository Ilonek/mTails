# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller recipe for the mTails desktop spreadsheet.

Build with:
    .venv/bin/pyinstaller --noconfirm --clean mTails.spec
The one-folder bundle lands in ``dist/mTails`` and is zipped for release.
"""

from PyInstaller.utils.hooks import collect_data_files

datas = [("ui", "ui"), ("assets", "assets")]

# Qt's own catalogues are needed for translated standard dialogs (Open/Save).
datas += collect_data_files("PySide6", subdir="translations", includes=["qtbase_*.qm", "qt_*.qm"])

a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=[
        "PySide6.QtWebEngineWidgets",
        "PySide6.QtWebEngineCore",
        "PySide6.QtWebChannel",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["tkinter", "unittest", "pydoc_data", "test", "tests"],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="mTails",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon="assets/mtails.png",
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="mTails",
)
