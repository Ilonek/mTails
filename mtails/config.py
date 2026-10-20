"""Persistent application settings (recent files, theme, autosave)."""

from __future__ import annotations

import json
import os
import shutil
import sys
import tempfile
from pathlib import Path
from typing import Any, Dict

from PySide6.QtCore import QStandardPaths

from . import APP_NAME, APP_ORG

# Earlier builds of the spreadsheet family used these profile folders.
LEGACY_APP_NAMES = ("mSheet", "mCell")

DEFAULTS: Dict[str, Any] = {
    "theme": "light",
    "autosave": True,
    "autosaveInterval": 120,
    "recent": [],
    "lastFolder": "",
    "zoom": 100,
    "uiLanguage": "system",
    "defaultFont": "Calibri",
    "defaultSize": 11,
    "showGridlines": True,
    "showHeadings": True,
    "showFormulaBar": True,
    "defaultSheets": 1,
    "autorecover": True,
}


def _legacy_data_dirs() -> list[Path]:
    """Profile locations used before the application was renamed to mTails."""
    if sys.platform.startswith("win"):
        base = Path(os.environ.get("APPDATA") or (Path.home() / "AppData" / "Roaming"))
    elif sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support"
    else:
        base = Path(os.environ.get("XDG_DATA_HOME") or (Path.home() / ".local" / "share"))
    # QStandardPaths used <org>/<app>; older builds also had a single-level fallback.
    paths = []
    for name in LEGACY_APP_NAMES:
        if name == APP_NAME:
            paths.append(base / name / name)
        else:
            paths.extend([base / name / name, base / name])
    return paths


def _migrate_legacy(target: Path) -> None:
    """Copy settings and recovery snapshots from a former profile once."""
    for legacy in _legacy_data_dirs():
        if not legacy.is_dir() or legacy == target:
            continue
        try:
            settings = legacy / "settings.json"
            if settings.is_file() and not (target / "settings.json").exists():
                shutil.copy2(settings, target / "settings.json")
            snapshots = legacy / "autosave"
            if snapshots.is_dir():
                destination = target / "autosave"
                destination.mkdir(parents=True, exist_ok=True)
                for snapshot in snapshots.glob("*.json"):
                    copy = destination / snapshot.name
                    if not copy.exists():
                        shutil.copy2(snapshot, copy)
        except OSError:
            continue
        return


def data_dir() -> Path:
    root = QStandardPaths.writableLocation(QStandardPaths.AppDataLocation)
    # Qt returns a relative location when no application name is registered
    # (for example inside test harnesses); never write next to the sources.
    candidates = [Path(root)] if root and os.path.isabs(root) else []
    generic = QStandardPaths.writableLocation(QStandardPaths.GenericDataLocation)
    if generic and os.path.isabs(generic):
        candidates.append(Path(generic) / APP_ORG / APP_NAME)
    candidates.append(Path.home() / ".local" / "share" / APP_ORG / APP_NAME)
    for path in candidates:
        try:
            path.mkdir(parents=True, exist_ok=True)
            if os.access(path, os.W_OK):
                _migrate_legacy(path)
                return path
        except OSError:
            continue
    fallback = Path(tempfile.gettempdir()) / APP_NAME
    fallback.mkdir(parents=True, exist_ok=True)
    _migrate_legacy(fallback)
    return fallback


def config_path() -> Path:
    return data_dir() / "settings.json"


def _load() -> Dict[str, Any]:
    try:
        with open(config_path(), "r", encoding="utf-8") as handle:
            stored = json.load(handle)
        merged = dict(DEFAULTS)
        merged.update(stored)
        return merged
    except Exception:
        return dict(DEFAULTS)


class Settings:
    """Small JSON backed settings store."""

    def __init__(self) -> None:
        self._data = _load()

    def get(self, key: str, default: Any = None) -> Any:
        return self._data.get(key, DEFAULTS.get(key, default))

    def set(self, key: str, value: Any) -> None:
        self._data[key] = value
        self.save()

    def override(self, key: str, value: Any) -> None:
        """Set a value for the current session without writing it to disk."""
        self._data[key] = value

    def update(self, values: Dict[str, Any]) -> None:
        self._data.update(values)
        self.save()

    def save(self) -> None:
        try:
            tmp = config_path().with_suffix(".tmp")
            with open(tmp, "w", encoding="utf-8") as handle:
                json.dump(self._data, handle, indent=2)
            os.replace(tmp, config_path())
        except Exception:
            pass

    def add_recent(self, path: str) -> None:
        path = os.path.abspath(path)
        recent = [p for p in self._data.get("recent", []) if os.path.abspath(p) != path]
        recent.insert(0, path)
        self._data["recent"] = recent[:20]
        self._data["lastFolder"] = os.path.dirname(path)
        self.save()

    def clear_recent(self) -> None:
        self._data["recent"] = []
        self.save()

    def autosave_dir(self) -> Path:
        path = data_dir() / "autosave"
        path.mkdir(parents=True, exist_ok=True)
        return path
