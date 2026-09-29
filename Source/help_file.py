"""Editable help file for the Qt application.

The default help ships with the app. A copy is placed in the user's app data
folder the first time Help is opened; that copy can be edited and is never
overwritten.
"""

from __future__ import annotations

import shutil
import sys
from pathlib import Path

HELP_FILE_NAME = "Accessible Ebook Library Manager Help.txt"
FALLBACK_TEXT = "Help could not be loaded. Press Alt to open the menu bar."


def bundled_help_path() -> Path:
    base = getattr(sys, "_MEIPASS", None)
    if base:
        return Path(base) / HELP_FILE_NAME
    return Path(__file__).resolve().parent.parent / "Documentation" / HELP_FILE_NAME


def user_help_path(app_data: Path) -> Path:
    return Path(app_data) / HELP_FILE_NAME


def ensure_user_help(app_data: Path, bundled: Path | None = None) -> Path:
    """Return the editable help file, copying the default in if it is missing."""
    target = user_help_path(app_data)
    source = bundled or bundled_help_path()
    if not target.exists() and source.exists():
        try:
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(source, target)
        except OSError:
            pass
    return target


def load_help_text(app_data: Path, bundled: Path | None = None) -> str:
    """Read the user's help file, falling back to the bundled default."""
    source = bundled or bundled_help_path()
    for path in (ensure_user_help(app_data, source), source):
        try:
            text = path.read_text(encoding="utf-8").strip()
        except OSError:
            continue
        if text:
            return text
    return FALLBACK_TEXT
