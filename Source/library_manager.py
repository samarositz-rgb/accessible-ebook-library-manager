"""
Accessible Ebook Library Manager
A screen-reader-friendly starter ebook manager for Windows.

Features:
- Standard Tkinter controls.
- Windows-style menu bar for screen-reader friendly command access.
- Plain listbox instead of a table so each book is spoken as one complete labeled row.
- Add EPUB, PDF, DOCX, TXT, MOBI, AZW, AZW3, KFX, HTML, ZIP, and other ebook/document files.
- Stores metadata in a simple SQLite database.
- For EPUB files, writes title, author, source, tags, and notes into the EPUB file itself.
- Copies imported books into a managed library folder.
- Search by title, author, source, tags, format, or notes.
- Opens selected book with the default Windows app.
- Opens Kindle for PC.
- Exports a selected book to another folder.

This app does not remove DRM. It manages files you are allowed to copy and read.
"""

import logging
import os
import filecmp
import shutil
import sqlite3
import subprocess
import sys
import zipfile
from logging.handlers import RotatingFileHandler
import re
import html
import posixpath
import tkinter as tk
import json
import queue
import tempfile
import threading
import time
import uuid
import urllib.parse
import urllib.request
import winsound
import ctypes
from io import BytesIO
from datetime import datetime, timedelta, timezone
from pathlib import Path
try:
    import winreg
except ImportError:
    winreg = None
from tkinter import (
    Tk,
    StringVar,
    END,
    SINGLE,
    LEFT,
    RIGHT,
    BOTH,
    Y,
    X,
    filedialog,
    messagebox,
    Menu,
    Listbox,
    Scrollbar,
)
from tkinter import ttk
from xml.etree import ElementTree as ET

from calibre_tools import CALIBRE_METADATA_EXTENSIONS, find_calibre_tool, read_calibre_metadata
from document_text import (
    detect_language_from_text,
    extract_text_for_indexing,
    read_text_from_epub_preserve_lines,
    read_text_for_metadata_detection,
    strip_xml_html_tags,
    strip_xml_html_tags_preserve_lines,
)
from book_flags import extract_source_book_id
from library_utils import (
    folder_file_stats,
    metadata_score_from_detection,
    metadata_score_from_row,
    normalize_duplicate_key,
    normalize_isbn_key,
    replace_folder_from_backup,
    safe_filename,
    sync_folder_contents,
    title_keys_look_same,
)
from online_metadata import lookup_online_metadata
from book_categories import normalize_import_tags


logger = logging.getLogger(__name__)

APP_NAME = "Accessible Ebook Library Manager"
DB_NAME = "library.db"
BOOKS_FOLDER = "Books"

SUPPORTED_EXTENSIONS = {
    ".epub", ".pdf", ".docx", ".doc", ".txt", ".rtf",
    ".mobi", ".azw", ".azw3", ".kfx", ".kfx-zip", ".prc", ".html", ".htm", ".zip"
}
VOICE_DREAM_LIBRARY_NOTICE = (
    "Please do not delete this file, or any files within these folders. "
    "This is your voice dream library, and we cannot recover deleted files."
)

BOOK_LIST_SPEECH_FIELDS = [
    ("title", "Title"),
    ("author", "Author"),
    ("series", "Series"),
    ("series_number", "Series number"),
    ("edition", "Edition"),
    ("year", "Year"),
    ("isbn", "ISBN"),
    ("publisher", "Publisher"),
    ("source", "Source"),
    ("tags", "Tags"),
    ("format", "Format"),
    ("added_at", "Date added"),
]
BOOK_LIST_SHORTCUT_FIELDS = [
    "title",
    "author",
    "series",
    "series_number",
    "edition",
    "year",
    "isbn",
    "publisher",
    "source",
    "tags",
]
DEFAULT_BOOK_LIST_SPEECH_FIELDS = ["title", "author"]
BOOK_LIST_DISPLAY_MODES = {
    "single": "Single list",
    "multiple": "Multiple alphabetical lists",
}
DEFAULT_BOOK_LIST_DISPLAY_MODE = "single"
BOOK_LIST_ALPHABET_GROUPS = [
    ("A-F", "abcdef"),
    ("G-L", "ghijkl"),
    ("M-R", "mnopqr"),
    ("S-Z", "stuvwxyz"),
    ("Other", ""),
]
NO_EDITION_VALUE = "No edition"
ACCESSIBILITY_METADATA_KEYS = [
    "accessibility_summary",
    "accessibility_features",
    "accessibility_hazards",
    "accessibility_access_modes",
    "accessibility_access_modes_sufficient",
    "accessibility_certified_by",
]
EXPLORER_CONTEXT_MENU_KEY = "AccessibleEbookLibraryManagerAdd"
EXPLORER_CONTEXT_MENU_LABEL = "Add to Accessible Ebook Library Manager"
BACKUP_SCHEDULES = {
    "on_demand": ("On Demand", None),
    "daily": ("Daily", timedelta(days=1)),
    "weekly": ("Weekly", timedelta(days=7)),
    "monthly": ("Monthly", timedelta(days=30)),
}
DEFAULT_BACKUP_SCHEDULE = "on_demand"
AUTOMATIC_DUPLICATE_CLEANUP_INTERVAL_MS = 30 * 60 * 1000
DEFAULT_AUTOMATIC_DUPLICATE_CLEANUP = "0"
MISSING_METADATA_SOUND_MODES = {
    "tags": ("Missing Tags Only", ["tags"]),
    "author": ("Missing Author or Tags", ["author", "tags"]),
    "useful": ("Missing Author, Tags, Edition, or Year", ["author", "tags", "edition", "year"]),
    "complete": ("Missing Author, Tags, Edition, Year, ISBN, or Publisher", ["author", "tags", "edition", "year", "isbn", "publisher"]),
}
DEFAULT_MISSING_METADATA_SOUND_MODE = "author"

OPF_NS = "http://www.idpf.org/2007/opf"
DC_NS = "http://purl.org/dc/elements/1.1/"
CONTAINER_NS = "urn:oasis:names:tc:opendocument:xmlns:container"

ET.register_namespace("opf", OPF_NS)
ET.register_namespace("dc", DC_NS)

WINDOWS_NO_CONSOLE_FLAGS = getattr(subprocess, "CREATE_NO_WINDOW", 0)
SCHEMA_VERSION = 4


def find_nvda_controller_dll() -> Path | None:
    names = ["nvdaControllerClient64.dll", "nvdaControllerClient32.dll"]
    folders = [
        Path(os.environ.get("PROGRAMFILES", "")) / "NVDA",
        Path(os.environ.get("PROGRAMFILES(X86)", "")) / "NVDA",
        Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "NVDA",
        Path.cwd(),
    ]
    for folder in folders:
        for name in names:
            path = folder / name
            if path.exists():
                return path
    for name in names:
        try:
            ctypes.WinDLL(name)
            return Path(name)
        except Exception:
            pass
    user_profile = Path(os.environ.get("USERPROFILE", ""))
    for root_name in ["Downloads", "Desktop", "Documents"]:
        root = user_profile / root_name
        if not root.exists():
            continue
        for name in names:
            try:
                match = next(root.rglob(name), None)
            except Exception:
                match = None
            if match and match.exists():
                return match
    return None


def app_data_folder() -> Path:
    base = os.environ.get("APPDATA")
    if base:
        folder = Path(base) / "AccessibleEbookLibraryManager"
    else:
        folder = Path.home() / "AccessibleEbookLibraryManager"
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def managed_books_folder(folder: Path) -> Path:
    preferred = folder / BOOKS_FOLDER
    try:
        preferred.mkdir(parents=True, exist_ok=True)
        if preferred.is_dir():
            return preferred
    except FileExistsError:
        pass

    for name in ["Library Books", "Managed Books", "Imported Books"]:
        candidate = folder / name
        candidate.mkdir(parents=True, exist_ok=True)
        if candidate.is_dir():
            return candidate

    raise RuntimeError("Could not create a folder for imported books.")


def utc_now_text():
    return datetime.now(timezone.utc).replace(tzinfo=None).replace(microsecond=0).isoformat() + "Z"


def parse_utc_text(value):
    if not value:
        return None
    try:
        return datetime.fromisoformat(value.replace("Z", ""))
    except ValueError:
        return None


def cloud_backup_subfolder(base_folder: Path) -> Path:
    return base_folder / "Accessible Ebook Library Manager Backups"


def format_byte_count(byte_count: int) -> str:
    size = float(max(byte_count, 0))
    for unit in ["bytes", "KB", "MB", "GB"]:
        if size < 1024 or unit == "GB":
            if unit == "bytes":
                return f"{int(size)} {unit}"
            return f"{size:.1f} {unit}"
        size /= 1024
    return f"{size:.1f} GB"


def app_launch_command_for_file_argument():
    if getattr(sys, "frozen", False):
        return f'"{sys.executable}" --import "%1"'

    executable = Path(sys.executable)
    pythonw = executable.with_name("pythonw.exe")
    launcher = pythonw if pythonw.exists() else executable
    script = Path(__file__).resolve()
    return f'"{launcher}" "{script}" --import "%1"'


class LibraryDatabase:
    def __init__(self):
        self.folder = app_data_folder()
        self.db_path = self.folder / DB_NAME
        self.books_path = managed_books_folder(self.folder)
        self.connection = sqlite3.connect(self.db_path)
        self.create_tables()

    def create_tables(self):
        self.connection.execute(
            """
            CREATE TABLE IF NOT EXISTS books (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT NOT NULL,
                author TEXT DEFAULT '',
                source TEXT DEFAULT '',
                tags TEXT DEFAULT '',
                notes TEXT DEFAULT '',
                format TEXT DEFAULT '',
                original_path TEXT DEFAULT '',
                stored_path TEXT NOT NULL UNIQUE,
                added_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        self.connection.execute(
            """
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT DEFAULT ''
            )
            """
        )
        self.connection.execute(
            """
            CREATE VIRTUAL TABLE IF NOT EXISTS books_fts USING fts5(
                book_text,
                tokenize='porter unicode61'
            )
            """
        )
        self.connection.execute("CREATE INDEX IF NOT EXISTS idx_books_title ON books(title COLLATE NOCASE)")
        self.connection.execute("CREATE INDEX IF NOT EXISTS idx_books_author ON books(author COLLATE NOCASE)")
        self.connection.execute("CREATE INDEX IF NOT EXISTS idx_books_added ON books(added_at)")
        self.connection.execute("CREATE INDEX IF NOT EXISTS idx_books_format ON books(format)")
        self.connection.commit()
        self.ensure_book_columns()
        self.connection.execute("CREATE INDEX IF NOT EXISTS idx_books_year ON books(year)")
        self.connection.commit()
        if int(self.get_setting("schema_version", "0")) < SCHEMA_VERSION:
            self.set_setting("schema_version", str(SCHEMA_VERSION))

    def ensure_book_columns(self):
        """Add newer metadata columns to existing libraries without deleting data."""
        existing = {row[1] for row in self.connection.execute("PRAGMA table_info(books)")}
        columns = {
            "edition": "TEXT DEFAULT ''",
            "year": "TEXT DEFAULT ''",
            "isbn": "TEXT DEFAULT ''",
            "publisher": "TEXT DEFAULT ''",
            "cover_url": "TEXT DEFAULT ''",
            "accessibility_summary": "TEXT DEFAULT ''",
            "accessibility_features": "TEXT DEFAULT ''",
            "accessibility_hazards": "TEXT DEFAULT ''",
            "accessibility_access_modes": "TEXT DEFAULT ''",
            "accessibility_access_modes_sufficient": "TEXT DEFAULT ''",
            "accessibility_certified_by": "TEXT DEFAULT ''",
            "content_indexed_at": "TEXT DEFAULT NULL",
            "series": "TEXT DEFAULT ''",
            "series_number": "TEXT DEFAULT ''",
            "content_flags": "TEXT DEFAULT ''",
            "source_book_id": "TEXT DEFAULT ''",
        }
        for name, column_type in columns.items():
            if name not in existing:
                self.connection.execute(f"ALTER TABLE books ADD COLUMN {name} {column_type}")
        self.connection.execute("CREATE INDEX IF NOT EXISTS idx_books_series ON books(series COLLATE NOCASE)")
        self.connection.commit()

    def get_setting(self, key, default=""):
        cursor = self.connection.execute("SELECT value FROM settings WHERE key = ?", (key,))
        row = cursor.fetchone()
        if not row:
            return default
        return row[0]

    def set_setting(self, key, value):
        self.connection.execute(
            "INSERT INTO settings (key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value",
            (key, value),
        )
        self.connection.commit()

    def backup_to(self, destination: Path):
        destination.parent.mkdir(parents=True, exist_ok=True)
        backup_connection = sqlite3.connect(destination)
        try:
            self.connection.commit()
            self.connection.backup(backup_connection)
            backup_connection.commit()
        finally:
            backup_connection.close()

    def close(self):
        self.connection.close()

    def index_book_content(self, book_id: int, text: str):
        self.connection.execute("DELETE FROM books_fts WHERE rowid = ?", (book_id,))
        if text.strip():
            self.connection.execute(
                "INSERT INTO books_fts(rowid, book_text) VALUES (?, ?)",
                (book_id, text[:500_000]),
            )
        self.connection.execute(
            "UPDATE books SET content_indexed_at = ? WHERE id = ?",
            (utc_now_text(), book_id),
        )
        self.connection.commit()

    def search_content(self, query: str) -> set:
        if not query.strip():
            return set()
        # Treat the user's words as one literal phrase, not an implicit AND
        # between unrelated words anywhere in the book.
        fts_query = '"' + query.strip().replace('"', '""') + '"'
        try:
            cursor = self.connection.execute(
                "SELECT rowid FROM books_fts WHERE books_fts MATCH ?",
                (fts_query,),
            )
            return {row[0] for row in cursor.fetchall()}
        except Exception:
            return set()

    def get_unindexed_books(self) -> list:
        cursor = self.connection.execute(
            """
            SELECT id, stored_path FROM books
            WHERE content_indexed_at IS NULL
            ORDER BY added_at DESC
            """
        )
        return cursor.fetchall()

    def clear_all_content_index(self):
        self.connection.execute("DELETE FROM books_fts")
        self.connection.execute("UPDATE books SET content_indexed_at = NULL")
        self.connection.commit()

    def add_book(self, title, author, source, tags, notes, original_path, stored_path):
        tags = normalize_import_tags(tags)
        ext = Path(stored_path).suffix.lower().replace(".", "")
        cursor = self.connection.execute(
            """
            INSERT INTO books
            (title, author, source, tags, notes, format, original_path, stored_path)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (title, author, source, tags, notes, ext, original_path, stored_path),
        )
        self.connection.commit()
        return cursor.lastrowid

    def link_book_file(self, book_id, path, file_format=None):
        """Point an existing entry at a file or folder that stays where it is (not copied)."""
        path = str(path)
        if file_format is None:
            file_format = Path(path).suffix.lower().lstrip(".")
        self.connection.execute(
            "UPDATE books SET stored_path = ?, original_path = ?, format = ? WHERE id = ?",
            (path, path, file_format, book_id),
        )
        self.connection.commit()

    def update_extra_fields(self, book_id, edition="", year="", isbn="", publisher=""):
        self.connection.execute(
            """
            UPDATE books
            SET edition = ?, year = ?, isbn = ?, publisher = ?
            WHERE id = ?
            """,
            (edition, year, isbn, publisher, book_id),
        )
        self.connection.commit()

    def update_series_fields(self, book_id, series="", series_number=""):
        self.connection.execute(
            """
            UPDATE books
            SET series = ?, series_number = ?
            WHERE id = ?
            """,
            (series, series_number, book_id),
        )
        self.connection.commit()

    def update_bulk_metadata_fields(self, book_ids, values):
        allowed = {
            "author", "source", "tags", "notes", "edition", "year",
            "publisher", "series",
        }
        fields = [key for key in allowed if key in values]
        if not fields:
            return 0
        assignments = ", ".join(f"{key} = ?" for key in fields)
        params = [values[key] for key in fields]
        count = 0
        for book_id in book_ids:
            self.connection.execute(
                f"UPDATE books SET {assignments} WHERE id = ?",
                params + [book_id],
            )
            count += 1
        self.connection.commit()
        return count

    def update_cover_url(self, book_id, cover_url=""):
        self.connection.execute(
            "UPDATE books SET cover_url = ? WHERE id = ?",
            (cover_url, book_id),
        )
        self.connection.commit()

    def update_accessibility_metadata(self, book_id, metadata):
        values = {key: str(metadata.get(key, "") or "") for key in ACCESSIBILITY_METADATA_KEYS}
        self.connection.execute(
            """
            UPDATE books
            SET accessibility_summary = ?,
                accessibility_features = ?,
                accessibility_hazards = ?,
                accessibility_access_modes = ?,
                accessibility_access_modes_sufficient = ?,
                accessibility_certified_by = ?
            WHERE id = ?
            """,
            (
                values["accessibility_summary"],
                values["accessibility_features"],
                values["accessibility_hazards"],
                values["accessibility_access_modes"],
                values["accessibility_access_modes_sufficient"],
                values["accessibility_certified_by"],
                book_id,
            ),
        )
        self.connection.commit()

    def update_book(self, book_id, title, author, source, tags, notes):
        self.connection.execute(
            """
            UPDATE books
            SET title = ?, author = ?, source = ?, tags = ?, notes = ?
            WHERE id = ?
            """,
            (title, author, source, tags, notes, book_id),
        )
        self.connection.commit()

    def delete_book(self, book_id, delete_file=False):
        row = self.get_book(book_id)
        if row and delete_file:
            stored_path = row[8]
            try:
                # Only files the library copied itself are deleted. Linked files, such
                # as audiobooks left in a Libation folder, are never touched.
                inside_library = Path(stored_path).resolve().is_relative_to(self.books_path.resolve())
                if inside_library and os.path.exists(stored_path):
                    os.remove(stored_path)
            except (OSError, ValueError):
                pass
        self.connection.execute("DELETE FROM books_fts WHERE rowid = ?", (book_id,))
        self.connection.execute("DELETE FROM books WHERE id = ?", (book_id,))
        self.connection.commit()

    def all_books_for_duplicate_check(self):
        cursor = self.connection.execute(
            """
            SELECT id, title, author, source, tags, notes, format, original_path, stored_path, added_at,
                   edition, year, isbn, publisher, cover_url,
                   accessibility_summary, accessibility_features, accessibility_hazards,
                   accessibility_access_modes, accessibility_access_modes_sufficient, accessibility_certified_by,
                   series, series_number
            FROM books
            ORDER BY title COLLATE NOCASE, author COLLATE NOCASE, added_at
            """
        )
        return cursor.fetchall()

    def get_book(self, book_id):
        cursor = self.connection.execute(
            """
            SELECT id, title, author, source, tags, notes, format, original_path, stored_path, added_at,
                   edition, year, isbn, publisher, cover_url,
                   accessibility_summary, accessibility_features, accessibility_hazards,
                   accessibility_access_modes, accessibility_access_modes_sufficient, accessibility_certified_by,
                   series, series_number
            FROM books WHERE id = ?
            """,
            (book_id,),
        )
        return cursor.fetchone()

    def get_book_by_original_path(self, original_path):
        cursor = self.connection.execute(
            """
            SELECT id, title, author, source, tags, notes, format, original_path, stored_path, added_at,
                   edition, year, isbn, publisher, cover_url,
                   accessibility_summary, accessibility_features, accessibility_hazards,
                   accessibility_access_modes, accessibility_access_modes_sufficient, accessibility_certified_by,
                   series, series_number
            FROM books WHERE original_path = ?
            """,
            (original_path,),
        )
        return cursor.fetchone()

    def set_content_flags(self, book_id, flags):
        """Store per-book notes such as image_only_pdf as a comma-separated list."""
        cleaned = ",".join(dict.fromkeys(flag.strip() for flag in flags if flag and flag.strip()))
        self.connection.execute("UPDATE books SET content_flags = ? WHERE id = ?", (cleaned, book_id))
        self.connection.commit()

    def get_content_flags(self, book_id):
        row = self.connection.execute("SELECT content_flags FROM books WHERE id = ?", (book_id,)).fetchone()
        return [flag for flag in (row[0] or "").split(",") if flag] if row else []

    def set_source_book_id(self, book_id, source_book_id):
        self.connection.execute("UPDATE books SET source_book_id = ? WHERE id = ?", (source_book_id or "", book_id))
        self.connection.commit()

    def get_source_book_id(self, book_id):
        row = self.connection.execute("SELECT source_book_id FROM books WHERE id = ?", (book_id,)).fetchone()
        return (row[0] or "") if row else ""

    def find_book_id_by_source_book_id(self, source_book_id):
        if not source_book_id:
            return None
        row = self.connection.execute(
            "SELECT id FROM books WHERE source_book_id = ? ORDER BY id LIMIT 1", (source_book_id,)
        ).fetchone()
        return row[0] if row else None

    def source_book_ids_by_book(self):
        return {
            book_id: source_book_id
            for book_id, source_book_id in self.connection.execute(
                "SELECT id, source_book_id FROM books WHERE source_book_id <> ''"
            )
        }

    def search_books(
        self, query="", sort_by="title", source_filter="", tag_filter="", format_filter="", extra_ids=None,
        accessibility_feature_filter="", exclude_hazard_filter="", content_flag_filter="",
    ):
        query = query.strip()
        source_filter = source_filter.strip()
        tag_filter = tag_filter.strip()
        format_filter = format_filter.strip().lower()
        accessibility_feature_filter = accessibility_feature_filter.strip()
        # Hazards are whole values such as flashing or noFlashingHazard, so match them
        # exactly; a plain substring test would treat "no flashing hazard" as flashing.
        exclude_hazard_filter = re.sub(r"[^a-z0-9]", "", exclude_hazard_filter.casefold())
        content_flag_filter = content_flag_filter.strip()

        order_map = {
            "title": "title COLLATE NOCASE ASC, author COLLATE NOCASE ASC",
            "title_desc": "title COLLATE NOCASE DESC, author COLLATE NOCASE ASC",
            "author": "author COLLATE NOCASE ASC, title COLLATE NOCASE ASC",
            "author_desc": "author COLLATE NOCASE DESC, title COLLATE NOCASE ASC",
            "series": "series COLLATE NOCASE ASC, CAST(series_number AS REAL) ASC, title COLLATE NOCASE ASC",
            "series_desc": "series COLLATE NOCASE DESC, CAST(series_number AS REAL) ASC, title COLLATE NOCASE ASC",
            "date": "year DESC, title COLLATE NOCASE ASC",
            "date_oldest": "year ASC, title COLLATE NOCASE ASC",
            "date_added": "added_at DESC, title COLLATE NOCASE ASC",
            "date_added_oldest": "added_at ASC, title COLLATE NOCASE ASC",
        }
        order_clause = order_map.get(sort_by, order_map["title"])

        conditions = []
        params = []

        if query:
            like = f"%{query}%"
            metadata_condition = (
                "(title LIKE ?"
                " OR author LIKE ?"
                " OR source LIKE ?"
                " OR tags LIKE ?"
                " OR notes LIKE ?"
                " OR format LIKE ?"
                " OR edition LIKE ?"
                " OR year LIKE ?"
                " OR isbn LIKE ?"
                " OR publisher LIKE ?"
                " OR series LIKE ?"
                " OR series_number LIKE ?"
                " OR accessibility_summary LIKE ?"
                " OR accessibility_features LIKE ?"
                " OR accessibility_hazards LIKE ?"
                " OR accessibility_access_modes LIKE ?"
                " OR accessibility_access_modes_sufficient LIKE ?"
                " OR accessibility_certified_by LIKE ?)"
            )
            if extra_ids:
                placeholders = ",".join("?" * len(extra_ids))
                conditions.append(f"({metadata_condition} OR id IN ({placeholders}))")
                params.extend([like] * 18)
                params.extend(extra_ids)
            else:
                conditions.append(metadata_condition)
                params.extend([like] * 18)

        if source_filter:
            conditions.append("source LIKE ?")
            params.append(f"%{source_filter}%")

        if tag_filter:
            conditions.append("tags LIKE ?")
            params.append(f"%{tag_filter}%")

        if format_filter:
            conditions.append("format LIKE ?")
            params.append(f"%{format_filter}%")

        if accessibility_feature_filter:
            conditions.append("accessibility_features LIKE ?")
            params.append(f"%{accessibility_feature_filter}%")

        if exclude_hazard_filter:
            conditions.append("(',' || REPLACE(LOWER(accessibility_hazards), ' ', '') || ',') NOT LIKE ?")
            params.append(f"%,{exclude_hazard_filter},%")

        if content_flag_filter:
            conditions.append("content_flags LIKE ?")
            params.append(f"%{content_flag_filter}%")

        where_clause = ""
        if conditions:
            where_clause = "WHERE " + " AND ".join(conditions)

        cursor = self.connection.execute(
            f"""
            SELECT id, title, author, source, tags, format, added_at,
                   edition, year, isbn, publisher, series, series_number, stored_path
            FROM books
            {where_clause}
            ORDER BY {order_clause}
            """,
            params,
        )
        return cursor.fetchall()

def first_or_empty(root, xpath, namespaces):
    item = root.find(xpath, namespaces)
    if item is None or item.text is None:
        return ""
    return item.text.strip()


def normalize_book_list_speech_fields(raw_value: str) -> list[str]:
    allowed = [key for key, _label in BOOK_LIST_SPEECH_FIELDS]
    selected = [part.strip() for part in raw_value.split(",") if part.strip()]
    selected = [key for key in selected if key in allowed]
    if "title" not in selected:
        selected.insert(0, "title")
    return selected


def get_epub_opf_path(epub_path: Path) -> str:
    with zipfile.ZipFile(epub_path, "r") as archive:
        container_xml = archive.read("META-INF/container.xml")
    root = ET.fromstring(container_xml)
    rootfile = root.find(".//container:rootfile", {"container": CONTAINER_NS})
    if rootfile is None:
        raise ValueError("EPUB container file does not point to an OPF package file.")
    opf_path = rootfile.attrib.get("full-path", "")
    if not opf_path:
        raise ValueError("EPUB OPF package path is missing.")
    return opf_path


def isbn_from_metadata_value(value: str) -> str:
    """Accept an ISBN identifier, not a number embedded in a source name or URL."""
    value = re.sub(r"[\u2010-\u2015]", "-", (value or "").strip())
    value = re.sub(r"^(?:urn:isbn:|isbn(?:-1[03])?\s*:?\s*)", "", value, flags=re.IGNORECASE)
    if not re.fullmatch(r"[0-9Xx\s-]+", value):
        return ""
    digits = re.sub(r"[\s-]", "", value).upper()
    if len(digits) == 13 and digits.isdigit() and digits.startswith(("978", "979")):
        if sum(int(char) * (1 if index % 2 == 0 else 3) for index, char in enumerate(digits)) % 10 == 0:
            return digits
    if len(digits) == 10 and digits[:9].isdigit() and (digits[-1].isdigit() or digits[-1] == "X"):
        if sum((10 - index) * (10 if char == "X" else int(char)) for index, char in enumerate(digits)) % 11 == 0:
            return digits
    return ""


def read_epub_metadata(epub_path: Path) -> dict:
    try:
        opf_path = get_epub_opf_path(epub_path)
        with zipfile.ZipFile(epub_path, "r") as archive:
            opf_xml = archive.read(opf_path)

        root = ET.fromstring(opf_xml)
        ns = {"opf": OPF_NS, "dc": DC_NS}
        subjects = []
        for subject in root.findall(".//dc:subject", ns):
            if subject.text:
                subjects.append(subject.text.strip())

        # dc:source may identify the print edition rather than the supplying
        # service. Keep that ISBN out of Source and prefer this EPUB's identifier.
        sources = [(element.text or "").strip() for element in root.findall(".//dc:source", ns)]
        source = next((value for value in sources if value and not isbn_from_metadata_value(value)), "")
        identifiers = []
        unique_id = root.get("unique-identifier")
        for element in root.findall(".//dc:identifier", ns):
            isbn = isbn_from_metadata_value(element.text or "")
            if isbn:
                identifiers.append((element.get("id") == unique_id, len(isbn), isbn))
        identifiers.sort(key=lambda item: (item[0], item[1]), reverse=True)

        return {
            "title": first_or_empty(root, ".//dc:title", ns),
            "author": first_or_empty(root, ".//dc:creator", ns),
            "source": source,
            "isbn": identifiers[0][2] if identifiers else "",
            "tags": normalize_import_tags(subjects),
            "notes": first_or_empty(root, ".//dc:description", ns),
            "publisher": first_or_empty(root, ".//dc:publisher", ns),
            "year": first_or_empty(root, ".//dc:date", ns)[:4],
        }
    except Exception:
        return {}


def metadata_values_by_property(root, property_name):
    values = []
    for item in root.findall(".//opf:meta", {"opf": OPF_NS}):
        if item.attrib.get("property") == property_name and item.text:
            values.append(item.text.strip())
    return values


def first_metadata_value_by_property(root, property_name):
    values = metadata_values_by_property(root, property_name)
    return values[0] if values else ""


def opf_relative_path(opf_path, href):
    base = posixpath.dirname(opf_path)
    return posixpath.normpath(posixpath.join(base, href))


def read_epub_accessibility_metadata(epub_path: Path) -> dict:
    result = {key: "" for key in ACCESSIBILITY_METADATA_KEYS}
    try:
        opf_path = get_epub_opf_path(epub_path)
        with zipfile.ZipFile(epub_path, "r") as archive:
            opf_xml = archive.read(opf_path)
            names = set(archive.namelist())
            root = ET.fromstring(opf_xml)
            ns = {"opf": OPF_NS}

            declared_features = metadata_values_by_property(root, "schema:accessibilityFeature")
            declared_hazards = metadata_values_by_property(root, "schema:accessibilityHazard")
            access_modes = metadata_values_by_property(root, "schema:accessMode")
            access_modes_sufficient = metadata_values_by_property(root, "schema:accessModeSufficient")
            summary = first_metadata_value_by_property(root, "schema:accessibilitySummary")
            certified_by = (
                first_metadata_value_by_property(root, "a11y:certifiedBy")
                or first_metadata_value_by_property(root, "schema:accessibilityAPI")
            )

            manifest_items = []
            nav_present = False
            page_list_present = False
            ncx_present = False
            for item in root.findall(".//opf:manifest/opf:item", ns):
                href = item.attrib.get("href", "")
                media_type = item.attrib.get("media-type", "")
                properties = item.attrib.get("properties", "")
                full_path = opf_relative_path(opf_path, href) if href else ""
                manifest_items.append((full_path, media_type, properties))
                if "nav" in properties.split():
                    nav_present = True
                if media_type == "application/x-dtbncx+xml":
                    ncx_present = True

            content_paths = [
                path for path, media_type, _properties in manifest_items
                if media_type in {"application/xhtml+xml", "text/html"} and path in names
            ]
            image_total = 0
            image_with_alt = 0
            heading_count = 0
            lang_present = False
            for content_path in content_paths[:80]:
                try:
                    text = archive.read(content_path).decode("utf-8", errors="ignore")
                except Exception:
                    continue
                for match in re.finditer(r"<img\b[^>]*>", text, flags=re.IGNORECASE):
                    image_total += 1
                    if re.search(r"\balt\s*=\s*(['\"])[\s\S]*?\1", match.group(0), flags=re.IGNORECASE):
                        image_with_alt += 1
                heading_count += len(re.findall(r"<h[1-6]\b", text, flags=re.IGNORECASE))
                if re.search(r"\b(?:xml:)?lang\s*=", text, flags=re.IGNORECASE):
                    lang_present = True
                if re.search(r"epub:type\s*=\s*(['\"])[^'\"]*\bpage-list\b", text, flags=re.IGNORECASE):
                    page_list_present = True

            inferred_features = []
            if nav_present or ncx_present:
                inferred_features.append("tableOfContents")
            if page_list_present:
                inferred_features.append("pageNavigation")
            if heading_count:
                inferred_features.append("structuralNavigation")
            if image_total and image_with_alt == image_total:
                inferred_features.append("alternativeText")

            hazards = declared_hazards or ["unknown"]
            features = sorted(set(declared_features + inferred_features), key=str.casefold)
            result["accessibility_features"] = ", ".join(features)
            result["accessibility_hazards"] = ", ".join(hazards)
            result["accessibility_access_modes"] = ", ".join(access_modes)
            result["accessibility_access_modes_sufficient"] = "; ".join(access_modes_sufficient)
            result["accessibility_certified_by"] = certified_by

            checks = []
            checks.append("navigation present" if nav_present or ncx_present else "navigation not found")
            checks.append("page list present" if page_list_present else "page list not found")
            if image_total:
                checks.append(f"image alt text {image_with_alt} of {image_total}")
            else:
                checks.append("no images found")
            checks.append(f"{heading_count} heading{'s' if heading_count != 1 else ''} found")
            checks.append("language declared" if lang_present else "language not found")
            if summary:
                result["accessibility_summary"] = f"{summary} Checked by app: {', '.join(checks)}."
            else:
                result["accessibility_summary"] = "Checked by app: " + ", ".join(checks) + "."
            return result
    except Exception:
        return result


def set_single_text(metadata_element, tag_name, value):
    existing = metadata_element.findall(tag_name)
    if value:
        if existing:
            existing[0].text = value
            for extra in existing[1:]:
                metadata_element.remove(extra)
        else:
            item = ET.SubElement(metadata_element, tag_name)
            item.text = value
    else:
        for item in existing:
            metadata_element.remove(item)


def set_single_meta_property(metadata_element, property_name, value):
    existing = [
        item for item in metadata_element.findall(f"{{{OPF_NS}}}meta")
        if item.attrib.get("property") == property_name
    ]
    if value:
        if existing:
            existing[0].text = value
            for extra in existing[1:]:
                metadata_element.remove(extra)
        else:
            item = ET.SubElement(metadata_element, f"{{{OPF_NS}}}meta")
            item.set("property", property_name)
            item.text = value
    else:
        for item in existing:
            metadata_element.remove(item)


def write_epub_metadata(epub_path: Path, title: str, author: str, source: str, tags: str, notes: str):
    opf_path = get_epub_opf_path(epub_path)
    temp_path = epub_path.with_suffix(epub_path.suffix + ".tmp")
    backup_path = epub_path.with_suffix(epub_path.suffix + ".bak")

    with zipfile.ZipFile(epub_path, "r") as source_archive:
        opf_xml = source_archive.read(opf_path)
        root = ET.fromstring(opf_xml)
        ns = {"opf": OPF_NS, "dc": DC_NS}

        metadata = root.find("opf:metadata", ns)
        if metadata is None:
            metadata = ET.SubElement(root, f"{{{OPF_NS}}}metadata")

        set_single_text(metadata, f"{{{DC_NS}}}title", title)
        set_single_text(metadata, f"{{{DC_NS}}}creator", author)
        set_single_text(metadata, f"{{{DC_NS}}}source", source)
        set_single_text(metadata, f"{{{DC_NS}}}description", notes)
        set_single_text(metadata, f"{{{DC_NS}}}language", "en")

        for subject in metadata.findall(f"{{{DC_NS}}}subject"):
            metadata.remove(subject)

        for tag in [part.strip() for part in tags.split(",") if part.strip()]:
            subject = ET.SubElement(metadata, f"{{{DC_NS}}}subject")
            subject.text = tag

        new_opf_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)

        with zipfile.ZipFile(temp_path, "w") as target_archive:
            for item in source_archive.infolist():
                if item.filename == opf_path:
                    target_archive.writestr(item, new_opf_xml)
                else:
                    target_archive.writestr(item, source_archive.read(item.filename))

    if not backup_path.exists():
        shutil.copy2(epub_path, backup_path)

    os.replace(temp_path, epub_path)


def natural_sort_key(path: Path):
    parts = re.split(r"(\d+)", path.name.casefold())
    return [int(part) if part.isdigit() else part for part in parts]


def book_part_sort_key(path: Path):
    name = path.stem.casefold()
    if re.search(r"(?:^|[_\-\s])fm(?:$|[_\-\s])", name):
        section_rank = 0
    elif re.search(r"(?:^|[_\-\s])em(?:$|[_\-\s])", name):
        section_rank = 2
    else:
        section_rank = 1

    chapter_match = re.search(r"(?:^|[_\-\s])ch(?:apter)?\s*0*(\d+)", name, flags=re.IGNORECASE)
    chapter_number = int(chapter_match.group(1)) if chapter_match else 9999
    return (section_rank, chapter_number, natural_sort_key(path))


def page_id_for_label(page_label):
    page_label = str(page_label or "").strip()
    page_id = "page-" + re.sub(r"[^0-9A-Za-z]+", "-", page_label).strip("-").lower()
    return page_id or "page"


def parse_text_toc_entries(text: str, max_entries: int = 350):
    lines = [clean_metadata_line(line) for line in re.split(r"[\r\n]+", text or "")]
    lines = [line for line in lines if line]
    start = None
    for index, line in enumerate(lines):
        if re.fullmatch(r"(?:table of )?contents", line, flags=re.IGNORECASE):
            start = index + 1
            break
    if start is None:
        return []

    entries = []
    seen = set()
    for line in lines[start:start + 900]:
        if re.match(r"^(?:(?:p?age)\s+)+[0-9ivxlcdm]+\s*$", line, flags=re.IGNORECASE):
            continue
        normalized_line = re.sub(r"\s+", " ", line).strip()
        match = re.match(
            r"^(?P<title>.+?)(?:\.{2,}|\s{2,}|\s)(?P<page>[0-9]+|[ivxlcdm]+)\s*$",
            normalized_line,
            flags=re.IGNORECASE,
        )
        if not match:
            continue
        title = re.sub(r"\.{2,}", " ", match.group("title")).strip(" .\t")
        title = re.sub(r"\s+", " ", title)
        page_label = match.group("page").strip()
        if len(title) < 3 or len(title) > 180:
            continue
        if title.casefold() in {"table of contents", "contents", "chapter", "part", "page"}:
            continue
        key = (title.casefold(), page_label.casefold())
        if key in seen:
            continue
        seen.add(key)
        entries.append({"title": title, "page": page_label, "page_id": page_id_for_label(page_label)})
        if len(entries) >= max_entries:
            break
    return entries


def looks_like_text_toc_entry_line(line: str):
    normalized_line = re.sub(r"\s+", " ", line or "").strip()
    if re.fullmatch(r"(?:table of )?contents", normalized_line, flags=re.IGNORECASE):
        return True
    match = re.match(
        r"^(?P<title>.+?)(?:\.{2,}|\s{2,}|\s)(?P<page>[0-9]+|[ivxlcdm]+)\s*$",
        normalized_line,
        flags=re.IGNORECASE,
    )
    if not match:
        return False
    title = re.sub(r"\.{2,}", " ", match.group("title")).strip(" .\t")
    if len(title) < 3 or len(title) > 180:
        return False
    return title.casefold() not in {"chapter", "part", "page"}


def looks_like_standalone_heading_line(line: str):
    normalized = re.sub(r"\s+", " ", line or "").strip()
    if not normalized:
        return False
    if len(normalized) > 140:
        return False
    if re.match(r"^(?:chapter|part|section|appendix|index|front matter)\b", normalized, flags=re.IGNORECASE):
        return True
    if re.match(r"^(?:\u00a7)+\s*\d+(?:\.\d+)*\.?\s+\S+", normalized):
        return True
    if re.match(r"^[A-Z][A-Z0-9 ,;:'\u2019\-\(\)&]{5,}$", normalized) and len(normalized.split()) <= 12:
        return True
    return False


def looks_like_heading_continuation_line(line: str):
    normalized = re.sub(r"\s+", " ", line or "").strip()
    if not normalized:
        return False
    if len(normalized) > 90:
        return False
    if normalized.endswith((".", "?", "!", ";", ":")):
        return False
    if re.search(r"\b(the|and|or|but|with|from|into|after|before|because|that|this|will|shall|must|have|has|was|were|are|is)\b", normalized, flags=re.IGNORECASE):
        return False
    return bool(re.match(r"^[A-Z0-9\u00a7][A-Za-z0-9\u00a7 ,'\u2019\-\(\)&:]+$", normalized))


def is_blank_page_notice(line: str):
    normalized = re.sub(r"\s+", " ", line or "").strip(" .;:-\t").casefold()
    return normalized in {
        "this page was intentionally left blank",
        "this page intentionally left blank",
        "intentionally left blank",
        "blank page",
    }


def is_restricted_access_notice(line: str):
    normalized = re.sub(r"\s+", " ", line or "").strip(" .;:-\t").casefold()
    normalized = normalized.replace("\u2019", "'")
    normalized = normalized.replace("publisher?s", "publisher's")
    return normalized in {
        "this document has been prepared exclusively for the use of a student with a print disability. it is protected by the publisher's original copyright. it may not be shared or transferred to any other person",
        "this document has been prepared exclusively for the use of a student with a print disability it is protected by the publisher's original copyright it may not be shared or transferred to any other person",
    }


def is_import_boilerplate_line(line: str):
    return is_blank_page_notice(line) or is_restricted_access_notice(line)


def is_page_label_line(line: str):
    line = re.sub(r"\s+", " ", line or "").strip()
    if re.fullmatch(r"(?:(?:p?age)\s+)+([0-9ivxlcdm]+)", line, flags=re.IGNORECASE):
        return True
    if re.fullmatch(r"\d{1,4}", line):
        return True
    return False


def page_label_from_line(line: str):
    line = re.sub(r"\s+", " ", line or "").strip()
    labeled = re.fullmatch(r"(?:(?:p?age)\s+)+([0-9ivxlcdm]+)", line, flags=re.IGNORECASE)
    if labeled:
        return labeled.group(1)
    if re.fullmatch(r"\d{1,4}", line):
        return line
    return ""


def cleaned_lines_for_reflow(text: str):
    raw_lines = [line.strip() for line in re.sub(r"\r\n?", "\n", text or "").split("\n")]
    nonempty = [re.sub(r"\s+", " ", line).strip() for line in raw_lines if line.strip()]
    counts = {}
    for line in nonempty:
        key = line.casefold()
        counts[key] = counts.get(key, 0) + 1

    repeated = set()
    for line in nonempty:
        key = line.casefold()
        if counts.get(key, 0) < 4:
            continue
        if len(line) > 120:
            continue
        if is_page_label_line(line) or looks_like_text_toc_entry_line(line):
            continue
        if re.match(r"^(chapter|part|section|appendix)\b", line, flags=re.IGNORECASE):
            continue
        repeated.add(key)

    cleaned = []
    removed = 0
    for line in raw_lines:
        normalized = re.sub(r"\s+", " ", line).strip()
        if normalized and (is_import_boilerplate_line(normalized) or normalized.casefold() in repeated):
            removed += 1
            continue
        cleaned.append(line)
    return cleaned, removed


def clean_metadata_line(line: str) -> str:
    line = html.unescape(line).strip()
    line = re.sub(r"\s+", " ", line)
    line = re.sub(r"^[#*\\\-\u2013\u2014: ]+", "", line)
    line = re.sub(r"^(?:bookshare|accessible book|daisy book)\s*[:\-]\s*", "", line, flags=re.IGNORECASE)
    return line.strip()


def is_bookshare_notice_line(line: str) -> bool:
    lowered = re.sub(r"\s+", " ", line or "").strip().lower()
    if not lowered:
        return False
    return (
        lowered == "notice"
        or "this accessible media has been made available to people with bona fide disabilities" in lowered
        or lowered.startswith("this notice tells you about restrictions on the use of this accessible media")
        or "bona fide disabilities that affect reading" in lowered
    )


def extract_isbn_from_text(text: str) -> str:
    """Find a likely ISBN, including Word superscript/check-digit artifacts."""
    if not text:
        return ""

    text = re.sub(r"[\u2010-\u2015]", "-", text)
    labeled_patterns = [
        r"\bISBN(?:-1[03])?[ \t]*:?[ \t]*([0-9Xx][0-9Xx \t\-\^\.]{8,35}[0-9Xx])",
        r"\bISBN(?:-1[03])?[ \t]*:?[ \t]*((?:97[89][-\t ]?)?[0-9][0-9Xx \t\-\^\.]{8,35})",
    ]
    general_patterns = [
        r"\b((?:97[89][-\s]?)?[0-9][-\s]?[0-9]{2,5}[-\s]?[0-9]{2,7}[-\s]?[0-9]{1,7}(?:[-\s\^]?[0-9Xx])?)\b",
    ]

    candidates = []
    for pattern in labeled_patterns:
        candidates.extend(re.findall(pattern, text, flags=re.IGNORECASE))
    for pattern in general_patterns:
        candidates.extend(re.findall(pattern, text, flags=re.IGNORECASE))

    fallback = ""
    isbn10 = ""
    for candidate in candidates:
        cleaned = re.sub(r"[^0-9Xx]", "", candidate)
        if len(cleaned) == 13:
            return cleaned
        if len(cleaned) == 10 and not isbn10:
            isbn10 = cleaned
        if not fallback and len(cleaned) in {9, 11, 12} and cleaned.startswith(("978", "979")):
            fallback = cleaned
    if isbn10:
        return isbn10
    return fallback


def extract_publication_year_from_text(text: str, require_label: bool = False, include_copyright: bool = True) -> str:
    if not text:
        return ""

    sample = text[:60000]
    edition_match = re.search(
        r"\b(19[5-9]\d|20[0-4]\d)(?:\s*[-/]\s*(?:19[5-9]\d|20[0-4]\d))?\s+(?:abridged\s+)?(?:edition|ed\.)\b",
        sample,
        flags=re.IGNORECASE,
    )
    if edition_match:
        return edition_match.group(1)

    labeled_patterns = [
        r"\b(?:published|publication date|date of publication|publication year)\D{0,80}(19[5-9]\d|20[0-4]\d)",
    ]
    if include_copyright:
        labeled_patterns.append(r"(?:\u00a9|\(c\)|copyright)\D{0,80}(19[5-9]\d|20[0-4]\d)")
    for pattern in labeled_patterns:
        match = re.search(pattern, sample, flags=re.IGNORECASE)
        if match:
            return match.group(1)

    if require_label:
        return ""

    match = re.search(r"\b(19[5-9]\d|20[0-4]\d)\b", sample, flags=re.IGNORECASE)
    return match.group(1) if match else ""


def extract_publisher_from_text(text: str) -> str:
    if not text:
        return ""

    sample = text[:60000]
    patterns = [
        r"(?:\u00a9|\(c\)|copyright)\s*(?:19[5-9]\d|20[0-4]\d)\s+([^\r\n]{2,80})",
        r"\bpublished by\s+([^\r\n]{2,100})",
        r"\bpublisher\s*[:\-]\s*([^\r\n]{2,100})",
    ]
    for pattern in patterns:
        match = re.search(pattern, sample, flags=re.IGNORECASE)
        if not match:
            continue
        publisher = clean_metadata_line(match.group(1))
        publisher = clean_publisher_value(publisher)
        if looks_like_publisher(publisher):
            return publisher

    if re.search(r"\bA\s+Thomson\s+Reuters\s+business\b", sample, flags=re.IGNORECASE):
        return "Thomson Reuters"
    return ""


def detect_series_fields(lines: list[str], existing: dict) -> tuple[str, str]:
    series = existing.get("series", "") or ""
    series_number = existing.get("series_number", "") or ""
    if not series:
        value = line_after_label(lines, ["series", "collection"], max_index=180)
        if value:
            series = clean_metadata_line(value)
    if not series_number:
        value = line_after_label(lines, ["series number", "book number", "volume", "vol"], max_index=180)
        if value:
            match = re.search(r"\d+(?:\.\d+)?", value)
            series_number = match.group(0) if match else clean_metadata_line(value)
    if (not series or not series_number):
        for line in lines[:180]:
            match = re.search(
                r"\b(?:book|volume|vol\.?)\s+(\d+(?:\.\d+)?)\s+(?:of|in)\s+(.+)",
                line,
                flags=re.IGNORECASE,
            )
            if match:
                if not series_number:
                    series_number = match.group(1)
                if not series:
                    series = clean_metadata_line(match.group(2))
                break
    return series, series_number


def clean_filename_title(path: Path) -> str:
    title = path.stem
    title = re.sub(r"[_]+", " ", title)
    title = re.sub(r"\s*-\s*", " - ", title)
    title = re.sub(r"\s+", " ", title)
    title = re.sub(r"\b(?:bookshare|daisy|epub|pdf|docx|accessible)\b", "", title, flags=re.IGNORECASE)
    title = re.sub(r"\b(?:bs|bookshare)[-_ ]?\d+\b", "", title, flags=re.IGNORECASE)
    title = re.sub(r"\s+", " ", title).strip(" -_.,")
    return title or path.stem.replace("_", " ").replace("-", " ")


def is_weak_title(title: str, path: Path) -> bool:
    if not title:
        return True
    cleaned_title = re.sub(r"\s+", " ", title).strip().lower()
    cleaned_stem = clean_filename_title(path).lower()
    raw_stem = path.stem.replace("_", " ").replace("-", " ").lower()
    return cleaned_title in {cleaned_stem, raw_stem, path.stem.lower()} or len(cleaned_title) < 3


def clean_author_value(value: str) -> str:
    value = clean_metadata_line(value)
    value = re.sub(r"^(?:by|authors?|written by|edited by)\b\s*[:\-]?\s*", "", value, flags=re.IGNORECASE)
    value = re.sub(r"\s+(?:narrated by|read by|bookshare|copyright|all rights reserved).*$", "", value, flags=re.IGNORECASE)
    value = re.sub(r"\s+", " ", value).strip(" .;,")
    if value.isupper() and len(value.split()) <= 8:
        value = value.title()
        value = re.sub(r"\b([A-Z])\.", lambda match: match.group(1).upper() + ".", value)
    return value


def clean_title_value(value: str) -> str:
    value = clean_metadata_line(value)
    value = re.sub(r"^(?:title|book title)\s*[:\-]?\s*", "", value, flags=re.IGNORECASE)
    value = re.sub(r"\s+", " ", value).strip(" .;,")
    return value


def title_words(value: str) -> list[str]:
    return [
        word.casefold()
        for word in re.findall(r"[A-Za-z][A-Za-z0-9']+", value or "")
        if len(word) > 2
    ]


def has_useful_filename_title(path: Path) -> bool:
    words = title_words(clean_filename_title(path))
    return len(words) >= 2


def looks_like_machine_pdf_title(value: str) -> bool:
    title = clean_title_value(value)
    if not title:
        return True

    lower = title.casefold()
    if re.search(r"\.(?:pdf|docx?|rtf|txt|epub)$", lower):
        return True
    if re.search(r"\b[a-z]\d{3,}\b", lower) or re.search(r"\b\d{4,}\b", lower):
        return True

    words = title_words(title)
    if not words:
        return True

    short_words = [word for word in words if len(word) <= 4]
    if len(words) >= 4 and len(short_words) / len(words) >= 0.75 and re.search(r"\d", title):
        return True

    machine_terms = {"matls", "contr", "ch", "fm", "em", "cb", "a012636"}
    if any(word in machine_terms for word in words):
        return True

    return False


def looks_like_boilerplate_title(value: str) -> bool:
    lower = clean_title_value(value).casefold()
    if not lower:
        return True
    if lower in {"published", "copyright", "contents", "table of contents", "chapter", "page"}:
        return True
    boilerplate = [
        "all rights reserved",
        "paragraph",
        "several pages",
        "photocopy",
        "recording",
        "licensed to",
        "permission",
        "publisher",
        "excerpted material",
        "reprinted by permission",
        "appearing in this book",
        "evaluation for use",
        "third persons",
        "respect for rights",
    ]
    if re.match(r"^\d+(?:\.\d+)+\s+", lower):
        return True
    return any(fragment in lower for fragment in boilerplate)


def title_candidate_matches_filename(candidate: str, path: Path) -> bool:
    filename_words = set(title_words(clean_filename_title(path)))
    candidate_words = set(title_words(candidate))
    if not filename_words or not candidate_words:
        return False
    overlap = filename_words & candidate_words
    needed = min(2, len(candidate_words), len(filename_words))
    return len(overlap) >= needed


def looks_like_useful_title_candidate(candidate: str, path: Path, prefer_filename: bool = False) -> bool:
    candidate = clean_title_value(candidate)
    if len(candidate) < 4 or looks_like_boilerplate_title(candidate):
        return False

    words = title_words(candidate)
    if len(words) > 18:
        return False

    if prefer_filename:
        if looks_like_machine_pdf_title(candidate):
            return False
        if has_useful_filename_title(path) and not title_candidate_matches_filename(candidate, path):
            return False

    return True


def should_replace_title(current_title: str, candidate: str, path: Path, prefer_filename: bool = False) -> bool:
    if not looks_like_useful_title_candidate(candidate, path, prefer_filename=prefer_filename):
        return False
    if not current_title:
        return True
    if prefer_filename and has_useful_filename_title(path) and not is_weak_title(current_title, path):
        return False
    return is_weak_title(current_title, path)


def title_page_candidate(lines: list[str], path: Path) -> str:
    filename_words = set(title_words(clean_filename_title(path)))
    for line in lines[:12]:
        candidate = clean_title_value(line)
        if not looks_like_useful_title_candidate(candidate, path, prefer_filename=False):
            continue
        candidate_words = set(title_words(candidate))
        if filename_words and filename_words.issubset(candidate_words):
            return candidate.title() if candidate.isupper() else candidate
    return ""


def clean_publisher_value(value: str) -> str:
    value = clean_metadata_line(value)
    value = re.sub(r"^(?:publisher|published by|imprint)\s*[:\-]?\s*", "", value, flags=re.IGNORECASE)
    value = re.sub(r"^(?:copyright|\(c\)|\u00a9)\s*(?:[a-z]\s*)?(?:19[5-9]\d|20[0-4]\d)(?:\s*,\s*(?:19[5-9]\d|20[0-4]\d))*\s*", "", value, flags=re.IGNORECASE)
    value = re.sub(r"^(?:[,;]\s*)?(?:19[5-9]\d|20[0-4]\d)(?:\s*,\s*(?:19[5-9]\d|20[0-4]\d))*\s*", "", value, flags=re.IGNORECASE)
    value = re.sub(r"[\s.]+(?:all rights.*|printed in .*)$", "", value, flags=re.IGNORECASE)
    value = re.sub(r"\s+", " ", value).strip(" .;,")
    return value


def looks_like_publisher(value: str) -> bool:
    if not value:
        return False
    lower = value.casefold()
    if len(value) > 120:
        return False
    if any(fragment in lower for fragment in [
        " across the street ", " friend", " gun", " saw ", " pages are indicated",
        "law school", "advisory board", "created this publication", "accurate and authoritative",
        "particular jurisdiction", "does not render", "professional advice",
        "reprinted by permission", " by permission", "excerpted material",
    ]):
        return False
    if value.rstrip().endswith("-"):
        return False
    publisher_pattern = (
        r"\b(?:press|publishing|publisher|publishers|books|house|pearson|mcgraw|"
        r"cengage|wiley|openstax|scholastic|harper|penguin|simon|houghton|"
        r"macmillan|oxford|cambridge|west|aspen|wolters|kluwer|lexisnexis)\b"
        r"|random house|thomson reuters|matthew bender|carolina academic"
    )
    return bool(re.search(publisher_pattern, lower, flags=re.IGNORECASE))


def labeled_value(line: str, labels: list[str]) -> str:
    label_pattern = "|".join(re.escape(label) for label in labels)
    match = re.match(rf"^(?:{label_pattern})\s*[:\-]\s*(.+)$", line, flags=re.IGNORECASE)
    if match:
        return clean_metadata_line(match.group(1))
    return ""


def line_after_label(lines: list[str], labels: list[str], max_index: int = 120) -> str:
    label_pattern = "|".join(re.escape(label) for label in labels)
    for index, line in enumerate(lines[:max_index]):
        value = labeled_value(line, labels)
        if value:
            return value
        if re.fullmatch(rf"(?:{label_pattern})\s*:?", line, flags=re.IGNORECASE):
            for next_line in lines[index + 1:index + 4]:
                if next_line and not re.fullmatch(rf"(?:{label_pattern})\s*:?", next_line, flags=re.IGNORECASE):
                    return clean_metadata_line(next_line)
    return ""


def looks_like_author(value: str) -> bool:
    if not value:
        return False
    lower = value.lower()
    blocked_fragments = [
        "chapter", "contents", "copyright", "isbn", "publisher", "bookshare",
        "licensed to", "persons licensed", "all rights reserved", "photocopy",
        "recording", "permission", "ellipsis", "several pages", "one of",
        "any means", "electronic or mechanical", "co-conspirators",
        "third persons", "evaluation for use", "respect for rights",
    ]
    if any(fragment in lower for fragment in blocked_fragments):
        return False
    words = [word for word in re.split(r"\s+", value) if word]
    if not 1 <= len(words) <= 8 or len(value) > 120:
        return False
    if lower.endswith(":") or re.search(r"(?<!\b[A-Z])\.(?!\s*[A-Z]\b)", value):
        return False
    if value[:1].islower():
        return False
    if re.fullmatch(r"[A-Z\s:.'-]+", value) and (len(words) <= 1 or value.rstrip().endswith(":") or value.casefold() in {"the court"}):
        return False
    if len(words) == 1 and not re.search(r"[A-Z][a-z]+", value):
        return False

    name_signal = bool(re.search(r"\b[A-Z]\.", value)) or len(re.findall(r"\b[A-Z][a-zA-Z'.-]+\b", value)) >= 2
    if not name_signal:
        return False
    return True


def detect_metadata_from_text(path: Path, existing: dict | None = None, max_chars: int = 50000) -> dict:
    """Best-effort local metadata detector.

    This is not a true AI model. It uses EPUB metadata, file names, and common
    title/author patterns found near the start of the book text. It is designed
    to help with sources such as Bookshare, where file names are often awkward.
    """
    existing = existing or {}
    result = {
        "title": existing.get("title", ""),
        "author": existing.get("author", ""),
        "series": existing.get("series", ""),
        "series_number": existing.get("series_number", ""),
        "source": existing.get("source", ""),
        "tags": existing.get("tags", ""),
        "notes": existing.get("notes", ""),
        "edition": existing.get("edition", ""),
        "year": existing.get("year", ""),
        "isbn": existing.get("isbn", ""),
        "publisher": existing.get("publisher", ""),
    }

    if path.suffix.lower() == ".epub":
        epub_metadata = read_epub_metadata(path)
        for key in result:
            if epub_metadata.get(key):
                result[key] = epub_metadata[key]

    prefer_filename_title = path.suffix.lower() in {".pdf", ".docx", ".doc"}
    text = read_text_for_metadata_detection(path, max_chars=max_chars)
    lines = [clean_metadata_line(line) for line in re.split(r"[\r\n]+", text)]
    lines = [
        line for line in lines
        if line and len(line) < 240 and not is_bookshare_notice_line(line) and not is_import_boilerplate_line(line)
    ]

    filename_title = clean_filename_title(path)
    if (
        is_weak_title(result["title"], path)
        or (prefer_filename_title and looks_like_machine_pdf_title(result["title"]) and has_useful_filename_title(path))
    ):
        result["title"] = filename_title

    if prefer_filename_title:
        front_title = title_page_candidate(lines, path)
        if front_title:
            result["title"] = front_title

    # Look harder for labeled Bookshare/front-matter fields.
    title_value = line_after_label(lines, ["title", "book title", "name"], max_index=180)
    if title_value:
        candidate = clean_title_value(title_value)
        if should_replace_title(result["title"], candidate, path, prefer_filename=prefer_filename_title):
            result["title"] = candidate

    author_value = line_after_label(
        lines,
        ["author", "authors", "creator", "creators", "by", "written by"],
        max_index=220,
    )
    if author_value and (not result["author"] or "unknown" in result["author"].lower()):
        cleaned_author = clean_author_value(author_value)
        if looks_like_author(cleaned_author):
            result["author"] = cleaned_author

    # Look for common author patterns and title-by-author front matter.
    author_patterns = [
        r"^by\s+(.+)$",
        r"^authors?[:\s]+(.+)$",
        r"^written by\s+(.+)$",
        r"^(.+?)\s*/\s*by\s+(.+)$",
        r"^(.+?)\s+by\s+([A-Z][A-Za-z .,'-]{2,120})$",
    ]
    for line in lines[:160]:
        for pattern in author_patterns:
            match = re.match(pattern, line, flags=re.IGNORECASE)
            if not match:
                continue
            if pattern.startswith("^(.+?)"):
                possible_title = clean_title_value(match.group(1))
                possible_author = clean_author_value(match.group(2))
                if should_replace_title(result["title"], possible_title, path, prefer_filename=prefer_filename_title):
                    result["title"] = possible_title
            else:
                possible_author = clean_author_value(match.group(1))
            if (not result["author"] or "unknown" in result["author"].lower()) and looks_like_author(possible_author):
                result["author"] = possible_author
                break
        if result["author"]:
            break

    # Look for common title labels.
    for line in lines[:160]:
        match = re.match(r"^(?:title|book title)[:\s]+(.+)$", line, flags=re.IGNORECASE)
        if match:
            candidate = clean_title_value(match.group(1))
            if should_replace_title(result["title"], candidate, path, prefer_filename=prefer_filename_title):
                result["title"] = candidate
            break

    # If there is still no good title, use the first substantial line that is not a boilerplate line.
    boilerplate_words = [
        "bookshare", "copyright", "all rights reserved", "dedication",
        "contents", "table of contents", "chapter", "isbn", "published"
    ]
    if not result["title"]:
        for line in lines[:100]:
            lower = line.lower()
            if len(line) < 4:
                continue
            if is_import_boilerplate_line(line):
                continue
            if any(word in lower for word in boilerplate_words):
                continue
            if lower.startswith("by "):
                continue
            result["title"] = line
            break

    if not result["title"]:
        result["title"] = filename_title

    # ISBN, year, publisher, and edition guesses.
    if text and not result.get("isbn"):
        result["isbn"] = extract_isbn_from_text(text)

    if text:
        text_year = extract_publication_year_from_text(
            text,
            require_label=prefer_filename_title or bool(result.get("year")),
            include_copyright=not prefer_filename_title,
        )
        if text_year:
            result["year"] = text_year

    if text and not result.get("publisher"):
        publisher_value = line_after_label(
            lines,
            ["publisher", "published by", "imprint"],
            max_index=220,
        )
        if publisher_value:
            cleaned_publisher = clean_publisher_value(publisher_value)
            if looks_like_publisher(cleaned_publisher):
                result["publisher"] = cleaned_publisher
        if not result.get("publisher"):
            result["publisher"] = extract_publisher_from_text(text)
        if not result.get("publisher"):
            for line in lines[:180]:
                cleaned_publisher = clean_publisher_value(line)
                if looks_like_publisher(cleaned_publisher):
                    result["publisher"] = cleaned_publisher
                    break

    result["series"], result["series_number"] = detect_series_fields(lines, result)

    if text and not result.get("edition"):
        edition_value = line_after_label(lines, ["edition"], max_index=220)
        if edition_value and "edition" in edition_value.lower():
            result["edition"] = clean_metadata_line(edition_value)
        for line in lines[:180]:
            if result.get("edition"):
                break
            match = re.search(r"\b(\d+(?:st|nd|rd|th)\s+edition|first edition|second edition|third edition|fourth edition|fifth edition|sixth edition|seventh edition|eighth edition|ninth edition|tenth edition|revised edition|international edition|teacher'?s edition|student edition)\b", line, flags=re.IGNORECASE)
            if match:
                result["edition"] = clean_metadata_line(match.group(1))
                break
            match = re.search(r"\b((?:19[5-9]\d|20[0-4]\d)(?:\s*[-/]\s*(?:19[5-9]\d|20[0-4]\d))?\s+(?:abridged\s+)?edition)\b", line, flags=re.IGNORECASE)
            if match:
                result["edition"] = clean_metadata_line(match.group(1))
                break

    if not result["author"]:
        result["author"] = catalog_author(text)

    # Source guess.
    combined = (str(path) + " " + text[:1000]).lower()
    if not result["source"]:
        if "bookshare" in combined:
            result["source"] = "Bookshare"
        elif "kindle" in combined or path.suffix.lower() in CALIBRE_METADATA_EXTENSIONS:
            result["source"] = "Kindle"

    # Only propose supported subject categories, never format or source tags.
    if not result["tags"]:
        tags = []
        keyword_value = line_after_label(lines, ["keywords", "subjects", "subject"], max_index=120)
        if keyword_value:
            tags.extend([part.strip() for part in re.split(r"[,;]", keyword_value) if part.strip()])
        deduped_tags = []
        for tag in tags:
            if tag and tag.lower() not in [existing_tag.lower() for existing_tag in deduped_tags]:
                deduped_tags.append(tag)
        result["tags"] = ", ".join(deduped_tags)

    result["tags"] = existing.get("tags") or normalize_import_tags(result["tags"])

    # Do not auto-fill notes. Notes remain available for manual editing only.
    return result


def catalog_author(text):
    # Cataloging statements identify responsibility after a slash and before
    # the physical-description marker, even when PDF paragraphs are flattened.
    match = re.search(r"\s/\s+([^/\r\n]{3,100}?)\.\s+p\.\s*cm\b", text[:50000], re.IGNORECASE)
    candidate = clean_author_value(match.group(1)) if match else ""
    return candidate if looks_like_author(candidate) else ""


def detect_metadata_from_text_content(text: str, existing: dict | None = None) -> dict:
    existing = existing or {}
    result = {
        "title": existing.get("title", ""),
        "author": existing.get("author", ""),
        "series": existing.get("series", ""),
        "series_number": existing.get("series_number", ""),
        "source": existing.get("source", ""),
        "tags": existing.get("tags", ""),
        "notes": existing.get("notes", ""),
        "edition": existing.get("edition", ""),
        "year": existing.get("year", ""),
        "isbn": existing.get("isbn", ""),
        "publisher": existing.get("publisher", ""),
    }

    lines = [clean_metadata_line(line) for line in re.split(r"[\r\n]+", text or "")]
    lines = [
        line for line in lines
        if line and len(line) < 240 and not is_bookshare_notice_line(line) and not is_import_boilerplate_line(line)
    ]

    title_value = line_after_label(lines, ["title", "book title", "name"], max_index=180)
    if title_value:
        result["title"] = clean_title_value(title_value)

    author_value = line_after_label(
        lines,
        ["author", "authors", "creator", "creators", "by", "written by"],
        max_index=220,
    )
    if author_value:
        cleaned_author = clean_author_value(author_value)
        if looks_like_author(cleaned_author):
            result["author"] = cleaned_author

    if not result["author"]:
        for line in lines[:160]:
            match = re.match(r"^by\s+(.+)$", line, flags=re.IGNORECASE)
            if match:
                cleaned_author = clean_author_value(match.group(1))
                if looks_like_author(cleaned_author):
                    result["author"] = cleaned_author
                    break

    if not result["author"]:
        result["author"] = catalog_author(text or "")

    if not result["title"]:
        for line in lines[:100]:
            lower = line.lower()
            if len(line) < 4 or lower.startswith("by "):
                continue
            if is_import_boilerplate_line(line):
                continue
            if any(word in lower for word in ["bookshare", "copyright", "contents", "table of contents", "chapter", "isbn", "published"]):
                continue
            result["title"] = clean_title_value(line)
            break

    if text and not result.get("isbn"):
        result["isbn"] = extract_isbn_from_text(text)

    if text:
        text_year = extract_publication_year_from_text(text, require_label=bool(result.get("year")))
        if text_year:
            result["year"] = text_year

    if text and not result.get("publisher"):
        publisher_value = line_after_label(lines, ["publisher", "published by", "imprint"], max_index=220)
        if publisher_value:
            cleaned_publisher = clean_publisher_value(publisher_value)
            if looks_like_publisher(cleaned_publisher):
                result["publisher"] = cleaned_publisher
        if not result.get("publisher"):
            result["publisher"] = extract_publisher_from_text(text)

    result["series"], result["series_number"] = detect_series_fields(lines, result)

    if text and not result.get("edition"):
        edition_value = line_after_label(lines, ["edition"], max_index=220)
        if edition_value:
            result["edition"] = clean_metadata_line(edition_value)
        for line in lines[:180]:
            if result.get("edition"):
                break
            match = re.search(r"\b(\d+(?:st|nd|rd|th)\s+edition|first edition|second edition|third edition|fourth edition|fifth edition|sixth edition|seventh edition|eighth edition|ninth edition|tenth edition|revised edition|international edition|teacher'?s edition|student edition)\b", line, flags=re.IGNORECASE)
            if match:
                result["edition"] = clean_metadata_line(match.group(1))
                break

    if not result["source"]:
        result["source"] = "Bookshare" if "bookshare" in (text or "").lower() else ""

    return result



def ask_single_field_with_windows_forms(heading, field_label, instructions, current_value, parent=None):
    if not sys.platform.startswith("win") or not shutil.which("powershell"):
        return None

    payload = {
        "field_label": field_label,
        "instructions": instructions,
        "value": current_value or "",
    }

    script = r'''
param(
    [string]$InputPath,
    [string]$OutputPath,
    [string]$WindowTitle
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$encoding = New-Object System.Text.UTF8Encoding($false)
$payload = Get-Content -LiteralPath $InputPath -Raw -Encoding UTF8 | ConvertFrom-Json

$form = New-Object System.Windows.Forms.Form
$form.Text = $WindowTitle
$form.StartPosition = 'CenterScreen'
$form.Width = 760
$form.Height = 260
$form.KeyPreview = $true

$main = New-Object System.Windows.Forms.TableLayoutPanel
$main.Dock = 'Fill'
$main.Padding = New-Object System.Windows.Forms.Padding(12)
$main.ColumnCount = 1
$main.RowCount = 3
$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
$form.Controls.Add($main)

$label = New-Object System.Windows.Forms.Label
$label.AutoSize = $true
$label.MaximumSize = New-Object System.Drawing.Size(700, 0)
$label.Text = [string]$payload.field_label + ". " + [string]$payload.instructions
$main.Controls.Add($label, 0, 0)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Dock = 'Top'
$textBox.Width = 700
$textBox.AccessibleName = [string]$payload.field_label
$textBox.AccessibleDescription = [string]$payload.instructions
$textBox.Text = [string]$payload.value
$main.Controls.Add($textBox, 0, 1)

$buttons = New-Object System.Windows.Forms.FlowLayoutPanel
$buttons.FlowDirection = 'RightToLeft'
$buttons.Dock = 'Fill'
$buttons.AutoSize = $true
$main.Controls.Add($buttons, 0, 2)

$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = '&OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$buttons.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$buttons.Controls.Add($cancelButton)

$form.AcceptButton = $okButton
$form.CancelButton = $cancelButton
$form.Add_Shown({
    $form.Activate()
    $textBox.Focus()
    $textBox.SelectAll()
})

$result = $form.ShowDialog()
if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    $output = [ordered]@{ ok = $true; value = $textBox.Text }
}
else {
    $output = [ordered]@{ ok = $false; value = '' }
}

$json = $output | ConvertTo-Json -Compress
[System.IO.File]::WriteAllText($OutputPath, $json, $encoding)
'''

    with tempfile.TemporaryDirectory(prefix="aelm_single_field_") as temp_folder:
        temp = Path(temp_folder)
        input_path = temp / "single_field_input.json"
        output_path = temp / "single_field_output.json"
        script_path = temp / "single_field_dialog.ps1"
        input_path.write_text(json.dumps(payload), encoding="utf-8")
        script_path.write_text(script, encoding="utf-8")

        process = subprocess.Popen(
            [
                "powershell",
                "-WindowStyle",
                "Hidden",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-STA",
                "-File",
                str(script_path),
                "-InputPath",
                str(input_path),
                "-OutputPath",
                str(output_path),
                "-WindowTitle",
                heading or field_label,
            ],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            creationflags=WINDOWS_NO_CONSOLE_FLAGS,
        )
        start_time = time.monotonic()
        while process.poll() is None:
            if time.monotonic() - start_time > 300:
                process.kill()
                raise TimeoutError("Windows dialog timed out.")
            if parent is not None:
                try:
                    parent.update()
                except Exception:
                    pass
            time.sleep(0.05)

        stdout, stderr = process.communicate()
        if process.returncode != 0 or not output_path.exists():
            raise RuntimeError(stderr.strip() or stdout.strip() or "Windows dialog failed.")
        result = json.loads(output_path.read_text(encoding="utf-8"))
        if not result.get("ok"):
            return None
        return str(result.get("value", ""))


class AccessibleSingleFieldDialog:
    """Simple single-value dialog used for settings and filters.

    Uses a standard Entry control because Windows screen readers usually expose
    it more predictably than custom multi-line Tk text fields.
    """

    def __init__(
        self,
        parent,
        field_label: str,
        instructions: str,
        current_value: str,
        include_field_prefix: bool = False,
        heading: str = "",
    ):
        self.result = None
        self.field_label = field_label
        self.include_field_prefix = include_field_prefix
        self.prefix = f"{field_label}: " if include_field_prefix else ""
        self.value_var = StringVar(value=f"{self.prefix}{current_value or ''}")
        self.window = tk.Toplevel(parent)
        self.window.title(heading or field_label)
        self.window.transient(parent)
        self.window.grab_set()

        main = ttk.Frame(self.window, padding=12)
        main.pack(fill=BOTH, expand=True)

        prompt_text = (
            f"{field_label}. {instructions} "
            "Type or edit the value in the edit area. "
            "Press Enter to accept, or Escape to cancel."
        )
        ttk.Label(main, text=prompt_text, wraplength=650).pack(anchor="w", pady=(0, 8))

        self.entry = ttk.Entry(main, textvariable=self.value_var, width=80)
        self.entry.pack(fill=X, expand=True, pady=(0, 8))

        button_frame = ttk.Frame(main)
        button_frame.pack(anchor="e")

        ok_button = ttk.Button(button_frame, text="Next", command=self.ok)
        ok_button.pack(side=LEFT, padx=(0, 8))
        cancel_button = ttk.Button(button_frame, text="Cancel", command=self.cancel)
        cancel_button.pack(side=LEFT)

        self.window.bind("<Return>", lambda event: self.ok())
        self.window.bind("<Escape>", lambda event: self.cancel())
        self.window.protocol("WM_DELETE_WINDOW", self.cancel)

        self.window.after(50, self.focus_entry)
        self.window.after(200, self.focus_entry)

    def focus_entry(self):
        self.window.lift()
        self.window.focus_force()
        self.entry.focus_force()
        if self.include_field_prefix:
            start = len(self.prefix)
            self.entry.selection_range(start, END)
            self.entry.icursor(END)
        else:
            self.entry.selection_range(0, END)
            self.entry.icursor(END)

    def ok(self):
        value = self.value_var.get().strip()
        if self.include_field_prefix and value.lower().startswith(self.prefix.lower().strip()):
            value = value[len(self.prefix.strip()):].strip()
        self.result = value
        self.window.destroy()

    def cancel(self):
        self.result = None
        self.window.destroy()

    @staticmethod
    def ask(
        parent,
        field_label: str,
        instructions: str,
        current_value: str,
        include_field_prefix: bool = False,
        heading: str = "",
    ):
        if sys.platform.startswith("win") and shutil.which("powershell"):
            try:
                value = ask_single_field_with_windows_forms(heading or field_label, field_label, instructions, current_value, parent=parent)
                if value is None:
                    return None
                value = value.strip()
                if include_field_prefix:
                    prefix = f"{field_label}:"
                    if value.lower().startswith(prefix.lower()):
                        value = value[len(prefix):].strip()
                return value
            except Exception:
                pass

        dialog = AccessibleSingleFieldDialog(
            parent,
            field_label,
            instructions,
            current_value,
            include_field_prefix=include_field_prefix,
            heading=heading,
        )
        parent.wait_window(dialog.window)
        return dialog.result


def ask_metadata_with_windows_forms(heading, fields, initial, initial_focus_key="title", require_title=True):
    if not sys.platform.startswith("win") or not shutil.which("powershell"):
        return None

    initial = initial or {}
    payload = {
        "fields": [{"key": key, "label": label, "hint": hint} for key, label, hint in fields],
        "values": {key: str(initial.get(key, "") or "") for key, _label, _hint in fields},
        "initialFocusKey": initial_focus_key or "title",
        "requireTitle": bool(require_title),
    }

    script = r'''
param(
    [string]$InputPath,
    [string]$OutputPath,
    [string]$WindowTitle
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$payload = Get-Content -LiteralPath $InputPath -Raw -Encoding UTF8 | ConvertFrom-Json
$form = New-Object System.Windows.Forms.Form
$form.Text = $WindowTitle
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(760, 620)
$form.MinimumSize = New-Object System.Drawing.Size(720, 520)
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font

$panel = New-Object System.Windows.Forms.Panel
$panel.Dock = [System.Windows.Forms.DockStyle]::Fill
$panel.AutoScroll = $true
$form.Controls.Add($panel)

$textBoxes = @{}
$y = 12
$tabIndex = 0

foreach ($field in $payload.fields) {
    $key = [string]$field.key
    $labelText = [string]$field.label
    $hint = [string]$field.hint
    if ($hint) { $labelText = "$labelText. $hint" }

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $labelText
    $label.Location = New-Object System.Drawing.Point(12, $y)
    $label.Size = New-Object System.Drawing.Size(700, 22)
    $label.UseMnemonic = $false
    $panel.Controls.Add($label)
    $y += 24

    $box = New-Object System.Windows.Forms.TextBox
    $box.Location = New-Object System.Drawing.Point(12, $y)
    $box.Width = 700
    $box.TabIndex = $tabIndex
    $box.AccessibleName = "$($field.label) edit"
    $box.AccessibleDescription = "Edit $($field.label)"
    if ($key -eq "notes") {
        $box.Multiline = $true
        $box.Height = 76
        $box.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
        $box.AcceptsReturn = $true
    } else {
        $box.Height = 24
    }
    $valueProperty = $payload.values.PSObject.Properties[$key]
    $value = $null
    if ($null -ne $valueProperty) { $value = $valueProperty.Value }
    if ($null -ne $value) { $box.Text = [string]$value }
    $textBoxes[$key] = $box
    $panel.Controls.Add($box)
    $y += $box.Height + 10
    $tabIndex += 1
}

$buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$buttonPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::RightToLeft
$buttonPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
$buttonPanel.Height = 46
$form.Controls.Add($buttonPanel)

$saveButton = New-Object System.Windows.Forms.Button
$saveButton.Text = "&Save"
$saveButton.Width = 90
$saveButton.TabIndex = $tabIndex
$noEditionButton = New-Object System.Windows.Forms.Button
$noEditionButton.Text = "No &Edition"
$noEditionButton.Width = 110
$noEditionButton.TabIndex = $tabIndex + 1
$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = "Cancel"
$cancelButton.Width = 90
$cancelButton.TabIndex = $tabIndex + 2

$buttonPanel.Controls.Add($cancelButton)
$buttonPanel.Controls.Add($noEditionButton)
$buttonPanel.Controls.Add($saveButton)

$noEditionButton.Add_Click({
    if ($textBoxes.ContainsKey("edition")) {
        $textBoxes["edition"].Text = "No edition"
        $textBoxes["edition"].Focus() | Out-Null
        $textBoxes["edition"].SelectAll()
    }
})

$saveButton.Add_Click({
    $titleBox = $textBoxes["title"]
    if ($payload.requireTitle -and ($null -eq $titleBox -or [string]::IsNullOrWhiteSpace($titleBox.Text))) {
        [System.Windows.Forms.MessageBox]::Show("Title is required.", "Missing title", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        if ($null -ne $titleBox) {
            $titleBox.Focus() | Out-Null
            $titleBox.SelectAll()
        }
        return
    }
    $result = [ordered]@{}
    foreach ($field in $payload.fields) {
        $key = [string]$field.key
        if ($textBoxes.ContainsKey($key)) {
            $result[$key] = $textBoxes[$key].Text.Trim()
        } else {
            $result[$key] = ""
        }
    }
    $json = $result | ConvertTo-Json -Depth 4
    [System.IO.File]::WriteAllText($OutputPath, $json, [System.Text.UTF8Encoding]::new($false))
    $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Close()
})

$cancelButton.Add_Click({
    $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Close()
})

$form.CancelButton = $cancelButton
$form.Add_Shown({
    $focusKey = [string]$payload.initialFocusKey
    if (-not $textBoxes.ContainsKey($focusKey)) {
        $focusKey = "title"
    }
    $textBoxes[$focusKey].Focus() | Out-Null
    $textBoxes[$focusKey].SelectAll()
})

[void]$form.ShowDialog()
'''

    with tempfile.TemporaryDirectory(prefix="aelm_metadata_") as temp_folder:
        temp = Path(temp_folder)
        input_path = temp / "metadata_input.json"
        output_path = temp / "metadata_output.json"
        script_path = temp / "metadata_editor.ps1"
        input_path.write_text(json.dumps(payload), encoding="utf-8")
        script_path.write_text(script, encoding="utf-8")

        completed = subprocess.run(
            [
                "powershell",
                "-WindowStyle",
                "Hidden",
                "-STA",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(script_path),
                str(input_path),
                str(output_path),
                heading,
            ],
            capture_output=True,
            text=True,
            creationflags=WINDOWS_NO_CONSOLE_FLAGS,
        )
        if completed.returncode != 0:
            raise RuntimeError(completed.stderr.strip() or completed.stdout.strip() or "Windows metadata editor failed.")
        if not output_path.exists():
            return None
        return json.loads(output_path.read_text(encoding="utf-8"))


class AccessibleMetadataFormDialog:
    """Screen-reader-oriented metadata editor.

    One persistent dialog shows one field at a time and provides Previous,
    Next, Save, and Read Current Value commands.
    """

    FIELDS = [
        ("title", "Title", "Required."),
        ("author", "Author", ""),
        ("series", "Series", "Series or collection name."),
        ("series_number", "Series number", "For example 1, 2, 3, or 2.5."),
        ("edition", "Edition", "For example third edition, revised edition, or No edition."),
        ("year", "Year", "Publication year."),
        ("isbn", "ISBN", ""),
        ("publisher", "Publisher", ""),
        ("source", "Source", "For example Bookshare, Kindle, Personal, or Web."),
        ("tags", "Tags", "Separate tags with commas."),
        ("notes", "Notes", "Optional."),
    ]

    def __init__(self, parent, heading="Book Metadata", initial=None, initial_focus_key="title"):
        self.result = None
        self.parent_app = getattr(parent, "_library_app", None)
        self.index = self.field_index(initial_focus_key)
        self.values = {key: str((initial or {}).get(key, "") or "") for key, _label, _hint in self.FIELDS}
        self.value_var = StringVar()
        self.prompt_var = StringVar()
        self.read_value_var = StringVar()

        self.window = tk.Toplevel(parent)
        self.window.title(heading)
        self.window.transient(parent)
        self.window.grab_set()

        main = ttk.Frame(self.window, padding=12)
        main.pack(fill=BOTH, expand=True)

        ttk.Label(main, textvariable=self.prompt_var, wraplength=760).pack(anchor="w", pady=(0, 8))

        self.read_entry = tk.Entry(main, textvariable=self.read_value_var, width=90, takefocus=1, state="readonly")
        self.read_entry.pack(fill=X, expand=True, pady=(0, 8))

        self.entry = tk.Entry(main, textvariable=self.value_var, width=90, takefocus=1)
        self.entry.pack(fill=X, expand=True, pady=(0, 8))

        command_text = (
            "Tab moves to the next metadata field. Shift+Tab moves to the previous metadata field. "
            "Alt+R reads the current value. Alt+S saves. Escape cancels."
        )
        ttk.Label(main, text=command_text, wraplength=760).pack(anchor="w", pady=(0, 8))

        button_frame = ttk.Frame(main)
        button_frame.pack(anchor="e")

        ttk.Button(button_frame, text="Previous", command=self.previous_field).pack(side=LEFT, padx=(0, 8))
        ttk.Button(button_frame, text="Read Current Value", command=self.read_current_value).pack(side=LEFT, padx=(0, 8))
        ttk.Button(button_frame, text="Edit Current Value", command=self.focus_entry).pack(side=LEFT, padx=(0, 8))
        ttk.Button(button_frame, text="Mark No Edition", command=self.mark_no_edition).pack(side=LEFT, padx=(0, 8))
        ttk.Button(button_frame, text="Next", command=self.next_field).pack(side=LEFT, padx=(0, 8))
        ttk.Button(button_frame, text="Save", command=self.save).pack(side=LEFT, padx=(0, 8))
        ttk.Button(button_frame, text="Cancel", command=self.cancel).pack(side=LEFT)

        self.entry.bind("<Return>", lambda event: self.next_field())
        self.entry.bind("<Tab>", lambda event: self.next_field())
        self.entry.bind("<Shift-Tab>", lambda event: self.previous_field())
        self.read_entry.bind("<Tab>", lambda event: self.next_field())
        self.read_entry.bind("<Shift-Tab>", lambda event: self.previous_field())
        self.window.bind("<Alt-r>", lambda event: self.read_current_value())
        self.window.bind("<Alt-R>", lambda event: self.read_current_value())
        self.window.bind("<Alt-s>", lambda event: self.save())
        self.window.bind("<Alt-S>", lambda event: self.save())
        self.window.bind("<Escape>", lambda event: self.cancel())
        self.window.protocol("WM_DELETE_WINDOW", self.cancel)

        self.load_field(self.index)
        self.window.after(50, self.focus_entry)
        self.window.after(250, self.focus_entry)

    def field_index(self, key):
        for index, (field_key, _label, _hint) in enumerate(self.FIELDS):
            if field_key == key:
                return index
        return 0

    def current_field(self):
        return self.FIELDS[self.index]

    def store_current_field(self):
        key, _label, _hint = self.current_field()
        self.values[key] = self.value_var.get().strip()

    def load_field(self, index):
        self.index = max(0, min(index, len(self.FIELDS) - 1))
        key, label, hint = self.current_field()
        value = self.values.get(key, "")
        spoken_value = value if value else "blank"
        hint_text = f" {hint}" if hint else ""
        self.window.title(f"Book Metadata: {label}")
        self.prompt_var.set(
            f"{label}. Field {self.index + 1} of {len(self.FIELDS)}.{hint_text} "
            f"Current value: {spoken_value}."
        )
        self.read_value_var.set(f"{label} current value: {spoken_value}")
        self.value_var.set(value)
        self.focus_entry()

    def focus_entry(self):
        self.window.lift()
        self.window.focus_force()
        self.entry.focus_force()
        self.entry.icursor(0)
        self.entry.selection_range(0, END)

    def read_current_value(self):
        key, label, _hint = self.current_field()
        value = self.value_var.get().strip()
        if not value:
            value = "blank"
        messagebox.showinfo(f"{label} current value", f"{label}: {value}")
        self.window.after(50, self.focus_entry)
        return "break"

    def mark_no_edition(self):
        self.store_current_field()
        edition_index = self.field_index("edition")
        self.values["edition"] = NO_EDITION_VALUE
        self.load_field(edition_index)
        if self.parent_app:
            self.parent_app.status_var.set("Edition marked as No edition.")
        return "break"

    def next_field(self):
        self.store_current_field()
        if self.index >= len(self.FIELDS) - 1:
            return self.save()
        self.load_field(self.index + 1)
        return "break"

    def previous_field(self):
        self.store_current_field()
        self.load_field(self.index - 1)
        return "break"

    def save(self):
        self.store_current_field()
        if not self.values.get("title", "").strip():
            messagebox.showerror("Missing title", "Title is required.")
            self.load_field(0)
            return "break"
        self.result = dict(self.values)
        self.window.destroy()
        return "break"

    def cancel(self):
        if messagebox.askyesno("Cancel metadata editing", "Cancel metadata editing and discard changes?"):
            self.result = None
            self.window.destroy()
        else:
            self.window.after(50, self.focus_entry)
        return "break"

    @staticmethod
    def ask(parent, heading="Book Metadata", initial=None, initial_focus_key="title"):
        if sys.platform.startswith("win") and shutil.which("powershell"):
            try:
                return ask_metadata_with_windows_forms(
                    heading,
                    AccessibleMetadataFormDialog.FIELDS,
                    initial or {},
                    initial_focus_key=initial_focus_key,
                )
            except Exception:
                pass
        dialog = AccessibleMetadataFormDialog(parent, heading, initial, initial_focus_key=initial_focus_key)
        parent.wait_window(dialog.window)
        return dialog.result


class TkMetadataDialog:
    """Compatibility wrapper for the metadata editor."""

    @staticmethod
    def ask(parent, heading="Book Metadata", initial=None, initial_focus_key="title"):
        return AccessibleMetadataFormDialog.ask(parent, heading, initial, initial_focus_key=initial_focus_key)


class BulkMetadataDialog:
    FIELDS = [
        ("author", "Author", "Leave blank to keep each book's current author."),
        ("series", "Series", "Leave blank to keep each book's current series."),
        ("edition", "Edition", "Leave blank to keep each book's current edition."),
        ("year", "Year", "Leave blank to keep each book's current publication year."),
        ("publisher", "Publisher", "Leave blank to keep each book's current publisher."),
        ("source", "Source", "Leave blank to keep each book's current source."),
        ("tags", "Tags", "Leave blank to keep each book's current tags."),
        ("notes", "Notes", "Leave blank to keep each book's current notes."),
    ]

    @staticmethod
    def ask(parent):
        if sys.platform.startswith("win") and shutil.which("powershell"):
            try:
                return ask_metadata_with_windows_forms(
                    "Bulk Edit Metadata",
                    BulkMetadataDialog.FIELDS,
                    {},
                    initial_focus_key="author",
                    require_title=False,
                )
            except Exception:
                pass
        messagebox.showerror(
            "Bulk editor unavailable",
            "The accessible bulk metadata editor requires Windows PowerShell forms."
        )
        return None


class LibraryApp:
    def __init__(self, root):
        self.root = root
        self.root._library_app = self
        self.db = LibraryDatabase()
        self.root.title(APP_NAME)
        self.root.geometry("1000x600")

        self.search_var = StringVar()
        self.status_var = StringVar(value="Ready")
        self.shortcut_readout_var = StringVar()
        self.book_list_ids = []
        self.book_list_titles = []
        self.book_list_announcement_texts = []
        self.multi_book_lists = []
        self.multi_book_list_frames = []
        self.multi_book_list_index_maps = []
        self.temporary_column_announcement_item = None
        self.active_multi_book_column = 0
        self.marked_book_ids = set()
        self.last_missing_metadata_sound_book_id = None
        self.nvda_controller = None
        self.nvda_controller_checked = False
        self.last_nvda_announcement = ""
        self.suppress_next_book_list_select_announcement = False
        self.last_alt_number_key = ""
        self.last_alt_number_time = 0.0
        self.shortcut_readout_return_after = None
        self.cover_image_cache = {}
        self.cover_preview_photo = None
        self.cover_preview_book_id = None
        self.cover_preview_after = None
        self.sort_by = "title"
        self.filter_source = ""
        self.filter_tag = ""
        self.filter_format = ""
        self.active_saved_view = ""
        self.backup_check_after = None
        self.watched_scan_after = None
        self.duplicate_cleanup_after = None
        self.watched_scan_running = False
        self.watched_scan_queue = queue.Queue()
        self.folder_import_running = False
        self.folder_discovery_running = False
        self._index_queue = queue.Queue()
        self._start_content_indexer()

        self.build_menu()
        self.build_ui()
        self.refresh_books()
        self.schedule_backup_check(5000)
        self.schedule_watched_folder_scan(15000)
        self.schedule_automatic_duplicate_cleanup(10000)
        self.root.after(750, self.import_command_line_files)

    def build_menu(self):
        menu_bar = Menu(self.root)

        file_menu = Menu(menu_bar, tearoff=False)
        file_menu.add_command(label="Open Book\tCtrl+O", command=self.open_book)
        file_menu.add_command(label="Add Book...\tCtrl+N", command=self.add_book)
        file_menu.add_command(label="Import Folder...\tCtrl+Shift+N", command=self.import_folder)
        file_menu.add_command(label="Export Book Copy...\tCtrl+E", command=self.export_book)
        send_to_menu = Menu(file_menu, tearoff=False)
        send_to_menu.add_command(label="Voice Dream...\tCtrl+Shift+V", command=self.send_to_voice_dream)
        send_to_menu.add_command(label="Dolphin EasyReader...", command=self.send_to_dolphin_easyreader)
        send_to_menu.add_command(label="Kindle...\tCtrl+Shift+K", command=self.send_to_kindle)
        send_to_menu.add_command(label="NLS eReader...\tCtrl+Shift+E", command=self.send_to_nls_ereader)
        send_to_menu.add_command(label="HumanWare Braille eReader (MTP)...", command=self.send_to_humanware_mtp)
        file_menu.add_cascade(label="Send To", menu=send_to_menu)
        file_menu.add_separator()
        file_menu.add_command(label="Exit\tAlt+F4", command=self.root.quit)
        menu_bar.add_cascade(label="File", menu=file_menu, underline=0)

        book_menu = Menu(menu_bar, tearoff=False)
        book_menu.add_command(label="Edit Metadata...\tF2", command=self.edit_book)
        book_menu.add_command(label="Bulk Edit Selected Metadata...", command=self.bulk_edit_selected_metadata)
        book_menu.add_command(label="Auto-Detect Metadata...\tCtrl+D", command=self.auto_detect_selected_metadata)
        repair_menu = Menu(book_menu, tearoff=False)
        repair_menu.add_command(label="Cleanup and Repair Wizard...", command=self.cleanup_repair_wizard)
        repair_menu.add_separator()
        repair_menu.add_command(label="Check EPUB Accessibility Metadata...", command=self.check_selected_epub_accessibility)
        repair_menu.add_command(label="Add EPUB Page Breaks from Page Labels...", command=self.add_page_breaks_to_selected_epub)
        repair_menu.add_command(label="Rebuild EPUB Table of Contents from Text TOC...", command=self.rebuild_selected_epub_toc_from_text)
        repair_menu.add_command(label="Clean Repaired EPUB Text...", command=self.clean_selected_epub_text)
        repair_menu.add_command(label="Combine Selected Books into One EPUB...", command=self.combine_selected_books_to_epub)
        repair_menu.add_separator()
        repair_menu.add_command(label="Remove Repair Backup Files...", command=self.remove_repair_backup_files)
        book_menu.add_cascade(label="Repair", menu=repair_menu)
        book_menu.add_command(label="Look Up Metadata Online...", command=self.lookup_selected_metadata_online)
        book_menu.add_command(label="View Cover Image...", command=self.view_selected_cover_image)
        book_menu.add_command(label="Convert to EPUB...\tCtrl+R", command=self.convert_selected_to_epub)
        book_menu.add_command(label="Show Book Details\tCtrl+I", command=self.show_selected_book_info)
        book_menu.add_command(label="Focus Books List\tCtrl+L", command=self.focus_books_list)
        book_menu.add_command(label="Select All Books\tCtrl+A", command=self.select_all_books)
        book_menu.add_command(label="Deselect All Books\tCtrl+Shift+A", command=self.deselect_all_books)
        book_menu.add_command(label="Delete from Library\tDelete", command=self.delete_book)
        menu_bar.add_cascade(label="Book", menu=book_menu, underline=0)

        organize_menu = Menu(menu_bar, tearoff=False)
        sort_menu = Menu(organize_menu, tearoff=False)
        sort_menu.add_command(label="Title A to Z", command=lambda: self.set_sort("title"))
        sort_menu.add_command(label="Title Z to A", command=lambda: self.set_sort("title_desc"))
        sort_menu.add_separator()
        sort_menu.add_command(label="Author A to Z", command=lambda: self.set_sort("author"))
        sort_menu.add_command(label="Author Z to A", command=lambda: self.set_sort("author_desc"))
        sort_menu.add_separator()
        sort_menu.add_command(label="Series A to Z", command=lambda: self.set_sort("series"))
        sort_menu.add_command(label="Series Z to A", command=lambda: self.set_sort("series_desc"))
        sort_menu.add_separator()
        sort_menu.add_command(label="Published Year Newest to Oldest", command=lambda: self.set_sort("date"))
        sort_menu.add_command(label="Published Year Oldest to Newest", command=lambda: self.set_sort("date_oldest"))
        sort_menu.add_separator()
        sort_menu.add_command(label="Date Added Newest to Oldest", command=lambda: self.set_sort("date_added"))
        sort_menu.add_command(label="Date Added Oldest to Newest", command=lambda: self.set_sort("date_added_oldest"))
        organize_menu.add_cascade(label="Sort", menu=sort_menu)
        organize_menu.add_separator()
        filter_menu = Menu(organize_menu, tearoff=False)
        filter_menu.add_command(label="By Source...", command=self.set_source_filter)
        filter_menu.add_command(label="By Tag...", command=self.set_tag_filter)
        filter_menu.add_command(label="By Format...", command=self.set_format_filter)
        filter_menu.add_command(label="Clear Filters", command=self.clear_filters)
        organize_menu.add_cascade(label="Filter", menu=filter_menu)
        organize_menu.add_separator()
        saved_views_menu = Menu(organize_menu, tearoff=False)
        saved_views_menu.add_command(label="Missing Metadata", command=lambda: self.apply_builtin_saved_view("missing_metadata"))
        saved_views_menu.add_command(label="Recently Added", command=lambda: self.apply_builtin_saved_view("recently_added"))
        saved_views_menu.add_command(label="EPUB Books", command=lambda: self.apply_builtin_saved_view("epub"))
        saved_views_menu.add_command(label="Bookshare Books", command=lambda: self.apply_builtin_saved_view("bookshare"))
        saved_views_menu.add_separator()
        saved_views_menu.add_command(label="Save Current View...", command=self.save_current_view)
        saved_views_menu.add_command(label="Choose Saved View...", command=self.choose_saved_view)
        saved_views_menu.add_command(label="Clear Saved View", command=self.clear_saved_view)
        organize_menu.add_cascade(label="Saved Views", menu=saved_views_menu)
        organize_menu.add_separator()
        organize_menu.add_command(label="Remove Duplicate Books, Prefer EPUB...", command=self.remove_duplicates_prefer_epub)
        organize_menu.add_command(label="Toggle Automatic Duplicate Cleanup", command=self.toggle_automatic_duplicate_cleanup)
        organize_menu.add_command(label="Show Sort and Filter Settings", command=self.show_organize_settings)
        menu_bar.add_cascade(label="Organize", menu=organize_menu, underline=0)

        search_menu = Menu(menu_bar, tearoff=False)
        search_menu.add_command(label="Search Library...\tCtrl+F", command=self.focus_search)
        search_menu.add_command(label="Focus Books List\tCtrl+L", command=self.focus_books_list)
        search_menu.add_command(label="Clear Search", command=self.clear_search)
        search_menu.add_command(label="Explain Search", command=self.explain_search)
        search_menu.add_separator()
        search_menu.add_command(label="Re-index Book Content...", command=self.reindex_library_content)
        menu_bar.add_cascade(label="Search", menu=search_menu, underline=0)

        settings_menu = Menu(menu_bar, tearoff=False)
        speech_menu = Menu(settings_menu, tearoff=False)
        speech_menu.add_command(label="Title Only", command=lambda: self.set_book_list_speech_fields(["title"]))
        speech_menu.add_command(label="Title and Author", command=lambda: self.set_book_list_speech_fields(["title", "author"]))
        speech_menu.add_command(label="Title, Author, and Edition", command=lambda: self.set_book_list_speech_fields(["title", "author", "edition"]))
        speech_menu.add_command(label="All Details", command=lambda: self.set_book_list_speech_fields([key for key, _label in BOOK_LIST_SPEECH_FIELDS]))
        speech_menu.add_command(label="Show Current Book List Reading", command=self.show_book_list_speech_fields)
        settings_menu.add_cascade(label="Book List Reading", menu=speech_menu)
        settings_menu.add_separator()
        display_menu = Menu(settings_menu, tearoff=False)
        display_menu.add_command(label="Single List", command=lambda: self.set_book_list_display_mode("single"))
        display_menu.add_command(label="Multiple Alphabetical Lists", command=lambda: self.set_book_list_display_mode("multiple"))
        display_menu.add_separator()
        display_menu.add_command(label="Show Current Display", command=self.show_book_list_display_mode)
        settings_menu.add_cascade(label="Book List Display", menu=display_menu)
        settings_menu.add_separator()
        missing_metadata_menu = Menu(settings_menu, tearoff=False)
        missing_metadata_menu.add_command(label="Off", command=lambda: self.set_missing_metadata_sound_mode("off"))
        for mode, (label, _fields) in MISSING_METADATA_SOUND_MODES.items():
            missing_metadata_menu.add_command(label=label, command=lambda mode=mode: self.set_missing_metadata_sound_mode(mode))
        missing_metadata_menu.add_separator()
        missing_metadata_menu.add_command(label="Show Current Setting", command=self.show_missing_metadata_sound_mode)
        missing_metadata_menu.add_command(label="Test Sound", command=self.test_missing_metadata_sound)
        settings_menu.add_cascade(label="Missing Metadata Alert Sound", menu=missing_metadata_menu)
        settings_menu.add_separator()
        backup_menu = Menu(settings_menu, tearoff=False)
        backup_menu.add_command(label="Use OneDrive Folder...", command=lambda: self.choose_cloud_backup_folder("onedrive"))
        backup_menu.add_command(label="Use Google Drive Folder...", command=lambda: self.choose_cloud_backup_folder("google_drive"))
        backup_menu.add_command(label="Use iCloud Drive Folder...", command=lambda: self.choose_cloud_backup_folder("icloud"))
        backup_menu.add_command(label="Choose Another Backup Folder...", command=lambda: self.choose_cloud_backup_folder("other"))
        backup_menu.add_separator()
        schedule_menu = Menu(backup_menu, tearoff=False)
        schedule_menu.add_command(label="On Demand", command=lambda: self.set_backup_schedule("on_demand"))
        schedule_menu.add_command(label="Daily", command=lambda: self.set_backup_schedule("daily"))
        schedule_menu.add_command(label="Weekly", command=lambda: self.set_backup_schedule("weekly"))
        schedule_menu.add_command(label="Monthly", command=lambda: self.set_backup_schedule("monthly"))
        backup_menu.add_cascade(label="Backup Schedule", menu=schedule_menu)
        backup_menu.add_separator()
        backup_menu.add_command(label="Back Up Now", command=self.backup_library_now)
        backup_menu.add_command(label="Restore From Backup...", command=self.restore_library_backup)
        backup_menu.add_command(label="Show Backup Status", command=self.show_backup_status)
        settings_menu.add_cascade(label="Library Backup", menu=backup_menu)
        settings_menu.add_separator()
        watch_menu = Menu(settings_menu, tearoff=False)
        watch_menu.add_command(label="Add Watched Folder...", command=self.add_watched_folder)
        watch_menu.add_command(label="Remove Watched Folder...", command=self.remove_watched_folder)
        watch_menu.add_command(label="Scan Watched Folders Now", command=self.scan_watched_folders_now)
        watch_menu.add_command(label="Toggle Automatic Watched Folder Scanning", command=self.toggle_watched_folder_auto_scan)
        watch_menu.add_command(label="Show Watched Folder Status", command=self.show_watched_folder_status)
        settings_menu.add_cascade(label="Watched Folders", menu=watch_menu)
        settings_menu.add_separator()
        explorer_menu = Menu(settings_menu, tearoff=False)
        explorer_menu.add_command(label="Add File Explorer Right-Click Command", command=self.install_file_explorer_context_menu)
        explorer_menu.add_command(label="Remove File Explorer Right-Click Command", command=self.remove_file_explorer_context_menu)
        explorer_menu.add_command(label="Show File Explorer Integration Status", command=self.show_file_explorer_context_menu_status)
        settings_menu.add_cascade(label="File Explorer Integration", menu=explorer_menu)
        settings_menu.add_separator()
        settings_menu.add_command(label="Set Voice Dream Loader Folder...", command=self.choose_voice_dream_folder)
        settings_menu.add_command(label="Set Dolphin EasyReader Folder...", command=self.choose_dolphin_easyreader_folder)
        settings_menu.add_command(label="Set NLS eReader Folder...", command=self.choose_nls_ereader_folder)
        settings_menu.add_command(label="Set Kindle Email Addresses...", command=self.set_kindle_email)
        settings_menu.add_separator()
        calibre_menu = Menu(settings_menu, tearoff=False)
        calibre_menu.add_command(label="Show Calibre Tools Status", command=self.show_calibre_tools_status)
        calibre_menu.add_command(label="Toggle Calibre Metadata Reading", command=self.toggle_calibre_metadata_reading)
        settings_menu.add_cascade(label="Kindle and Calibre", menu=calibre_menu)
        settings_menu.add_separator()
        settings_menu.add_command(label="Set Default Ebook Reader...", command=self.choose_default_reader)
        settings_menu.add_command(label="Use System Default Reader", command=self.clear_default_reader)
        settings_menu.add_command(label="Open Library Folder", command=self.open_library_folder)
        menu_bar.add_cascade(label="Settings", menu=settings_menu, underline=0)

        help_menu = Menu(menu_bar, tearoff=False)
        help_menu.add_command(label="Help\tF1", command=self.show_help)
        menu_bar.add_cascade(label="Help", menu=help_menu, underline=0)

        self.root.config(menu=menu_bar)

    def build_ui(self):
        main = ttk.Frame(self.root, padding=10)
        main.pack(fill=BOTH, expand=True)

        search_frame = ttk.Frame(main)
        search_frame.pack(fill=X, pady=(0, 8))

        ttk.Label(search_frame, text="Search library metadata").pack(side=LEFT, padx=(0, 8))
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_var)
        self.search_entry.pack(side=LEFT, fill=X, expand=True, padx=(0, 8))
        self.search_entry.bind("<Return>", lambda event: self.search_and_focus())
        ttk.Button(search_frame, text="Search", command=self.search_and_focus).pack(side=LEFT)

        list_frame = ttk.Frame(main)
        list_frame.pack(fill=BOTH, expand=True)

        ttk.Label(
            list_frame,
            text="Books list. Use Settings to choose which details are read while navigating."
        ).pack(anchor="w")

        cover_frame = ttk.Frame(list_frame, padding=(12, 0, 0, 0))
        cover_frame.pack(side=RIGHT, fill=Y)
        ttk.Label(cover_frame, text="Cover preview").pack(anchor="w")
        self.cover_preview_label = ttk.Label(
            cover_frame,
            text="No cover",
            anchor="center",
            relief="solid",
            width=22,
        )
        self.cover_preview_label.pack(fill=X, pady=(4, 6), ipady=48)
        self.cover_preview_title_var = StringVar(value="")
        ttk.Label(
            cover_frame,
            textvariable=self.cover_preview_title_var,
            wraplength=170,
            justify="center",
        ).pack(fill=X)
        ttk.Button(cover_frame, text="View Cover", command=self.view_selected_cover_image).pack(fill=X, pady=(8, 0))
        ttk.Button(cover_frame, text="Find Cover", command=self.find_cover_for_selected_book).pack(fill=X, pady=(4, 0))

        self.book_display_area = ttk.Frame(list_frame)
        self.book_display_area.pack(side=LEFT, fill=BOTH, expand=True)

        self.single_book_list_frame = ttk.Frame(self.book_display_area)
        self.multi_book_list_frame = ttk.Frame(self.book_display_area)

        self.book_list = Listbox(self.single_book_list_frame, selectmode="browse", height=20, exportselection=False, takefocus=1)
        scrollbar = Scrollbar(self.single_book_list_frame, orient="vertical", command=self.book_list.yview)
        self.book_list.configure(yscrollcommand=scrollbar.set)
        self.book_list.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=LEFT, fill=Y)
        self.bind_book_list_widget(self.book_list)

        for group_label, _letters in BOOK_LIST_ALPHABET_GROUPS:
            column = ttk.Frame(self.multi_book_list_frame, padding=(0, 0, 8, 0))
            ttk.Label(column, text=group_label).pack(anchor="w")
            listbox = Listbox(
                column,
                selectmode="browse",
                height=20,
                exportselection=False,
                takefocus=1,
                name=f"{group_label.lower().replace('-', '_')}_book_list",
            )
            column_scrollbar = Scrollbar(column, orient="vertical", command=listbox.yview)
            listbox.configure(yscrollcommand=column_scrollbar.set)
            listbox.pack(side=LEFT, fill=BOTH, expand=True)
            column_scrollbar.pack(side=LEFT, fill=Y)
            self.bind_book_list_widget(listbox, multiple_column=True)
            self.multi_book_list_frames.append(column)
            self.multi_book_lists.append(listbox)
            self.multi_book_list_index_maps.append([])

        self.apply_book_list_display_mode(initial=True)

        self.root.bind("<Control-n>", lambda event: self.add_book())
        self.root.bind("<Control-N>", lambda event: self.import_folder())
        self.root.bind("<Control-o>", lambda event: self.open_book())
        self.root.bind("<Control-e>", lambda event: self.export_book())
        self.root.bind("<Control-V>", lambda event: self.send_to_voice_dream())
        self.root.bind("<Control-k>", lambda event: self.open_kindle())
        self.root.bind("<Control-r>", lambda event: self.convert_selected_to_epub())
        self.root.bind("<Control-K>", lambda event: self.send_to_kindle())
        self.root.bind("<Control-E>", lambda event: self.send_to_nls_ereader())
        self.root.bind("<Control-f>", self.focus_search_from_keyboard)
        self.root.bind("<Control-i>", lambda event: self.show_selected_book_info())
        self.root.bind("<Control-d>", lambda event: self.auto_detect_selected_metadata())
        self.root.bind("<Control-l>", lambda event: self.focus_books_list())
        self.root.bind("<Control-a>", self.select_all_books)
        self.root.bind("<Control-A>", lambda event: self.deselect_all_books())
        self.root.bind("<Escape>", self.clear_search_from_keyboard)
        self.root.bind("<F1>", lambda event: self.show_help())

        self.shortcut_readout = tk.Entry(
            main,
            textvariable=self.shortcut_readout_var,
            state="readonly",
            takefocus=1,
        )
        self.shortcut_readout.pack(fill=X, pady=(8, 0))
        self.bind_alt_number_shortcuts(self.shortcut_readout)
        self.shortcut_readout.bind("<Up>", self.return_to_book_list_from_readout)
        self.shortcut_readout.bind("<Down>", self.return_to_book_list_from_readout)
        self.shortcut_readout.bind("<Prior>", self.return_to_book_list_from_readout)
        self.shortcut_readout.bind("<Next>", self.return_to_book_list_from_readout)
        self.shortcut_readout.bind("<Home>", self.return_to_book_list_from_readout)
        self.shortcut_readout.bind("<End>", self.return_to_book_list_from_readout)
        self.shortcut_readout.bind("<Escape>", self.return_to_book_list_from_readout)
        status = ttk.Label(main, textvariable=self.status_var, relief="sunken", anchor="w")
        status.pack(fill=X, pady=(8, 0))

        self.focus_books_list()

    def bind_book_list_widget(self, widget, multiple_column=False):
        widget.bind("<Return>", lambda event: self.open_book())
        widget.bind("<Alt-Return>", lambda event: self.open_book())
        widget.bind("<Delete>", lambda event: self.delete_book())
        widget.bind("<F2>", lambda event: self.edit_book())
        widget.bind("<Control-e>", lambda event: self.export_book())
        widget.bind("<Control-V>", lambda event: self.send_to_voice_dream())
        widget.bind("<Control-o>", lambda event: self.open_book())
        widget.bind("<Control-k>", lambda event: self.open_kindle())
        widget.bind("<Control-r>", lambda event: self.convert_selected_to_epub())
        widget.bind("<Control-K>", lambda event: self.send_to_kindle())
        widget.bind("<Control-E>", lambda event: self.send_to_nls_ereader())
        widget.bind("<Control-n>", lambda event: self.add_book())
        widget.bind("<Control-f>", self.focus_search_from_keyboard)
        widget.bind("<Control-i>", lambda event: self.show_selected_book_info())
        widget.bind("<Control-d>", lambda event: self.auto_detect_selected_metadata())
        widget.bind("<Control-a>", self.select_all_books)
        widget.bind("<Control-A>", lambda event: self.deselect_all_books())
        widget.bind("<Control-space>", self.toggle_mark_current_book)
        widget.bind("<Escape>", self.clear_search_from_keyboard)
        widget.bind("<Home>", self.move_to_first_book)
        widget.bind("<End>", self.move_to_last_book)
        if multiple_column:
            widget.bind("<Left>", lambda event: self.move_between_book_columns(event, -1))
            widget.bind("<Right>", lambda event: self.move_between_book_columns(event, 1))
        self.bind_alt_number_shortcuts(widget)
        widget.bind("<<ListboxSelect>>", self.on_book_list_select)
        widget.bind("<KeyPress>", self.on_book_list_keypress)

    def bind_alt_number_shortcuts(self, widget):
        for digit in "1234567890":
            widget.bind(f"<Alt-KeyPress-{digit}>", self.on_book_list_alt_number)

    def get_book_list_display_mode(self):
        mode = self.db.get_setting("book_list_display_mode", DEFAULT_BOOK_LIST_DISPLAY_MODE)
        if mode not in BOOK_LIST_DISPLAY_MODES:
            return DEFAULT_BOOK_LIST_DISPLAY_MODE
        return mode

    def book_list_display_mode_label(self):
        return BOOK_LIST_DISPLAY_MODES[self.get_book_list_display_mode()]

    def set_book_list_display_mode(self, mode):
        if mode not in BOOK_LIST_DISPLAY_MODES:
            mode = DEFAULT_BOOK_LIST_DISPLAY_MODE
        current_book_id = self.selected_book_id_quiet()
        self.db.set_setting("book_list_display_mode", mode)
        self.apply_book_list_display_mode()
        self.refresh_books(selected_book_id=current_book_id)
        self.status_var.set(f"Book list display set to {BOOK_LIST_DISPLAY_MODES[mode]}.")

    def show_book_list_display_mode(self):
        messagebox.showinfo(
            "Current Book List Display",
            f"The book list is currently shown as:\n\n{self.book_list_display_mode_label()}"
        )

    def apply_book_list_display_mode(self, initial=False):
        if not hasattr(self, "single_book_list_frame") or not hasattr(self, "multi_book_list_frame"):
            return
        self.single_book_list_frame.pack_forget()
        self.multi_book_list_frame.pack_forget()
        if self.get_book_list_display_mode() == "multiple":
            self.multi_book_list_frame.pack(fill=BOTH, expand=True)
            self.show_active_multi_book_column()
        else:
            self.single_book_list_frame.pack(fill=BOTH, expand=True)
        if not initial:
            self.focus_books_list()

    def is_multiple_book_list_mode(self):
        return self.get_book_list_display_mode() == "multiple"

    def show_active_multi_book_column(self):
        if not hasattr(self, "multi_book_list_frames"):
            return
        for frame in self.multi_book_list_frames:
            frame.pack_forget()
        if not self.multi_book_list_frames:
            return
        if (
            self.multi_book_lists
            and 0 <= self.active_multi_book_column < len(self.multi_book_lists)
            and self.multi_book_lists[self.active_multi_book_column].size() == 0
        ):
            for index, listbox in enumerate(self.multi_book_lists):
                if listbox.size() > 0:
                    self.active_multi_book_column = index
                    break
        self.active_multi_book_column = max(
            0,
            min(self.active_multi_book_column, len(self.multi_book_list_frames) - 1),
        )
        self.multi_book_list_frames[self.active_multi_book_column].pack(fill=BOTH, expand=True)

    def sort_label(self):
        labels = {
            "title": "Title A to Z",
            "title_desc": "Title Z to A",
            "author": "Author A to Z",
            "author_desc": "Author Z to A",
            "series": "Series A to Z",
            "series_desc": "Series Z to A",
            "date": "Published Year Newest to Oldest",
            "date_oldest": "Published Year Oldest to Newest",
            "date_added": "Date Added Newest to Oldest",
            "date_added_oldest": "Date Added Oldest to Newest",
        }
        return labels.get(self.sort_by, "Title A to Z")

    def active_filter_summary(self):
        parts = []
        if self.filter_source:
            parts.append(f"source contains {self.filter_source}")
        if self.filter_tag:
            parts.append(f"tag contains {self.filter_tag}")
        if self.filter_format:
            parts.append(f"format contains {self.filter_format}")
        if self.active_saved_view:
            parts.append(f"saved view {self.saved_view_label(self.active_saved_view)}")
        if not parts:
            return "No filters"
        return "; ".join(parts)

    def set_sort(self, sort_by):
        self.sort_by = sort_by
        self.refresh_books()
        self.focus_books_list()
        self.status_var.set(f"Sorted by {self.sort_label()}. {self.book_list.size()} book{'s' if self.book_list.size() != 1 else ''} shown.")

    def set_source_filter(self):
        value = AccessibleSingleFieldDialog.ask(
            self.root,
            "Source Filter",
            "Enter source text to filter by, such as Bookshare, Kindle, or Personal. Leave blank to clear the source filter.",
            self.filter_source,
        )
        if value is None:
            return
        self.filter_source = value.strip()
        self.refresh_books()
        self.focus_books_list()
        self.status_var.set(f"Source filter set. {self.book_list.size()} book{'s' if self.book_list.size() != 1 else ''} shown.")

    def set_tag_filter(self):
        value = AccessibleSingleFieldDialog.ask(
            self.root,
            "Tag Filter",
            "Enter tag text to filter by, such as textbook, fiction, or unread. Leave blank to clear the tag filter.",
            self.filter_tag,
        )
        if value is None:
            return
        self.filter_tag = value.strip()
        self.refresh_books()
        self.focus_books_list()
        self.status_var.set(f"Tag filter set. {self.book_list.size()} book{'s' if self.book_list.size() != 1 else ''} shown.")

    def set_format_filter(self):
        value = AccessibleSingleFieldDialog.ask(
            self.root,
            "Format Filter",
            "Enter format text to filter by, such as epub, pdf, or docx. Leave blank to clear the format filter.",
            self.filter_format,
        )
        if value is None:
            return
        self.filter_format = value.strip().lower()
        self.refresh_books()
        self.focus_books_list()
        self.status_var.set(f"Format filter set. {self.book_list.size()} book{'s' if self.book_list.size() != 1 else ''} shown.")

    def clear_filters(self):
        self.filter_source = ""
        self.filter_tag = ""
        self.filter_format = ""
        self.active_saved_view = ""
        self.refresh_books()
        self.focus_books_list()
        self.status_var.set("Filters cleared.")

    def saved_views(self):
        raw = self.db.get_setting("saved_views", "{}")
        try:
            data = json.loads(raw)
            return data if isinstance(data, dict) else {}
        except Exception:
            return {}

    def set_saved_views(self, views):
        self.db.set_setting("saved_views", json.dumps(views, sort_keys=True))

    def apply_builtin_saved_view(self, view_name):
        self.active_saved_view = view_name
        if view_name == "recently_added":
            self.sort_by = "date_added"
            self.filter_source = ""
            self.filter_tag = ""
            self.filter_format = ""
            self.search_var.set("")
        elif view_name == "epub":
            self.filter_format = "epub"
            self.filter_source = ""
            self.filter_tag = ""
            self.search_var.set("")
        elif view_name == "bookshare":
            self.filter_source = "Bookshare"
            self.filter_tag = ""
            self.filter_format = ""
            self.search_var.set("")
        self.refresh_books()
        self.focus_books_list()
        self.status_var.set(f"Saved view applied: {self.saved_view_label(view_name)}.")

    def saved_view_label(self, view_name):
        labels = {
            "missing_metadata": "Missing Metadata",
            "recently_added": "Recently Added",
            "epub": "EPUB Books",
            "bookshare": "Bookshare Books",
        }
        return labels.get(view_name, view_name)

    def save_current_view(self):
        name = AccessibleSingleFieldDialog.ask(
            self.root,
            "View name",
            "Type a short name for this saved view.",
            "",
            heading="Save Current View",
        )
        if not name:
            return
        views = self.saved_views()
        views[name] = {
            "query": self.search_var.get(),
            "sort_by": self.sort_by,
            "source": self.filter_source,
            "tag": self.filter_tag,
            "format": self.filter_format,
            "view": self.active_saved_view,
        }
        self.set_saved_views(views)
        self.status_var.set(f"Saved view {name}.")
        messagebox.showinfo("Saved View", f"Saved view:\n\n{name}")

    def choose_saved_view(self):
        views = self.saved_views()
        if not views:
            messagebox.showinfo("No Saved Views", "No custom saved views have been saved yet.")
            return
        names = sorted(views)
        prompt = "\n".join(f"{index + 1}. {name}" for index, name in enumerate(names))
        choice = AccessibleSingleFieldDialog.ask(
            self.root,
            "Saved view number",
            "Type the number of the saved view to open.\n\n" + prompt,
            "",
            heading="Choose Saved View",
        )
        if not choice:
            return
        try:
            selected = names[int(choice) - 1]
        except Exception:
            messagebox.showerror("Invalid choice", "Please enter a valid saved view number.")
            return
        view = views[selected]
        self.search_var.set(view.get("query", ""))
        self.sort_by = view.get("sort_by", "title")
        self.filter_source = view.get("source", "")
        self.filter_tag = view.get("tag", "")
        self.filter_format = view.get("format", "")
        self.active_saved_view = view.get("view", "")
        self.refresh_books()
        self.focus_books_list()
        self.status_var.set(f"Saved view applied: {selected}.")

    def clear_saved_view(self):
        self.active_saved_view = ""
        self.refresh_books()
        self.focus_books_list()
        self.status_var.set("Saved view cleared.")

    def row_missing_metadata_for_view(self, row):
        values = {
            "tags": row[4],
            "author": row[2],
            "edition": row[7] if len(row) > 7 else "",
            "year": row[8] if len(row) > 8 else "",
            "isbn": row[9] if len(row) > 9 else "",
            "publisher": row[10] if len(row) > 10 else "",
        }
        return any(not (value or "").strip() for value in values.values())

    def show_organize_settings(self):
        messagebox.showinfo(
            "Current Sort and Filter Settings",
            f"Sort: {self.sort_label()}\n"
            f"Filters: {self.active_filter_summary()}\n"
            f"Books shown: {self.book_list.size()}"
        )

    def duplicate_group_key(self, row):
        title_key = normalize_duplicate_key(row[1])
        author_key = normalize_duplicate_key(row[2])
        isbn_key = normalize_isbn_key(row[12] or "")
        if isbn_key:
            return ("isbn", isbn_key)
        if title_key and author_key:
            return ("title_author", title_key, author_key)
        return None

    def duplicate_rows_match(self, left, right):
        left_isbn = normalize_isbn_key(left[12] or "")
        right_isbn = normalize_isbn_key(right[12] or "")
        if left_isbn and right_isbn:
            return left_isbn == right_isbn
        if left_isbn or right_isbn:
            # One copy may have a detected ISBN while an older duplicate does not.
            pass

        left_title = left[1] or ""
        right_title = right[1] or ""
        if not title_keys_look_same(left_title, right_title, allow_richer_metadata_match=True):
            return False

        left_author = normalize_duplicate_key(left[2] or "")
        right_author = normalize_duplicate_key(right[2] or "")
        if left_author and right_author and left_author != right_author:
            return False
        return True

    def duplicate_keep_sort_key(self, row):
        book_format = (row[6] or "").casefold()
        stored_path = str(row[8] or "")
        path_suffix = Path(stored_path).suffix.lower().lstrip(".")
        is_epub = book_format == "epub" or path_suffix == "epub"
        metadata_score = sum(1 for index in [1, 2, 3, 4, 5, 10, 11, 12, 13] if row[index])
        file_exists = 1 if stored_path and Path(stored_path).exists() else 0
        return (
            0 if is_epub else 1,
            -file_exists,
            -metadata_score,
            row[9] or "",
            row[0],
        )

    def duplicate_groups(self, database=None):
        database = database or self.db
        rows = list(database.all_books_for_duplicate_check())
        parent = {row[0]: row[0] for row in rows}

        def find(book_id):
            while parent[book_id] != book_id:
                parent[book_id] = parent[parent[book_id]]
                book_id = parent[book_id]
            return book_id

        def union(left_id, right_id):
            left_root = find(left_id)
            right_root = find(right_id)
            if left_root != right_root:
                parent[right_root] = left_root

        for index, left in enumerate(rows):
            for right in rows[index + 1:]:
                if self.duplicate_rows_match(left, right):
                    union(left[0], right[0])

        grouped = {}
        for row in rows:
            grouped.setdefault(find(row[0]), []).append(row)
        return [group for group in grouped.values() if len(group) > 1]

    def duplicate_removal_plan(self, database=None):
        plan = []
        for rows in self.duplicate_groups(database):
            ranked = sorted(rows, key=self.duplicate_keep_sort_key)
            keep = ranked[0]
            remove = ranked[1:]
            if remove:
                plan.append((keep, remove))
        return plan

    def summarize_duplicate_plan(self, plan, limit=12):
        lines = []
        total_remove = sum(len(remove) for _keep, remove in plan)
        lines.append(f"Duplicate groups found: {len(plan)}")
        lines.append(f"Books that would be removed from the library list: {total_remove}")
        lines.append("")
        for keep, remove in plan[:limit]:
            lines.append(f"Keep: {keep[1]} by {keep[2] or 'Unknown'} ({keep[6] or Path(keep[8]).suffix.lstrip('.') or 'unknown'})")
            for row in remove:
                lines.append(f"Remove: {row[1]} by {row[2] or 'Unknown'} ({row[6] or Path(row[8]).suffix.lstrip('.') or 'unknown'})")
            lines.append("")
        if len(plan) > limit:
            lines.append(f"And {len(plan) - limit} more duplicate group{'s' if len(plan) - limit != 1 else ''}.")
            lines.append("")
        lines.append("EPUB versions are preferred when available. This first confirmation removes duplicates from the app's library list only.")
        return "\n".join(lines)

    def remove_duplicates_prefer_epub(self):
        plan = self.duplicate_removal_plan()
        if not plan:
            messagebox.showinfo(
                "No duplicates found",
                "No duplicate books were found using ISBN, title, and compatible author information."
            )
            self.status_var.set("No duplicates found.")
            return

        summary = self.summarize_duplicate_plan(plan)
        if not messagebox.askyesno("Remove duplicate books", summary + "\n\nContinue?"):
            return

        delete_files = messagebox.askyesno(
            "Delete duplicate files too?",
            "Remove the duplicate book files from disk too?\n\n"
            "Choose No to remove duplicates from the app's library list only. This is safer."
        )

        removed = 0
        errors = []
        keep_id = plan[0][0][0] if plan else None
        for _keep, remove_rows in plan:
            for row in remove_rows:
                try:
                    self.db.delete_book(row[0], delete_file=delete_files)
                    removed += 1
                except Exception as exc:
                    errors.append(f"{row[1]} -- {exc}")

        self.refresh_books(selected_book_id=keep_id)
        self.focus_books_list()
        if errors:
            report_path = self.write_import_report(removed, errors)
            messagebox.showwarning(
                "Duplicates removed with warnings",
                f"Removed {removed} duplicate book{'s' if removed != 1 else ''}. "
                f"{len(errors)} item{'s' if len(errors) != 1 else ''} could not be removed.\n\n"
                f"Report saved at:\n{report_path}"
            )
        else:
            messagebox.showinfo(
                "Duplicates removed",
                f"Removed {removed} duplicate book{'s' if removed != 1 else ''}."
            )
        self.status_var.set(f"Removed {removed} duplicate book{'s' if removed != 1 else ''}.")

    def automatic_duplicate_cleanup_enabled(self):
        return self.db.get_setting(
            "automatic_duplicate_cleanup",
            DEFAULT_AUTOMATIC_DUPLICATE_CLEANUP,
        ) == "1"

    def toggle_automatic_duplicate_cleanup(self):
        enabled = not self.automatic_duplicate_cleanup_enabled()
        if enabled and not messagebox.askyesno(
            "Enable Automatic Duplicate Cleanup",
            "Automatic cleanup removes likely duplicate entries from the library list without asking each time. "
            "It keeps the book files, but a mistaken match can hide a different book from the list.\n\nEnable it?",
        ):
            return
        self.db.set_setting("automatic_duplicate_cleanup", "1" if enabled else "0")
        state = "on" if enabled else "off"
        self.status_var.set(f"Automatic duplicate cleanup turned {state}.")
        messagebox.showinfo(
            "Automatic Duplicate Cleanup",
            f"Automatic duplicate cleanup is now {state}.\n\n"
            "When on, the app removes duplicate entries from the library list on startup and periodically. "
            "It does not delete book files from disk."
        )
        self.schedule_automatic_duplicate_cleanup(1000 if enabled else None)

    def schedule_automatic_duplicate_cleanup(self, delay_ms=None):
        if self.duplicate_cleanup_after is not None:
            try:
                self.root.after_cancel(self.duplicate_cleanup_after)
            except Exception:
                pass
            self.duplicate_cleanup_after = None
        if not self.automatic_duplicate_cleanup_enabled():
            return
        delay_ms = AUTOMATIC_DUPLICATE_CLEANUP_INTERVAL_MS if delay_ms is None else delay_ms
        self.duplicate_cleanup_after = self.root.after(delay_ms, self.automatic_duplicate_cleanup)

    def automatic_duplicate_cleanup(self):
        self.duplicate_cleanup_after = None
        if not self.automatic_duplicate_cleanup_enabled():
            return
        events = queue.Queue()

        def worker():
            database = None
            try:
                database = LibraryDatabase()
                plan = self.duplicate_removal_plan(database)
                ids = [row[0] for _keep, remove_rows in plan for row in remove_rows]
                if ids:
                    database.connection.executemany(
                        "DELETE FROM books_fts WHERE rowid = ?", ((book_id,) for book_id in ids)
                    )
                    database.connection.executemany(
                        "DELETE FROM books WHERE id = ?", ((book_id,) for book_id in ids)
                    )
                    database.connection.commit()
                events.put((len(ids), ""))
            except Exception as exc:
                logger.exception("Automatic duplicate cleanup failed")
                events.put((0, str(exc)))
            finally:
                if database is not None:
                    database.close()

        def poll():
            try:
                removed, error = events.get_nowait()
            except queue.Empty:
                self.root.after(100, poll)
                return
            if error:
                self.log_error("Automatic duplicate cleanup", RuntimeError(error))
                self.status_var.set("Automatic duplicate cleanup failed.")
            elif removed:
                self.refresh_books()
                self.status_var.set(
                    f"Automatic duplicate cleanup removed {removed} duplicate book entr{'y' if removed == 1 else 'ies'}."
                )
            self.schedule_automatic_duplicate_cleanup()

        threading.Thread(target=worker, daemon=True, name="duplicate-cleanup").start()
        self.root.after(100, poll)

    def get_book_list_speech_fields(self):
        raw = self.db.get_setting(
            "book_list_speech_fields",
            ",".join(DEFAULT_BOOK_LIST_SPEECH_FIELDS),
        )
        return normalize_book_list_speech_fields(raw)

    def book_list_speech_summary(self):
        selected = set(self.get_book_list_speech_fields())
        labels = [label for key, label in BOOK_LIST_SPEECH_FIELDS if key in selected]
        return ", ".join(labels)

    def set_book_list_speech_fields(self, selected):
        selected = normalize_book_list_speech_fields(",".join(selected))
        self.db.set_setting("book_list_speech_fields", ",".join(selected))
        current_book_id = None
        if self.book_list.size() > 0:
            index = self.current_book_index()
            if index is not None and index < len(self.book_list_ids):
                current_book_id = self.book_list_ids[index]
        self.refresh_books(selected_book_id=current_book_id)
        self.focus_books_list()
        self.status_var.set(f"Book list reading saved: {self.book_list_speech_summary()}.")

    def show_book_list_speech_fields(self):
        messagebox.showinfo(
            "Current Book List Reading",
            f"The book list currently reads:\n\n{self.book_list_speech_summary()}"
        )

    def missing_metadata_sound_enabled(self):
        legacy_enabled = self.db.get_setting("missing_metadata_sound", "0") == "1"
        return legacy_enabled or self.missing_metadata_sound_mode() != "off"

    def missing_metadata_sound_mode(self):
        mode = self.db.get_setting("missing_metadata_sound_mode", "")
        if mode in MISSING_METADATA_SOUND_MODES or mode == "off":
            return mode
        if self.db.get_setting("missing_metadata_sound", "0") == "1":
            return "complete"
        return DEFAULT_MISSING_METADATA_SOUND_MODE

    def missing_metadata_sound_mode_label(self):
        mode = self.missing_metadata_sound_mode()
        if mode == "off":
            return "Off"
        return MISSING_METADATA_SOUND_MODES.get(mode, MISSING_METADATA_SOUND_MODES[DEFAULT_MISSING_METADATA_SOUND_MODE])[0]

    def set_missing_metadata_sound_mode(self, mode):
        if mode != "off" and mode not in MISSING_METADATA_SOUND_MODES:
            mode = DEFAULT_MISSING_METADATA_SOUND_MODE
        self.db.set_setting("missing_metadata_sound_mode", mode)
        self.db.set_setting("missing_metadata_sound", "0" if mode == "off" else "1")
        self.last_missing_metadata_sound_book_id = None
        label = self.missing_metadata_sound_mode_label()
        self.status_var.set(f"Missing metadata alert sound set to {label}.")
        messagebox.showinfo("Missing Metadata Alert Sound", f"Missing metadata alert sound is now set to:\n\n{label}")
        if mode != "off":
            self.play_missing_metadata_sound()

    def show_missing_metadata_sound_mode(self):
        messagebox.showinfo(
            "Missing Metadata Alert Sound",
            f"Current setting:\n\n{self.missing_metadata_sound_mode_label()}"
        )

    def toggle_missing_metadata_sound(self):
        enabled = not self.missing_metadata_sound_enabled()
        self.set_missing_metadata_sound_mode(DEFAULT_MISSING_METADATA_SOUND_MODE if enabled else "off")
        state = "on" if enabled else "off"
        self.status_var.set(f"Missing metadata alert sound turned {state}.")

    def test_missing_metadata_sound(self):
        self.play_missing_metadata_sound()
        messagebox.showinfo(
            "Missing Metadata Alert Sound",
            "If your Windows system sounds are enabled, you should have heard the missing metadata alert sound."
        )

    def book_has_missing_metadata(self, book_id):
        row = self.db.get_book(book_id)
        if not row:
            return False
        mode = self.missing_metadata_sound_mode()
        if mode == "off":
            return False
        fields = MISSING_METADATA_SOUND_MODES.get(mode, MISSING_METADATA_SOUND_MODES[DEFAULT_MISSING_METADATA_SOUND_MODE])[1]
        values = {
            "author": row[2],
            "edition": row[10],
            "year": row[11],
            "isbn": row[12],
            "publisher": row[13],
        }
        return any(not (values.get(field) or "").strip() for field in fields)

    def play_missing_metadata_sound_if_needed(self, book_id):
        if not self.missing_metadata_sound_enabled():
            return
        if self.book_has_missing_metadata(book_id) and book_id != self.last_missing_metadata_sound_book_id:
            self.last_missing_metadata_sound_book_id = book_id
            self.play_missing_metadata_sound()
        elif not self.book_has_missing_metadata(book_id):
            self.last_missing_metadata_sound_book_id = None

    def play_missing_metadata_sound(self):
        if not sys.platform.startswith("win"):
            return
        def play():
            try:
                winsound.Beep(880, 70)
            except Exception:
                pass
        threading.Thread(target=play, daemon=True).start()

    def nvda_book_list_announcements_enabled(self):
        controller = self.get_nvda_controller()
        if not controller:
            return False
        try:
            return controller.nvdaController_testIfRunning() == 0
        except Exception:
            return False

    def get_nvda_controller(self):
        if self.nvda_controller_checked:
            return self.nvda_controller
        self.nvda_controller_checked = True
        if not sys.platform.startswith("win"):
            return None
        path = find_nvda_controller_dll()
        if not path:
            return None
        try:
            controller = ctypes.WinDLL(str(path))
            controller.nvdaController_testIfRunning.restype = ctypes.c_int
            controller.nvdaController_speakText.argtypes = [ctypes.c_wchar_p]
            controller.nvdaController_speakText.restype = ctypes.c_int
            self.nvda_controller = controller
            return controller
        except Exception:
            self.nvda_controller = None
            return None

    def announce_current_book_to_nvda(self, index=None):
        if not self.nvda_book_list_announcements_enabled():
            return
        if self.book_list.size() == 0:
            return
        if index is None:
            selected = self.book_list.curselection()
            if selected:
                index = selected[0]
            else:
                try:
                    index = self.book_list.index("active")
                except Exception:
                    return
        if index < 0 or index >= self.book_list.size():
            return
        controller = self.get_nvda_controller()
        if not controller:
            return
        try:
            if controller.nvdaController_testIfRunning() != 0:
                return
            text = self.book_list_announcement_text(index)
            if not text or text == self.last_nvda_announcement:
                return
            self.last_nvda_announcement = text
            controller.nvdaController_speakText(text)
        except Exception:
            pass

    def speak_text(self, text):
        text = re.sub(r"\s+", " ", text or "").strip()
        if not text:
            return
        self.status_var.set(text)
        try:
            selected_index = self.current_book_index_quiet()
            self.shortcut_readout_var.set(text)
            self.shortcut_readout.focus_force()
            self.shortcut_readout.selection_range(0, END)
            self.shortcut_readout.icursor(END)
            if self.shortcut_readout_return_after is not None:
                self.root.after_cancel(self.shortcut_readout_return_after)
                self.shortcut_readout_return_after = None
        except Exception:
            pass

    def return_to_book_list_from_readout(self, event=None):
        selected_index = self.current_book_index_quiet()
        self.return_focus_from_shortcut_readout(selected_index)
        if event and event.keysym in {"Up", "Down", "Prior", "Next", "Home", "End"}:
            self.root.after(1, lambda keysym=event.keysym: self.move_book_list_after_readout(keysym))
        return "break"

    def move_book_list_after_readout(self, keysym):
        if self.book_list.size() == 0:
            return
        index = self.current_book_index_quiet()
        if index is None:
            index = 0
        if keysym == "Up":
            index -= 1
        elif keysym == "Down":
            index += 1
        elif keysym == "Prior":
            index -= 10
        elif keysym == "Next":
            index += 10
        elif keysym == "Home":
            index = 0
        elif keysym == "End":
            index = self.book_list.size() - 1
        self.select_book_list_index(index)

    def return_focus_from_shortcut_readout(self, selected_index=None):
        self.shortcut_readout_return_after = None
        try:
            if self.root.focus_get() == self.shortcut_readout:
                if self.book_list.size() == 0:
                    self.book_list.focus_set()
                    return
                if selected_index is None:
                    selected_index = self.current_book_index()
                if selected_index is None:
                    selected_index = 0
                selected_index = max(0, min(selected_index, self.book_list.size() - 1))
                self.select_book_list_index(selected_index, focus=True)
        except Exception:
            pass

    def announce_text_to_nvda(self, text):
        if not self.nvda_book_list_announcements_enabled():
            return
        controller = self.get_nvda_controller()
        if not controller:
            return
        try:
            if controller.nvdaController_testIfRunning() != 0:
                return
            text = re.sub(r"\s+", " ", text or "").strip()
            if not text or text == self.last_nvda_announcement:
                return
            self.last_nvda_announcement = text
            controller.nvdaController_speakText(text)
        except Exception:
            pass

    def navigation_title_key(self, title):
        title = re.sub(r"\s+", " ", title or "").strip()
        title = re.sub(r"^(?:the|a)\s+", "", title, flags=re.IGNORECASE)
        return title.casefold()

    def book_list_group_index(self, title_key):
        first = (title_key or "").strip()[:1]
        for index, (_label, letters) in enumerate(BOOK_LIST_ALPHABET_GROUPS):
            if letters and first in letters:
                return index
        return len(BOOK_LIST_ALPHABET_GROUPS) - 1

    def book_list_group_label_for_index(self, global_index):
        if global_index is None or global_index < 0 or global_index >= len(self.book_list_titles):
            return ""
        group_index = self.book_list_group_index(self.book_list_titles[global_index])
        if 0 <= group_index < len(BOOK_LIST_ALPHABET_GROUPS):
            return BOOK_LIST_ALPHABET_GROUPS[group_index][0]
        return ""

    def normalize_book_list_row(self, row):
        if not row:
            return row
        if len(row) == 13:
            return row
        if len(row) == 11:
            return tuple(row) + ("", "")
        return (
            row[0],
            row[1],
            row[2],
            row[3],
            row[4],
            row[6] if len(row) > 6 else "",
            row[9] if len(row) > 9 else "",
            row[10] if len(row) > 10 else "",
            row[11] if len(row) > 11 else "",
            row[12] if len(row) > 12 else "",
            row[13] if len(row) > 13 else "",
            row[21] if len(row) > 21 else "",
            row[22] if len(row) > 22 else "",
        )

    def format_book_list_row(self, row):
        row = self.normalize_book_list_row(row)
        (
            _book_id,
            title,
            author,
            source,
            tags,
            book_format,
            added_at,
            edition,
            year,
            isbn,
            publisher,
            series,
            series_number,
        ) = row
        values = {
            "title": title or "Untitled",
            "author": author or "Unknown",
            "edition": edition or "Not specified",
            "year": year or "Not specified",
            "isbn": isbn or "Not specified",
            "publisher": publisher or "Not specified",
            "series": series or "Not specified",
            "series_number": series_number or "Not specified",
            "source": source or "Not specified",
            "tags": tags or "None",
            "format": book_format or "Unknown",
            "added_at": added_at or "Unknown",
        }
        labels = dict(BOOK_LIST_SPEECH_FIELDS)
        parts = []
        for key in self.get_book_list_speech_fields():
            if key in {"title", "author", "edition", "series", "series_number"}:
                parts.append(f"{values[key]}.")
            else:
                parts.append(f"{labels[key]}: {values[key]}.")
        prefix = "Selected. " if _book_id in self.marked_book_ids else ""
        return prefix + " ".join(parts)

    def book_list_announcement_text(self, index):
        if index is None or index < 0 or index >= len(self.book_list_ids):
            return ""
        if index < len(self.book_list_announcement_texts):
            text = self.book_list_announcement_texts[index]
        else:
            text = self.book_list.get(index)
        return re.sub(r"\s+", " ", text or "").strip()

    def book_list_column_change_announcement_text(self, index):
        text = self.book_list_announcement_text(index)
        group_label = self.book_list_group_label_for_index(index)
        if group_label:
            return f"{group_label}. {text}"
        return text

    def book_field_values_for_shortcuts(self, row):
        labels = dict(BOOK_LIST_SPEECH_FIELDS)
        return {
            "title": (labels["title"], row[1] or "Untitled"),
            "author": (labels["author"], row[2] or "Unknown"),
            "series": (labels["series"], row[21] if len(row) > 21 and row[21] else "Not specified"),
            "series_number": (labels["series_number"], row[22] if len(row) > 22 and row[22] else "Not specified"),
            "edition": (labels["edition"], row[10] or "Not specified"),
            "year": (labels["year"], row[11] or "Not specified"),
            "isbn": (labels["isbn"], row[12] or "Not specified"),
            "publisher": (labels["publisher"], row[13] or "Not specified"),
            "source": (labels["source"], row[3] or "Not specified"),
            "tags": (labels["tags"], row[4] or "None"),
            "format": (labels["format"], row[6] or "Unknown"),
            "added_at": (labels["added_at"], row[9] or "Unknown"),
        }

    def backup_schedule_key(self):
        schedule = self.db.get_setting("backup_schedule", DEFAULT_BACKUP_SCHEDULE)
        if schedule not in BACKUP_SCHEDULES:
            return DEFAULT_BACKUP_SCHEDULE
        return schedule

    def backup_schedule_label(self):
        return BACKUP_SCHEDULES[self.backup_schedule_key()][0]

    def backup_folder(self):
        raw = self.db.get_setting("backup_folder", "").strip()
        if not raw:
            return None
        return Path(raw)

    def backup_paths(self):
        folder = self.backup_folder()
        if not folder:
            return None, None, None
        return folder, folder / "library_backup.db", folder / "library_backup_manifest.json"

    def backup_books_folder(self):
        folder = self.backup_folder()
        if not folder:
            return None
        return folder / BOOKS_FOLDER

    def detected_cloud_folder(self, service):
        home = Path.home()
        if service == "onedrive":
            candidates = [
                os.environ.get("OneDrive"),
                os.environ.get("OneDriveConsumer"),
                os.environ.get("OneDriveCommercial"),
                str(home / "OneDrive"),
            ]
        elif service == "google_drive":
            candidates = [
                str(home / "Google Drive"),
                str(home / "My Drive"),
                str(home / "Google Drive" / "My Drive"),
            ]
        elif service == "icloud":
            candidates = [
                str(home / "iCloudDrive"),
                str(home / "iCloud Drive"),
                str(home / "iCloudPhotos"),
            ]
        else:
            candidates = []
        for candidate in candidates:
            if candidate and Path(candidate).exists():
                return Path(candidate)
        return None

    def choose_cloud_backup_folder(self, service="other"):
        labels = {
            "onedrive": "OneDrive",
            "google_drive": "Google Drive",
            "icloud": "iCloud Drive",
            "other": "cloud",
        }
        label = labels.get(service, "cloud")
        folder = None
        detected = self.detected_cloud_folder(service)
        if detected:
            use_detected = messagebox.askyesno(
                "Use detected cloud folder?",
                f"I found this {label} folder:\n\n{detected}\n\nUse it for library database and book file backups?"
            )
            if use_detected:
                folder = detected
        if folder is None:
            if service != "other" and not detected:
                messagebox.showinfo(
                    "Cloud folder not found",
                    f"I could not find a {label} folder automatically. Choose the synced folder to use for backups."
                )
            chosen = filedialog.askdirectory(title=f"Choose {label} backup folder")
            if not chosen:
                self.focus_books_list()
                return
            folder = Path(chosen)

        backup_folder = folder
        if backup_folder.name != "Accessible Ebook Library Manager Backups":
            backup_folder = cloud_backup_subfolder(backup_folder)
        try:
            backup_folder.mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            messagebox.showerror("Backup folder failed", f"Could not create the backup folder.\n\n{exc}")
            return

        self.db.set_setting("backup_folder", str(backup_folder))
        self.status_var.set(f"Library backup folder set to {backup_folder}.")
        messagebox.showinfo(
            "Library Backup",
            f"Backups will be saved here:\n\n{backup_folder}\n\nThe backup includes the library metadata database and the imported book files. Set a schedule or choose Back Up Now from Settings, Library Backup."
        )
        self.schedule_backup_check(1000)

    def set_backup_schedule(self, schedule):
        if schedule not in BACKUP_SCHEDULES:
            schedule = DEFAULT_BACKUP_SCHEDULE
        self.db.set_setting("backup_schedule", schedule)
        label = BACKUP_SCHEDULES[schedule][0]
        self.status_var.set(f"Library backup schedule set to {label}.")
        messagebox.showinfo("Library Backup Schedule", f"Backup schedule set to {label}.")
        self.schedule_backup_check(1000)

    def backup_manifest(self):
        _folder, _backup_file, manifest_file = self.backup_paths()
        if not manifest_file or not manifest_file.exists():
            return {}
        try:
            return json.loads(manifest_file.read_text(encoding="utf-8"))
        except Exception:
            return {}

    def backup_due(self):
        schedule = self.backup_schedule_key()
        interval = BACKUP_SCHEDULES[schedule][1]
        if interval is None:
            return False
        folder, backup_file, _manifest_file = self.backup_paths()
        if not folder:
            return False
        if not backup_file.exists():
            return True
        last_backup = parse_utc_text(self.db.get_setting("last_backup_at", ""))
        if last_backup is None:
            return True
        db_mtime = str(self.db.db_path.stat().st_mtime)
        backed_mtime = self.db.get_setting("last_backup_db_mtime", "")
        books_folder = self.db.books_path
        book_file_count, book_byte_count = folder_file_stats(books_folder)
        backed_book_signature = self.db.get_setting("last_backup_books_signature", "")
        book_signature = f"{book_file_count}:{book_byte_count}"
        changed = db_mtime != backed_mtime or book_signature != backed_book_signature
        return changed and datetime.now(timezone.utc).replace(tzinfo=None) - last_backup >= interval

    def _start_content_indexer(self):
        threading.Thread(target=self._content_index_worker, daemon=True).start()

    def _content_index_worker(self):
        database = LibraryDatabase()
        while True:
            try:
                book_id, stored_path = self._index_queue.get(timeout=10)
                try:
                    self._index_one_book(book_id, stored_path, database)
                finally:
                    self._index_queue.task_done()
            except queue.Empty:
                unindexed = database.get_unindexed_books()
                for book_id, stored_path in unindexed[:5]:
                    self._index_one_book(book_id, stored_path, database)

    def _index_one_book(self, book_id: int, stored_path: str, database=None):
        database = database or self.db
        try:
            path = Path(stored_path)
            if not path.exists():
                return
            text = extract_text_for_indexing(path)
            database.index_book_content(book_id, text)
        except Exception as exc:
            logger.warning("Could not index book %s at %s: %s", book_id, stored_path, exc)

    def queue_book_for_indexing(self, book_id: int, stored_path: str):
        self._index_queue.put((book_id, stored_path))

    def reindex_library_content(self):
        answer = messagebox.askyesno(
            "Re-index Library Content",
            "This will re-index the text content of all books so they can be found by search. "
            "Indexing runs in the background and may take several minutes for large libraries. "
            "You can continue using the app while it runs.\n\nContinue?",
        )
        if not answer:
            return
        self.db.clear_all_content_index()
        self.status_var.set("Content re-indexing started in background.")
        messagebox.showinfo(
            "Re-indexing started",
            "Library content re-indexing has started in the background. "
            "Search will include book content as indexing progresses.",
        )

    def schedule_backup_check(self, delay_ms=None):
        if self.backup_check_after is not None:
            try:
                self.root.after_cancel(self.backup_check_after)
            except Exception:
                pass
        if delay_ms is None:
            delay_ms = 60 * 60 * 1000
        self.backup_check_after = self.root.after(delay_ms, self.check_library_backup)

    def check_library_backup(self):
        self.backup_check_after = None
        try:
            self.notice_if_cloud_backup_changed()
            if self.backup_due():
                self.backup_library_now(automatic=True)
        finally:
            self.schedule_backup_check()

    def notice_if_cloud_backup_changed(self):
        _folder, backup_file, _manifest_file = self.backup_paths()
        if not backup_file or not backup_file.exists():
            return
        current_mtime = str(backup_file.stat().st_mtime)
        last_seen = self.db.get_setting("last_seen_backup_file_mtime", "")
        if last_seen and current_mtime != last_seen:
            self.status_var.set("The cloud library backup changed since the last check. Use Settings, Library Backup, Restore From Backup if you need it.")
        self.db.set_setting("last_seen_backup_file_mtime", current_mtime)

    def backup_library_now(self, automatic=False):
        folder, backup_file, manifest_file = self.backup_paths()
        backup_books_folder = self.backup_books_folder()
        if not folder:
            if automatic:
                return
            messagebox.showinfo(
                "Choose a backup folder",
                "Choose a Google Drive, OneDrive, iCloud Drive, or other synced folder before backing up the library database and imported book files."
            )
            self.choose_cloud_backup_folder("other")
            folder, backup_file, manifest_file = self.backup_paths()
            backup_books_folder = self.backup_books_folder()
            if not folder:
                return

        try:
            folder.mkdir(parents=True, exist_ok=True)
            self.db.backup_to(backup_file)
            sync_folder_contents(self.db.books_path, backup_books_folder)
            db_mtime = str(self.db.db_path.stat().st_mtime)
            book_count = self.db.connection.execute("SELECT COUNT(*) FROM books").fetchone()[0]
            book_file_count, book_byte_count = folder_file_stats(self.db.books_path)
            backup_file_count, backup_byte_count = folder_file_stats(backup_books_folder)
            book_signature = f"{book_file_count}:{book_byte_count}"
            manifest = {
                "app": APP_NAME,
                "created_at": utc_now_text(),
                "source_database": str(self.db.db_path),
                "backup_database": str(backup_file),
                "source_books_folder": str(self.db.books_path),
                "backup_books_folder": str(backup_books_folder),
                "source_database_mtime": db_mtime,
                "book_count": book_count,
                "book_file_count": book_file_count,
                "book_byte_count": book_byte_count,
                "backup_book_file_count": backup_file_count,
                "backup_book_byte_count": backup_byte_count,
            }
            manifest_file.write_text(json.dumps(manifest, indent=2), encoding="utf-8")
            self.db.set_setting("last_backup_at", manifest["created_at"])
            self.db.set_setting("last_backup_db_mtime", db_mtime)
            self.db.set_setting("last_backup_books_signature", book_signature)
            self.db.set_setting("last_seen_backup_file_mtime", str(backup_file.stat().st_mtime))
            message = f"Library database and {backup_file_count} book file{'s' if backup_file_count != 1 else ''} backed up to {folder}."
            self.status_var.set(message)
            if not automatic:
                messagebox.showinfo("Library Backup Complete", message)
        except Exception as exc:
            self.status_var.set("Library backup failed.")
            if not automatic:
                messagebox.showerror("Library Backup Failed", f"Could not back up the library database and imported book files.\n\n{exc}")

    def backup_file_is_valid(self, backup_file):
        try:
            connection = sqlite3.connect(backup_file)
            try:
                row = connection.execute(
                    "SELECT name FROM sqlite_master WHERE type='table' AND name='books'"
                ).fetchone()
                if row is None:
                    return False
                integrity = connection.execute("PRAGMA integrity_check").fetchone()
                if not integrity or integrity[0] != "ok":
                    return False
                columns = {item[1] for item in connection.execute("PRAGMA table_info(books)")}
                return {"id", "title", "stored_path"}.issubset(columns)
            finally:
                connection.close()
        except Exception:
            return False

    def validate_backup_contents(self, backup_file, manifest_file, books_folder):
        if not self.backup_file_is_valid(backup_file):
            return "The backup database is incomplete or failed its integrity check."

        connection = sqlite3.connect(backup_file)
        try:
            book_count = connection.execute("SELECT COUNT(*) FROM books").fetchone()[0]
            stored_names = {
                Path(row[0] or "").name
                for row in connection.execute("SELECT stored_path FROM books")
                if Path(row[0] or "").name
            }
        finally:
            connection.close()

        manifest = {}
        if manifest_file and manifest_file.exists():
            try:
                manifest = json.loads(manifest_file.read_text(encoding="utf-8"))
            except Exception:
                return "The backup manifest could not be read. The backup may still be syncing or damaged."
            if manifest.get("app") not in {None, APP_NAME}:
                return "The backup manifest belongs to a different application."
            try:
                if "book_count" in manifest and int(manifest["book_count"]) != book_count:
                    return "The database book count does not match the backup manifest. The backup may be only partly synced."
            except (TypeError, ValueError):
                return "The backup manifest contains an invalid book count."

        if books_folder and books_folder.is_dir():
            actual_count, actual_bytes = folder_file_stats(books_folder)
            expected_count = manifest.get("backup_book_file_count")
            expected_bytes = manifest.get("backup_book_byte_count")
            try:
                if expected_count is not None and int(expected_count) != actual_count:
                    return "The backed up book files do not match the manifest count. The backup may be only partly synced."
                if expected_bytes is not None and int(expected_bytes) != actual_bytes:
                    return "The backed up book files do not match the manifest size. The backup may be only partly synced."
            except (TypeError, ValueError):
                return "The backup manifest contains invalid book file counts or sizes."
            available_names = {path.name for path in books_folder.rglob("*") if path.is_file()}
            missing = stored_names - available_names
            if missing:
                return f"The backup is missing {len(missing)} book file(s) referenced by its database. No restore was started."
            if book_count and not actual_count:
                return "The backup database contains books, but the backed up Books folder is empty."
        elif book_count and manifest.get("backup_book_file_count") is not None:
            return "The backup contains book records but its Books folder is missing."

        return ""

    def repair_restored_stored_paths(self):
        rows = self.db.connection.execute("SELECT id, stored_path FROM books").fetchall()
        for book_id, stored_path in rows:
            filename = Path(stored_path or "").name
            if not filename:
                continue
            restored_path = self.db.books_path / filename
            if restored_path.exists() and str(restored_path) != stored_path:
                self.db.connection.execute(
                    "UPDATE books SET stored_path = ?, format = ? WHERE id = ?",
                    (str(restored_path), restored_path.suffix.lower().lstrip("."), book_id),
                )
        self.db.connection.commit()

    def restore_library_backup(self):
        folder, backup_file, _manifest_file = self.backup_paths()
        backup_books_folder = self.backup_books_folder()
        if not folder:
            messagebox.showinfo("Choose a backup folder", "Choose a backup folder before restoring.")
            self.choose_cloud_backup_folder("other")
            folder, backup_file, _manifest_file = self.backup_paths()
            backup_books_folder = self.backup_books_folder()
            if not folder:
                return
        if not backup_file.exists():
            messagebox.showerror("Backup not found", f"No library backup was found here:\n\n{backup_file}")
            return
        validation_error = self.validate_backup_contents(backup_file, _manifest_file, backup_books_folder)
        if validation_error:
            messagebox.showerror("Backup not ready to restore", validation_error)
            return

        books_backup_available = backup_books_folder and backup_books_folder.exists() and backup_books_folder.is_dir()
        restore_scope = (
            "the cloud backup over the current local library database and imported book files"
            if books_backup_available
            else "the cloud backup over the current local library database"
        )
        books_note = (
            "The current local database and Books folder will be saved as safety copies first."
            if books_backup_available
            else "This backup does not include a Books folder, so only the database will be restored. The current local database will be saved as a safety copy first."
        )
        if not messagebox.askyesno(
            "Restore Library Backup",
            f"Restore {restore_scope}?\n\n{books_note}"
        ):
            self.focus_books_list()
            return

        preserved_folder = self.db.get_setting("backup_folder", "")
        preserved_schedule = self.db.get_setting("backup_schedule", DEFAULT_BACKUP_SCHEDULE)
        safety_copy = self.db.folder / f"library_before_restore_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
        safety_books_folder = self.db.folder / f"{BOOKS_FOLDER}_before_restore_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        try:
            self.db.connection.commit()
            shutil.copy2(self.db.db_path, safety_copy)
            if books_backup_available and self.db.books_path.exists():
                sync_folder_contents(self.db.books_path, safety_books_folder)
            self.db.close()
            shutil.copy2(backup_file, self.db.db_path)
            self.db = LibraryDatabase()
            if books_backup_available:
                replace_folder_from_backup(backup_books_folder, self.db.books_path)
                self.repair_restored_stored_paths()
            self.db.set_setting("backup_folder", preserved_folder)
            self.db.set_setting("backup_schedule", preserved_schedule)
            self.refresh_books()
            self.settle_book_list_focus()
            if books_backup_available:
                restored_file_count, _restored_bytes = folder_file_stats(self.db.books_path)
                message = (
                    f"Library database and {restored_file_count} book file{'s' if restored_file_count != 1 else ''} restored from backup. "
                    f"Safety copies saved at {safety_copy} and {safety_books_folder}."
                )
            else:
                message = f"Library database restored from backup. Safety copy saved at {safety_copy}."
            self.status_var.set(message)
            messagebox.showinfo("Library Restored", message)
        except Exception as exc:
            try:
                self.db = LibraryDatabase()
            except Exception:
                pass
            messagebox.showerror("Restore Failed", f"Could not restore the library backup.\n\n{exc}")

    def show_backup_status(self):
        folder, backup_file, manifest_file = self.backup_paths()
        backup_books_folder = self.backup_books_folder()
        manifest = self.backup_manifest()
        current_backup_book_count, current_backup_book_bytes = folder_file_stats(backup_books_folder) if backup_books_folder else (0, 0)
        text = (
            f"Backup folder: {folder or 'Not set'}\n"
            f"Schedule: {self.backup_schedule_label()}\n"
            f"Last backup: {self.db.get_setting('last_backup_at', 'Never') or 'Never'}\n"
            f"Database backup file: {backup_file if backup_file and backup_file.exists() else 'Not found'}\n"
            f"Books backup folder: {backup_books_folder if backup_books_folder and backup_books_folder.exists() else 'Not found'}\n"
            f"Book files in backup folder: {current_backup_book_count}\n"
            f"Manifest file: {manifest_file if manifest_file and manifest_file.exists() else 'Not found'}\n"
            f"Books in last manifest: {manifest.get('book_count', 'Unknown')}\n"
            f"Book files in last manifest: {manifest.get('backup_book_file_count', 'Unknown')}\n"
            f"Book backup size in last manifest: {manifest.get('backup_book_byte_count', current_backup_book_bytes)} bytes"
        )
        messagebox.showinfo("Library Backup Status", text)

    def watched_folders(self):
        raw = self.db.get_setting("watched_folders", "[]")
        try:
            folders = json.loads(raw)
        except Exception:
            folders = []
        clean = []
        seen = set()
        for folder in folders:
            try:
                path = str(Path(folder).expanduser())
            except Exception:
                continue
            key = path.casefold()
            if path and key not in seen:
                clean.append(path)
                seen.add(key)
        return clean

    def save_watched_folders(self, folders):
        self.db.set_setting("watched_folders", json.dumps(folders))

    def watched_folder_auto_scan_enabled(self):
        return self.db.get_setting("watched_folder_auto_scan", "1") == "1"

    def watched_file_signatures(self):
        raw = self.db.get_setting("watched_file_signatures", "{}")
        try:
            data = json.loads(raw)
        except Exception:
            data = {}
        return data if isinstance(data, dict) else {}

    def save_watched_file_signatures(self, signatures):
        self.db.set_setting("watched_file_signatures", json.dumps(signatures, sort_keys=True))

    def file_signature(self, path: Path):
        try:
            stat = path.stat()
            return f"{stat.st_mtime_ns}:{stat.st_size}"
        except Exception:
            return ""

    def canonical_path_text(self, path: Path):
        try:
            return str(path.resolve())
        except Exception:
            return str(path.absolute())

    def path_is_inside(self, path: Path, parent: Path):
        try:
            path.resolve().relative_to(parent.resolve())
            return True
        except Exception:
            return False

    def add_watched_folder(self):
        folder = filedialog.askdirectory(title="Choose folder to watch for books")
        if not folder:
            return
        path = self.canonical_path_text(Path(folder))
        folders = self.watched_folders()
        if path.casefold() not in [existing.casefold() for existing in folders]:
            folders.append(path)
            self.save_watched_folders(folders)
        self.db.set_setting("watched_folder_auto_scan", "1")
        self.status_var.set(f"Watched folder added: {path}")
        messagebox.showinfo(
            "Watched Folder Added",
            f"The app will scan this folder for new or changed books:\n\n{path}"
        )
        self.schedule_watched_folder_scan(1000)

    def remove_watched_folder(self):
        folders = self.watched_folders()
        if not folders:
            messagebox.showinfo("Watched Folders", "No watched folders are set.")
            return
        lines = [f"{index}. {folder}" for index, folder in enumerate(folders, start=1)]
        value = AccessibleSingleFieldDialog.ask(
            self.root,
            "Remove Watched Folder",
            "Type the number of the watched folder to remove.\n\n" + "\n".join(lines),
            "1",
            heading="Remove Watched Folder",
        )
        if value is None:
            return
        try:
            index = int(value.strip())
        except ValueError:
            messagebox.showerror("Invalid choice", "Please enter one of the listed numbers.")
            return
        if index < 1 or index > len(folders):
            messagebox.showerror("Invalid choice", "Please enter one of the listed numbers.")
            return
        removed = folders.pop(index - 1)
        self.save_watched_folders(folders)
        self.status_var.set(f"Watched folder removed: {removed}")
        messagebox.showinfo("Watched Folder Removed", f"Removed watched folder:\n\n{removed}")

    def toggle_watched_folder_auto_scan(self):
        enabled = not self.watched_folder_auto_scan_enabled()
        self.db.set_setting("watched_folder_auto_scan", "1" if enabled else "0")
        state = "on" if enabled else "off"
        self.status_var.set(f"Automatic watched folder scanning turned {state}.")
        messagebox.showinfo("Watched Folders", f"Automatic watched folder scanning is now {state}.")
        if enabled:
            self.schedule_watched_folder_scan(1000)

    def show_watched_folder_status(self):
        folders = self.watched_folders()
        state = "On" if self.watched_folder_auto_scan_enabled() else "Off"
        window = tk.Toplevel(self.root)
        window.title("Watched Folder Status")
        window.transient(self.root)
        window.grab_set()

        main = ttk.Frame(window, padding=12)
        main.pack(fill=BOTH, expand=True)
        ttk.Label(main, text=f"Automatic scanning: {state}").pack(anchor="w", pady=(0, 10))

        if not folders:
            ttk.Label(main, text="No watched folders are configured.").pack(anchor="w")
        else:
            ttk.Label(main, text=f"Watched folders: {len(folders)}").pack(anchor="w", pady=(0, 6))
            path_entries = []
            for index, folder in enumerate(folders, start=1):
                ttk.Label(main, text=f"Folder {index} of {len(folders)}").pack(anchor="w")
                path_var = StringVar(value=folder)
                entry = ttk.Entry(main, textvariable=path_var, state="readonly", width=80)
                entry.pack(fill=X, expand=True, pady=(2, 8))
                path_entries.append(entry)

        button = ttk.Button(main, text="Close", command=window.destroy)
        button.pack(anchor="e", pady=(4, 0))
        window.bind("<Escape>", lambda _event: window.destroy())
        window.bind("<Return>", lambda _event: window.destroy())

        def focus_status():
            window.lift()
            if folders:
                path_entries[0].focus_force()
                path_entries[0].selection_range(0, END)
                path_entries[0].icursor(END)
            else:
                button.focus_force()

        window.after(50, focus_status)
        window.after(200, focus_status)

    def schedule_watched_folder_scan(self, delay_ms=None):
        if self.watched_scan_after is not None:
            try:
                self.root.after_cancel(self.watched_scan_after)
            except Exception:
                pass
        if delay_ms is None:
            delay_ms = 5 * 60 * 1000
        self.watched_scan_after = self.root.after(delay_ms, self.check_watched_folders)

    def check_watched_folders(self):
        self.watched_scan_after = None
        try:
            if self.watched_folder_auto_scan_enabled() and self.watched_folders():
                self.scan_watched_folders(automatic=True)
        finally:
            self.schedule_watched_folder_scan()

    def scan_watched_folders_now(self):
        self.scan_watched_folders(automatic=False)

    def iter_watched_book_files(self, folders=None, managed_folder=None):
        folders = self.watched_folders() if folders is None else folders
        managed_folder = self.db.books_path if managed_folder is None else managed_folder
        for folder_text in folders:
            folder = Path(folder_text)
            if not folder.exists():
                yield None, f"{folder} -- watched folder not found"
                continue
            try:
                for path in folder.rglob("*"):
                    if self.path_is_inside(path, managed_folder):
                        continue
                    if self.is_ignored_watched_file(path):
                        continue
                    if path.is_file() and path.suffix.lower() in SUPPORTED_EXTENSIONS:
                        yield path, ""
            except Exception as exc:
                yield None, f"{folder} -- {exc}"

    def is_ignored_watched_file(self, path: Path):
        return self.is_ignored_import_file(path)

    def is_ignored_import_file(self, path: Path):
        if not path.is_file():
            return False

        if path.stem.casefold().startswith("quickstart"):
            return True

        if path.name.casefold() in {
            "please do not delete this file.txt",
            "do not delete.txt",
        }:
            return True

        if path.suffix.lower() not in {".txt", ".rtf", ".html", ".htm"}:
            return False

        try:
            if path.stat().st_size > 65536:
                return False
            text = path.read_text(encoding="utf-8", errors="ignore")
        except Exception:
            return False

        normalized_text = re.sub(r"\s+", " ", text).strip().casefold()
        normalized_notice = re.sub(r"\s+", " ", VOICE_DREAM_LIBRARY_NOTICE).strip().casefold()
        return normalized_notice in normalized_text

    def update_existing_watched_book(self, row, source_path: Path, database=None):
        database = database or self.db
        stored_path = Path(row[8])
        if not stored_path.exists():
            destination = self.unique_destination(
                safe_filename(f"{row[2]} - {row[1]}").strip(" -") or safe_filename(row[1]),
                source_path.suffix.lower(),
            )
            shutil.copy2(source_path, destination)
            database.connection.execute(
                "UPDATE books SET stored_path = ?, format = ? WHERE id = ?",
                (str(destination), destination.suffix.lower().lstrip("."), row[0]),
            )
            database.connection.commit()
            stored_path = destination
        else:
            shutil.copy2(source_path, stored_path)
        if stored_path.suffix.lower() == ".epub":
            self.update_accessibility_from_epub(row[0], stored_path, database=database)

    def watched_file_matches_existing_book(self, source_path: Path, database=None):
        database = database or self.db
        try:
            source_stat = source_path.stat()
        except Exception:
            return None

        for row in database.all_books_for_duplicate_check():
            stored_path = Path(row[8])
            try:
                if source_path.resolve() == stored_path.resolve():
                    return row
            except Exception:
                pass

            try:
                stored_stat = stored_path.stat()
            except Exception:
                continue

            if source_stat.st_size != stored_stat.st_size:
                continue

            try:
                if filecmp.cmp(source_path, stored_path, shallow=False):
                    return row
            except Exception:
                continue
        return None

    def watched_file_matches_existing_metadata(self, source_path: Path, database=None):
        database = database or self.db
        if source_path.suffix.lower() == ".zip":
            return None, ""

        if source_path.suffix.lower() == ".epub":
            # A Bookshare or Learning Ally title keeps its library ID however the
            # file is renamed, so it is a more reliable match than the title.
            library_id = extract_source_book_id(source_path)
            matching_id = database.find_book_id_by_source_book_id(library_id)
            matching_row = database.get_book(matching_id) if matching_id else None
            if matching_row:
                return matching_row, f"same library book ID {library_id}"

        detected = self.guess_metadata_from_file(source_path, database=database)
        detected_isbn = normalize_isbn_key(detected.get("isbn", ""))
        source_title = detected.get("title", "")
        filename_title = clean_filename_title(source_path)
        candidate_titles = [title for title in [source_title, filename_title] if title]
        candidate_title_keys = {normalize_duplicate_key(title) for title in candidate_titles if normalize_duplicate_key(title)}
        detected_metadata_score = metadata_score_from_detection(detected)
        source_extension = source_path.suffix.lower().lstrip(".")

        for row in database.all_books_for_duplicate_check():
            row_isbn = normalize_isbn_key(row[12] or "")
            if detected_isbn and row_isbn and detected_isbn == row_isbn:
                return row, f"same ISBN {detected_isbn}"

            row_title_key = normalize_duplicate_key(row[1])
            if row_title_key and row_title_key in candidate_title_keys:
                return row, "same normalized title"

            row_metadata_score = metadata_score_from_row(row)
            allow_richer_metadata_match = row_metadata_score > detected_metadata_score
            for candidate_title in candidate_titles:
                if title_keys_look_same(row[1], candidate_title, allow_richer_metadata_match=allow_richer_metadata_match):
                    if allow_richer_metadata_match:
                        return row, "same title with richer corrected metadata"
                    return row, "same likely title"

            path_keys = set()
            path_titles = []
            for path_text in [row[7], row[8]]:
                if not path_text:
                    continue
                path = Path(path_text)
                if source_extension and path.suffix.lower().lstrip(".") != source_extension:
                    continue
                path_title = clean_filename_title(path)
                path_titles.append(path_title)
                path_keys.add(normalize_duplicate_key(path_title))
            if candidate_title_keys & path_keys:
                return row, "same filename title"
            for path_title in path_titles:
                for candidate_title in candidate_titles:
                    if title_keys_look_same(path_title, candidate_title, allow_richer_metadata_match=allow_richer_metadata_match):
                        if allow_richer_metadata_match:
                            return row, "same filename title with richer corrected metadata"
                        return row, "same likely filename title"

        return None, ""

    def scan_watched_folders(self, automatic=False):
        if self.watched_scan_running:
            return
        folders = self.watched_folders()
        if not folders:
            if not automatic:
                messagebox.showinfo("Watched Folders", "No watched folders are set. Add one from Settings, Watched Folders.")
            return

        self.watched_scan_running = True
        events = queue.Queue()
        progress_window = None
        message_var = None
        progress = None
        if not automatic:
            progress_window = tk.Toplevel(self.root)
            progress_window.title("Scanning Watched Folders")
            progress_window.transient(self.root)
            progress_window.resizable(False, False)
            progress_window.protocol("WM_DELETE_WINDOW", lambda: None)
            message_var = StringVar(value="Scanning watched folders.")
            ttk.Label(progress_window, textvariable=message_var, padding=12).pack(fill=X)
            progress = ttk.Progressbar(progress_window, mode="indeterminate", length=360)
            progress.pack(fill=X, padx=12, pady=(0, 12))
            progress.start(12)
        self.status_var.set("Scanning watched folders in the background.")
        signatures = self.watched_file_signatures()

        def worker():
            database = None
            imported = 0
            updated = 0
            skipped = []
            scan_signatures = dict(signatures)
            try:
                database = LibraryDatabase()
                for path, error in self.iter_watched_book_files(folders, database.books_path):
                    if error:
                        skipped.append(error)
                        continue
                    if path is None:
                        continue
                    canonical = self.canonical_path_text(path)
                    signature = self.file_signature(path)
                    if signature and scan_signatures.get(canonical) == signature:
                        continue
                    events.put(("progress", path.name))
                    try:
                        existing = database.get_book_by_original_path(canonical)
                        if existing and path.suffix.lower() != ".zip":
                            self.update_existing_watched_book(existing, path, database=database)
                            updated += 1
                        elif self.watched_file_matches_existing_book(path, database=database):
                            skipped.append(f"{path} -- already in library")
                        else:
                            matched_row, match_reason = self.watched_file_matches_existing_metadata(path, database=database)
                            if matched_row:
                                skipped.append(f"{path} -- already in library as {matched_row[1]} ({match_reason})")
                            elif path.suffix.lower() == ".zip":
                                zip_imported, zip_skipped = self.import_zip_file_without_prompt(
                                    path, default_source="Watched Folder", database=database
                                )
                                imported += zip_imported
                                skipped.extend(zip_skipped)
                            elif self.import_one_book_without_prompt(path, default_source="Watched Folder", database=database):
                                imported += 1
                        if signature:
                            scan_signatures[canonical] = signature
                    except sqlite3.IntegrityError:
                        if signature:
                            scan_signatures[canonical] = signature
                        skipped.append(f"{path} -- already imported")
                    except Exception as exc:
                        skipped.append(f"{path} -- {exc}")
                        logger.exception("Watched folder scan failed for %s", path)
                report_path = self.write_import_report(imported + updated, skipped) if (skipped or imported or updated) else ""
                events.put(("done", imported, updated, skipped, scan_signatures, report_path, ""))
            except Exception as exc:
                logger.exception("Watched folder scan failed")
                events.put(("done", imported, updated, skipped, scan_signatures, "", str(exc)))
            finally:
                if database is not None:
                    database.close()

        def poll():
            try:
                while True:
                    event = events.get_nowait()
                    if event[0] == "progress":
                        message = f"Checking: {event[1]}"
                        self.status_var.set(message)
                        if message_var is not None:
                            message_var.set(message)
                    else:
                        _kind, imported, updated, skipped, signatures, report_path, error = event
                        self.watched_scan_running = False
                        if progress_window is not None:
                            progress.stop()
                            progress_window.destroy()
                        if error:
                            self.status_var.set("Watched folder scan failed.")
                            if not automatic:
                                messagebox.showerror("Watched Folder Scan Failed", error)
                            return
                        self.save_watched_file_signatures(signatures)
                        if imported or updated:
                            self.refresh_books()
                        message = f"Watched folder scan complete. Imported {imported}. Updated {updated}. Skipped {len(skipped)}."
                        self.status_var.set(message)
                        if not automatic and skipped:
                            message += f"\n\nImport report saved at:\n{report_path}"
                        if not automatic:
                            messagebox.showinfo("Watched Folder Scan", message)
                        return
            except queue.Empty:
                pass
            if self.watched_scan_running:
                self.root.after(100, poll)

        threading.Thread(target=worker, daemon=True, name="watched-folder-scan").start()
        self.root.after(100, poll)

    def explorer_context_menu_registry_paths(self):
        paths = []
        for extension in sorted(SUPPORTED_EXTENSIONS):
            paths.append(
                rf"Software\Classes\SystemFileAssociations\{extension}\shell\{EXPLORER_CONTEXT_MENU_KEY}"
            )
        return paths

    def delete_registry_tree(self, root, subkey):
        if winreg is None:
            return
        try:
            with winreg.OpenKey(root, subkey, 0, winreg.KEY_READ | winreg.KEY_WRITE) as key:
                while True:
                    try:
                        child = winreg.EnumKey(key, 0)
                    except OSError:
                        break
                    self.delete_registry_tree(root, subkey + "\\" + child)
            winreg.DeleteKey(root, subkey)
        except FileNotFoundError:
            pass

    def install_file_explorer_context_menu(self):
        if not sys.platform.startswith("win") or winreg is None:
            messagebox.showinfo(
                "Windows only",
                "File Explorer right-click integration is only available on Windows."
            )
            return
        command = app_launch_command_for_file_argument()
        try:
            for menu_path in self.explorer_context_menu_registry_paths():
                with winreg.CreateKey(winreg.HKEY_CURRENT_USER, menu_path) as key:
                    winreg.SetValueEx(key, "", 0, winreg.REG_SZ, EXPLORER_CONTEXT_MENU_LABEL)
                    winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, sys.executable)
                with winreg.CreateKey(winreg.HKEY_CURRENT_USER, menu_path + r"\command") as key:
                    winreg.SetValueEx(key, "", 0, winreg.REG_SZ, command)
            self.status_var.set("File Explorer right-click command installed.")
            messagebox.showinfo(
                "File Explorer Integration",
                "Installed the File Explorer right-click command for supported ebook and document files.\n\n"
                f"Command name: {EXPLORER_CONTEXT_MENU_LABEL}"
            )
        except Exception as exc:
            messagebox.showerror(
                "File Explorer integration failed",
                f"Could not install the right-click command.\n\n{exc}"
            )

    def remove_file_explorer_context_menu(self):
        if not sys.platform.startswith("win") or winreg is None:
            messagebox.showinfo(
                "Windows only",
                "File Explorer right-click integration is only available on Windows."
            )
            return
        try:
            for menu_path in self.explorer_context_menu_registry_paths():
                self.delete_registry_tree(winreg.HKEY_CURRENT_USER, menu_path)
            self.status_var.set("File Explorer right-click command removed.")
            messagebox.showinfo(
                "File Explorer Integration",
                "Removed the File Explorer right-click command."
            )
        except Exception as exc:
            messagebox.showerror(
                "File Explorer integration failed",
                f"Could not remove the right-click command.\n\n{exc}"
            )

    def file_explorer_context_menu_installed_count(self):
        if not sys.platform.startswith("win") or winreg is None:
            return 0
        count = 0
        for menu_path in self.explorer_context_menu_registry_paths():
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, menu_path):
                    count += 1
            except FileNotFoundError:
                pass
        return count

    def show_file_explorer_context_menu_status(self):
        total = len(SUPPORTED_EXTENSIONS)
        installed = self.file_explorer_context_menu_installed_count()
        command = app_launch_command_for_file_argument()
        messagebox.showinfo(
            "File Explorer Integration Status",
            f"Right-click entries installed: {installed} of {total}\n\n"
            f"Command name: {EXPLORER_CONTEXT_MENU_LABEL}\n\n"
            f"Launch command:\n{command}"
        )

    def focus_search(self):
        value = AccessibleSingleFieldDialog.ask(
            self.root,
            "Search Library",
            "Enter text to search in title, author, source, tags, format, notes, edition, year, ISBN, publisher, and book content. Leave blank to show all books.",
            self.search_var.get(),
            heading="Search Library",
        )
        if value is None:
            self.focus_books_list()
            return
        self.search_var.set(value.strip())
        self.search_and_focus()

    def focus_search_from_keyboard(self, event=None):
        self.focus_search()
        return "break"

    def focus_books_list(self):
        if self.book_list.size() == 0:
            target = self.first_visible_book_list_widget()
            if target is not None:
                target.focus_set()
            self.status_var.set("No books are shown.")
            return
        index = self.current_book_index_quiet()
        if index is None:
            index = 0
        self.select_book_list_index(index, focus=True)
        self.status_var.set("Books list focused. Use up and down arrow to choose a book.")

    def settle_book_list_focus(self, selected_index=None):
        if self.book_list.size() == 0:
            target = self.first_visible_book_list_widget()
            if target is not None:
                target.focus_force()
            return
        if selected_index is None:
            selected_index = self.current_book_index_quiet()
            if selected_index is None:
                selected_index = 0
        selected_index = max(0, min(selected_index, self.book_list.size() - 1))
        self.select_book_list_index(selected_index, focus=True)

    def first_visible_book_list_widget(self):
        if self.is_multiple_book_list_mode() and self.multi_book_lists:
            if 0 <= self.active_multi_book_column < len(self.multi_book_lists):
                return self.multi_book_lists[self.active_multi_book_column]
            return self.multi_book_lists[0]
        return self.book_list

    def focused_multi_book_list(self):
        focused = self.root.focus_get()
        if focused in self.multi_book_lists:
            return focused
        return None

    def multi_book_list_position_for_global_index(self, global_index):
        for column_index, index_map in enumerate(self.multi_book_list_index_maps):
            if global_index in index_map:
                return column_index, index_map.index(global_index)
        return None, None

    def global_index_from_multi_widget(self, widget):
        if widget not in self.multi_book_lists:
            return None
        column_index = self.multi_book_lists.index(widget)
        index_map = self.multi_book_list_index_maps[column_index]
        selected = widget.curselection()
        if selected:
            local_index = selected[0]
        else:
            try:
                local_index = widget.index("active")
            except Exception:
                return None
        if local_index < 0 or local_index >= len(index_map):
            return None
        return index_map[local_index]

    def select_multi_book_list_index(self, global_index, focus_widget=None, focus=False):
        if not self.multi_book_lists:
            return
        target_widget = focus_widget
        for listbox in self.multi_book_lists:
            listbox.selection_clear(0, END)
        column_index, local_index = self.multi_book_list_position_for_global_index(global_index)
        if column_index is None:
            return
        self.active_multi_book_column = column_index
        self.show_active_multi_book_column()
        target_widget = target_widget or self.multi_book_lists[column_index]
        listbox = self.multi_book_lists[column_index]
        listbox.selection_set(local_index)
        listbox.activate(local_index)
        listbox.see(local_index)
        if focus and target_widget in self.multi_book_lists:
            target_widget.focus_set()

    def restore_temporary_column_announcement_item(self):
        item = self.temporary_column_announcement_item
        self.temporary_column_announcement_item = None
        if not item:
            return
        listbox, local_index, original_text = item
        try:
            if 0 <= local_index < listbox.size():
                listbox.delete(local_index)
                listbox.insert(local_index, original_text)
        except Exception:
            pass

    def set_temporary_column_announcement_item(self, listbox, local_index, text):
        self.restore_temporary_column_announcement_item()
        try:
            if local_index < 0 or local_index >= listbox.size():
                return
            original_text = listbox.get(local_index)
            listbox.delete(local_index)
            listbox.insert(local_index, text)
            self.temporary_column_announcement_item = (listbox, local_index, original_text)
            self.root.after(1500, self.restore_temporary_column_announcement_item)
        except Exception:
            self.temporary_column_announcement_item = None

    def select_book_list_index(self, index, focus=None, focus_widget=None, announce=True):
        if announce:
            self.restore_temporary_column_announcement_item()
        if self.book_list.size() == 0:
            return
        index = max(0, min(index, self.book_list.size() - 1))
        book_id = self.book_list_ids[index] if index < len(self.book_list_ids) else None
        self.book_list.selection_clear(0, END)
        self.book_list.selection_set(index)
        self.book_list.activate(index)
        self.book_list.see(index)
        if self.is_multiple_book_list_mode():
            if focus is None:
                focus = True
            self.select_multi_book_list_index(index, focus_widget=focus_widget, focus=focus)
        elif focus or focus is None:
            self.book_list.focus_set()
        if book_id is not None:
            if announce:
                self.status_var.set(self.book_list_announcement_text(index))
            else:
                self.suppress_next_book_list_select_announcement = True
            self.play_missing_metadata_sound_if_needed(book_id)
            if announce:
                self.announce_current_book_to_nvda(index)
            self.schedule_cover_preview_update(book_id)
        else:
            self.status_var.set(self.book_list.get(index))
            self.clear_cover_preview()

    def move_to_first_book(self, event=None):
        if self.book_list.size() == 0:
            self.status_var.set("No books are shown.")
            return "break"
        self.select_book_list_index(0)
        return "break"

    def move_to_last_book(self, event=None):
        if self.book_list.size() == 0:
            self.status_var.set("No books are shown.")
            return "break"
        self.select_book_list_index(self.book_list.size() - 1)
        return "break"

    def move_between_book_columns(self, event, direction):
        if not self.is_multiple_book_list_mode() or event.widget not in self.multi_book_lists:
            return None
        if self.book_list.size() == 0:
            self.status_var.set("No books are shown.")
            return "break"

        current_column = self.multi_book_lists.index(event.widget)
        selected = event.widget.curselection()
        if selected:
            current_local_index = selected[0]
        else:
            try:
                current_local_index = event.widget.index("active")
            except Exception:
                current_local_index = 0

        column_count = len(self.multi_book_lists)
        for offset in range(1, column_count + 1):
            next_column = (current_column + (offset * direction)) % column_count
            index_map = self.multi_book_list_index_maps[next_column]
            if not index_map:
                continue
            next_local_index = min(current_local_index, len(index_map) - 1)
            next_global_index = index_map[next_local_index]
            text = self.book_list_column_change_announcement_text(next_global_index)
            self.set_temporary_column_announcement_item(
                self.multi_book_lists[next_column],
                next_local_index,
                text,
            )
            self.select_book_list_index(
                next_global_index,
                focus=True,
                focus_widget=self.multi_book_lists[next_column],
                announce=False,
            )
            self.status_var.set(text)
            self.announce_text_to_nvda(text)
            return "break"

        return "break"

    def toggle_mark_current_book(self, event=None):
        index = self.current_book_index()
        if index is None or index >= len(self.book_list_ids):
            return "break"
        book_id = self.book_list_ids[index]
        if book_id in self.marked_book_ids:
            self.marked_book_ids.remove(book_id)
            state = "not selected"
        else:
            self.marked_book_ids.add(book_id)
            state = "selected"
        self.refresh_books(selected_book_id=book_id)
        self.root.after(75, lambda selected_index=index: self.settle_book_list_focus(selected_index))
        self.status_var.set(f"Book {state}. {len(self.marked_book_ids)} book{'s' if len(self.marked_book_ids) != 1 else ''} selected.")
        return "break"

    def select_all_books(self, event=None):
        if event is not None and isinstance(event.widget, (tk.Entry, tk.Text)):
            return None
        if not self.book_list_ids:
            self.status_var.set("No books are shown to select.")
            self.focus_books_list()
            return "break"
        index = self.current_book_index()
        self.marked_book_ids.update(self.book_list_ids)
        self.refresh_books()
        self.root.after(75, lambda: self.settle_book_list_focus(index))
        count = len(self.book_list_ids)
        self.status_var.set(f"Selected all {count} shown book{'s' if count != 1 else ''}.")
        return "break"

    def deselect_all_books(self, event=None):
        if not self.marked_book_ids:
            self.status_var.set("No books are selected for batch actions.")
            self.focus_books_list()
            return "break"
        index = self.current_book_index()
        self.marked_book_ids.clear()
        self.refresh_books()
        self.root.after(75, lambda: self.settle_book_list_focus(index))
        self.status_var.set("All books deselected.")
        return "break"

    def on_book_list_select(self, event=None):
        if event is not None and event.widget in self.multi_book_lists:
            index = self.global_index_from_multi_widget(event.widget)
            if index is not None:
                self.select_book_list_index(index, focus=False, focus_widget=event.widget)
        self.root.after(1, self.sound_for_current_selection)
        book_id = self.selected_book_id_quiet()
        if book_id is not None:
            self.schedule_cover_preview_update(book_id)
        return None

    def on_book_list_keypress(self, event):
        char = (event.char or "").casefold()
        if not char or len(char) != 1 or not char.isalpha():
            return None
        if event.state & 0x0004 or event.state & 0x0008:
            return None
        if self.book_list.size() == 0:
            return "break"

        current = self.current_book_index()
        if current is None:
            current = -1

        total = self.book_list.size()
        for offset in range(1, total + 1):
            index = (current + offset) % total
            title_key = self.book_list_titles[index] if index < len(self.book_list_titles) else ""
            if title_key.startswith(char):
                self.select_book_list_index(index)
                return "break"

        self.status_var.set(f"No book title starts with {char}.")
        return "break"

    def on_book_list_alt_number(self, event):
        digit = event.keysym or event.char or ""
        if digit == "0":
            field_index = 9
        elif digit.isdigit():
            field_index = int(digit) - 1
        else:
            return "break"

        if field_index < 0 or field_index >= len(BOOK_LIST_SHORTCUT_FIELDS):
            return "break"

        book_id = self.selected_book_id_quiet()
        if book_id is None:
            return "break"

        row = self.db.get_book(book_id)
        if not row:
            messagebox.showerror("Book not found", "The selected book was not found.")
            return "break"

        field_key = BOOK_LIST_SHORTCUT_FIELDS[field_index]
        values = self.book_field_values_for_shortcuts(row)
        label, value = values[field_key]
        now = time.monotonic()
        is_quick_repeat = (
            digit == self.last_alt_number_key
            and now - self.last_alt_number_time <= 0.7
        )
        self.last_alt_number_key = digit
        self.last_alt_number_time = now

        self.speak_text(f"{label}: {value}")
        editable_fields = {key for key, _label, _hint in AccessibleMetadataFormDialog.FIELDS}
        if is_quick_repeat:
            if field_key in editable_fields:
                self.edit_book(initial_focus_field=field_key)
            else:
                self.speak_text(f"{label} is read only. {label}: {value}")
        return "break"

    def sound_for_current_selection(self):
        if self.book_list.size() == 0:
            return
        if self.suppress_next_book_list_select_announcement:
            self.suppress_next_book_list_select_announcement = False
            return
        index = self.current_book_index_quiet()
        if index is None:
            return
        if index < 0 or index >= len(self.book_list_ids):
            return
        self.play_missing_metadata_sound_if_needed(self.book_list_ids[index])
        self.announce_current_book_to_nvda(index)

    def clear_search(self):
        self.search_var.set("")
        self.refresh_books()
        self.root.after(50, self.settle_book_list_focus)
        self.root.after(200, self.settle_book_list_focus)
        self.status_var.set("Search cleared. Full book list shown.")

    def clear_search_from_keyboard(self, event=None):
        if not self.search_var.get().strip():
            return None
        self.clear_search()
        return "break"

    def explain_search(self):
        messagebox.showinfo(
            "Search Library",
            "Search looks through both library metadata and the full text content of your books.\n\n"
            "Metadata searched: title, author, series, series number, source, tags, format, notes, edition, year, ISBN, and publisher.\n\n"
            "Book content is indexed in the background when you add books. "
            "If a book was added before content indexing was available, use Search, Re-index Book Content to index your existing library."
        )

    def search_and_focus(self):
        self.refresh_books()
        if self.book_list.size() > 0:
            self.root.after(50, self.settle_book_list_focus)
            self.root.after(200, self.settle_book_list_focus)
        else:
            target = self.first_visible_book_list_widget()
            if target is not None:
                target.focus_force()
            self.status_var.set("No books matched the search.")

    def refresh_books(self, selected_book_id=None):
        self.temporary_column_announcement_item = None
        self.book_list.delete(0, END)
        for listbox in self.multi_book_lists:
            listbox.delete(0, END)
        self.book_list_ids = []
        self.book_list_titles = []
        self.book_list_announcement_texts = []
        self.multi_book_list_index_maps = [[] for _listbox in self.multi_book_lists]

        query = self.search_var.get()
        content_ids = self.db.search_content(query) if query.strip() else None
        rows = self.db.search_books(
            query,
            sort_by=self.sort_by,
            source_filter=self.filter_source,
            tag_filter=self.filter_tag,
            format_filter=self.filter_format,
            extra_ids=content_ids,
        )
        if self.active_saved_view == "missing_metadata":
            rows = [row for row in rows if self.row_missing_metadata_for_view(row)]
        for row in rows:
            book_id = row[0]
            row_text = self.format_book_list_row(row)
            title_key = self.navigation_title_key(row[1])
            global_index = len(self.book_list_ids)
            self.book_list.insert(END, row_text)
            self.book_list_ids.append(book_id)
            self.book_list_titles.append(title_key)
            self.book_list_announcement_texts.append(row_text)
            if self.multi_book_lists:
                group_index = self.book_list_group_index(title_key)
                self.multi_book_lists[group_index].insert(END, row_text)
                self.multi_book_list_index_maps[group_index].append(global_index)

        count = len(rows)
        self.status_var.set(
            f"{count} book{'s' if count != 1 else ''} shown. "
            f"Sorted by {self.sort_label()}. {self.active_filter_summary()}."
        )

        if count:
            selected_index = 0
            if selected_book_id is not None:
                try:
                    selected_index = self.book_list_ids.index(int(selected_book_id))
                except ValueError:
                    selected_index = 0

            self.select_book_list_index(selected_index)
        else:
            self.show_active_multi_book_column()
            self.clear_cover_preview()

    def clear_cover_preview(self, text="No cover"):
        if not hasattr(self, "cover_preview_label"):
            return
        self.cover_preview_book_id = None
        self.cover_preview_photo = None
        self.cover_preview_label.configure(image="", text=text)
        self.cover_preview_title_var.set("")

    def schedule_cover_preview_update(self, book_id):
        if not hasattr(self, "cover_preview_label"):
            return
        if self.cover_preview_after is not None:
            try:
                self.root.after_cancel(self.cover_preview_after)
            except Exception:
                pass
        self.cover_preview_after = self.root.after(150, lambda: self.update_cover_preview(book_id))

    def update_cover_preview(self, book_id):
        self.cover_preview_after = None
        self.cover_preview_book_id = book_id
        row = self.db.get_book(book_id)
        if not row:
            self.clear_cover_preview()
            return

        title = row[1] or "Selected book"
        cover_url = row[14] if len(row) > 14 else ""
        self.cover_preview_title_var.set(title)
        if not cover_url:
            self.cover_preview_photo = None
            self.cover_preview_label.configure(image="", text="No cover saved")
            return

        cached = self.cover_image_cache.get(cover_url)
        if cached is not None:
            self.cover_preview_photo = cached
            self.cover_preview_label.configure(image=cached, text="")
            return

        self.cover_preview_photo = None
        self.cover_preview_label.configure(image="", text="Loading cover")

        def worker(expected_book_id, url):
            data = None
            error = None
            try:
                request = urllib.request.Request(url, headers={"User-Agent": APP_NAME})
                with urllib.request.urlopen(request, timeout=8) as response:
                    data = response.read(2_000_000)
            except Exception as exc:
                error = exc
            self.root.after(0, lambda: self.finish_cover_preview_load(expected_book_id, url, data, error))

        threading.Thread(target=worker, args=(book_id, cover_url), daemon=True).start()

    def finish_cover_preview_load(self, expected_book_id, cover_url, data, error):
        if self.cover_preview_book_id != expected_book_id:
            return
        photo = None
        if data is not None:
            try:
                from PIL import Image, ImageTk
                image = Image.open(BytesIO(data))
                image.thumbnail((170, 240))
                photo = ImageTk.PhotoImage(image)
            except Exception:
                photo = None
        if photo is None:
            self.cover_preview_photo = None
            self.cover_preview_label.configure(image="", text="Cover available")
            return
        self.cover_image_cache[cover_url] = photo
        self.cover_preview_photo = photo
        self.cover_preview_label.configure(image=photo, text="")

    def current_book_index(self):
        """Return the active list index, and make it the selection."""
        if self.book_list.size() == 0:
            return None

        if self.is_multiple_book_list_mode():
            focused_listbox = self.focused_multi_book_list()
            index = self.global_index_from_multi_widget(focused_listbox) if focused_listbox is not None else None
            if index is None:
                index = self.current_book_index_quiet()
            if index is None:
                index = 0
            index = max(0, min(index, self.book_list.size() - 1))
            self.select_book_list_index(index, focus=False, focus_widget=focused_listbox)
            return index

        try:
            index = self.book_list.index("active")
        except Exception:
            selected = self.book_list.curselection()
            index = selected[0] if selected else 0

        if index < 0 or index >= self.book_list.size():
            index = 0

        self.book_list.selection_clear(0, END)
        self.book_list.selection_set(index)
        self.book_list.activate(index)
        self.book_list.see(index)
        return index

    def current_book_index_quiet(self):
        if self.book_list.size() == 0:
            return None

        if self.is_multiple_book_list_mode():
            focused_listbox = self.focused_multi_book_list()
            if focused_listbox is not None:
                index = self.global_index_from_multi_widget(focused_listbox)
                if index is not None:
                    return index

        selected = self.book_list.curselection()
        if selected:
            index = selected[0]
        else:
            try:
                index = self.book_list.index("active")
            except Exception:
                index = 0

        if index < 0 or index >= self.book_list.size():
            return 0
        return index

    def selected_book_id_quiet(self):
        index = self.current_book_index_quiet()
        if index is None or index >= len(self.book_list_ids):
            return None
        return int(self.book_list_ids[index])

    def selected_book_id(self):
        if self.book_list.size() == 0:
            messagebox.showinfo("No books", "There are no books in the current list.")
            return None

        index = self.current_book_index()
        if index is None or index >= len(self.book_list_ids):
            messagebox.showinfo("No selection", "Please select a book first.")
            return None

        return int(self.book_list_ids[index])

    def selected_book_ids(self):
        visible_marked = [book_id for book_id in self.book_list_ids if book_id in self.marked_book_ids]
        if visible_marked:
            return visible_marked
        book_id = self.selected_book_id()
        return [book_id] if book_id is not None else []

    def current_book_list_text(self):
        if self.book_list.size() == 0:
            return ""
        index = self.current_book_index()
        if index is None or index >= self.book_list.size():
            return ""
        return self.book_list.get(index)

    def show_selected_book_info(self):
        if self.book_list.size() == 0:
            messagebox.showinfo("No books", "There are no books in the current list.")
            return

        index = self.current_book_index()
        if index is None:
            messagebox.showinfo("No selection", "Please select a book first.")
            return

        book_id = self.selected_book_id()
        row = self.db.get_book(book_id) if book_id is not None else None
        if not row:
            text = self.book_list.get(index)
        else:
            text = (
                f"Title: {row[1]}\n"
                f"Author: {row[2] or 'Unknown'}\n"
                f"Series: {row[21] if len(row) > 21 and row[21] else 'Not specified'}\n"
                f"Series number: {row[22] if len(row) > 22 and row[22] else 'Not specified'}\n"
                f"Edition: {row[10] or 'Not specified'}\n"
                f"Year: {row[11] or 'Not specified'}\n"
                f"ISBN: {row[12] or 'Not specified'}\n"
                f"Publisher: {row[13] or 'Not specified'}\n"
                f"Cover image: {'Available' if len(row) > 14 and row[14] else 'Not saved'}\n"
                f"EPUB accessibility:\n{self.accessibility_text_from_row(row)}\n"
                f"Source: {row[3] or 'Not specified'}\n"
                f"Tags: {row[4] or 'None'}\n"
                f"Format: {row[6] or 'Unknown'}\n"
                f"Stored path: {row[8]}"
            )
        self.status_var.set(text.replace("\n", " "))
        messagebox.showinfo("Selected book information", text)

    def crash_log_path(self):
        return self.db.folder / "crash_log.txt"

    def log_error(self, context, exc):
        logger.exception(context)

    def safe_message_error(self, title, message):
        try:
            messagebox.showerror(title, message)
        except Exception:
            pass

    def update_accessibility_from_epub(self, book_id, stored_path, database=None):
        database = database or self.db
        path = Path(stored_path)
        if path.suffix.lower() != ".epub" or not path.exists():
            return {}
        metadata = read_epub_accessibility_metadata(path)
        database.update_accessibility_metadata(book_id, metadata)
        return metadata

    def accessibility_text_from_row(self, row):
        if not row or len(row) <= 20:
            return "Not checked"
        labels = [
            ("Accessibility summary", row[15]),
            ("Accessibility features", row[16]),
            ("Accessibility hazards", row[17]),
            ("Access modes", row[18]),
            ("Access modes sufficient", row[19]),
            ("Certified by", row[20]),
        ]
        parts = [f"{label}: {value}" for label, value in labels if value]
        return "\n".join(parts) if parts else "Not checked"

    def extract_supported_files_from_zip(self, zip_path: Path) -> list[Path]:
        """Extract a ZIP file and return supported ebook/document files inside it.

        Bookshare ZIP files often contain an EPUB. EPUB files are imported as
        whole files, which preserves images, navigation, and structure.
        """
        import uuid
        extract_root = self.db.folder / "Extracted_Zips" / f"{safe_filename(zip_path.stem)}_{uuid.uuid4().hex[:8]}"
        extract_root.mkdir(parents=True, exist_ok=True)

        try:
            with zipfile.ZipFile(zip_path, "r") as archive:
                for member in archive.infolist():
                    member_path = Path(member.filename)
                    if member.is_dir():
                        continue
                    if member_path.is_absolute() or ".." in member_path.parts:
                        continue
                    archive.extract(member, extract_root)
        except zipfile.BadZipFile:
            raise OSError("ZIP file is damaged or not a valid ZIP file")
        except Exception as exc:
            raise OSError(f"Could not extract ZIP file: {exc}")

        extracted_files = [path for path in extract_root.rglob("*") if path.is_file()]
        supported_inside = [
            path for path in extracted_files
            if path.suffix.lower() in SUPPORTED_EXTENSIONS and path.suffix.lower() != ".zip"
            and not self.is_ignored_import_file(path)
        ]

        epubs = [path for path in supported_inside if path.suffix.lower() == ".epub"]
        if epubs:
            return epubs

        return supported_inside

    def import_zip_file_without_prompt(self, zip_path: Path, default_source: str = "", database=None) -> tuple[int, list[str]]:
        database = database or self.db
        imported = 0
        skipped = []

        ready, reason = self.is_file_ready_for_import(zip_path)
        if not ready:
            return 0, [f"{zip_path} -- {reason}"]

        try:
            inner_files = self.extract_supported_files_from_zip(zip_path)
        except Exception as exc:
            return 0, [f"{zip_path} -- {exc}"]

        if not inner_files:
            return 0, [f"{zip_path} -- ZIP contained no supported ebook files"]

        for inner_file in inner_files:
            try:
                if self.import_one_book_without_prompt(inner_file, default_source=default_source, database=database):
                    imported += 1
                else:
                    skipped.append(f"{inner_file} -- unsupported or not imported")
            except sqlite3.IntegrityError:
                skipped.append(f"{inner_file} -- already imported or duplicate stored path")
            except Exception as exc:
                skipped.append(f"{inner_file} -- {exc}")

        return imported, skipped

    def guess_metadata_from_file(self, source_path: Path, database=None) -> dict:
        database = database or getattr(self, "db", None)
        default_title = clean_filename_title(source_path)
        metadata = {"title": default_title, "author": "", "series": "", "series_number": "", "edition": "", "year": "", "isbn": "", "publisher": "", "source": "", "tags": "", "notes": ""}

        if source_path.suffix.lower() == ".epub":
            epub_metadata = read_epub_metadata(source_path)
            metadata.update({key: value for key, value in epub_metadata.items() if value})

        if source_path.suffix.lower() in CALIBRE_METADATA_EXTENSIONS and self.calibre_metadata_reading_enabled(database):
            calibre_metadata = read_calibre_metadata(source_path)
            metadata.update({key: value for key, value in calibre_metadata.items() if value})

        metadata["tags"] = normalize_import_tags(metadata["tags"])
        return detect_metadata_from_text(source_path, existing=metadata)

    def is_file_ready_for_import(self, source_path: Path) -> tuple[bool, str]:
        """Return whether a file can be read now.

        This helps with iCloud, OneDrive, Dropbox, Google Drive, and other cloud-sync folders
        where a file may appear in Windows Explorer but not actually be
        downloaded locally yet.
        """
        try:
            if not source_path.exists():
                return False, "file does not exist"

            if not source_path.is_file():
                return False, "not a file"

            if source_path.suffix.lower() not in SUPPORTED_EXTENSIONS:
                return False, "unsupported file type"

            size = source_path.stat().st_size
            if size == 0:
                return False, "empty file or cloud placeholder"

            with open(source_path, "rb") as test_file:
                test_file.read(1)

            return True, ""
        except PermissionError:
            return False, "permission denied or cloud file not downloaded"
        except OSError as exc:
            return False, f"not readable, possibly not downloaded: {exc}"
        except Exception as exc:
            return False, f"not readable: {exc}"

    def write_import_report(self, imported: int, skipped_items: list[str]) -> Path:
        report_path = self.db.folder / "last_import_report.txt"
        lines = [
            "Accessible Ebook Library Manager import report",
            "",
            f"Imported: {imported}",
            f"Skipped: {len(skipped_items)}",
            "",
        ]

        if skipped_items:
            lines.append("Skipped files:")
            lines.extend(skipped_items)
            lines.append("")
            lines.append("Tip: For iCloud Drive, OneDrive, Dropbox, or Google Drive, select the files or folder in File Explorer and choose the option that keeps them available offline, such as Always keep on this device or Available offline, then try importing again.")

        report_path.write_text("\n".join(lines), encoding="utf-8")
        return report_path

    def command_line_import_paths(self):
        args = sys.argv[1:]
        paths = []
        index = 0
        while index < len(args):
            arg = args[index]
            if arg == "--import":
                index += 1
                if index < len(args):
                    paths.append(Path(args[index]))
            elif arg.startswith("--import="):
                paths.append(Path(arg.split("=", 1)[1]))
            elif not arg.startswith("-"):
                candidate = Path(arg)
                if candidate.suffix.lower() in SUPPORTED_EXTENSIONS:
                    paths.append(candidate)
            index += 1
        return paths

    def import_command_line_files(self):
        paths = self.command_line_import_paths()
        if not paths:
            return
        self.start_folder_import(paths, "File Explorer", [])

    def import_one_book_without_prompt(self, source_path: Path, default_source: str = "", database=None) -> bool:
        database = database or self.db
        if self.is_ignored_import_file(source_path):
            raise OSError("ignored helper file")

        if source_path.suffix.lower() == ".zip":
            imported, skipped = self.import_zip_file_without_prompt(source_path, default_source=default_source, database=database)
            if imported:
                return True
            raise OSError("; ".join(skipped) if skipped else "ZIP contained no importable books")

        ready, reason = self.is_file_ready_for_import(source_path)
        if not ready:
            raise OSError(reason)

        metadata = self.guess_metadata_from_file(source_path, database=database)
        if default_source and not metadata.get("source"):
            metadata["source"] = default_source

        filename = safe_filename(f"{metadata['author']} - {metadata['title']}").strip(" -")
        destination = self.unique_destination(filename, source_path.suffix.lower())
        shutil.copy2(source_path, destination)

        if destination.suffix.lower() == ".epub":
            try:
                write_epub_metadata(
                    destination,
                    metadata["title"],
                    metadata["author"],
                    metadata["source"],
                    metadata["tags"],
                    metadata["notes"],
                )
            except Exception:
                # Folder import should not stop because one EPUB metadata write failed.
                pass

        new_book_id = database.add_book(
            metadata["title"],
            metadata["author"],
            metadata["source"],
            metadata["tags"],
            metadata["notes"],
            self.canonical_path_text(source_path),
            str(destination),
        )
        database.update_extra_fields(
            new_book_id,
            metadata.get("edition", ""),
            metadata.get("year", ""),
            metadata.get("isbn", ""),
            metadata.get("publisher", ""),
        )
        database.update_series_fields(new_book_id, metadata.get("series", ""), metadata.get("series_number", ""))
        self.update_accessibility_from_epub(new_book_id, destination, database=database)
        self.queue_book_for_indexing(new_book_id, str(destination))
        return True

    def import_folder(self):
        try:
            folder = filedialog.askdirectory(title="Choose folder of books to import")
            if not folder:
                return

            include_subfolders = messagebox.askyesno(
                "Include subfolders",
                "Import books from subfolders too?"
            )

            default_source = AccessibleSingleFieldDialog.ask(
                self.root,
                "Source",
                "Optional: enter a source for these books, such as Bookshare, Personal, Kindle, or leave blank.",
                "",
            )
            if default_source is None:
                default_source = ""

            root_folder = Path(folder)
            self.start_folder_discovery(root_folder, include_subfolders, default_source.strip())
        except Exception as exc:
            self.log_error("Fatal folder import error", exc)
            messagebox.showerror(
                "Folder import crashed",
                f"Folder import hit an unexpected error, but the app caught it this time.\n\n{exc}\n\n"
                f"Crash details saved at:\n{self.crash_log_path()}"
            )

    def start_folder_discovery(self, root_folder, include_subfolders, default_source):
        if self.folder_discovery_running or self.folder_import_running:
            return
        self.folder_discovery_running = True
        progress_window = tk.Toplevel(self.root)
        progress_window.title("Searching for Books")
        progress_window.transient(self.root)
        progress_window.resizable(False, False)
        progress_window.protocol("WM_DELETE_WINDOW", lambda: None)
        message_var = StringVar(value="Searching the selected folder.")
        ttk.Label(progress_window, textvariable=message_var, padding=12).pack(fill=X)
        progress = ttk.Progressbar(progress_window, mode="indeterminate", length=360)
        progress.pack(fill=X, padx=12, pady=(0, 12))
        progress.start(12)
        events = queue.Queue()

        def worker():
            supported = []
            skipped = []
            try:
                candidates = root_folder.rglob("*") if include_subfolders else root_folder.iterdir()
                for candidate in candidates:
                    try:
                        if not candidate.is_file():
                            continue
                        if self.is_ignored_import_file(candidate):
                            skipped.append(f"{candidate} -- ignored helper file")
                        elif candidate.suffix.lower() in SUPPORTED_EXTENSIONS:
                            supported.append(candidate)
                    except Exception as exc:
                        skipped.append(f"{candidate} -- could not inspect file: {exc}")
                events.put(("done", supported, skipped, ""))
            except Exception as exc:
                logger.exception("Listing import folder failed: %s", root_folder)
                events.put(("done", supported, skipped, str(exc)))

        def poll():
            try:
                _kind, supported, skipped, error = events.get_nowait()
            except queue.Empty:
                self.root.after(100, poll)
                return
            self.folder_discovery_running = False
            progress.stop()
            progress_window.destroy()
            if error:
                messagebox.showerror(
                    "Folder import failed",
                    f"I could not read this folder.\n\n{error}\n\nCrash details saved at:\n{self.crash_log_path()}"
                )
                return
            if not supported:
                report_path = self.write_import_report(0, skipped)
                messagebox.showinfo(
                    "No supported books found",
                    f"No supported ebook or document files were found in that folder.\n\nReport saved at:\n{report_path}"
                )
                return
            if len(supported) > 1 and messagebox.askyesno(
                "One book in multiple parts?",
                "Does this folder contain one book split into multiple part files?\n\n"
                "Choose Yes to combine the files into one EPUB and add one book to the library. "
                "Choose No to import each file as a separate book."
            ):
                self.start_combined_folder_import(
                    root_folder, sorted(supported, key=book_part_sort_key), default_source, skipped
                )
                return
            if messagebox.askyesno(
                "Confirm folder import",
                f"Found {len(supported)} supported file{'s' if len(supported) != 1 else ''}. Import them now?"
            ):
                self.start_folder_import(supported, default_source, skipped)

        self.status_var.set("Searching the selected folder for books.")
        threading.Thread(target=worker, daemon=True, name="folder-discovery").start()
        self.root.after(100, poll)

    def start_folder_import(self, paths, default_source, skipped_items):
        if self.folder_import_running:
            return
        self.folder_import_running = True
        progress_window = tk.Toplevel(self.root)
        progress_window.title("Importing Books")
        progress_window.transient(self.root)
        progress_window.resizable(False, False)
        progress_window.protocol("WM_DELETE_WINDOW", lambda: None)
        message_var = StringVar(value="Preparing folder import.")
        ttk.Label(progress_window, textvariable=message_var, padding=12).pack(fill=X)
        progress = ttk.Progressbar(progress_window, mode="determinate", maximum=max(1, len(paths)), length=360)
        progress.pack(fill=X, padx=12, pady=(0, 12))
        events = queue.Queue()

        def worker():
            imported = 0
            skipped = list(skipped_items)
            database = None
            try:
                database = LibraryDatabase()
                for index, path in enumerate(paths, start=1):
                    events.put(("progress", index, len(paths), path.name))
                    try:
                        if path.suffix.lower() == ".zip":
                            zip_imported, zip_skipped = self.import_zip_file_without_prompt(
                                path, default_source=default_source, database=database
                            )
                            imported += zip_imported
                            skipped.extend(zip_skipped)
                        elif self.import_one_book_without_prompt(path, default_source=default_source, database=database):
                            imported += 1
                        else:
                            skipped.append(f"{path} -- unsupported or not imported")
                    except sqlite3.IntegrityError as exc:
                        skipped.append(f"{path} -- already imported or duplicate stored path")
                        logger.warning("Duplicate during folder import: %s", path, exc_info=exc)
                    except Exception as exc:
                        skipped.append(f"{path} -- {exc}")
                        logger.exception("Importing file failed: %s", path)
                report_path = self.write_import_report(imported, skipped)
                events.put(("done", imported, skipped, report_path, ""))
            except Exception as exc:
                logger.exception("Background folder import failed")
                events.put(("done", imported, skipped, "", str(exc)))
            finally:
                if database is not None:
                    database.close()

        def poll():
            try:
                while True:
                    event = events.get_nowait()
                    if event[0] == "progress":
                        _kind, index, total, name = event
                        progress.configure(value=index)
                        message = f"Importing {index} of {total}: {name}"
                        message_var.set(message)
                        self.status_var.set(message)
                    else:
                        _kind, imported, skipped, report_path, error = event
                        self.folder_import_running = False
                        progress_window.destroy()
                        if error:
                            messagebox.showerror("Folder import failed", f"{error}\n\nCrash details saved at:\n{self.crash_log_path()}")
                            return
                        self.refresh_books()
                        if imported:
                            self.focus_books_list()
                        count = len(skipped)
                        messagebox.showinfo(
                            "Folder import complete",
                            f"Imported {imported} book{'s' if imported != 1 else ''}. "
                            f"Skipped {count} file{'s' if count != 1 else ''}.\n\n"
                            f"Import report saved at:\n{report_path}",
                        )
                        self.status_var.set(f"Folder import complete. Imported {imported}. Skipped {count}.")
                        return
            except queue.Empty:
                pass
            if self.folder_import_running:
                self.root.after(100, poll)

        threading.Thread(target=worker, daemon=True, name="folder-import").start()
        self.root.after(100, poll)

    def start_combined_folder_import(self, root_folder: Path, parts: list[Path], default_source: str, skipped_items: list[str]):
        progress_window = tk.Toplevel(self.root)
        progress_window.title("Combining Book Parts")
        progress_window.transient(self.root)
        progress_window.resizable(False, False)
        progress_window.protocol("WM_DELETE_WINDOW", lambda: None)

        message_var = StringVar(value="Preparing to read book parts.")
        ttk.Label(progress_window, textvariable=message_var, padding=12).pack(fill=X)
        progress = ttk.Progressbar(progress_window, mode="indeterminate", length=320)
        progress.pack(fill=X, padx=12, pady=(0, 12))
        progress.start(12)

        work_queue = queue.Queue()
        total = len(parts)
        self.status_var.set(f"Combining {total} book part{'s' if total != 1 else ''}.")

        def worker():
            local_skipped = list(skipped_items)
            readable_parts = []
            combined_sample = []
            try:
                for index, part in enumerate(parts, start=1):
                    work_queue.put(("progress", index, total, part.name))
                    if part.suffix.lower() == ".zip":
                        local_skipped.append(f"{part} -- ZIP files cannot be combined into one EPUB")
                        continue
                    text = read_text_for_metadata_detection(part, max_chars=250000)
                    if not text.strip():
                        local_skipped.append(f"{part} -- no readable text found for combined EPUB")
                        continue
                    readable_parts.append((part, text))
                    if len(" ".join(combined_sample)) < 50000:
                        combined_sample.append(text[:10000])
                work_queue.put(("done", readable_parts, combined_sample, local_skipped, None))
            except Exception as exc:
                work_queue.put(("done", [], [], local_skipped, exc))

        def poll_worker():
            try:
                while True:
                    item = work_queue.get_nowait()
                    if item[0] == "progress":
                        _kind, index, count, name = item
                        message = f"Reading part {index} of {count}: {name}"
                        message_var.set(message)
                        self.status_var.set(message)
                    elif item[0] == "done":
                        _kind, readable_parts, combined_sample, final_skipped, error = item
                        progress.stop()
                        progress_window.destroy()
                        if error:
                            self.log_error("Reading combined folder parts", error)
                            report_path = self.write_import_report(0, final_skipped)
                            messagebox.showerror(
                                "Combined EPUB import failed",
                                f"I could not finish reading the book parts.\n\n{error}\n\n"
                                f"Import report saved at:\n{report_path}"
                            )
                            self.status_var.set("Combined EPUB import failed.")
                            return
                        self.finish_combined_folder_import(
                            root_folder,
                            readable_parts,
                            combined_sample,
                            default_source,
                            final_skipped,
                        )
                        return
            except queue.Empty:
                pass
            self.root.after(200, poll_worker)

        threading.Thread(target=worker, daemon=True).start()
        poll_worker()

    def finish_combined_folder_import(self, root_folder: Path, readable_parts, combined_sample, default_source: str, skipped_items: list[str]):
        imported = self.finish_combined_epub_import(root_folder, readable_parts, combined_sample, default_source, skipped_items)
        report_path = self.write_import_report(imported, skipped_items)
        self.refresh_books()
        if imported:
            self.focus_books_list()
        messagebox.showinfo(
            "Combined EPUB import complete",
            f"Imported {imported} combined EPUB book{'s' if imported != 1 else ''}. "
            f"Skipped {len(skipped_items)} file{'s' if len(skipped_items) != 1 else ''}.\n\n"
            f"Import report saved at:\n{report_path}"
        )
        self.status_var.set(f"Combined EPUB import complete. Imported {imported}. Skipped {len(skipped_items)}.")

    def import_folder_as_combined_epub(self, root_folder: Path, parts: list[Path], default_source: str, skipped_items: list[str]):
        readable_parts = []
        combined_sample = []
        for part in parts:
            if part.suffix.lower() == ".zip":
                skipped_items.append(f"{part} -- ZIP files cannot be combined into one EPUB")
                continue
            text = read_text_for_metadata_detection(part, max_chars=250000)
            if not text.strip():
                skipped_items.append(f"{part} -- no readable text found for combined EPUB")
                continue
            readable_parts.append((part, text))
            if len(" ".join(combined_sample)) < 50000:
                combined_sample.append(text[:10000])

        return self.finish_combined_epub_import(root_folder, readable_parts, combined_sample, default_source, skipped_items)

    def finish_combined_epub_import(self, root_folder: Path, readable_parts, combined_sample, default_source: str, skipped_items: list[str]):
        if not readable_parts:
            messagebox.showerror(
                "Could not combine folder",
                "I could not find readable text in the selected folder parts."
            )
            return 0

        initial_metadata = {
            "title": clean_filename_title(root_folder),
            "author": "",
            "series": "",
            "series_number": "",
            "edition": "",
            "year": "",
            "isbn": "",
            "publisher": "",
            "source": default_source,
            "tags": "",
            "notes": "",
        }
        detected = detect_metadata_from_text_content("\n".join(combined_sample), existing=initial_metadata)
        initial_metadata.update({key: value for key, value in detected.items() if value})

        metadata = TkMetadataDialog.ask(self.root, "Combined Book Metadata", initial_metadata)
        if not metadata:
            return 0

        output_name = safe_filename(f"{metadata['author']} - {metadata['title']}").strip(" -") or safe_filename(metadata["title"])
        destination = self.unique_destination(output_name, ".epub")
        self.create_combined_epub(destination, readable_parts, metadata, root_folder=root_folder)

        new_book_id = self.db.add_book(
            metadata["title"],
            metadata["author"],
            metadata["source"],
            metadata["tags"],
            metadata["notes"],
            self.canonical_path_text(root_folder),
            str(destination),
        )
        self.db.update_extra_fields(
            new_book_id,
            metadata.get("edition", ""),
            metadata.get("year", ""),
            metadata.get("isbn", ""),
            metadata.get("publisher", ""),
        )
        self.db.update_series_fields(new_book_id, metadata.get("series", ""), metadata.get("series_number", ""))
        self.update_accessibility_from_epub(new_book_id, destination)
        self.queue_book_for_indexing(new_book_id, str(destination))
        return 1

    def combine_selected_books_to_epub(self):
        book_ids = self.selected_book_ids()
        if len(book_ids) < 2:
            messagebox.showinfo(
                "Select Multiple Books",
                "Select at least two book parts first. Use Control+Space on each book part, then run this command again."
            )

            return

        rows = [self.db.get_book(book_id) for book_id in book_ids]
        rows = [row for row in rows if row]
        if len(rows) < 2:
            messagebox.showinfo("Select Multiple Books", "I could not find at least two selected books to combine.")
            return

        if not messagebox.askyesno(
            "Combine Selected Books",
            f"Combine {len(rows)} selected books into one EPUB using the current list order?\n\n"
            "The original library entries will be kept unless you choose to remove them after the combined EPUB is created."
        ):
            return

        readable_parts = []
        combined_sample = []
        skipped = []
        for row in rows:
            stored_path = Path(row[8])
            if not stored_path.exists():
                skipped.append(f"{row[1]} -- stored file not found")
                continue
            if stored_path.suffix.lower() == ".zip":
                skipped.append(f"{row[1]} -- ZIP files cannot be combined into one EPUB")
                continue
            text = read_text_for_metadata_detection(stored_path, max_chars=250000)
            if not text.strip():
                skipped.append(f"{row[1]} -- no readable text found")
                continue
            readable_parts.append((stored_path, text))
            if len(" ".join(combined_sample)) < 50000:
                combined_sample.append(text[:10000])

        if len(readable_parts) < 2:
            report_path = self.write_import_report(0, skipped)
            messagebox.showerror(
                "Could Not Combine Books",
                f"I found readable text in fewer than two selected books.\n\nReport saved at:\n{report_path}"
            )
            return

        authors = {row[2].strip() for row in rows if row[2].strip()}
        sources = {row[3].strip() for row in rows if row[3].strip()}
        first_row = rows[0]
        initial_metadata = {
            "title": first_row[1] or "Combined Book",
            "author": next(iter(authors)) if len(authors) == 1 else "",
            "series": first_row[21] if len(first_row) > 21 else "",
            "series_number": first_row[22] if len(first_row) > 22 else "",
            "edition": "",
            "year": "",
            "isbn": "",
            "publisher": "",
            "source": next(iter(sources)) if len(sources) == 1 else "",
            "tags": "",
            "notes": "Combined from selected library books.",
        }
        detected = detect_metadata_from_text_content("\n".join(combined_sample), existing=initial_metadata)
        initial_metadata.update({key: value for key, value in detected.items() if value})

        metadata = TkMetadataDialog.ask(self.root, "Combined Book Metadata", initial_metadata)
        if not metadata:
            return

        output_name = safe_filename(f"{metadata['author']} - {metadata['title']}").strip(" -") or safe_filename(metadata["title"])
        destination = self.unique_destination(output_name, ".epub")
        self.create_combined_epub(destination, readable_parts, metadata, root_folder=None)

        original_note = "Combined selected library books: " + ", ".join(str(row[0]) for row in rows)
        new_book_id = self.db.add_book(
            metadata["title"],
            metadata["author"],
            metadata["source"],
            metadata["tags"],
            metadata["notes"],
            original_note,
            str(destination),
        )
        self.db.update_extra_fields(
            new_book_id,
            metadata.get("edition", ""),
            metadata.get("year", ""),
            metadata.get("isbn", ""),
            metadata.get("publisher", ""),
        )
        self.db.update_series_fields(new_book_id, metadata.get("series", ""), metadata.get("series_number", ""))
        self.update_accessibility_from_epub(new_book_id, destination)
        self.queue_book_for_indexing(new_book_id, str(destination))

        removed_parts = 0
        if messagebox.askyesno(
            "Remove Original Part Entries?",
            "The combined EPUB was created. Remove the original selected part entries from the library list?\n\n"
            "Their stored files will not be deleted."
        ):
            for row in rows:
                try:
                    self.db.delete_book(row[0], delete_file=False)
                    self.marked_book_ids.discard(row[0])
                    removed_parts += 1
                except Exception as exc:
                    skipped.append(f"{row[1]} -- could not remove original entry: {exc}")

        self.refresh_books(selected_book_id=new_book_id)
        self.focus_books_list()
        report_path = self.write_import_report(1, skipped) if skipped else None
        message = (
            f"Created one combined EPUB from {len(readable_parts)} selected book parts.\n\n"
            f"Removed {removed_parts} original part entr{'y' if removed_parts == 1 else 'ies'} from the library list."
        )
        if report_path:
            message += f"\n\nReport saved at:\n{report_path}"
        messagebox.showinfo("Combined EPUB Created", message)
        self.status_var.set("Combined selected books into one EPUB.")

    def create_combined_epub(self, destination: Path, readable_parts, metadata, root_folder: Path | None = None):
        book_uuid = f"urn:uuid:{uuid.uuid4()}"
        title = metadata.get("title", "Untitled") or "Untitled"
        author = metadata.get("author", "") or "Unknown"
        publisher = metadata.get("publisher", "") or ""
        year = metadata.get("year", "") or ""
        isbn = metadata.get("isbn", "") or ""
        combined_text = "\n".join(text for _part_path, text in readable_parts)
        language = detect_language_from_text(combined_text, default="en")
        modified = datetime.now(timezone.utc).replace(tzinfo=None).strftime("%Y-%m-%dT%H:%M:%SZ")
        chapters = []
        text_toc_entries = parse_text_toc_entries(combined_text)
        source_note_parts = []
        if root_folder:
            source_note_parts.append(f"Combined source folder: {root_folder}")
        source_note_parts.append("Combined source files:")
        source_note_parts.extend(f"- {part_path.name}" for part_path, _text in readable_parts)
        combined_source_note = "\n".join(source_note_parts)

        for index, (part_path, text) in enumerate(readable_parts, start=1):
            chapter_title = clean_filename_title(part_path) or f"Part {index}"
            body = self.xhtml_body_from_text(text)
            filename = f"chapter{index:04d}.xhtml"
            chapters.append((filename, chapter_title, body))

        container_xml = """<?xml version="1.0" encoding="UTF-8"?>
<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
  <rootfiles>
    <rootfile full-path="EPUB/package.opf" media-type="application/oebps-package+xml"/>
  </rootfiles>
</container>
"""
        page_targets = {}
        for filename, _chapter_title, body in chapters:
            for page_id in re.findall(r'id=["\'](page-[^"\']+)["\']', body, flags=re.IGNORECASE):
                page_targets.setdefault(page_id.casefold(), filename)

        toc_nav_entries = []
        for entry in text_toc_entries:
            target_filename = page_targets.get(entry["page_id"].casefold())
            if target_filename:
                toc_nav_entries.append((target_filename + "#" + entry["page_id"], entry["title"]))

        if toc_nav_entries:
            nav_items = "\n".join(
                f'      <li><a href="{html.escape(href)}">{html.escape(label)}</a></li>'
                for href, label in toc_nav_entries
            )
        else:
            nav_items = "\n".join(
                f'      <li><a href="{filename}">{html.escape(chapter_title)}</a></li>'
                for filename, chapter_title, _body in chapters
            )
        nav_xhtml = f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops" lang="{language}" xml:lang="{language}">
  <head><title>Table of Contents</title></head>
  <body>
    <nav epub:type="toc" id="toc">
      <h1>Table of Contents</h1>
      <ol>
{nav_items}
      </ol>
    </nav>
  </body>
</html>
"""
        manifest_items = ['    <item id="nav" href="nav.xhtml" media-type="application/xhtml+xml" properties="nav"/>']
        spine_items = []
        for index, (filename, _chapter_title, _body) in enumerate(chapters, start=1):
            manifest_items.append(f'    <item id="chapter{index}" href="{filename}" media-type="application/xhtml+xml"/>')
            spine_items.append(f'    <itemref idref="chapter{index}"/>')

        package_opf = f"""<?xml version="1.0" encoding="UTF-8"?>
<package xmlns="http://www.idpf.org/2007/opf" version="3.0" unique-identifier="book-id">
  <metadata xmlns:dc="http://purl.org/dc/elements/1.1/">
    <dc:identifier id="book-id">{html.escape(book_uuid)}</dc:identifier>
    <dc:title>{html.escape(title)}</dc:title>
    <dc:creator>{html.escape(author)}</dc:creator>
    <dc:language>{language}</dc:language>
{f'    <dc:publisher>{html.escape(publisher)}</dc:publisher>' if publisher else ''}
{f'    <dc:date>{html.escape(year)}</dc:date>' if year else ''}
{f'    <dc:identifier>{html.escape(isbn)}</dc:identifier>' if isbn else ''}
    <dc:relation>{html.escape(combined_source_note)}</dc:relation>
    <meta property="schema:accessibilityFeature">tableOfContents</meta>
    <meta property="dcterms:modified">{modified}</meta>
  </metadata>
  <manifest>
{chr(10).join(manifest_items)}
  </manifest>
  <spine>
{chr(10).join(spine_items)}
  </spine>
</package>
"""
        destination.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(destination, "w") as archive:
            archive.writestr("mimetype", "application/epub+zip", compress_type=zipfile.ZIP_STORED)
            archive.writestr("META-INF/container.xml", container_xml)
            archive.writestr("EPUB/package.opf", package_opf)
            archive.writestr("EPUB/nav.xhtml", nav_xhtml)
            for filename, chapter_title, body in chapters:
                chapter_xhtml = f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops" lang="{language}" xml:lang="{language}">
  <head><title>{html.escape(chapter_title)}</title></head>
  <body>
    <h1>{html.escape(chapter_title)}</h1>
{body}
  </body>
</html>
"""
                archive.writestr(f"EPUB/{filename}", chapter_xhtml)

    def xhtml_body_from_text(self, text: str):
        cleaned_lines, _removed = cleaned_lines_for_reflow(text)
        output = []
        paragraph_lines = []
        used_page_ids = set()
        previous_output_was_heading = False

        def flush_paragraph():
            nonlocal previous_output_was_heading
            if not paragraph_lines:
                return
            paragraph = re.sub(r"\s+", " ", " ".join(paragraph_lines)).strip()
            paragraph_lines.clear()
            if paragraph:
                output.append(f"    <p>{html.escape(paragraph)}</p>")
                previous_output_was_heading = False

        for raw_line in cleaned_lines:
            line = raw_line.strip()
            if not line:
                flush_paragraph()
                previous_output_was_heading = False
                continue

            page_label = page_label_from_line(line)
            if page_label:
                flush_paragraph()
                output.append("    " + self.epub_pagebreak_span(page_label, used_page_ids))
                previous_output_was_heading = False
                continue

            if looks_like_text_toc_entry_line(line):
                flush_paragraph()
                output.append(f"    <p>{html.escape(line)}</p>")
                previous_output_was_heading = False
                continue

            if looks_like_standalone_heading_line(line) or (previous_output_was_heading and looks_like_heading_continuation_line(line)):
                flush_paragraph()
                if re.match(r"^(?:chapter|part|section|appendix|index|front matter|\u00a7)", line, flags=re.IGNORECASE) or previous_output_was_heading:
                    output.append(f"    <h2>{html.escape(line)}</h2>")
                    previous_output_was_heading = True
                else:
                    output.append(f"    <p>{html.escape(line)}</p>")
                    previous_output_was_heading = False
                continue

            if line.endswith("?"):
                flush_paragraph()
                output.append(f"    <p>{html.escape(line)}</p>")
                previous_output_was_heading = False
                continue

            paragraph_lines.append(line)

        flush_paragraph()
        if not output:
            output.append("    <p>No text found.</p>")
        return "\n".join(output)

    def epub_pagebreak_span(self, page_label, used_page_ids):
        page_label = str(page_label or "").strip()
        page_id_base = page_id_for_label(page_label)
        page_id = page_id_base
        suffix = 2
        while page_id in used_page_ids:
            page_id = f"{page_id_base}-{suffix}"
            suffix += 1
        used_page_ids.add(page_id)
        return (
            f'<span epub:type="pagebreak" role="doc-pagebreak" id="{html.escape(page_id)}" '
            f'aria-label="Page {html.escape(page_label)}"></span>'
        )

    def ensure_xhtml_epub_namespace(self, xhtml_text):
        if "xmlns:epub=" in xhtml_text:
            return xhtml_text
        return re.sub(
            r"<html\b([^>]*)>",
            r'<html\1 xmlns:epub="http://www.idpf.org/2007/ops">',
            xhtml_text,
            count=1,
            flags=re.IGNORECASE,
        )

    def insert_pagebreaks_in_xhtml_text(self, xhtml_text):
        used_page_ids = set(re.findall(r'id=["\'](page-[^"\']+)["\']', xhtml_text, flags=re.IGNORECASE))
        inserted = 0

        def replace_paragraph_start(match):
            nonlocal inserted
            opening = match.group(1)
            label = match.group(2) or match.group(3)
            rest = match.group(4).strip()
            inserted += 1
            return self.epub_pagebreak_span(label, used_page_ids) + "\n" + opening + rest

        xhtml_text = re.sub(
            r"(<p\b[^>]*>)\s*(?:(?:(?:p?age)\s+)+([0-9ivxlcdm]+)|(\d{1,4}))\s+([\s\S]*?</p>)",
            replace_paragraph_start,
            xhtml_text,
            flags=re.IGNORECASE,
        )

        def replace_standalone(match):
            nonlocal inserted
            label = match.group(1)
            inserted += 1
            return ">" + self.epub_pagebreak_span(label, used_page_ids) + "<"

        xhtml_text = re.sub(
            r">\s*(?:(?:p?age)\s+)+([0-9ivxlcdm]+)\s*<",
            replace_standalone,
            xhtml_text,
            flags=re.IGNORECASE,
        )

        if inserted:
            xhtml_text = self.ensure_xhtml_epub_namespace(xhtml_text)
        return xhtml_text, inserted

    def add_book(self):
        paths = filedialog.askopenfilenames(
            title="Choose books to add",
            filetypes=[
                ("Ebook and document files", "*.epub *.pdf *.docx *.doc *.txt *.rtf *.mobi *.azw *.azw3 *.kfx *.kfx-zip *.prc *.html *.htm *.zip"),
                ("All files", "*.*"),
            ],
        )
        if not paths:
            return

        added = 0
        for path in paths:
            source_path = Path(path)

            if self.is_ignored_import_file(source_path):
                continue

            if source_path.suffix.lower() not in SUPPORTED_EXTENSIONS:
                if not messagebox.askyesno(
                    "Unsupported extension",
                    f"{source_path.name} is not a common ebook extension. Add it anyway?",
                ):
                    continue

            if source_path.suffix.lower() == ".zip":
                zip_imported, zip_skipped = self.import_zip_file_without_prompt(source_path, default_source="")
                added += zip_imported
                if zip_skipped:
                    report_path = self.write_import_report(zip_imported, zip_skipped)
                    messagebox.showwarning(
                        "ZIP import warnings",
                        f"Imported {zip_imported} book{'s' if zip_imported != 1 else ''} from the ZIP. "
                        f"Skipped {len(zip_skipped)} item{'s' if len(zip_skipped) != 1 else ''}.\n\n"
                        f"Report saved at:\n{report_path}"
                    )
                continue

            initial_metadata = self.guess_metadata_from_file(source_path)

            metadata = TkMetadataDialog.ask(self.root, "Add Book Metadata", initial_metadata)
            if not metadata:
                continue

            filename = safe_filename(f"{metadata['author']} - {metadata['title']}").strip(" -")
            destination = self.unique_destination(filename, source_path.suffix.lower())
            shutil.copy2(source_path, destination)

            if destination.suffix.lower() == ".epub":
                try:
                    write_epub_metadata(
                        destination,
                        metadata["title"],
                        metadata["author"],
                        metadata["source"],
                        metadata["tags"],
                        metadata["notes"],
                    )
                except Exception as exc:
                    messagebox.showwarning(
                        "EPUB metadata not written",
                        f"The book was imported, but I could not write metadata into the EPUB file itself.\n\n{exc}"
                    )

            new_book_id = self.db.add_book(
                metadata["title"],
                metadata["author"],
                metadata["source"],
                metadata["tags"],
                metadata["notes"],
                self.canonical_path_text(source_path),
                str(destination),
            )
            self.db.update_extra_fields(
                new_book_id,
                metadata.get("edition", ""),
                metadata.get("year", ""),
                metadata.get("isbn", ""),
                metadata.get("publisher", ""),
            )
            self.db.update_series_fields(new_book_id, metadata.get("series", ""), metadata.get("series_number", ""))
            self.update_accessibility_from_epub(new_book_id, destination)
            self.queue_book_for_indexing(new_book_id, str(destination))
            added += 1

        self.refresh_books()
        self.status_var.set(f"Added {added} book{'s' if added != 1 else ''}.")
        if added:
            self.focus_books_list()

    def unique_destination(self, base_name, extension):
        candidate = self.db.books_path / f"{base_name}{extension}"
        number = 2
        while candidate.exists():
            candidate = self.db.books_path / f"{base_name} ({number}){extension}"
            number += 1
        return candidate

    def repair_backup_files(self):
        backup_suffixes = (
            ".bak",
            ".before_pagebreaks.bak",
            ".before_toc_rebuild.bak",
            ".before_text_cleanup.bak",
        )
        try:
            candidates = [
                path for path in self.db.books_path.rglob("*")
                if path.is_file() and any(path.name.lower().endswith(suffix) for suffix in backup_suffixes)
            ]
        except Exception as exc:
            self.log_error("Listing repair backup files", exc)
            return []
        return sorted(candidates, key=lambda path: str(path).casefold())

    def remove_repair_backup_files(self):
        backups = self.repair_backup_files()
        if not backups:
            self.status_var.set("No repair backup files found.")
            messagebox.showinfo(
                "No Backup Files Found",
                "No repair backup files were found in the managed Books folder."
            )
            return

        total_bytes = 0
        for path in backups:
            try:
                total_bytes += path.stat().st_size
            except OSError:
                pass

        if not messagebox.askyesno(
            "Remove Repair Backup Files",
            f"Remove {len(backups)} repair backup file{'s' if len(backups) != 1 else ''} from the managed Books folder?\n\n"
            f"Total size: {format_byte_count(total_bytes)}\n\n"
            "This only removes files whose names end in .bak, .before_pagebreaks.bak, .before_toc_rebuild.bak, or .before_text_cleanup.bak."
        ):
            return

        removed = 0
        skipped = []
        for path in backups:
            try:
                path.unlink()
                removed += 1
            except Exception as exc:
                skipped.append(f"{path} -- {exc}")

        if skipped:
            report_path = self.write_import_report(removed, skipped)
            messagebox.showwarning(
                "Backup Cleanup Finished With Warnings",
                f"Removed {removed} backup file{'s' if removed != 1 else ''}. "
                f"Could not remove {len(skipped)} file{'s' if len(skipped) != 1 else ''}.\n\n"
                f"Report saved at:\n{report_path}"
            )
        else:
            messagebox.showinfo(
                "Backup Cleanup Complete",
                f"Removed {removed} backup file{'s' if removed != 1 else ''}."
            )
        self.status_var.set(f"Removed {removed} repair backup file{'s' if removed != 1 else ''}.")

    def row_to_metadata(self, row):
        return {
            "title": row[1],
            "author": row[2],
            "edition": row[10],
            "year": row[11],
            "isbn": row[12],
            "publisher": row[13],
            "series": row[21] if len(row) > 21 else "",
            "series_number": row[22] if len(row) > 22 else "",
            "source": row[3],
            "tags": row[4],
            "notes": row[5],
            "cover_url": row[14] if len(row) > 14 else "",
            "accessibility_summary": row[15] if len(row) > 15 else "",
            "accessibility_features": row[16] if len(row) > 16 else "",
            "accessibility_hazards": row[17] if len(row) > 17 else "",
            "accessibility_access_modes": row[18] if len(row) > 18 else "",
            "accessibility_access_modes_sufficient": row[19] if len(row) > 19 else "",
            "accessibility_certified_by": row[20] if len(row) > 20 else "",
        }

    def merge_online_metadata(self, existing, online, replace_existing=False):
        merged = dict(existing)
        for key in ["title", "author", "series", "series_number", "edition", "year", "isbn", "publisher", "tags", "notes", "cover_url"]:
            incoming = (online.get(key, "") or "").strip()
            if not incoming:
                continue
            if replace_existing or not (merged.get(key, "") or "").strip():
                merged[key] = incoming
        return merged

    def save_metadata_for_book(self, book_id, row, metadata, status_prefix="Metadata updated"):
        self.db.update_book(
            book_id,
            metadata["title"],
            metadata["author"],
            metadata["source"],
            metadata["tags"],
            metadata["notes"],
        )
        self.db.update_extra_fields(
            book_id,
            metadata.get("edition", ""),
            metadata.get("year", ""),
            metadata.get("isbn", ""),
            metadata.get("publisher", ""),
        )
        self.db.update_series_fields(
            book_id,
            metadata.get("series", ""),
            metadata.get("series_number", ""),
        )
        if metadata.get("cover_url"):
            self.db.update_cover_url(book_id, metadata.get("cover_url", ""))

        stored_path = Path(row[8])
        if stored_path.suffix.lower() == ".epub" and stored_path.exists():
            self.update_accessibility_from_epub(book_id, stored_path)
            try:
                write_epub_metadata(
                    stored_path,
                    metadata["title"],
                    metadata["author"],
                    metadata["source"],
                    metadata["tags"],
                    metadata["notes"],
                )
                self.status_var.set(f"{status_prefix} in the database and EPUB file.")
            except Exception as exc:
                messagebox.showwarning(
                    "EPUB metadata not written",
                    f"The database was updated, but I could not write metadata into the EPUB file itself.\n\n{exc}"
                )
                self.status_var.set(f"{status_prefix} in the database only.")
        else:
            self.status_var.set(f"{status_prefix} in the database only.")

    def auto_detect_selected_metadata(self):
        book_id = self.selected_book_id()
        if book_id is None:
            return

        row = self.db.get_book(book_id)
        if not row:
            messagebox.showerror("Book not found", "The selected book was not found.")
            return

        stored_path = Path(row[8])
        if not stored_path.exists():
            messagebox.showerror("Book file missing", "The stored book file could not be found.")
            return

        existing = self.row_to_metadata(row)

        detected = detect_metadata_from_text(stored_path, existing=existing)

        summary = (
            "Detected metadata:\n\n"
            f"Title: {detected.get('title', '')}\n"
            f"Author: {detected.get('author', '')}\n"
            f"Series: {detected.get('series', '')}\n"
            f"Series number: {detected.get('series_number', '')}\n"
            f"Edition: {detected.get('edition', '')}\n"
            f"Year: {detected.get('year', '')}\n"
            f"ISBN: {detected.get('isbn', '')}\n"
            f"Publisher: {detected.get('publisher', '')}\n"
            f"Source: {detected.get('source', '')}\n"
            f"Tags: {detected.get('tags', '')}\n\n"
            "Review and edit these fields before saving?"
        )

        if not messagebox.askyesno("Auto-detect metadata", summary):
            return

        metadata = TkMetadataDialog.ask(self.root, "Review Detected Metadata", detected)
        if not metadata:
            return

        self.save_metadata_for_book(book_id, row, metadata, status_prefix="Detected metadata saved")
        self.update_accessibility_from_epub(book_id, row[8])
        self.refresh_books(selected_book_id=book_id)
        self.focus_books_list()

    def bulk_edit_selected_metadata(self):
        book_ids = self.selected_book_ids()
        if not book_ids:
            return
        if len(book_ids) < 2:
            if not messagebox.askyesno(
                "Only One Book Selected",
                "Only one book is selected. Continue with bulk metadata edit for this one book?"
            ):
                return
        metadata = BulkMetadataDialog.ask(self.root)
        if not metadata:
            return
        updates = {
            key: value.strip()
            for key, value in metadata.items()
            if key in {
                "author", "source", "tags", "notes", "edition", "year",
                "publisher", "series",
            }
            and value.strip()
        }
        if not updates:
            messagebox.showinfo("No Changes", "No fields were entered, so no books were changed.")
            return
        summary = "\n".join(f"{key.replace('_', ' ').title()}: {value}" for key, value in updates.items())
        if not messagebox.askyesno(
            "Apply Bulk Metadata",
            f"Apply these metadata changes to {len(book_ids)} selected book{'s' if len(book_ids) != 1 else ''}?\n\n{summary}"
        ):
            return
        changed = self.db.update_bulk_metadata_fields(book_ids, updates)
        self.refresh_books(selected_book_id=book_ids[0])
        self.focus_books_list()
        self.status_var.set(f"Bulk metadata updated for {changed} book{'s' if changed != 1 else ''}.")
        messagebox.showinfo(
            "Bulk Metadata Updated",
            f"Updated {changed} book{'s' if changed != 1 else ''}."
        )

    def check_selected_epub_accessibility(self):
        book_id = self.selected_book_id()
        if book_id is None:
            return

        row = self.db.get_book(book_id)
        if not row:
            messagebox.showerror("Book not found", "The selected book was not found.")
            return

        stored_path = Path(row[8])
        if stored_path.suffix.lower() != ".epub":
            messagebox.showinfo("EPUB only", "Accessibility checking is currently available for EPUB books.")
            return
        if not stored_path.exists():
            messagebox.showerror("Book file missing", "The stored EPUB file could not be found.")
            return

        self.status_var.set("Checking EPUB accessibility metadata.")
        metadata = self.update_accessibility_from_epub(book_id, stored_path)
        self.refresh_books(selected_book_id=book_id)
        updated_row = self.db.get_book(book_id)
        text = self.accessibility_text_from_row(updated_row)
        self.status_var.set("EPUB accessibility metadata checked.")
        messagebox.showinfo(
            "EPUB Accessibility Metadata",
            text if metadata else "No accessibility metadata could be checked for this EPUB."
        )

    def cleanup_repair_wizard(self):
        book_id = self.selected_book_id()
        if book_id is None:
            return
        row = self.db.get_book(book_id)
        if not row:
            messagebox.showerror("Book not found", "The selected book was not found.")
            return
        stored_path = Path(row[8])
        if stored_path.suffix.lower() != ".epub":
            messagebox.showinfo("EPUB only", "The cleanup and repair wizard currently works on EPUB books.")
            return
        if not stored_path.exists():
            messagebox.showerror("Book file missing", "The stored EPUB file could not be found.")
            return

        if not messagebox.askyesno(
            "Cleanup and Repair Wizard",
            f"Run the cleanup and repair wizard for:\n\n{row[1]}\n\n"
            "The wizard can check accessibility metadata, clean repeated text, add page breaks, and rebuild the table of contents. "
            "Backup copies are made before EPUB file repairs.\n\nContinue?"
        ):
            return

        run_accessibility = messagebox.askyesno("Wizard Step", "Check EPUB accessibility metadata?")
        run_clean = messagebox.askyesno("Wizard Step", "Clean repeated headers, footers, blank-page notices, language metadata, and page labels?")
        run_pages = False if run_clean else messagebox.askyesno("Wizard Step", "Add EPUB page breaks from written page labels?")
        run_toc = messagebox.askyesno("Wizard Step", "Rebuild the EPUB table of contents from a text table of contents?")

        results = []
        try:
            if run_accessibility:
                metadata = self.update_accessibility_from_epub(book_id, stored_path)
                results.append("Accessibility metadata checked." if metadata else "No accessibility metadata was found.")
            if run_clean:
                changed = self.clean_epub_text_file(stored_path)
                results.append(f"Text cleanup made {changed} change{'s' if changed != 1 else ''}.")
            if run_pages:
                inserted = self.add_page_breaks_to_epub_file(stored_path)
                results.append(f"Added {inserted} page-break marker{'s' if inserted != 1 else ''}.")
            if run_toc:
                inserted = self.add_page_breaks_to_epub_file(stored_path)
                added, skipped = self.rebuild_epub_toc_from_text(stored_path)
                results.append(
                    f"Before TOC rebuild, added {inserted} page-break marker{'s' if inserted != 1 else ''}. "
                    f"Built {added} TOC item{'s' if added != 1 else ''}; skipped {skipped}."
                )
            if run_clean or run_pages or run_toc:
                self.update_accessibility_from_epub(book_id, stored_path)
        except Exception as exc:
            self.log_error("Cleanup and repair wizard", exc)
            messagebox.showerror(
                "Cleanup Wizard Failed",
                f"The cleanup wizard could not finish.\n\n{exc}\n\nCrash details saved at:\n{self.crash_log_path()}"
            )
            return

        if not results:
            results.append("No wizard steps were selected.")
        self.refresh_books(selected_book_id=book_id)
        self.focus_books_list()
        text = "\n".join(results)
        self.status_var.set("Cleanup and repair wizard complete.")
        messagebox.showinfo("Cleanup and Repair Complete", text)

    def add_page_breaks_to_selected_epub(self):
        book_id = self.selected_book_id()
        if book_id is None:
            return

        row = self.db.get_book(book_id)
        if not row:
            messagebox.showerror("Book not found", "The selected book was not found.")
            return

        stored_path = Path(row[8])
        if stored_path.suffix.lower() != ".epub":
            messagebox.showinfo("EPUB only", "Page-break repair is available for EPUB books.")
            return
        if not stored_path.exists():
            messagebox.showerror("Book file missing", "The stored EPUB file could not be found.")
            return

        if not messagebox.askyesno(
            "Add EPUB Page Breaks",
            "Scan this EPUB for written page labels such as Page 277 or age 277 and turn them into EPUB page-break markers?\n\n"
            "A backup copy of the EPUB will be made first."
        ):
            return

        try:
            inserted = self.add_page_breaks_to_epub_file(stored_path)
        except Exception as exc:
            self.log_error("Adding EPUB page breaks", exc)
            messagebox.showerror(
                "Page-break repair failed",
                f"Could not add page breaks to this EPUB.\n\n{exc}\n\nCrash details saved at:\n{self.crash_log_path()}"
            )
            return

        if inserted:
            self.update_accessibility_from_epub(book_id, stored_path)
            self.refresh_books(selected_book_id=book_id)
            self.focus_books_list()
            self.status_var.set(f"Added {inserted} EPUB page break{'s' if inserted != 1 else ''}.")
            messagebox.showinfo(
                "Page Breaks Added",
                f"Added {inserted} EPUB page break{'s' if inserted != 1 else ''}."
            )
        else:
            self.status_var.set("No written page labels were found.")
            messagebox.showinfo(
                "No Page Labels Found",
                "I did not find written page labels like Page 277 or age 277 in this EPUB."
            )

    def add_page_breaks_to_epub_file(self, epub_path: Path):
        temp_path = epub_path.with_suffix(epub_path.suffix + ".pagebreaks.tmp")
        backup_path = epub_path.with_suffix(epub_path.suffix + ".before_pagebreaks.bak")
        inserted_total = 0

        with zipfile.ZipFile(epub_path, "r") as source_archive:
            with zipfile.ZipFile(temp_path, "w") as target_archive:
                for item in source_archive.infolist():
                    data = source_archive.read(item.filename)
                    if item.filename.lower().endswith((".xhtml", ".html", ".htm")):
                        try:
                            text = data.decode("utf-8")
                            updated, inserted = self.insert_pagebreaks_in_xhtml_text(text)
                            if inserted:
                                data = updated.encode("utf-8")
                                inserted_total += inserted
                        except UnicodeDecodeError:
                            try:
                                text = data.decode("latin-1")
                                updated, inserted = self.insert_pagebreaks_in_xhtml_text(text)
                                if inserted:
                                    data = updated.encode("utf-8")
                                    inserted_total += inserted
                            except Exception:
                                pass
                    target_archive.writestr(item, data)

        if inserted_total:
            if not backup_path.exists():
                shutil.copy2(epub_path, backup_path)
            os.replace(temp_path, epub_path)
        else:
            try:
                temp_path.unlink()
            except OSError:
                pass
        return inserted_total

    def rebuild_selected_epub_toc_from_text(self):
        book_id = self.selected_book_id()
        if book_id is None:
            return

        row = self.db.get_book(book_id)
        if not row:
            messagebox.showerror("Book not found", "The selected book was not found.")
            return

        stored_path = Path(row[8])
        if stored_path.suffix.lower() != ".epub":
            messagebox.showinfo("EPUB only", "Table of contents repair is available for EPUB books.")
            return
        if not stored_path.exists():
            messagebox.showerror("Book file missing", "The stored EPUB file could not be found.")
            return

        if not messagebox.askyesno(
            "Rebuild EPUB Table of Contents",
            "Look for a text Table of Contents in this EPUB and rebuild the EPUB navigation from it?\n\n"
            "The app will first scan for written page labels and add EPUB page-break markers if needed. "
            "Backup copies of the EPUB will be made before repairs."
        ):
            return

        try:
            self.status_var.set("Checking page labels before rebuilding the table of contents.")
            inserted = self.add_page_breaks_to_epub_file(stored_path)
            added, skipped = self.rebuild_epub_toc_from_text(stored_path)
        except Exception as exc:
            self.log_error("Rebuilding EPUB TOC from text", exc)
            messagebox.showerror(
                "TOC repair failed",
                f"Could not rebuild the EPUB table of contents.\n\n{exc}\n\nCrash details saved at:\n{self.crash_log_path()}"
            )
            return

        if added:
            self.update_accessibility_from_epub(book_id, stored_path)
            self.refresh_books(selected_book_id=book_id)
            self.focus_books_list()
            self.status_var.set(f"Rebuilt EPUB table of contents with {added} item{'s' if added != 1 else ''}.")
            messagebox.showinfo(
                "Table of Contents Rebuilt",
                f"Added {inserted} page-break marker{'s' if inserted != 1 else ''} before rebuilding.\n\n"
                f"Built EPUB navigation with {added} item{'s' if added != 1 else ''} from the text table of contents.\n\n"
                f"Skipped {skipped} item{'s' if skipped != 1 else ''} that did not have matching page-break anchors."
            )
        else:
            self.status_var.set("No usable text table of contents was found.")
            messagebox.showinfo(
                "No TOC Built",
                f"Added {inserted} page-break marker{'s' if inserted != 1 else ''}, but I could not build navigation from the text table of contents."
            )

    def rebuild_epub_toc_from_text(self, epub_path: Path):
        text = read_text_from_epub_preserve_lines(epub_path, max_chars=750000)
        toc_entries = parse_text_toc_entries(text)
        if not toc_entries:
            return 0, 0

        opf_path = get_epub_opf_path(epub_path)
        temp_path = epub_path.with_suffix(epub_path.suffix + ".toc.tmp")
        backup_path = epub_path.with_suffix(epub_path.suffix + ".before_toc_rebuild.bak")

        with zipfile.ZipFile(epub_path, "r") as source_archive:
            opf_xml = source_archive.read(opf_path)
            root = ET.fromstring(opf_xml)
            ns = {"opf": OPF_NS}
            manifest = root.find("opf:manifest", ns)
            if manifest is None:
                manifest = ET.SubElement(root, f"{{{OPF_NS}}}manifest")

            nav_item = None
            for item in manifest.findall("opf:item", ns):
                properties = item.attrib.get("properties", "")
                if "nav" in properties.split():
                    nav_item = item
                    break

            if nav_item is None:
                existing_ids = {item.attrib.get("id", "") for item in manifest.findall("opf:item", ns)}
                nav_id = "nav"
                suffix = 2
                while nav_id in existing_ids:
                    nav_id = f"nav{suffix}"
                    suffix += 1
                nav_item = ET.SubElement(manifest, f"{{{OPF_NS}}}item")
                nav_item.set("id", nav_id)
                nav_item.set("href", "nav.xhtml")
                nav_item.set("media-type", "application/xhtml+xml")
                nav_item.set("properties", "nav")

            nav_href = nav_item.attrib.get("href") or "nav.xhtml"
            nav_item.set("media-type", "application/xhtml+xml")
            nav_item.set("properties", "nav")
            opf_dir = posixpath.dirname(opf_path)
            nav_path = posixpath.normpath(posixpath.join(opf_dir, nav_href))
            nav_dir = posixpath.dirname(nav_path)

            page_targets = {}
            for item in source_archive.infolist():
                name = item.filename
                if name == nav_path or not name.lower().endswith((".xhtml", ".html", ".htm")):
                    continue
                try:
                    content = source_archive.read(name).decode("utf-8", errors="ignore")
                except Exception:
                    continue
                for page_id in re.findall(r'id=["\'](page-[^"\']+)["\']', content, flags=re.IGNORECASE):
                    href = posixpath.relpath(name, nav_dir) if nav_dir else name
                    page_targets.setdefault(page_id.casefold(), href + "#" + page_id)

            nav_entries = []
            for entry in toc_entries:
                href = page_targets.get(entry["page_id"].casefold())
                if href:
                    nav_entries.append((href, entry["title"]))

            if not nav_entries:
                return 0, len(toc_entries)

            title = ""
            try:
                row_title = first_or_empty(root, ".//dc:title", {"dc": DC_NS})
                title = row_title or "Table of Contents"
            except Exception:
                title = "Table of Contents"
            nav_xhtml = self.build_nav_xhtml(title, nav_entries)
            new_opf_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)

            with zipfile.ZipFile(temp_path, "w") as target_archive:
                wrote_nav = False
                for item in source_archive.infolist():
                    if item.filename == opf_path:
                        target_archive.writestr(item, new_opf_xml)
                    elif item.filename == nav_path:
                        target_archive.writestr(item, nav_xhtml.encode("utf-8"))
                        wrote_nav = True
                    else:
                        target_archive.writestr(item, source_archive.read(item.filename))
                if not wrote_nav:
                    target_archive.writestr(nav_path, nav_xhtml.encode("utf-8"))

        if not backup_path.exists():
            shutil.copy2(epub_path, backup_path)
        os.replace(temp_path, epub_path)
        return len(nav_entries), len(toc_entries) - len(nav_entries)

    def build_nav_xhtml(self, title, nav_entries):
        nav_items = "\n".join(
            f'      <li><a href="{html.escape(href)}">{html.escape(label)}</a></li>'
            for href, label in nav_entries
        )
        return f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops" lang="en" xml:lang="en">
  <head><title>{html.escape(title or "Table of Contents")}</title></head>
  <body>
    <nav epub:type="toc" id="toc">
      <h1>Table of Contents</h1>
      <ol>
{nav_items}
      </ol>
    </nav>
  </body>
</html>
"""

    def clean_selected_epub_text(self):
        book_id = self.selected_book_id()
        if book_id is None:
            return

        row = self.db.get_book(book_id)
        if not row:
            messagebox.showerror("Book not found", "The selected book was not found.")
            return

        stored_path = Path(row[8])
        if stored_path.suffix.lower() != ".epub":
            messagebox.showinfo("EPUB only", "Text cleanup is available for EPUB books.")
            return
        if not stored_path.exists():
            messagebox.showerror("Book file missing", "The stored EPUB file could not be found.")
            return

        if not messagebox.askyesno(
            "Clean EPUB Text",
            "Clean this EPUB by removing repeated short headers or footers, removing blank-page notices, adding language metadata, and adding page breaks from page labels?\n\n"
            "A backup copy of the EPUB will be made first."
        ):
            return

        try:
            changed = self.clean_epub_text_file(stored_path)
        except Exception as exc:
            self.log_error("Cleaning EPUB text", exc)
            messagebox.showerror(
                "EPUB cleanup failed",
                f"Could not clean this EPUB.\n\n{exc}\n\nCrash details saved at:\n{self.crash_log_path()}"
            )
            return

        self.update_accessibility_from_epub(book_id, stored_path)
        self.refresh_books(selected_book_id=book_id)
        self.focus_books_list()
        self.status_var.set(f"EPUB text cleanup complete. {changed} change{'s' if changed != 1 else ''} made.")
        messagebox.showinfo(
            "EPUB Cleanup Complete",
            f"Made {changed} text cleanup change{'s' if changed != 1 else ''}."
        )

    def paragraph_texts_from_xhtml(self, xhtml_text):
        texts = []
        for match in re.finditer(r"<p\b[^>]*>([\s\S]*?)</p>", xhtml_text, flags=re.IGNORECASE):
            text = strip_xml_html_tags_preserve_lines(match.group(1))
            text = re.sub(r"\s+", " ", text).strip()
            if text:
                texts.append(text)
        return texts

    def repeated_paragraphs_for_cleanup(self, xhtml_documents):
        counts = {}
        originals = {}
        for _name, text in xhtml_documents:
            for paragraph in self.paragraph_texts_from_xhtml(text):
                key = paragraph.casefold()
                counts[key] = counts.get(key, 0) + 1
                originals[key] = paragraph

        repeated = set()
        for key, count in counts.items():
            paragraph = originals[key]
            if count < 4:
                continue
            if len(paragraph) > 120:
                continue
            if is_page_label_line(paragraph) or looks_like_text_toc_entry_line(paragraph):
                continue
            if re.match(r"^(chapter|part|section|appendix)\b", paragraph, flags=re.IGNORECASE):
                continue
            repeated.add(key)
        return repeated

    def clean_xhtml_text(self, xhtml_text, repeated_paragraphs):
        removed = 0

        def replace_paragraph(match):
            nonlocal removed
            opening = match.group(1)
            inner = match.group(2)
            closing = match.group(3)
            text = strip_xml_html_tags_preserve_lines(inner)
            normalized = re.sub(r"\s+", " ", text).strip()
            if normalized and (is_import_boilerplate_line(normalized) or normalized.casefold() in repeated_paragraphs):
                removed += 1
                return ""
            return opening + inner + closing

        xhtml_text = re.sub(
            r"(<p\b[^>]*>)([\s\S]*?)(</p>)",
            replace_paragraph,
            xhtml_text,
            flags=re.IGNORECASE,
        )
        xhtml_text, inserted = self.insert_pagebreaks_in_xhtml_text(xhtml_text)
        return xhtml_text, removed + inserted

    def clean_epub_text_file(self, epub_path: Path):
        temp_path = epub_path.with_suffix(epub_path.suffix + ".clean.tmp")
        backup_path = epub_path.with_suffix(epub_path.suffix + ".before_text_cleanup.bak")
        changed_total = 0

        with zipfile.ZipFile(epub_path, "r") as source_archive:
            xhtml_documents = []
            for item in source_archive.infolist():
                if item.filename.lower().endswith((".xhtml", ".html", ".htm")):
                    try:
                        xhtml_documents.append((item.filename, source_archive.read(item.filename).decode("utf-8", errors="ignore")))
                    except Exception:
                        pass
            repeated = self.repeated_paragraphs_for_cleanup(xhtml_documents)
            language_text = "\n".join(
                strip_xml_html_tags_preserve_lines(text)
                for _name, text in xhtml_documents[:40]
            )
            detected_language = detect_language_from_text(language_text, default="en")

            opf_path = get_epub_opf_path(epub_path)
            opf_xml = source_archive.read(opf_path)
            root = ET.fromstring(opf_xml)
            metadata = root.find(f"{{{OPF_NS}}}metadata")
            if metadata is None:
                metadata = ET.SubElement(root, f"{{{OPF_NS}}}metadata")
            language_before = first_or_empty(root, ".//dc:language", {"dc": DC_NS})
            if not language_before:
                set_single_text(metadata, f"{{{DC_NS}}}language", detected_language)
                changed_total += 1
            new_opf_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)

            with zipfile.ZipFile(temp_path, "w") as target_archive:
                for item in source_archive.infolist():
                    data = source_archive.read(item.filename)
                    if item.filename == opf_path:
                        data = new_opf_xml
                    elif item.filename.lower().endswith((".xhtml", ".html", ".htm")):
                        try:
                            text = data.decode("utf-8")
                            updated, changed = self.clean_xhtml_text(text, repeated)
                            if changed:
                                data = updated.encode("utf-8")
                                changed_total += changed
                        except UnicodeDecodeError:
                            try:
                                text = data.decode("latin-1")
                                updated, changed = self.clean_xhtml_text(text, repeated)
                                if changed:
                                    data = updated.encode("utf-8")
                                    changed_total += changed
                            except Exception:
                                pass
                    target_archive.writestr(item, data)

        if changed_total:
            if not backup_path.exists():
                shutil.copy2(epub_path, backup_path)
            os.replace(temp_path, epub_path)
        else:
            try:
                temp_path.unlink()
            except OSError:
                pass
        return changed_total

    def lookup_selected_metadata_online(self):
        book_id = self.selected_book_id()
        if book_id is None:
            return

        row = self.db.get_book(book_id)
        if not row:
            messagebox.showerror("Book not found", "The selected book was not found.")
            return

        existing = self.row_to_metadata(row)
        if not messagebox.askyesno(
            "Lookup metadata online",
            "This will send the selected book's ISBN, title, and author to Open Library and Google Books to look for missing metadata.\n\nContinue?"
        ):
            return

        self.status_var.set("Looking up metadata online.")
        self.root.update_idletasks()

        try:
            online, service = lookup_online_metadata(existing)
        except Exception as exc:
            messagebox.showerror("Online metadata lookup failed", f"Could not look up metadata online.\n\n{exc}")
            self.status_var.set("Online metadata lookup failed.")
            return

        if not online:
            messagebox.showinfo("No metadata found", "No online metadata match was found.")
            self.status_var.set("No online metadata match found.")
            return

        replace_existing = messagebox.askyesno(
            "Replace existing fields?",
            f"Found metadata from {service}.\n\nChoose Yes to replace existing fields with online values.\nChoose No to fill blank fields only."
        )
        merged = self.merge_online_metadata(existing, online, replace_existing=replace_existing)
        if online.get("cover_url"):
            merged["cover_url"] = online["cover_url"]

        summary = (
            f"Found metadata from {service}:\n\n"
            f"Title: {online.get('title', '')}\n"
            f"Author: {online.get('author', '')}\n"
            f"Series: {online.get('series', '')}\n"
            f"Series number: {online.get('series_number', '')}\n"
            f"Edition: {online.get('edition', '')}\n"
            f"Year: {online.get('year', '')}\n"
            f"ISBN: {online.get('isbn', '')}\n"
            f"Publisher: {online.get('publisher', '')}\n"
            f"Tags: {online.get('tags', '')}\n\n"
            "Review and edit before saving?"
        )
        if not messagebox.askyesno("Review online metadata", summary):
            return

        metadata = TkMetadataDialog.ask(self.root, "Review Online Metadata", merged)
        if not metadata:
            return
        if merged.get("cover_url"):
            metadata["cover_url"] = merged["cover_url"]

        self.save_metadata_for_book(book_id, row, metadata, status_prefix=f"Online metadata from {service} saved")
        self.refresh_books(selected_book_id=book_id)
        self.focus_books_list()

    def view_selected_cover_image(self):
        book_id = self.selected_book_id()
        if book_id is None:
            return

        row = self.db.get_book(book_id)
        if not row:
            messagebox.showerror("Book not found", "The selected book was not found.")
            return

        cover_url = row[14] if len(row) > 14 else ""
        if not cover_url:
            if not self.find_cover_for_selected_book(confirm=True):
                return
            updated_row = self.db.get_book(book_id)
            cover_url = updated_row[14] if updated_row and len(updated_row) > 14 else ""
            if not cover_url:
                return

        try:
            import webbrowser
            webbrowser.open(cover_url)
            self.status_var.set("Opened cover image.")
        except Exception as exc:
            messagebox.showerror("Open cover failed", f"Could not open the cover image.\n\n{exc}")

    def find_cover_for_selected_book(self, confirm=False):
        book_id = self.selected_book_id()
        if book_id is None:
            return False

        row = self.db.get_book(book_id)
        if not row:
            messagebox.showerror("Book not found", "The selected book was not found.")
            return False

        if confirm and not messagebox.askyesno(
            "No cover saved",
            "No cover image is saved for this book yet. Look online now?"
        ):
            return False

        try:
            online, service = lookup_online_metadata(self.row_to_metadata(row))
        except Exception as exc:
            messagebox.showerror("Cover lookup failed", f"Could not look up a cover image online.\n\n{exc}")
            return False

        cover_url = online.get("cover_url", "")
        if not cover_url:
            messagebox.showinfo("No cover found", "No cover image was found online for this book.")
            self.status_var.set("No cover image found.")
            return False

        self.db.update_cover_url(book_id, cover_url)
        self.status_var.set(f"Cover image found from {service}.")
        self.schedule_cover_preview_update(book_id)
        return True

    def edit_book(self, initial_focus_field="title"):
        book_id = self.selected_book_id()
        if book_id is None:
            return

        row = self.db.get_book(book_id)
        if not row:
            messagebox.showerror("Book not found", "The selected book was not found.")
            return

        try:
            metadata = TkMetadataDialog.ask(
                self.root,
                "Edit Book Metadata",
                {
                    "title": row[1],
                    "author": row[2],
                    "series": row[21] if len(row) > 21 else "",
                    "series_number": row[22] if len(row) > 22 else "",
                    "edition": row[10],
                    "year": row[11],
                    "isbn": row[12],
                    "publisher": row[13],
                    "source": row[3],
                    "tags": row[4],
                    "notes": row[5],
                },
                initial_focus_key=initial_focus_field,
            )
        except Exception as exc:
            self.log_error("Opening edit metadata dialog", exc)
            messagebox.showerror(
                "Metadata editor failed",
                f"The metadata editor could not be opened.\n\n{exc}\n\nCrash details saved at:\n{self.crash_log_path()}"
            )
            self.focus_books_list()
            return

        if not metadata:
            return

        self.db.update_book(
            book_id,
            metadata["title"],
            metadata["author"],
            metadata["source"],
            metadata["tags"],
            metadata["notes"],
        )
        self.db.update_extra_fields(
            book_id,
            metadata.get("edition", ""),
            metadata.get("year", ""),
            metadata.get("isbn", ""),
            metadata.get("publisher", ""),
        )
        self.db.update_series_fields(
            book_id,
            metadata.get("series", ""),
            metadata.get("series_number", ""),
        )

        stored_path = Path(row[8])
        if stored_path.suffix.lower() == ".epub" and stored_path.exists():
            self.update_accessibility_from_epub(book_id, stored_path)
            try:
                write_epub_metadata(
                    stored_path,
                    metadata["title"],
                    metadata["author"],
                    metadata["source"],
                    metadata["tags"],
                    metadata["notes"],
                )
                self.status_var.set("Metadata updated in the database and EPUB file.")
            except Exception as exc:
                messagebox.showwarning(
                    "EPUB metadata not written",
                    f"The database was updated, but I could not write metadata into the EPUB file itself.\n\n{exc}"
                )
                self.status_var.set("Metadata updated in the database only.")
        else:
            self.status_var.set("Metadata updated in the database only. File metadata writing is currently supported for EPUB files.")

        self.refresh_books(selected_book_id=book_id)
        self.focus_books_list()

    def open_book(self):
        # Make sure the action uses the book currently highlighted or active in the list.
        book_id = self.selected_book_id()
        if book_id is None:
            return

        row = self.db.get_book(book_id)
        if not row:
            messagebox.showerror("Book not found", "The selected book was not found in the library database.")
            return

        path = str(row[8])
        if not os.path.exists(path):
            messagebox.showerror(
                "Book file missing",
                "The stored book file could not be found. It may have been moved or deleted.\n\n"
                f"Stored path: {path}"
            )
            return

        try:
            reader_path = self.db.get_setting("default_reader_path", "")
            if reader_path and Path(reader_path).exists():
                subprocess.Popen([reader_path, path])
            elif sys.platform.startswith("win"):
                os.startfile(path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", path])
            else:
                subprocess.Popen(["xdg-open", path])
            self.status_var.set(f"Opened {row[1]}.")
        except Exception as exc:
            messagebox.showerror("Open failed", f"Could not open this book.\n\n{exc}")

    def open_kindle(self):
        if not sys.platform.startswith("win"):
            messagebox.showinfo("Windows only", "Kindle for PC launching is only supported on Windows in this starter app.")
            return

        possible_paths = [
            Path(os.environ.get("LOCALAPPDATA", "")) / "Amazon" / "Kindle" / "application" / "Kindle.exe",
            Path(os.environ.get("PROGRAMFILES", "")) / "Amazon" / "Kindle" / "Kindle.exe",
            Path(os.environ.get("PROGRAMFILES(X86)", "")) / "Amazon" / "Kindle" / "Kindle.exe",
        ]

        for kindle_path in possible_paths:
            if kindle_path.exists():
                try:
                    subprocess.Popen([str(kindle_path)], shell=False)
                    self.status_var.set("Opened Kindle for PC.")
                    return
                except Exception as exc:
                    messagebox.showerror("Open Kindle failed", f"Could not open Kindle for PC.\n\n{exc}")
                    return

        try:
            subprocess.Popen("start kindle:", shell=True)
            self.status_var.set("Tried to open Kindle for PC.")
        except Exception:
            messagebox.showerror(
                "Kindle not found",
                "I could not find Kindle for PC. Install Kindle for PC, then try again."
            )

    def calibre_metadata_reading_enabled(self, database=None):
        db = database or getattr(self, "db", None)
        if db is None or not hasattr(db, "get_setting"):
            return True
        return db.get_setting("calibre_metadata_reading", "1") != "0"

    def toggle_calibre_metadata_reading(self):
        enabled = not self.calibre_metadata_reading_enabled()
        self.db.set_setting("calibre_metadata_reading", "1" if enabled else "0")
        state = "on" if enabled else "off"
        self.status_var.set(f"Calibre metadata reading is now {state}.")
        messagebox.showinfo(
            "Calibre Metadata Reading",
            f"Calibre metadata reading is now {state}.\n\n"
            "When this is on, the manager uses Calibre's command-line metadata tool quietly for Kindle files. "
            "You still use this app as the interface."
        )

    def show_calibre_tools_status(self):
        ebook_meta = find_calibre_tool("ebook-meta")
        ebook_convert = find_calibre_tool("ebook-convert")
        metadata_state = "On" if self.calibre_metadata_reading_enabled() else "Off"
        lines = [
            f"Calibre metadata reading: {metadata_state}",
            "",
            f"Metadata tool: {ebook_meta or 'Not found'}",
            f"Conversion tool: {ebook_convert or 'Not found'}",
            "",
            "Calibre is used only as a background helper. The manager does not open Calibre's interface and does not remove DRM.",
        ]
        messagebox.showinfo("Kindle and Calibre", "\n".join(lines))
        self.status_var.set("Calibre tools status shown.")

    def find_ebook_convert(self):
        return find_calibre_tool("ebook-convert")

    def convert_selected_to_epub(self):
        book_id = self.selected_book_id()
        if book_id is None:
            return

        row = self.db.get_book(book_id)
        if not row:
            messagebox.showerror("Book not found", "The selected book was not found.")
            return

        source_file = Path(row[8])
        if not source_file.exists():
            messagebox.showerror("Book file missing", "The stored book file could not be found.")
            return

        if source_file.suffix.lower() == ".epub":
            messagebox.showinfo("Already EPUB", "This book is already an EPUB file.")
            return

        converter = self.find_ebook_convert()
        if not converter:
            messagebox.showerror(
                "Calibre not found",
                "I could not find Calibre's ebook-convert tool. Install Calibre first, then try again."
            )
            return

        output_name = safe_filename(f"{row[2]} - {row[1]}").strip(" -") or safe_filename(row[1])
        output_path = self.unique_destination(output_name, ".epub")

        command = [converter, str(source_file), str(output_path), "--title", row[1]]
        if row[2]:
            command.extend(["--authors", row[2]])
        if row[4]:
            command.extend(["--tags", row[4]])

        self.status_var.set("Converting to EPUB.")
        self.root.update_idletasks()

        try:
            completed = subprocess.run(command, capture_output=True, text=True, check=False, timeout=300)
        except subprocess.TimeoutExpired:
            messagebox.showerror("Conversion timed out", "The conversion took too long and was stopped.")
            self.status_var.set("Conversion timed out.")
            return
        except Exception as exc:
            messagebox.showerror("Conversion failed", f"Could not convert this book to EPUB.\n\n{exc}")
            self.status_var.set("Conversion failed.")
            return

        if completed.returncode != 0 or not output_path.exists():
            details = completed.stderr.strip() or completed.stdout.strip() or "No details were returned by ebook-convert."
            messagebox.showerror("Conversion failed", details[:3000])
            self.status_var.set("Conversion failed.")
            return

        try:
            write_epub_metadata(output_path, row[1], row[2], row[3], row[4], row[5])
        except Exception:
            pass

        new_book_id = self.db.add_book(row[1], row[2], row[3], row[4], row[5], str(source_file), str(output_path))
        self.db.update_extra_fields(new_book_id, row[10], row[11], row[12], row[13])
        self.db.update_series_fields(
            new_book_id,
            row[21] if len(row) > 21 else "",
            row[22] if len(row) > 22 else "",
        )
        self.update_accessibility_from_epub(new_book_id, output_path)
        self.queue_book_for_indexing(new_book_id, str(output_path))
        self.refresh_books()
        self.focus_books_list()
        messagebox.showinfo("Conversion complete", "The EPUB was created and added to your library.")

    def set_kindle_email(self):
        current = self.db.get_setting("kindle_email", "")
        value = AccessibleSingleFieldDialog.ask(
            self.root,
            "Kindle Email Addresses",
            "Enter one or more Send to Kindle email addresses. Separate multiple addresses with commas, semicolons, or spaces.",
            current,
        )
        if value is None:
            return
        emails = self.parse_kindle_emails(value)
        if value.strip() and not emails:
            messagebox.showerror("Invalid email", "Please enter at least one valid Kindle email address.")
            return
        invalid = [part for part in re.split(r"[,;\s]+", value.strip()) if part and "@" not in part]
        if invalid:
            messagebox.showerror("Invalid email", f"This does not look like a valid email address:\n\n{invalid[0]}")
            return
        self.db.set_setting("kindle_email", ", ".join(emails))
        count = len(emails)
        self.status_var.set(f"{count} Kindle email address{'es' if count != 1 else ''} saved.")

    def parse_kindle_emails(self, value):
        emails = []
        for part in re.split(r"[,;\s]+", value or ""):
            email = part.strip()
            if not email or "@" not in email:
                continue
            if email.lower() not in [existing.lower() for existing in emails]:
                emails.append(email)
        return emails

    def choose_kindle_recipients(self, emails):
        if not emails:
            return []
        if len(emails) == 1:
            return emails

        lines = ["0. All Kindle addresses"]
        lines.extend(f"{index}. {email}" for index, email in enumerate(emails, start=1))
        value = AccessibleSingleFieldDialog.ask(
            self.root,
            "Kindle Recipients",
            "Type 0 for all addresses, or type one or more numbers separated by commas.\n\n" + "\n".join(lines),
            "0",
            heading="Choose Kindle Recipients",
        )
        if value is None:
            return []
        value = value.strip()
        if not value or value == "0":
            return emails
        selected = []
        for part in re.split(r"[,;\s]+", value):
            if not part:
                continue
            try:
                index = int(part)
            except ValueError:
                messagebox.showerror("Invalid choice", "Please enter recipient numbers from the list.")
                return []
            if index == 0:
                return emails
            if index < 1 or index > len(emails):
                messagebox.showerror("Invalid choice", "Please enter recipient numbers from the list.")
                return []
            email = emails[index - 1]
            if email not in selected:
                selected.append(email)
        return selected

    def send_to_kindle(self):
        book_ids = self.selected_book_ids()
        if not book_ids:
            return

        rows = []
        missing = []
        for book_id in book_ids:
            row = self.db.get_book(book_id)
            if not row:
                continue
            source = Path(row[8])
            if source.exists():
                rows.append((row, source))
            else:
                missing.append(row[1])
        if not rows:
            messagebox.showerror("Book file missing", "The stored book file could not be found.")
            return

        emails = self.parse_kindle_emails(self.db.get_setting("kindle_email", ""))
        if not emails:
            self.set_kindle_email()
            emails = self.parse_kindle_emails(self.db.get_setting("kindle_email", ""))
            if not emails:
                return
        recipients = self.choose_kindle_recipients(emails)
        if not recipients:
            return

        try:
            # Standard mailto cannot reliably attach files. This opens a draft and shows the file path to attach.
            import webbrowser
            subject = "Send to Kindle"
            if len(rows) == 1:
                subject = f"Send to Kindle: {rows[0][0][1]}"
            body_lines = ["Attach these files before sending to Kindle:", ""]
            body_lines.extend(str(source) for _row, source in rows)
            if missing:
                body_lines.extend(["", "These selected books were skipped because their stored files were missing:"])
                body_lines.extend(missing)
            mailto = (
                "mailto:"
                + ",".join(urllib.parse.quote(email) for email in recipients)
                + "?subject="
                + urllib.parse.quote(subject)
                + "&body="
                + urllib.parse.quote("\r\n".join(body_lines))
            )
            webbrowser.open(mailto)
            messagebox.showinfo(
                "Send to Kindle",
                "Your email app should open. Attach the selected book file or files if they are not already attached, then send the message.\n\n"
                + "\n".join(str(source) for _row, source in rows)
            )
        except Exception as exc:
            messagebox.showerror("Send to Kindle failed", f"Could not send this book to Kindle.\n\n{exc}")

    def choose_default_reader(self):
        path = filedialog.askopenfilename(
            title="Choose default ebook reader program",
            filetypes=[("Programs", "*.exe"), ("All files", "*.*")]
        )
        if not path:
            return
        self.db.set_setting("default_reader_path", path)
        self.status_var.set(f"Default reader set to {path}.")

    def clear_default_reader(self):
        self.db.set_setting("default_reader_path", "")
        self.status_var.set("Default reader cleared. System default will be used.")

    def open_library_folder(self):
        folder = str(self.db.folder)
        try:
            if sys.platform.startswith("win"):
                os.startfile(folder)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", folder])
            else:
                subprocess.Popen(["xdg-open", folder])
        except Exception as exc:
            messagebox.showerror("Open library folder failed", f"Could not open the library folder.\n\n{exc}")

    def removable_windows_drives(self):
        if not sys.platform.startswith("win"):
            return []

        try:
            import ctypes
            from ctypes import wintypes

            kernel32 = ctypes.windll.kernel32
            DRIVE_REMOVABLE = 2
            drives = []
            for letter in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
                root = f"{letter}:\\"
                if kernel32.GetDriveTypeW(ctypes.c_wchar_p(root)) != DRIVE_REMOVABLE:
                    continue

                volume_buffer = ctypes.create_unicode_buffer(261)
                filesystem_buffer = ctypes.create_unicode_buffer(261)
                serial_number = wintypes.DWORD()
                max_component_length = wintypes.DWORD()
                filesystem_flags = wintypes.DWORD()
                ok = kernel32.GetVolumeInformationW(
                    ctypes.c_wchar_p(root),
                    volume_buffer,
                    len(volume_buffer),
                    ctypes.byref(serial_number),
                    ctypes.byref(max_component_length),
                    ctypes.byref(filesystem_flags),
                    filesystem_buffer,
                    len(filesystem_buffer),
                )
                label = volume_buffer.value if ok else ""
                drives.append((root, label))
            return drives
        except Exception:
            return []

    def likely_nls_ereader_drives(self):
        keywords = ["nls", "ereader", "e-reader", "bard", "humanware", "cartridge"]
        candidates = []
        for root, label in self.removable_windows_drives():
            haystack = f"{root} {label}".lower()
            score = 0
            if any(keyword in haystack for keyword in keywords):
                score += 10
            try:
                names = {path.name.lower() for path in Path(root).iterdir()}
                if any(keyword in " ".join(names) for keyword in keywords):
                    score += 5
            except Exception:
                pass
            if score:
                candidates.append((root, label, score))
        candidates.sort(key=lambda item: item[2], reverse=True)
        return candidates

    def choose_nls_ereader_folder(self):
        folder = filedialog.askdirectory(
            title="Choose NLS eReader folder"
        )
        if not folder:
            return None

        self.db.set_setting("nls_ereader_folder", folder)
        self.status_var.set(f"NLS eReader folder set to {folder}.")
        messagebox.showinfo(
            "NLS eReader folder saved",
            f"NLS eReader folder saved:\n\n{folder}"
        )
        return folder

    def choose_detected_nls_drive(self, candidates):
        if len(candidates) == 1:
            root, label, _score = candidates[0]
            name = f"{root} {label}".strip()
            if messagebox.askyesno(
                "NLS eReader found",
                f"I found this likely NLS eReader drive:\n\n{name}\n\nUse this drive?"
            ):
                return root
            return None

        lines = []
        for index, (root, label, _score) in enumerate(candidates, start=1):
            name = f"{root} {label}".strip()
            lines.append(f"{index}. {name}")

        value = AccessibleSingleFieldDialog.ask(
            self.root,
            "NLS eReader Drive",
            "Several likely eReader drives were found. Type the number to use, or press Escape to choose a folder manually.\n\n"
            + "\n".join(lines),
            "1",
        )
        if value is None:
            return None
        try:
            selected = int(value.strip())
        except ValueError:
            messagebox.showerror("Invalid choice", "Please enter one of the listed numbers.")
            return None
        if selected < 1 or selected > len(candidates):
            messagebox.showerror("Invalid choice", "Please enter one of the listed numbers.")
            return None
        return candidates[selected - 1][0]

    def get_nls_ereader_folder(self):
        folder = self.db.get_setting("nls_ereader_folder", "")
        if folder and Path(folder).exists():
            return folder

        candidates = self.likely_nls_ereader_drives()
        if candidates:
            selected = self.choose_detected_nls_drive(candidates)
            if selected:
                self.db.set_setting("nls_ereader_folder", selected)
                return selected

        messagebox.showinfo(
            "Choose NLS eReader folder",
            "I could not automatically find an NLS eReader drive. Choose the eReader drive or the folder where books should be copied."
        )
        return self.choose_nls_ereader_folder()

    def send_to_nls_ereader(self):
        book_id = self.selected_book_id()
        if book_id is None:
            return

        row = self.db.get_book(book_id)
        if not row:
            messagebox.showerror("Book not found", "The selected book was not found.")
            return

        source = Path(row[8])
        if not source.exists():
            messagebox.showerror(
                "File missing",
                "The stored book file could not be found, so it cannot be sent to the NLS eReader."
            )
            return

        ereader_folder = self.get_nls_ereader_folder()
        if not ereader_folder:
            return

        destination_folder = Path(ereader_folder)
        if not destination_folder.exists():
            messagebox.showerror(
                "Folder missing",
                "The NLS eReader folder does not exist. Please choose it again."
            )
            self.db.set_setting("nls_ereader_folder", "")
            return

        destination = destination_folder / source.name
        if destination.exists():
            replace = messagebox.askyesno(
                "Replace existing file",
                f"{destination.name} already exists on the NLS eReader. Replace it?"
            )
            if not replace:
                return

        try:
            shutil.copy2(source, destination)
        except Exception as exc:
            messagebox.showerror(
                "Send to NLS eReader failed",
                f"Could not copy the book to the NLS eReader.\n\n{exc}"
            )
            return

        self.status_var.set(f"Sent to NLS eReader: {source.name}")
        messagebox.showinfo(
            "Sent to NLS eReader",
            f"Copied this book to the NLS eReader:\n\n{destination}"
        )

    def copy_to_humanware_mtp(self, source: Path, device_name: str = "") -> dict:
        if not sys.platform.startswith("win"):
            return {
                "ok": False,
                "message": "HumanWare MTP sending is only supported on Windows.",
                "device": "",
                "folder": "",
            }

        script = r'''
param(
    [Parameter(Mandatory=$true)][string]$SourcePath,
    [Parameter(Mandatory=$true)][string]$OutputPath,
    [string]$DeviceName = ''
)

$ErrorActionPreference = 'Stop'

function Write-Result([bool]$Ok, [string]$Message, [string]$Device = '', [string]$Folder = '') {
    $result = [ordered]@{
        ok = $Ok
        message = $Message
        device = $Device
        folder = $Folder
    }
    $json = $result | ConvertTo-Json -Depth 4 -Compress
    $encoding = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($OutputPath, $json, $encoding)
}

function Get-ShellFolder($Item) {
    try {
        return $Item.GetFolder
    }
    catch {
        return $null
    }
}

function Get-ShellItems($Folder) {
    try {
        return @($Folder.Items())
    }
    catch {
        return @()
    }
}

function Find-PreferredFolder($Folder, [int]$Depth) {
    if ($null -eq $Folder -or $Depth -lt 0) {
        return $null
    }

    $items = Get-ShellItems $Folder
    $preferredPattern = '(?i)^(Books|Book|Documents|Downloads|Download|My Books|Digital Editions)$'
    foreach ($item in $items) {
        if ($item.IsFolder -and $item.Name -match $preferredPattern) {
            return $item
        }
    }

    $storageFallback = $null
    foreach ($item in $items) {
        if (-not $item.IsFolder) {
            continue
        }
        if ($item.Name -match '(?i)internal|shared|storage|sd card|memory|card') {
            if ($null -eq $storageFallback) {
                $storageFallback = $item
            }
            $child = Find-PreferredFolder (Get-ShellFolder $item) ($Depth - 1)
            if ($null -ne $child) {
                return $child
            }
        }
    }

    foreach ($item in $items) {
        if (-not $item.IsFolder) {
            continue
        }
        if ($item.Name -match '(?i)android|data|dcim|music|pictures|podcasts|recordings|system') {
            continue
        }
        $child = Find-PreferredFolder (Get-ShellFolder $item) ($Depth - 1)
        if ($null -ne $child) {
            return $child
        }
    }

    if ($null -ne $storageFallback) {
        return $storageFallback
    }
    return $null
}

try {
    if (-not (Test-Path -LiteralPath $SourcePath)) {
        Write-Result $false 'The selected book file could not be found.'
        exit 0
    }

    $shell = New-Object -ComObject Shell.Application
    $thisPc = $shell.Namespace(17)
    if ($null -eq $thisPc) {
        Write-Result $false 'Windows did not expose the This PC device list.'
        exit 0
    }

    $devicePattern = '(?i)humanware|brailliant|mantis|chameleon|victor|stream|e.?reader|nls|aph'
    $rootItems = Get-ShellItems $thisPc
    $devices = @()
    if (-not [string]::IsNullOrWhiteSpace($DeviceName)) {
        foreach ($item in $rootItems) {
            if ($item.IsFolder -and $item.Name -eq $DeviceName) {
                $devices += $item
            }
        }
    }

    foreach ($item in $rootItems) {
        if ($devices.Count -eq 0 -and $item.IsFolder -and $item.Name -match $devicePattern) {
            $devices += $item
        }
    }

    if ($devices.Count -eq 0) {
        $fallbackDevices = @()
        foreach ($item in $rootItems) {
            if (-not $item.IsFolder) {
                continue
            }
            if ($item.Name -match '(?i)\(C:\)|windows\s*\(C:\)|local disk') {
                continue
            }
            if ($item.Name -notmatch '(?i)usb|removable|humanware|brailliant|mantis|chameleon|victor|stream|e.?reader|nls|aph') {
                continue
            }
            $fallbackDevices += $item.Name
        }
        if ($fallbackDevices.Count -gt 0) {
            $result = [ordered]@{
                ok = $false
                code = 'choose_device'
                message = 'No HumanWare-named eReader was found. Choose the connected eReader from the Windows device list.'
                device = ''
                folder = ''
                devices = @($fallbackDevices)
            }
            $json = $result | ConvertTo-Json -Depth 5 -Compress
            $encoding = New-Object System.Text.UTF8Encoding($false)
            [System.IO.File]::WriteAllText($OutputPath, $json, $encoding)
            exit 0
        }
        Write-Result $false "No likely HumanWare MTP eReader was found. Connect and unlock the eReader, choose MTP or File Transfer mode if prompted, then try again."
        exit 0
    }

    $device = $devices[0]
    $deviceFolder = Get-ShellFolder $device
    if ($null -eq $deviceFolder) {
        Write-Result $false ("Windows found " + $device.Name + ", but did not allow access to its storage.") $device.Name ''
        exit 0
    }

    $targetItem = Find-PreferredFolder $deviceFolder 5
    if ($null -eq $targetItem) {
        $targetFolder = $deviceFolder
        $targetName = $device.Name
    }
    else {
        $targetFolder = Get-ShellFolder $targetItem
        $targetName = $targetItem.Name
    }

    if ($null -eq $targetFolder) {
        Write-Result $false ("Windows found " + $device.Name + ", but could not open the destination folder.") $device.Name ''
        exit 0
    }

    $fileName = [System.IO.Path]::GetFileName($SourcePath)
    $targetFolder.CopyHere($SourcePath, 16)

    $copied = $false
    for ($index = 0; $index -lt 120; $index++) {
        Start-Sleep -Milliseconds 500
        foreach ($item in (Get-ShellItems $targetFolder)) {
            if ($item.Name -eq $fileName) {
                $copied = $true
                break
            }
        }
        if ($copied) {
            break
        }
    }

    if ($copied) {
        Write-Result $true ("Copied " + $fileName + " to " + $device.Name + ", " + $targetName + ".") $device.Name $targetName
    }
    else {
        Write-Result $true ("Windows started copying " + $fileName + " to " + $device.Name + ". If the file is large, it may still be finishing in the background.") $device.Name $targetName
    }
}
catch {
    Write-Result $false $_.Exception.Message
}
'''

        with tempfile.TemporaryDirectory(prefix="aelm_mtp_") as temp_folder:
            temp = Path(temp_folder)
            script_path = temp / "send_to_humanware_mtp.ps1"
            output_path = temp / "send_to_humanware_mtp_result.json"
            script_path.write_text(script, encoding="utf-8")

            completed = subprocess.run(
            [
                "powershell",
                "-WindowStyle",
                "Hidden",
                "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-STA",
                    "-File",
                    str(script_path),
                    "-SourcePath",
                    str(source),
                    "-OutputPath",
                    str(output_path),
                    "-DeviceName",
                    device_name,
                ],
                capture_output=True,
                text=True,
                timeout=90,
                creationflags=WINDOWS_NO_CONSOLE_FLAGS,
            )
            if completed.returncode != 0:
                return {
                    "ok": False,
                    "message": completed.stderr.strip() or completed.stdout.strip() or "Windows MTP copy failed.",
                    "device": "",
                    "folder": "",
                }
            if not output_path.exists():
                return {
                    "ok": False,
                    "message": "Windows MTP copy did not return a result.",
                    "device": "",
                    "folder": "",
                }
            return json.loads(output_path.read_text(encoding="utf-8"))

    def send_to_humanware_mtp(self):
        book_id = self.selected_book_id()
        if book_id is None:
            return

        row = self.db.get_book(book_id)
        if not row:
            messagebox.showerror("Book not found", "The selected book was not found.")
            return

        source = Path(row[8])
        if not source.exists():
            messagebox.showerror(
                "File missing",
                "The stored book file could not be found, so it cannot be sent to the HumanWare eReader."
            )
            return

        self.status_var.set("Looking for a HumanWare MTP eReader.")
        try:
            result = self.copy_to_humanware_mtp(source)
            if result.get("code") == "choose_device":
                devices = result.get("devices") or []
                if devices:
                    lines = [f"{index}. {name}" for index, name in enumerate(devices, start=1)]
                    value = AccessibleSingleFieldDialog.ask(
                        self.root,
                        "Choose HumanWare eReader",
                        "Windows did not provide a HumanWare device name. Type the number for the connected eReader, or press Escape to cancel.\n\n"
                        + "\n".join(lines),
                        "1",
                    )
                    if value is None:
                        self.status_var.set("HumanWare MTP send canceled.")
                        return
                    try:
                        selected = int(value.strip())
                    except ValueError:
                        messagebox.showerror("Invalid choice", "Please enter one of the listed numbers.")
                        return
                    if selected < 1 or selected > len(devices):
                        messagebox.showerror("Invalid choice", "Please enter one of the listed numbers.")
                        return
                    self.status_var.set(f"Sending to {devices[selected - 1]}.")
                    result = self.copy_to_humanware_mtp(source, device_name=devices[selected - 1])
        except subprocess.TimeoutExpired:
            messagebox.showerror(
                "Send to HumanWare eReader failed",
                "Windows did not finish the MTP copy in time. The eReader may be locked, busy, or waiting for File Transfer mode."
            )
            self.status_var.set("HumanWare MTP send timed out.")
            return
        except Exception as exc:
            self.log_error("Sending to HumanWare MTP eReader", exc)
            messagebox.showerror(
                "Send to HumanWare eReader failed",
                f"Could not send the book through Windows MTP.\n\n{exc}"
            )
            self.status_var.set("HumanWare MTP send failed.")
            return

        if result.get("ok"):
            message = result.get("message", "The book was sent to the HumanWare eReader.")
            self.status_var.set(message)
            messagebox.showinfo("Sent to HumanWare eReader", message)
        else:
            message = result.get("message", "No HumanWare MTP eReader was found.")
            self.status_var.set("HumanWare MTP eReader was not found.")
            messagebox.showerror("HumanWare eReader not found", message)

    def choose_voice_dream_folder(self):
        folder = filedialog.askdirectory(
            title="Choose Voice Dream Loader folder"
        )
        if not folder:
            return None

        self.db.set_setting("voice_dream_loader_folder", folder)
        self.status_var.set(f"Voice Dream Loader folder set to {folder}.")
        messagebox.showinfo(
            "Voice Dream folder saved",
            f"Voice Dream Loader folder saved:\n\n{folder}"
        )
        return folder

    def get_voice_dream_folder(self):
        folder = self.db.get_setting("voice_dream_loader_folder", "")
        if folder and Path(folder).exists():
            return folder

        messagebox.showinfo(
            "Choose Voice Dream Loader folder",
            "Choose the Voice Dream Loader folder. This is usually inside iCloud Drive, in the Voice Dream folder."
        )
        return self.choose_voice_dream_folder()

    def send_to_voice_dream(self):
        self.copy_selected_books_to_folder_target(
            app_name="Voice Dream",
            folder_name="Voice Dream Loader folder",
            folder_getter=self.get_voice_dream_folder,
            setting_key="voice_dream_loader_folder",
        )

    def choose_dolphin_easyreader_folder(self):
        folder = filedialog.askdirectory(
            title="Choose Dolphin EasyReader folder"
        )
        if not folder:
            return None

        self.db.set_setting("dolphin_easyreader_folder", folder)
        self.status_var.set(f"Dolphin EasyReader folder set to {folder}.")
        messagebox.showinfo(
            "Dolphin EasyReader folder saved",
            f"Dolphin EasyReader folder saved:\n\n{folder}"
        )
        return folder

    def get_dolphin_easyreader_folder(self):
        folder = self.db.get_setting("dolphin_easyreader_folder", "")
        if folder and Path(folder).exists():
            return folder

        messagebox.showinfo(
            "Choose Dolphin EasyReader folder",
            "Choose the Dolphin EasyReader import or library folder where books should be copied."
        )
        return self.choose_dolphin_easyreader_folder()

    def send_to_dolphin_easyreader(self):
        self.copy_selected_books_to_folder_target(
            app_name="Dolphin EasyReader",
            folder_name="Dolphin EasyReader folder",
            folder_getter=self.get_dolphin_easyreader_folder,
            setting_key="dolphin_easyreader_folder",
        )

    def copy_batch_to_folder(self, rows, destination_folder):
        planned_names = set()
        destinations = []
        for row in rows:
            source = Path(row[8])
            if not source.exists():
                destinations.append((row, source, None, "stored file missing"))
                continue

            destination = destination_folder / source.name
            key = destination.name.casefold()
            if key in planned_names:
                base = safe_filename(f"{row[2]} - {row[1]}").strip(" -") or safe_filename(row[1])
                destination = destination_folder / f"{base}{source.suffix}"
                suffix = 2
                while destination.name.casefold() in planned_names:
                    destination = destination_folder / f"{base} ({suffix}){source.suffix}"
                    suffix += 1
            planned_names.add(destination.name.casefold())
            try:
                if source.resolve() == destination.resolve():
                    destinations.append((row, source, None, "source and destination are the same file"))
                    continue
            except OSError:
                pass
            destinations.append((row, source, destination, ""))

        existing = [destination for _row, _source, destination, error in destinations if destination and destination.exists()]
        replace_existing = False
        if existing:
            preview = "\n".join(path.name for path in existing[:10])
            if len(existing) > 10:
                preview += f"\nAnd {len(existing) - 10} more."
            replace_existing = messagebox.askyesno(
                "Replace existing files?",
                f"{len(existing)} destination file(s) already exist. Replace those files?\n\n{preview}",
            )

        copied = 0
        skipped = []
        for row, source, destination, error in destinations:
            if error:
                skipped.append(f"{row[1]}: {error}")
                continue
            if destination.exists() and not replace_existing:
                skipped.append(f"{row[1]}: skipped because {destination.name} already exists")
                continue
            try:
                shutil.copy2(source, destination)
                copied += 1
            except Exception as exc:
                skipped.append(f"{row[1]}: {exc}")
        return copied, skipped

    def copy_selected_books_to_folder_target(self, app_name, folder_name, folder_getter, setting_key):
        book_ids = self.selected_book_ids()
        if not book_ids:
            return

        rows = [self.db.get_book(book_id) for book_id in book_ids]
        rows = [row for row in rows if row]
        if not rows:
            messagebox.showerror("Book not found", "The selected book or books were not found.")
            return

        target_folder = folder_getter()
        if not target_folder:
            return

        destination_folder = Path(target_folder)
        if not destination_folder.exists():
            messagebox.showerror(
                "Folder missing",
                f"The {folder_name} does not exist. Please choose it again."
            )
            self.db.set_setting(setting_key, "")
            return

        sent, skipped = self.copy_batch_to_folder(rows, destination_folder)

        self.status_var.set(f"Sent {sent} book{'s' if sent != 1 else ''} to {app_name}.")
        if skipped:
            messagebox.showwarning(
                f"Send to {app_name} finished with warnings",
                f"Copied {sent} book{'s' if sent != 1 else ''} to the {folder_name}.\n\n"
                + "\n".join(skipped[:20])
            )
        else:
            messagebox.showinfo(
                f"Sent to {app_name}",
                f"Copied {sent} book{'s' if sent != 1 else ''} to the {folder_name}."
            )

    def export_book(self):
        book_ids = self.selected_book_ids()
        if not book_ids:
            return

        rows = [self.db.get_book(book_id) for book_id in book_ids]
        rows = [row for row in rows if row]
        if not rows:
            return

        folder = filedialog.askdirectory(title="Choose export folder")
        if not folder:
            return

        export_format = self.choose_export_format()
        if export_format is None:
            return

        if export_format == "epub":
            exported, skipped = self.export_batch_as_epub(rows, Path(folder))
        else:
            exported, skipped = self.copy_batch_to_folder(rows, Path(folder))

        self.status_var.set(f"Exported {exported} book{'s' if exported != 1 else ''}.")
        if skipped:
            messagebox.showwarning(
                "Export finished with warnings",
                f"Exported {exported} book{'s' if exported != 1 else ''}.\n\n"
                + "\n".join(skipped[:20])
            )
        else:
            messagebox.showinfo("Export complete", f"Exported {exported} book{'s' if exported != 1 else ''}.")

    def choose_export_format(self):
        result = {"format": None}
        window = tk.Toplevel(self.root)
        window.title("Export Format")
        window.transient(self.root)
        window.grab_set()

        main = ttk.Frame(window, padding=12)
        main.pack(fill=BOTH, expand=True)
        ttk.Label(
            main,
            text="Choose the format for the exported book copies. The originals in your library will not be changed.",
            wraplength=500,
        ).pack(anchor="w", pady=(0, 8))

        format_var = StringVar(value="original")
        ttk.Radiobutton(
            main,
            text="Keep original format",
            variable=format_var,
            value="original",
        ).pack(anchor="w", pady=3)
        ttk.Radiobutton(
            main,
            text="Convert exported copies to EPUB",
            variable=format_var,
            value="epub",
        ).pack(anchor="w", pady=3)

        buttons = ttk.Frame(main)
        buttons.pack(anchor="e", pady=(10, 0))

        def accept():
            result["format"] = format_var.get()
            window.destroy()

        ttk.Button(buttons, text="Export", command=accept).pack(side=LEFT, padx=(0, 8))
        ttk.Button(buttons, text="Cancel", command=window.destroy).pack(side=LEFT)
        window.bind("<Return>", lambda _event: accept())
        window.bind("<Escape>", lambda _event: window.destroy())
        window.protocol("WM_DELETE_WINDOW", window.destroy)

        def focus_choice():
            window.lift()
            window.focus_force()
            for child in main.winfo_children():
                if isinstance(child, ttk.Radiobutton):
                    child.focus_force()
                    break

        window.after(50, focus_choice)
        window.after(200, focus_choice)
        self.root.wait_window(window)
        return result["format"]

    def export_batch_as_epub(self, rows, destination_folder):
        needs_conversion = any(
            row and Path(row[8]).suffix.lower() != ".epub" and Path(row[8]).exists()
            for row in rows
        )
        converter = self.find_ebook_convert() if needs_conversion else None
        if needs_conversion and not converter:
            messagebox.showerror(
                "Calibre not found",
                "I could not find Calibre's ebook-convert tool. Install Calibre to export books as EPUB.",
            )
            return 0, ["EPUB export was not started because Calibre could not be found."]

        destination_folder.mkdir(parents=True, exist_ok=True)
        exported = 0
        skipped = []
        planned = set()
        for row in rows:
            source = Path(row[8])
            if not source.exists():
                skipped.append(f"{row[1]}: stored book file is missing")
                continue

            base_name = safe_filename(f"{row[2]} - {row[1]}").strip(" -") or safe_filename(row[1])
            output_path = destination_folder / f"{base_name}.epub"
            suffix = 2
            while output_path.exists() or output_path.name.casefold() in planned:
                output_path = destination_folder / f"{base_name} ({suffix}).epub"
                suffix += 1
            planned.add(output_path.name.casefold())

            if source.suffix.lower() == ".epub":
                try:
                    shutil.copy2(source, output_path)
                    exported += 1
                except Exception as exc:
                    skipped.append(f"{row[1]}: {exc}")
                continue

            command = [converter, str(source), str(output_path), "--title", row[1]]
            if row[2]:
                command.extend(["--authors", row[2]])
            if row[4]:
                command.extend(["--tags", row[4]])
            self.status_var.set(f"Converting {row[1]} to EPUB for export.")
            self.root.update_idletasks()
            try:
                completed = subprocess.run(command, capture_output=True, text=True, check=False, timeout=300)
                if completed.returncode != 0 or not output_path.exists():
                    details = completed.stderr.strip() or completed.stdout.strip() or "Conversion did not create an EPUB file."
                    skipped.append(f"{row[1]}: {details[:500]}")
                    try:
                        output_path.unlink(missing_ok=True)
                    except OSError:
                        pass
                    continue
                try:
                    write_epub_metadata(output_path, row[1], row[2], row[3], row[4], row[5])
                except Exception:
                    pass
                exported += 1
            except subprocess.TimeoutExpired:
                skipped.append(f"{row[1]}: conversion timed out")
                try:
                    output_path.unlink(missing_ok=True)
                except OSError:
                    pass
            except Exception as exc:
                skipped.append(f"{row[1]}: {exc}")
        return exported, skipped

    def delete_book(self):
        book_ids = self.selected_book_ids()
        if not book_ids:
            return

        rows = [self.db.get_book(book_id) for book_id in book_ids]
        rows = [row for row in rows if row]
        if not rows:
            return

        count = len(rows)
        title_preview = rows[0][1] if count == 1 else f"{count} selected books"
        answer = messagebox.askyesnocancel(
            "Delete book" if count == 1 else "Delete selected books",
            f"Remove {title_preview} from the library database?\n\n"
            "Choose Yes to continue. Choose No or Cancel to stop.",
        )
        if answer is not True:
            return

        delete_file = messagebox.askyesno(
            "Delete stored file",
            "Also delete the stored copy of the book file or files? Choose No to keep the files but remove them from the library list.",
        )

        removed = 0
        for row in rows:
            self.db.delete_book(row[0], delete_file=delete_file)
            self.marked_book_ids.discard(row[0])
            removed += 1
        self.refresh_books()
        self.status_var.set(f"Removed {removed} book{'s' if removed != 1 else ''} from library.")

    def help_text(self):
        return (
            "Accessible Ebook Library Manager keyboard commands:\n\n"
            "Alt: Open the menu bar.\n"
            "Control+N: Add book.\n"
            "Control+Shift+N: Import a folder of books, including Bookshare ZIP files.\n"
            "F2: Edit selected book metadata. On Windows, the metadata editor uses native edit boxes so screen readers can read field names, contents, and typed text. Use Tab and Shift+Tab to move between fields.\n"
            "Use Book, Bulk Edit Selected Metadata, to apply author, series, edition, year, ISBN, publisher, source, tags, or notes to selected books. Blank fields are left unchanged.\n"
            "Control+D: Auto-detect metadata from the selected book.\n"
            "Use Book, Repair, Cleanup and Repair Wizard, to run common EPUB checks and repairs from one guided flow.\n"
            "Use Book, Repair, Check EPUB Accessibility Metadata, to inspect the selected EPUB for package accessibility metadata, navigation, page list, heading structure, language, and image alt text coverage.\n"
            "Use Book, Repair, Add EPUB Page Breaks from Page Labels, to repair an already-imported EPUB that has written page labels like Page 277 or age 277 but no real page-break markers.\n"
            "Use Book, Repair, Rebuild EPUB Table of Contents from Text TOC, to align EPUB navigation with a text table of contents when page-break anchors are available.\n"
            "Use Book, Repair, Clean Repaired EPUB Text, to remove obvious repeated headers or footers, remove blank-page notices, auto-detect missing language metadata, and add page breaks from page labels.\n"
            "Use Book, Repair, Combine Selected Books into One EPUB, to combine selected part-books into a new EPUB in the current list order.\n"
            "Use the Book menu, Look Up Book Metadata from Internet, to search Open Library and Google Books for metadata.\n"
            "Use the Book menu, View Cover Image, to open a visual cover image when one is available or look one up online. The cover preview beside the book list is visual only and does not change book list speech.\n"
            "Enter or Control+O: Open selected book.\n"
            "Control+E: Export selected book.\n"
            "Control+R: Convert selected book to EPUB.\n"
            "Control+Shift+K: Send selected book to Kindle.\n"
            "Control+Shift+V: Send selected book to Voice Dream Loader folder.\n"
            "Use File, Send To, Dolphin EasyReader, to copy the selected book or selected books to a Dolphin EasyReader import or library folder.\n"
            "Control+Shift+E: Send selected book to an NLS eReader if it is connected.\n"
            "Use File, Send To, HumanWare Braille eReader MTP, for HumanWare devices that appear under This PC but do not have a normal drive letter.\n"
            "Control+Space: Select or unselect the current book for batch actions. Kindle, Export, and Delete use selected books when any are selected.\n"
            "Control+A: Select all shown books for batch actions.\n"
            "Control+Shift+A: Deselect all books selected for batch actions.\n"
            "Control+K: Open Kindle for PC.\n"
            "Delete: Remove selected book from library.\n"
            "Control+F: Search library metadata.\n"
            "Escape: Clear the current search and return to the full book list.\n"
            "Home and End: Move to the first or last shown book.\n"
            "Control+I: Show book details.\n"
            "In the books list, Alt+1 reads title, Alt+2 reads author, Alt+3 reads series, Alt+4 reads series number, Alt+5 reads edition, Alt+6 reads year, Alt+7 reads ISBN, Alt+8 reads publisher, Alt+9 reads source, and Alt+0 reads tags. Press the same Alt+number twice quickly to edit that field when it is editable.\n"
            "Use Organize, Sort, to sort title, author, or series A to Z or Z to A, and to sort published year or date added newest to oldest or oldest to newest.\n"
            "Use Organize, Filter, to filter by source, tag, or format, or to clear filters.\n"
            "Use Organize, Saved Views, to open built-in views like Missing Metadata, Recently Added, EPUB Books, and Bookshare Books, or to save and reopen your current search/filter/sort view.\n"
            "Use Organize, Remove Duplicates Prefer EPUB, to remove likely duplicate library entries while keeping an EPUB version when one exists.\n"
            "Use Organize, Toggle Automatic Duplicate Cleanup, to turn startup and periodic duplicate cleanup on or off. Automatic cleanup removes duplicate library entries only, not book files.\n"
            "Use Settings, Book List Reading, to choose title only, title and author, title author and edition, or all details.\n"
            "Use Settings, Missing Metadata Alert Sound, to choose whether the alert means missing author only, missing useful textbook details, or more complete metadata.\n"
            "Use Settings, Library Backup, to choose a Google Drive, OneDrive, iCloud Drive, or other synced folder for database and imported book file backups. You can back up on demand, daily, weekly, or monthly, and restore from the cloud backup if the local library is lost.\n"
            "Use Settings, Watched Folders, to add folders that are scanned automatically for new or changed ebook files.\n"
            "Use Settings, File Explorer Integration, to add or remove a Windows right-click command for adding supported files directly from File Explorer.\n"
            "If NVDA is running and its controller is available, the app automatically uses NVDA book list announcements. No setting is needed.\n"
            "Use Settings, Kindle and Calibre, to check whether Calibre's background command-line tools are available. The manager can use those tools for Kindle metadata without opening Calibre's interface.\n"
            "Use Settings, Set Kindle Email Addresses, to save more than one Send to Kindle address.\n"
            "Use File, Send To, NLS eReader, to copy the selected book to a connected NLS eReader. If the app cannot detect it, you can choose the eReader folder manually.\n"
            "Use File, Send To, HumanWare Braille eReader MTP, when Windows shows the device under This PC but File Explorer does not give it a pasteable folder path.\n"
            "Control+L: Move to the books list.\n"
            "Use Up and Down Arrow in the books list to choose a book.\n"
            "In the books list, press a letter to jump to the next title starting with that letter. Leading The and A are ignored for this jump.\n"
            "F1: Help.\n\n"
            "The book list is not a table. Each book is one list item. By default it reads the title and author without saying the labels Title and Author. You can make it shorter or more detailed in Settings. The app automatically focuses the books list when it opens and after searches/imports.\n\n"
            "EPUB metadata changes are written into the EPUB file itself, with a .bak backup made first. "
            "Other file formats are updated in the library database only. This app does not remove DRM."
        )

    def help_file_path(self):
        return app_data_folder() / "Accessible Ebook Library Manager Help.txt"

    def show_help(self):
        help_path = self.help_file_path()
        try:
            help_path.parent.mkdir(parents=True, exist_ok=True)
            help_path.write_text(self.help_text(), encoding="utf-8")
            if sys.platform.startswith("win"):
                os.startfile(help_path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(help_path)])
            else:
                subprocess.Popen(["xdg-open", str(help_path)])
            self.status_var.set(f"Help opened at {help_path}.")
        except Exception as exc:
            messagebox.showerror(
                "Help could not be opened",
                f"I could not open the help file in the default text editor.\n\n"
                f"Help file path:\n{help_path}\n\n{exc}"
            )


def _setup_logging():
    log_path = app_data_folder() / "crash_log.txt"
    handler = RotatingFileHandler(log_path, maxBytes=1_000_000, backupCount=3, encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(threadName)s %(message)s"))
    logging.basicConfig(level=logging.WARNING, handlers=[handler])


def main():
    _setup_logging()
    try:
        root = Tk()
        LibraryApp(root)
        root.mainloop()
    except Exception:
        logger.exception("Fatal application crash")
        try:
            log_path = app_data_folder() / "crash_log.txt"
            messagebox.showerror(
                "Application error",
                f"The app hit an unexpected error.\n\nCrash details saved at:\n{log_path}"
            )
        except Exception:
            raise


if __name__ == "__main__":
    main()
