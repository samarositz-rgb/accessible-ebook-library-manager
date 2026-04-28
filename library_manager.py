"""
Accessible Ebook Library Manager
A screen-reader-friendly starter ebook manager for Windows.

Features:
- Standard Tkinter controls.
- Windows-style menu bar for screen-reader friendly command access.
- Plain listbox instead of a table so each book is spoken as one complete labeled row.
- Add EPUB, PDF, DOCX, TXT, MOBI, AZW, AZW3, HTML, ZIP, and other ebook/document files.
- Stores metadata in a simple SQLite database.
- For EPUB files, writes title, author, source, tags, and notes into the EPUB file itself.
- Copies imported books into a managed library folder.
- Search by title, author, source, tags, format, or notes.
- Opens selected book with the default Windows app.
- Opens Kindle for PC.
- Exports a selected book to another folder.

This app does not remove DRM. It manages files you are allowed to copy and read.
"""

import os
import shutil
import sqlite3
import subprocess
import sys
import traceback
import zipfile
import re
import html
import tkinter as tk
import json
import tempfile
import time
import urllib.parse
import urllib.request
import winsound
import ctypes
from datetime import datetime, timedelta
from pathlib import Path
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


APP_NAME = "Accessible Ebook Library Manager"
DB_NAME = "library.db"
BOOKS_FOLDER = "Books"

SUPPORTED_EXTENSIONS = {
    ".epub", ".pdf", ".docx", ".doc", ".txt", ".rtf",
    ".mobi", ".azw", ".azw3", ".html", ".htm", ".zip"
}

BOOK_LIST_SPEECH_FIELDS = [
    ("title", "Title"),
    ("author", "Author"),
    ("edition", "Edition"),
    ("year", "Year"),
    ("isbn", "ISBN"),
    ("publisher", "Publisher"),
    ("source", "Source"),
    ("tags", "Tags"),
    ("format", "Format"),
    ("added_at", "Date added"),
]
DEFAULT_BOOK_LIST_SPEECH_FIELDS = ["title", "author"]
BACKUP_SCHEDULES = {
    "on_demand": ("On Demand", None),
    "daily": ("Daily", timedelta(days=1)),
    "weekly": ("Weekly", timedelta(days=7)),
    "monthly": ("Monthly", timedelta(days=30)),
}
DEFAULT_BACKUP_SCHEDULE = "on_demand"
MISSING_METADATA_SOUND_MODES = {
    "author": ("Missing Author Only", ["author"]),
    "useful": ("Missing Author, Edition, or Year", ["author", "edition", "year"]),
    "complete": ("Missing Author, Edition, Year, ISBN, or Publisher", ["author", "edition", "year", "isbn", "publisher"]),
}
DEFAULT_MISSING_METADATA_SOUND_MODE = "author"

OPF_NS = "http://www.idpf.org/2007/opf"
DC_NS = "http://purl.org/dc/elements/1.1/"
CONTAINER_NS = "urn:oasis:names:tc:opendocument:xmlns:container"

ET.register_namespace("opf", OPF_NS)
ET.register_namespace("dc", DC_NS)

WINDOWS_NO_CONSOLE_FLAGS = getattr(subprocess, "CREATE_NO_WINDOW", 0)


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
    return datetime.utcnow().replace(microsecond=0).isoformat() + "Z"


def parse_utc_text(value):
    if not value:
        return None
    try:
        return datetime.fromisoformat(value.replace("Z", ""))
    except ValueError:
        return None


def cloud_backup_subfolder(base_folder: Path) -> Path:
    return base_folder / "Accessible Ebook Library Manager Backups"


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
        self.connection.commit()
        self.ensure_book_columns()

    def ensure_book_columns(self):
        """Add newer metadata columns to existing libraries without deleting data."""
        existing = {row[1] for row in self.connection.execute("PRAGMA table_info(books)")}
        columns = {
            "edition": "TEXT DEFAULT ''",
            "year": "TEXT DEFAULT ''",
            "isbn": "TEXT DEFAULT ''",
            "publisher": "TEXT DEFAULT ''",
            "cover_url": "TEXT DEFAULT ''",
        }
        for name, column_type in columns.items():
            if name not in existing:
                self.connection.execute(f"ALTER TABLE books ADD COLUMN {name} {column_type}")
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

    def add_book(self, title, author, source, tags, notes, original_path, stored_path):
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

    def update_cover_url(self, book_id, cover_url=""):
        self.connection.execute(
            "UPDATE books SET cover_url = ? WHERE id = ?",
            (cover_url, book_id),
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
                if os.path.exists(stored_path):
                    os.remove(stored_path)
            except OSError:
                pass
        self.connection.execute("DELETE FROM books WHERE id = ?", (book_id,))
        self.connection.commit()

    def all_books_for_duplicate_check(self):
        cursor = self.connection.execute(
            """
            SELECT id, title, author, source, tags, notes, format, original_path, stored_path, added_at,
                   edition, year, isbn, publisher, cover_url
            FROM books
            ORDER BY title COLLATE NOCASE, author COLLATE NOCASE, added_at
            """
        )
        return cursor.fetchall()

    def get_book(self, book_id):
        cursor = self.connection.execute(
            """
            SELECT id, title, author, source, tags, notes, format, original_path, stored_path, added_at,
                   edition, year, isbn, publisher, cover_url
            FROM books WHERE id = ?
            """,
            (book_id,),
        )
        return cursor.fetchone()

    def search_books(self, query="", sort_by="title", source_filter="", tag_filter="", format_filter=""):
        query = query.strip()
        source_filter = source_filter.strip()
        tag_filter = tag_filter.strip()
        format_filter = format_filter.strip().lower()

        order_map = {
            "title": "title COLLATE NOCASE, author COLLATE NOCASE",
            "author": "author COLLATE NOCASE, title COLLATE NOCASE",
            "date": "year DESC, title COLLATE NOCASE",
            "date_added": "added_at DESC, title COLLATE NOCASE",
        }
        order_clause = order_map.get(sort_by, order_map["title"])

        conditions = []
        params = []

        if query:
            like = f"%{query}%"
            conditions.append(
                """(
                   title LIKE ?
                   OR author LIKE ?
                   OR source LIKE ?
                   OR tags LIKE ?
                   OR notes LIKE ?
                   OR format LIKE ?
                   OR edition LIKE ?
                   OR year LIKE ?
                   OR isbn LIKE ?
                   OR publisher LIKE ?
                )"""
            )
            params.extend([like] * 10)

        if source_filter:
            conditions.append("source LIKE ?")
            params.append(f"%{source_filter}%")

        if tag_filter:
            conditions.append("tags LIKE ?")
            params.append(f"%{tag_filter}%")

        if format_filter:
            conditions.append("format LIKE ?")
            params.append(f"%{format_filter}%")

        where_clause = ""
        if conditions:
            where_clause = "WHERE " + " AND ".join(conditions)

        cursor = self.connection.execute(
            f"""
            SELECT id, title, author, source, tags, format, added_at,
                   edition, year, isbn, publisher
            FROM books
            {where_clause}
            ORDER BY {order_clause}
            """,
            params,
        )
        return cursor.fetchall()

def safe_filename(text: str) -> str:
    bad = '<>:"/\\|?*'
    cleaned = "".join("_" if ch in bad else ch for ch in text).strip()
    return cleaned or "Untitled"


def normalize_duplicate_key(text: str) -> str:
    text = (text or "").casefold()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    text = re.sub(r"\b(?:the|a|an)\b", " ", text)
    return re.sub(r"\s+", " ", text).strip()


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

        return {
            "title": first_or_empty(root, ".//dc:title", ns),
            "author": first_or_empty(root, ".//dc:creator", ns),
            "source": first_or_empty(root, ".//dc:source", ns),
            "tags": ", ".join(subjects),
            "notes": first_or_empty(root, ".//dc:description", ns),
            "publisher": first_or_empty(root, ".//dc:publisher", ns),
            "year": first_or_empty(root, ".//dc:date", ns)[:4],
        }
    except Exception:
        return {}


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


def strip_xml_html_tags(text: str) -> str:
    text = re.sub(r"<[^>]+>", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def read_text_from_epub(epub_path: Path, max_chars: int = 50000) -> str:
    chunks = []
    try:
        with zipfile.ZipFile(epub_path, "r") as archive:
            names = archive.namelist()
            text_names = [
                name for name in names
                if name.lower().endswith((".xhtml", ".html", ".htm", ".xml"))
                and "nav" not in name.lower()
                and "toc" not in name.lower()
                and "container.xml" not in name.lower()
            ]

            for name in text_names[:35]:
                try:
                    raw = archive.read(name)
                    decoded = raw.decode("utf-8", errors="ignore")
                    cleaned = strip_xml_html_tags(decoded)
                    if cleaned:
                        chunks.append(cleaned)
                    if sum(len(chunk) for chunk in chunks) >= max_chars:
                        break
                except Exception:
                    continue
    except Exception:
        return ""

    return "\n".join(chunks)[:max_chars]


def read_text_from_docx(docx_path: Path, max_chars: int = 50000) -> str:
    try:
        with zipfile.ZipFile(docx_path, "r") as archive:
            raw = archive.read("word/document.xml")
        decoded = raw.decode("utf-8", errors="ignore")
        return strip_xml_html_tags(decoded)[:max_chars]
    except Exception:
        return ""


def read_text_from_plain_file(path: Path, max_chars: int = 50000) -> str:
    try:
        raw = path.read_bytes()[:max_chars * 2]
        return raw.decode("utf-8", errors="ignore")[:max_chars]
    except Exception:
        return ""


def read_metadata_text_from_pdf(pdf_path: Path, max_chars: int = 12000) -> str:
    try:
        raw = pdf_path.read_bytes()[: max_chars * 20]
    except Exception:
        return ""

    text = raw.decode("latin-1", errors="ignore")
    lines = []
    for key in ["Title", "Author", "Subject", "Keywords", "Creator", "Producer"]:
        for match in re.finditer(rf"/{key}\s*\((.*?)\)", text, flags=re.DOTALL):
            value = match.group(1)
            value = value.replace(r"\(", "(").replace(r"\)", ")").replace(r"\\", "\\")
            value = re.sub(r"\s+", " ", value).strip()
            if value:
                lines.append(f"{key}: {value}")
    return "\n".join(lines)[:max_chars]


def read_text_for_metadata_detection(path: Path, max_chars: int = 50000) -> str:
    suffix = path.suffix.lower()
    if suffix == ".epub":
        return read_text_from_epub(path, max_chars=max_chars)
    if suffix == ".docx":
        return read_text_from_docx(path, max_chars=max_chars)
    if suffix == ".pdf":
        return read_metadata_text_from_pdf(path, max_chars=max_chars)
    if suffix in {".txt", ".rtf", ".html", ".htm"}:
        return read_text_from_plain_file(path, max_chars=max_chars)
    return ""


def clean_metadata_line(line: str) -> str:
    line = line.strip()
    line = re.sub(r"\s+", " ", line)
    line = re.sub(r"^[#*\\-–—: ]+", "", line)
    return line.strip()


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
    value = re.sub(r"^(?:by|author|authors|written by|edited by)\s*[:\-]?\s*", "", value, flags=re.IGNORECASE)
    value = re.sub(r"\s+(?:narrated by|read by|bookshare|copyright|all rights reserved).*$", "", value, flags=re.IGNORECASE)
    value = re.sub(r"\s+", " ", value).strip(" .;,")
    return value


def clean_title_value(value: str) -> str:
    value = clean_metadata_line(value)
    value = re.sub(r"^(?:title|book title)\s*[:\-]?\s*", "", value, flags=re.IGNORECASE)
    value = re.sub(r"\s+", " ", value).strip(" .;,")
    return value


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
    if any(word in lower for word in ["chapter", "contents", "copyright", "isbn", "publisher", "bookshare"]):
        return False
    words = [word for word in re.split(r"\s+", value) if word]
    return 1 <= len(words) <= 8 and len(value) <= 120


def fetch_json(url: str, timeout: int = 12) -> dict:
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": "AccessibleEbookLibraryManager/1.0",
            "Accept": "application/json",
        },
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8", errors="ignore"))


def normalize_online_metadata(metadata: dict) -> dict:
    result = {
        "title": metadata.get("title", ""),
        "author": metadata.get("author", ""),
        "edition": metadata.get("edition", ""),
        "year": metadata.get("year", ""),
        "isbn": metadata.get("isbn", ""),
        "publisher": metadata.get("publisher", ""),
        "source": metadata.get("source", ""),
        "tags": metadata.get("tags", ""),
        "notes": metadata.get("notes", ""),
        "cover_url": metadata.get("cover_url", ""),
    }
    for key, value in list(result.items()):
        if isinstance(value, list):
            result[key] = ", ".join(str(item) for item in value if item)
        elif value is None:
            result[key] = ""
        else:
            result[key] = str(value).strip()
    if result["year"]:
        match = re.search(r"\b(1[5-9]\d{2}|20\d{2})\b", result["year"])
        if match:
            result["year"] = match.group(1)
    return result


def metadata_from_open_library_doc(doc: dict) -> dict:
    isbn_values = doc.get("isbn") or []
    tags = doc.get("subject") or []
    cover_id = doc.get("cover_i")
    cover_url = f"https://covers.openlibrary.org/b/id/{cover_id}-L.jpg" if cover_id else ""
    return normalize_online_metadata({
        "title": doc.get("title", ""),
        "author": ", ".join(doc.get("author_name") or []),
        "year": str(doc.get("first_publish_year") or ""),
        "isbn": isbn_values[0] if isbn_values else "",
        "publisher": (doc.get("publisher") or [""])[0],
        "tags": ", ".join(tags[:8]),
        "cover_url": cover_url,
    })


def metadata_from_google_volume(volume: dict) -> dict:
    info = volume.get("volumeInfo", {})
    identifiers = info.get("industryIdentifiers") or []
    isbn = ""
    for identifier in identifiers:
        if identifier.get("type") == "ISBN_13":
            isbn = identifier.get("identifier", "")
            break
    if not isbn and identifiers:
        isbn = identifiers[0].get("identifier", "")
    categories = info.get("categories") or []
    image_links = info.get("imageLinks") or {}
    cover_url = image_links.get("large") or image_links.get("medium") or image_links.get("thumbnail") or ""
    return normalize_online_metadata({
        "title": info.get("title", ""),
        "author": ", ".join(info.get("authors") or []),
        "year": info.get("publishedDate", ""),
        "isbn": isbn,
        "publisher": info.get("publisher", ""),
        "tags": ", ".join(categories[:8]),
        "notes": info.get("description", ""),
        "cover_url": cover_url,
    })


def lookup_online_metadata(existing: dict) -> tuple[dict, str]:
    isbn = re.sub(r"[^0-9Xx]", "", existing.get("isbn", ""))
    title = existing.get("title", "").strip()
    author = existing.get("author", "").strip()

    queries = []
    if isbn:
        queries.append(("openlibrary", f"https://openlibrary.org/search.json?isbn={urllib.parse.quote(isbn)}&limit=1"))
        queries.append(("google", f"https://www.googleapis.com/books/v1/volumes?q=isbn:{urllib.parse.quote(isbn)}&maxResults=1"))
    if title:
        open_query = f"title:{title}"
        google_query = f"intitle:{title}"
        if author:
            open_query += f" author:{author}"
            google_query += f"+inauthor:{author}"
        queries.append(("openlibrary", "https://openlibrary.org/search.json?" + urllib.parse.urlencode({"q": open_query, "limit": "1"})))
        queries.append(("google", "https://www.googleapis.com/books/v1/volumes?" + urllib.parse.urlencode({"q": google_query, "maxResults": "1"})))

    last_error = ""
    for service, url in queries:
        try:
            data = fetch_json(url)
            if service == "openlibrary":
                docs = data.get("docs") or []
                if docs:
                    return metadata_from_open_library_doc(docs[0]), "Open Library"
            else:
                items = data.get("items") or []
                if items:
                    return metadata_from_google_volume(items[0]), "Google Books"
        except Exception as exc:
            last_error = str(exc)
            continue

    if last_error:
        raise RuntimeError(last_error)
    return {}, ""


def detect_metadata_from_text(path: Path, existing: dict | None = None) -> dict:
    """Best-effort local metadata detector.

    This is not a true AI model. It uses EPUB metadata, file names, and common
    title/author patterns found near the start of the book text. It is designed
    to help with sources such as Bookshare, where file names are often awkward.
    """
    existing = existing or {}
    result = {
        "title": existing.get("title", ""),
        "author": existing.get("author", ""),
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

    text = read_text_for_metadata_detection(path)
    lines = [clean_metadata_line(line) for line in re.split(r"[\r\n]+", text)]
    lines = [line for line in lines if line and len(line) < 240 and not is_bookshare_notice_line(line)]

    filename_title = clean_filename_title(path)
    if is_weak_title(result["title"], path):
        result["title"] = filename_title

    # Look harder for labeled Bookshare/front-matter fields.
    title_value = line_after_label(lines, ["title", "book title", "name"], max_index=180)
    if title_value and is_weak_title(result["title"], path):
        result["title"] = clean_title_value(title_value)

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
                if is_weak_title(result["title"], path) and possible_title:
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
            if candidate:
                result["title"] = candidate
            break

    # If there is still no good title, use the first substantial line that is not a boilerplate line.
    boilerplate_words = [
        "bookshare", "copyright", "all rights reserved", "dedication",
        "contents", "table of contents", "chapter", "isbn", "published"
    ]
    if is_weak_title(result["title"], path):
        for line in lines[:100]:
            lower = line.lower()
            if len(line) < 4:
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
        isbn_matches = re.findall(
            r"\b(?:ISBN(?:-1[03])?:?\s*)?((?:97[89][-\s]?)?\d[-\s]?\d{2,5}[-\s]?\d{2,7}[-\s]?\d{1,7}[-\s]?[\dXx])\b",
            text,
            flags=re.IGNORECASE,
        )
        for isbn in isbn_matches:
            cleaned_isbn = re.sub(r"[^0-9Xx]", "", isbn)
            if len(cleaned_isbn) in {10, 13}:
                result["isbn"] = cleaned_isbn
                break

    if text and not result.get("year"):
        year_match = re.search(
            r"\b(?:copyright|published|publication date|date|year)?\s*(19[5-9]\d|20[0-4]\d)\b",
            text[:12000],
            flags=re.IGNORECASE,
        )
        if year_match:
            result["year"] = year_match.group(1)

    if text and not result.get("publisher"):
        publisher_value = line_after_label(
            lines,
            ["publisher", "published by", "imprint"],
            max_index=220,
        )
        if publisher_value:
            result["publisher"] = clean_metadata_line(publisher_value)
        if not result.get("publisher"):
            for line in lines[:180]:
                if re.search(r"\b(press|publishing|publishers|books|house|pearson|mcgraw|cengage|wiley|openstax|scholastic|harper|penguin|random house|simon|houghton|macmillan|oxford|cambridge)\b", line, flags=re.IGNORECASE):
                    result["publisher"] = line
                    break

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

    # Source guess.
    combined = (str(path) + " " + text[:1000]).lower()
    if not result["source"]:
        if "bookshare" in combined:
            result["source"] = "Bookshare"
        elif "kindle" in combined or path.suffix.lower() in {".azw", ".azw3", ".mobi"}:
            result["source"] = "Kindle"
        else:
            result["source"] = "Personal"

    # Tags guess.
    if not result["tags"]:
        tags = []
        keyword_value = line_after_label(lines, ["keywords", "subjects", "subject"], max_index=120)
        if keyword_value:
            tags.extend([part.strip() for part in re.split(r"[,;]", keyword_value) if part.strip()])
        if "bookshare" in combined:
            tags.append("Bookshare")
        if path.suffix.lower() == ".epub":
            tags.append("EPUB")
        elif path.suffix.lower() == ".docx":
            tags.append("DOCX")
        elif path.suffix.lower() == ".pdf":
            tags.append("PDF")
        deduped_tags = []
        for tag in tags:
            if tag and tag.lower() not in [existing_tag.lower() for existing_tag in deduped_tags]:
                deduped_tags.append(tag)
        result["tags"] = ", ".join(deduped_tags)

    # Do not auto-fill notes. Notes remain available for manual editing only.
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


def ask_metadata_with_windows_forms(heading, fields, initial, initial_focus_key="title"):
    if not sys.platform.startswith("win") or not shutil.which("powershell"):
        return None

    initial = initial or {}
    payload = {
        "fields": [{"key": key, "label": label, "hint": hint} for key, label, hint in fields],
        "values": {key: str(initial.get(key, "") or "") for key, _label, _hint in fields},
        "initialFocusKey": initial_focus_key or "title",
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
$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = "Cancel"
$cancelButton.Width = 90
$cancelButton.TabIndex = $tabIndex + 1

$buttonPanel.Controls.Add($cancelButton)
$buttonPanel.Controls.Add($saveButton)

$saveButton.Add_Click({
    $titleBox = $textBoxes["title"]
    if ($null -eq $titleBox -or [string]::IsNullOrWhiteSpace($titleBox.Text)) {
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
        ("edition", "Edition", "For example third edition or revised edition."),
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
        self.marked_book_ids = set()
        self.last_missing_metadata_sound_book_id = None
        self.nvda_controller = None
        self.nvda_controller_checked = False
        self.last_nvda_announcement = ""
        self.last_alt_number_key = ""
        self.last_alt_number_time = 0.0
        self.shortcut_readout_return_after = None
        self.sort_by = "title"
        self.filter_source = ""
        self.filter_tag = ""
        self.filter_format = ""
        self.backup_check_after = None

        self.build_menu()
        self.build_ui()
        self.refresh_books()
        self.schedule_backup_check(5000)

    def build_menu(self):
        menu_bar = Menu(self.root)

        file_menu = Menu(menu_bar, tearoff=False)
        file_menu.add_command(label="Open Book\tCtrl+O", command=self.open_book)
        file_menu.add_command(label="Add Book...\tCtrl+N", command=self.add_book)
        file_menu.add_command(label="Import Folder...\tCtrl+Shift+N", command=self.import_folder)
        file_menu.add_command(label="Export Copy...\tCtrl+E", command=self.export_book)
        send_to_menu = Menu(file_menu, tearoff=False)
        send_to_menu.add_command(label="Voice Dream...\tCtrl+Shift+V", command=self.send_to_voice_dream)
        send_to_menu.add_command(label="Kindle...\tCtrl+Shift+K", command=self.send_to_kindle)
        send_to_menu.add_command(label="NLS eReader...\tCtrl+Shift+E", command=self.send_to_nls_ereader)
        send_to_menu.add_command(label="HumanWare Braille eReader (MTP)...", command=self.send_to_humanware_mtp)
        file_menu.add_cascade(label="Send To", menu=send_to_menu)
        file_menu.add_separator()
        file_menu.add_command(label="Exit\tAlt+F4", command=self.root.quit)
        menu_bar.add_cascade(label="File", menu=file_menu, underline=0)

        book_menu = Menu(menu_bar, tearoff=False)
        book_menu.add_command(label="Edit Metadata...\tF2", command=self.edit_book)
        book_menu.add_command(label="Auto-Detect Metadata...\tCtrl+D", command=self.auto_detect_selected_metadata)
        book_menu.add_command(label="Look Up Book Metadata from Internet...", command=self.lookup_selected_metadata_online)
        book_menu.add_command(label="View Cover Image...", command=self.view_selected_cover_image)
        book_menu.add_command(label="Convert to EPUB...\tCtrl+R", command=self.convert_selected_to_epub)
        book_menu.add_command(label="Show Selected Book Information\tCtrl+I", command=self.show_selected_book_info)
        book_menu.add_command(label="Read Current Book\tCtrl+Shift+I", command=self.read_current_book)
        book_menu.add_command(label="Focus Books List\tCtrl+L", command=self.focus_books_list)
        book_menu.add_command(label="Deselect All Books\tCtrl+Shift+A", command=self.deselect_all_books)
        book_menu.add_command(label="Delete from Library\tDelete", command=self.delete_book)
        menu_bar.add_cascade(label="Book", menu=book_menu, underline=0)

        organize_menu = Menu(menu_bar, tearoff=False)
        organize_menu.add_command(label="Sort by Title", command=lambda: self.set_sort("title"))
        organize_menu.add_command(label="Sort by Author", command=lambda: self.set_sort("author"))
        organize_menu.add_command(label="Sort by Published Year", command=lambda: self.set_sort("date"))
        organize_menu.add_command(label="Sort by Date Added", command=lambda: self.set_sort("date_added"))
        organize_menu.add_separator()
        organize_menu.add_command(label="Filter by Source...", command=self.set_source_filter)
        organize_menu.add_command(label="Filter by Tag...", command=self.set_tag_filter)
        organize_menu.add_command(label="Filter by Format...", command=self.set_format_filter)
        organize_menu.add_command(label="Clear Filters", command=self.clear_filters)
        organize_menu.add_separator()
        organize_menu.add_command(label="Remove Duplicates, Prefer EPUB...", command=self.remove_duplicates_prefer_epub)
        organize_menu.add_command(label="Show Current Organize Settings", command=self.show_organize_settings)
        menu_bar.add_cascade(label="Organize", menu=organize_menu, underline=0)

        search_menu = Menu(menu_bar, tearoff=False)
        search_menu.add_command(label="Search Metadata...\tCtrl+F", command=self.focus_search)
        search_menu.add_command(label="Move to Books List\tCtrl+L", command=self.focus_books_list)
        search_menu.add_command(label="Clear Search", command=self.clear_search)
        search_menu.add_command(label="Explain Search", command=self.explain_search)
        menu_bar.add_cascade(label="Search", menu=search_menu, underline=0)

        settings_menu = Menu(menu_bar, tearoff=False)
        speech_menu = Menu(settings_menu, tearoff=False)
        speech_menu.add_command(label="Title Only", command=lambda: self.set_book_list_speech_fields(["title"]))
        speech_menu.add_command(label="Title and Author", command=lambda: self.set_book_list_speech_fields(["title", "author"]))
        speech_menu.add_command(label="Title, Author, and Edition", command=lambda: self.set_book_list_speech_fields(["title", "author", "edition"]))
        speech_menu.add_command(label="Full Details", command=lambda: self.set_book_list_speech_fields([key for key, _label in BOOK_LIST_SPEECH_FIELDS]))
        speech_menu.add_command(label="Show Current Speech Details", command=self.show_book_list_speech_fields)
        settings_menu.add_cascade(label="Book List Speech", menu=speech_menu)
        settings_menu.add_separator()
        missing_metadata_menu = Menu(settings_menu, tearoff=False)
        missing_metadata_menu.add_command(label="Off", command=lambda: self.set_missing_metadata_sound_mode("off"))
        missing_metadata_menu.add_command(label="Missing Author Only", command=lambda: self.set_missing_metadata_sound_mode("author"))
        missing_metadata_menu.add_command(label="Missing Author, Edition, or Year", command=lambda: self.set_missing_metadata_sound_mode("useful"))
        missing_metadata_menu.add_command(label="Missing Author, Edition, Year, ISBN, or Publisher", command=lambda: self.set_missing_metadata_sound_mode("complete"))
        missing_metadata_menu.add_separator()
        missing_metadata_menu.add_command(label="Show Current Setting", command=self.show_missing_metadata_sound_mode)
        missing_metadata_menu.add_command(label="Test Sound", command=self.test_missing_metadata_sound)
        settings_menu.add_cascade(label="Missing Metadata Sound", menu=missing_metadata_menu)
        settings_menu.add_separator()
        backup_menu = Menu(settings_menu, tearoff=False)
        backup_menu.add_command(label="Use OneDrive Folder...", command=lambda: self.choose_cloud_backup_folder("onedrive"))
        backup_menu.add_command(label="Use Google Drive Folder...", command=lambda: self.choose_cloud_backup_folder("google_drive"))
        backup_menu.add_command(label="Use iCloud Drive Folder...", command=lambda: self.choose_cloud_backup_folder("icloud"))
        backup_menu.add_command(label="Choose Other Backup Folder...", command=lambda: self.choose_cloud_backup_folder("other"))
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
        settings_menu.add_command(label="Toggle NVDA Book List Announcements", command=self.toggle_nvda_book_list_announcements)
        settings_menu.add_command(label="Set Voice Dream Loader Folder...", command=self.choose_voice_dream_folder)
        settings_menu.add_command(label="Set NLS eReader Folder...", command=self.choose_nls_ereader_folder)
        settings_menu.add_command(label="Set Kindle Email Addresses...", command=self.set_kindle_email)
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

        ttk.Label(search_frame, text="Search metadata").pack(side=LEFT, padx=(0, 8))
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

        self.book_list = Listbox(list_frame, selectmode="browse", height=20, exportselection=False, takefocus=1)
        scrollbar = Scrollbar(list_frame, orient="vertical", command=self.book_list.yview)
        self.book_list.configure(yscrollcommand=scrollbar.set)

        self.book_list.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)

        self.book_list.bind("<Return>", lambda event: self.open_book())
        self.book_list.bind("<Alt-Return>", lambda event: self.open_book())
        self.book_list.bind("<Delete>", lambda event: self.delete_book())
        self.book_list.bind("<F2>", lambda event: self.edit_book())
        self.book_list.bind("<Control-e>", lambda event: self.export_book())
        self.book_list.bind("<Control-V>", lambda event: self.send_to_voice_dream())
        self.book_list.bind("<Control-o>", lambda event: self.open_book())
        self.book_list.bind("<Control-k>", lambda event: self.open_kindle())
        self.book_list.bind("<Control-r>", lambda event: self.convert_selected_to_epub())
        self.book_list.bind("<Control-K>", lambda event: self.send_to_kindle())
        self.book_list.bind("<Control-E>", lambda event: self.send_to_nls_ereader())
        self.book_list.bind("<Control-n>", lambda event: self.add_book())
        self.book_list.bind("<Control-f>", self.focus_search_from_keyboard)
        self.book_list.bind("<Control-i>", lambda event: self.show_selected_book_info())
        self.book_list.bind("<Control-I>", lambda event: self.read_current_book())
        self.book_list.bind("<Control-d>", lambda event: self.auto_detect_selected_metadata())
        self.book_list.bind("<Control-A>", lambda event: self.deselect_all_books())
        self.book_list.bind("<Control-space>", self.toggle_mark_current_book)
        self.book_list.bind("<Escape>", self.clear_search_from_keyboard)
        self.bind_alt_number_shortcuts(self.book_list)
        self.book_list.bind("<<ListboxSelect>>", self.on_book_list_select)
        self.book_list.bind("<KeyPress>", self.on_book_list_keypress)

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
        self.root.bind("<Control-I>", lambda event: self.read_current_book())
        self.root.bind("<Control-d>", lambda event: self.auto_detect_selected_metadata())
        self.root.bind("<Control-l>", lambda event: self.focus_books_list())
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

        status = ttk.Label(main, textvariable=self.status_var, relief="sunken", anchor="w")
        status.pack(fill=X, pady=(8, 0))

        self.book_list.focus_set()

    def bind_alt_number_shortcuts(self, widget):
        for digit in "1234567890":
            widget.bind(f"<Alt-KeyPress-{digit}>", self.on_book_list_alt_number)

    def sort_label(self):
        labels = {
            "title": "Title",
            "author": "Author",
            "date": "Published Year",
            "date_added": "Date Added",
        }
        return labels.get(self.sort_by, "Title")

    def active_filter_summary(self):
        parts = []
        if self.filter_source:
            parts.append(f"source contains {self.filter_source}")
        if self.filter_tag:
            parts.append(f"tag contains {self.filter_tag}")
        if self.filter_format:
            parts.append(f"format contains {self.filter_format}")
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
            "Enter source text to filter by, such as Bookshare, Kindle, Personal, or leave blank to clear source filter.",
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
            "Enter tag text to filter by, such as textbook, fiction, unread, or leave blank to clear tag filter.",
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
            "Enter format text to filter by, such as epub, pdf, docx, or leave blank to clear format filter.",
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
        self.refresh_books()
        self.focus_books_list()
        self.status_var.set("Filters cleared.")

    def show_organize_settings(self):
        messagebox.showinfo(
            "Current Organize Settings",
            f"Sort: {self.sort_label()}\n"
            f"Filters: {self.active_filter_summary()}\n"
            f"Books shown: {self.book_list.size()}"
        )

    def duplicate_group_key(self, row):
        title_key = normalize_duplicate_key(row[1])
        author_key = normalize_duplicate_key(row[2])
        isbn_key = re.sub(r"[^0-9Xx]", "", row[12] or "").casefold()
        if len(isbn_key) in {10, 13}:
            return ("isbn", isbn_key)
        if title_key and author_key:
            return ("title_author", title_key, author_key)
        return None

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

    def duplicate_groups(self):
        grouped = {}
        for row in self.db.all_books_for_duplicate_check():
            key = self.duplicate_group_key(row)
            if key is None:
                continue
            grouped.setdefault(key, []).append(row)
        return [rows for rows in grouped.values() if len(rows) > 1]

    def duplicate_removal_plan(self):
        plan = []
        for rows in self.duplicate_groups():
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
                "No duplicate books were found using title plus author, or ISBN when available."
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
        self.status_var.set(f"Book list speech details saved: {self.book_list_speech_summary()}.")

    def show_book_list_speech_fields(self):
        messagebox.showinfo(
            "Current Book List Speech Details",
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
        self.status_var.set(f"Missing metadata sound set to {label}.")
        messagebox.showinfo("Missing Metadata Sound", f"Missing metadata sound is now set to:\n\n{label}")
        if mode != "off":
            self.play_missing_metadata_sound()

    def show_missing_metadata_sound_mode(self):
        messagebox.showinfo(
            "Missing Metadata Sound",
            f"Current setting:\n\n{self.missing_metadata_sound_mode_label()}"
        )

    def toggle_missing_metadata_sound(self):
        enabled = not self.missing_metadata_sound_enabled()
        self.set_missing_metadata_sound_mode(DEFAULT_MISSING_METADATA_SOUND_MODE if enabled else "off")
        state = "on" if enabled else "off"
        self.status_var.set(f"Missing metadata sound turned {state}.")

    def test_missing_metadata_sound(self):
        self.play_missing_metadata_sound()
        messagebox.showinfo(
            "Missing Metadata Sound",
            "If your Windows system sounds are enabled, you should have heard the missing metadata sound."
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
        try:
            winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS | winsound.SND_ASYNC)
        except Exception:
            try:
                winsound.MessageBeep()
            except Exception:
                try:
                    winsound.Beep(880, 70)
                except Exception:
                    pass

    def nvda_book_list_announcements_enabled(self):
        return self.db.get_setting("nvda_book_list_announcements", "1") == "1"

    def toggle_nvda_book_list_announcements(self):
        enabled = not self.nvda_book_list_announcements_enabled()
        self.db.set_setting("nvda_book_list_announcements", "1" if enabled else "0")
        state = "on" if enabled else "off"
        self.status_var.set(f"NVDA book list announcements turned {state}.")
        messagebox.showinfo("NVDA Book List Announcements", f"NVDA book list announcements are now {state}.")

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
            text = self.book_list.get(index)
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
            selected_index = self.current_book_index()
            self.shortcut_readout_var.set(text)
            self.shortcut_readout.focus_force()
            self.shortcut_readout.selection_range(0, END)
            self.shortcut_readout.icursor(END)
            if self.shortcut_readout_return_after is not None:
                self.root.after_cancel(self.shortcut_readout_return_after)
            self.shortcut_readout_return_after = self.root.after(
                1600,
                lambda index=selected_index: self.return_focus_from_shortcut_readout(index),
            )
        except Exception:
            pass

    def return_focus_from_shortcut_readout(self, selected_index=None):
        self.shortcut_readout_return_after = None
        try:
            if self.root.focus_get() == self.shortcut_readout:
                self.settle_book_list_focus(selected_index)
        except Exception:
            pass

    def navigation_title_key(self, title):
        title = re.sub(r"\s+", " ", title or "").strip()
        title = re.sub(r"^(?:the|a)\s+", "", title, flags=re.IGNORECASE)
        return title.casefold()

    def format_book_list_row(self, row):
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
        ) = row
        values = {
            "title": title or "Untitled",
            "author": author or "Unknown",
            "edition": edition or "Not specified",
            "year": year or "Not specified",
            "isbn": isbn or "Not specified",
            "publisher": publisher or "Not specified",
            "source": source or "Not specified",
            "tags": tags or "None",
            "format": book_format or "Unknown",
            "added_at": added_at or "Unknown",
        }
        labels = dict(BOOK_LIST_SPEECH_FIELDS)
        parts = []
        for key in self.get_book_list_speech_fields():
            if key in {"title", "author", "edition"}:
                parts.append(f"{values[key]}.")
            else:
                parts.append(f"{labels[key]}: {values[key]}.")
        prefix = "Selected. " if _book_id in self.marked_book_ids else ""
        return prefix + " ".join(parts)

    def book_field_values_for_shortcuts(self, row):
        labels = dict(BOOK_LIST_SPEECH_FIELDS)
        return {
            "title": (labels["title"], row[1] or "Untitled"),
            "author": (labels["author"], row[2] or "Unknown"),
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
                f"I found this {label} folder:\n\n{detected}\n\nUse it for library database backups?"
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
            f"Backups will be saved here:\n\n{backup_folder}\n\nSet a schedule or choose Back Up Now from Settings, Library Backup."
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
        return db_mtime != backed_mtime and datetime.utcnow() - last_backup >= interval

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
        if not folder:
            if automatic:
                return
            messagebox.showinfo(
                "Choose backup folder",
                "Choose a Google Drive, OneDrive, iCloud Drive, or other synced folder before backing up."
            )
            self.choose_cloud_backup_folder("other")
            folder, backup_file, manifest_file = self.backup_paths()
            if not folder:
                return

        try:
            folder.mkdir(parents=True, exist_ok=True)
            self.db.backup_to(backup_file)
            db_mtime = str(self.db.db_path.stat().st_mtime)
            book_count = self.db.connection.execute("SELECT COUNT(*) FROM books").fetchone()[0]
            manifest = {
                "app": APP_NAME,
                "created_at": utc_now_text(),
                "source_database": str(self.db.db_path),
                "backup_database": str(backup_file),
                "source_database_mtime": db_mtime,
                "book_count": book_count,
            }
            manifest_file.write_text(json.dumps(manifest, indent=2), encoding="utf-8")
            self.db.set_setting("last_backup_at", manifest["created_at"])
            self.db.set_setting("last_backup_db_mtime", db_mtime)
            self.db.set_setting("last_seen_backup_file_mtime", str(backup_file.stat().st_mtime))
            message = f"Library database backed up to {backup_file}."
            self.status_var.set(message)
            if not automatic:
                messagebox.showinfo("Library Backup Complete", message)
        except Exception as exc:
            self.status_var.set("Library backup failed.")
            if not automatic:
                messagebox.showerror("Library Backup Failed", f"Could not back up the library database.\n\n{exc}")

    def backup_file_is_valid(self, backup_file):
        try:
            connection = sqlite3.connect(backup_file)
            try:
                row = connection.execute(
                    "SELECT name FROM sqlite_master WHERE type='table' AND name='books'"
                ).fetchone()
                return row is not None
            finally:
                connection.close()
        except Exception:
            return False

    def restore_library_backup(self):
        folder, backup_file, _manifest_file = self.backup_paths()
        if not folder:
            messagebox.showinfo("No backup folder", "Choose a backup folder before restoring.")
            self.choose_cloud_backup_folder("other")
            folder, backup_file, _manifest_file = self.backup_paths()
            if not folder:
                return
        if not backup_file.exists():
            messagebox.showerror("No backup found", f"No library backup was found here:\n\n{backup_file}")
            return
        if not self.backup_file_is_valid(backup_file):
            messagebox.showerror("Backup not valid", "The backup file does not look like an Accessible Ebook Library Manager database.")
            return

        if not messagebox.askyesno(
            "Restore Library Backup",
            "Restore the cloud backup over the current local library database?\n\n"
            "The current local database will be saved as a safety copy first."
        ):
            self.focus_books_list()
            return

        preserved_folder = self.db.get_setting("backup_folder", "")
        preserved_schedule = self.db.get_setting("backup_schedule", DEFAULT_BACKUP_SCHEDULE)
        safety_copy = self.db.folder / f"library_before_restore_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
        try:
            self.db.connection.commit()
            shutil.copy2(self.db.db_path, safety_copy)
            self.db.close()
            shutil.copy2(backup_file, self.db.db_path)
            self.db = LibraryDatabase()
            self.db.set_setting("backup_folder", preserved_folder)
            self.db.set_setting("backup_schedule", preserved_schedule)
            self.refresh_books()
            self.settle_book_list_focus()
            message = f"Library restored from backup. Safety copy saved at {safety_copy}."
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
        manifest = self.backup_manifest()
        text = (
            f"Backup folder: {folder or 'Not set'}\n"
            f"Schedule: {self.backup_schedule_label()}\n"
            f"Last backup: {self.db.get_setting('last_backup_at', 'Never') or 'Never'}\n"
            f"Backup file: {backup_file if backup_file and backup_file.exists() else 'Not found'}\n"
            f"Manifest file: {manifest_file if manifest_file and manifest_file.exists() else 'Not found'}\n"
            f"Books in last manifest: {manifest.get('book_count', 'Unknown')}"
        )
        messagebox.showinfo("Library Backup Status", text)

    def focus_search(self):
        value = AccessibleSingleFieldDialog.ask(
            self.root,
            "Search Metadata",
            "Enter text to search in title, author, source, tags, format, notes, edition, year, ISBN, and publisher. Leave blank to show all books.",
            self.search_var.get(),
            heading="Search Metadata",
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
        self.book_list.focus_set()
        if self.book_list.size() > 0 and not self.book_list.curselection():
            self.book_list.selection_set(0)
            self.book_list.activate(0)
        self.status_var.set("Books list focused. Use up and down arrow to choose a book.")

    def settle_book_list_focus(self, selected_index=None):
        if self.book_list.size() == 0:
            self.book_list.focus_force()
            return
        if selected_index is None:
            selected = self.book_list.curselection()
            if selected:
                selected_index = selected[0]
            else:
                try:
                    selected_index = self.book_list.index("active")
                except Exception:
                    selected_index = 0
        selected_index = max(0, min(selected_index, self.book_list.size() - 1))
        self.book_list.focus_set()
        self.select_book_list_index(selected_index)

    def select_book_list_index(self, index):
        if self.book_list.size() == 0:
            return
        index = max(0, min(index, self.book_list.size() - 1))
        book_id = self.book_list_ids[index] if index < len(self.book_list_ids) else None
        self.book_list.selection_clear(0, END)
        self.book_list.selection_set(index)
        self.book_list.activate(index)
        self.book_list.see(index)
        self.book_list.focus_set()
        if book_id is not None:
            prefix = "Selected. " if book_id in self.marked_book_ids else ""
            self.status_var.set(prefix + self.book_list.get(index))
            self.play_missing_metadata_sound_if_needed(book_id)
            self.announce_current_book_to_nvda(index)
        else:
            self.status_var.set(self.book_list.get(index))

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
        self.root.after(1, self.sound_for_current_selection)
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

        if field_index < 0 or field_index >= len(BOOK_LIST_SPEECH_FIELDS):
            return "break"

        book_id = self.selected_book_id()
        if book_id is None:
            return "break"

        row = self.db.get_book(book_id)
        if not row:
            messagebox.showerror("Not found", "The selected book was not found.")
            return "break"

        field_key, _default_label = BOOK_LIST_SPEECH_FIELDS[field_index]
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
        selected = self.book_list.curselection()
        if selected:
            index = selected[0]
        else:
            try:
                index = self.book_list.index("active")
            except Exception:
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
            "Search metadata",
            "Search looks through library metadata only: title, author, source, tags, format, notes, edition, year, ISBN, and publisher.\n\n"
            "It does not search inside the full text of every book yet."
        )

    def search_and_focus(self):
        self.refresh_books()
        if self.book_list.size() > 0:
            self.root.after(50, self.settle_book_list_focus)
            self.root.after(200, self.settle_book_list_focus)
        else:
            self.book_list.focus_force()
            self.status_var.set("No books matched the search.")

    def refresh_books(self, selected_book_id=None):
        self.book_list.delete(0, END)
        self.book_list_ids = []
        self.book_list_titles = []

        rows = self.db.search_books(
            self.search_var.get(),
            sort_by=self.sort_by,
            source_filter=self.filter_source,
            tag_filter=self.filter_tag,
            format_filter=self.filter_format,
        )
        for row in rows:
            book_id = row[0]
            self.book_list.insert(END, self.format_book_list_row(row))
            self.book_list_ids.append(book_id)
            self.book_list_titles.append(self.navigation_title_key(row[1]))

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

    def current_book_index(self):
        """Return the active list index, and make it the selection."""
        if self.book_list.size() == 0:
            return None

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

    def read_current_book(self):
        text = self.current_book_list_text()
        if not text:
            messagebox.showinfo("No books", "There are no books in the current list.")
            return
        self.status_var.set(text)
        messagebox.showinfo("Current book", text)

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
                f"Edition: {row[10] or 'Not specified'}\n"
                f"Year: {row[11] or 'Not specified'}\n"
                f"ISBN: {row[12] or 'Not specified'}\n"
                f"Publisher: {row[13] or 'Not specified'}\n"
                f"Cover image: {'Available' if len(row) > 14 and row[14] else 'Not saved'}\n"
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
        try:
            with open(self.crash_log_path(), "a", encoding="utf-8") as log:
                log.write("\n" + "=" * 60 + "\n")
                log.write(f"Context: {context}\n")
                log.write(traceback.format_exc())
                log.write("\n")
        except Exception:
            pass

    def safe_message_error(self, title, message):
        try:
            messagebox.showerror(title, message)
        except Exception:
            pass

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
        ]

        epubs = [path for path in supported_inside if path.suffix.lower() == ".epub"]
        if epubs:
            return epubs

        return supported_inside

    def import_zip_file_without_prompt(self, zip_path: Path, default_source: str = "") -> tuple[int, list[str]]:
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
                if self.import_one_book_without_prompt(inner_file, default_source=default_source):
                    imported += 1
                else:
                    skipped.append(f"{inner_file} -- unsupported or not imported")
            except sqlite3.IntegrityError:
                skipped.append(f"{inner_file} -- already imported or duplicate stored path")
            except Exception as exc:
                skipped.append(f"{inner_file} -- {exc}")

        return imported, skipped

    def guess_metadata_from_file(self, source_path: Path) -> dict:
        default_title = clean_filename_title(source_path)
        metadata = {"title": default_title, "author": "", "edition": "", "year": "", "isbn": "", "publisher": "", "source": "", "tags": "", "notes": ""}

        if source_path.suffix.lower() == ".epub":
            epub_metadata = read_epub_metadata(source_path)
            metadata.update({key: value for key, value in epub_metadata.items() if value})

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

    def import_one_book_without_prompt(self, source_path: Path, default_source: str = "") -> bool:
        if source_path.suffix.lower() == ".zip":
            imported, skipped = self.import_zip_file_without_prompt(source_path, default_source=default_source)
            if imported:
                return True
            raise OSError("; ".join(skipped) if skipped else "ZIP contained no importable books")

        ready, reason = self.is_file_ready_for_import(source_path)
        if not ready:
            raise OSError(reason)

        metadata = self.guess_metadata_from_file(source_path)
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

        new_book_id = self.db.add_book(
            metadata["title"],
            metadata["author"],
            metadata["source"],
            metadata["tags"],
            metadata["notes"],
            str(source_path),
            str(destination),
        )
        self.db.update_extra_fields(
            new_book_id,
            metadata.get("edition", ""),
            metadata.get("year", ""),
            metadata.get("isbn", ""),
            metadata.get("publisher", ""),
        )
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

            try:
                if include_subfolders:
                    candidates_iter = root_folder.rglob("*")
                else:
                    candidates_iter = root_folder.iterdir()

                candidates = []
                skipped_items = []
                for candidate in candidates_iter:
                    try:
                        if candidate.is_file():
                            candidates.append(candidate)
                    except Exception as exc:
                        skipped_items.append(f"{candidate} -- could not inspect file: {exc}")

            except Exception as exc:
                self.log_error("Listing import folder", exc)
                messagebox.showerror(
                    "Folder import failed",
                    f"I could not read this folder.\n\n{exc}\n\nCrash details saved at:\n{self.crash_log_path()}"
                )
                return

            supported = []
            for path in candidates:
                try:
                    if path.suffix.lower() in SUPPORTED_EXTENSIONS:
                        supported.append(path)
                except Exception as exc:
                    skipped_items.append(f"{path} -- could not check extension: {exc}")

            if not supported:
                report_path = self.write_import_report(0, skipped_items)
                messagebox.showinfo(
                    "No supported books found",
                    f"No supported ebook or document files were found in that folder.\n\nReport saved at:\n{report_path}"
                )
                return

            if not messagebox.askyesno(
                "Confirm folder import",
                f"Found {len(supported)} supported file{'s' if len(supported) != 1 else ''}. Import them now?"
            ):
                return

            imported = 0

            for path in supported:
                try:
                    # Keep UI responsive enough during large imports.
                    self.status_var.set(f"Importing: {path.name}")
                    self.root.update_idletasks()

                    if path.suffix.lower() == ".zip":
                        zip_imported, zip_skipped = self.import_zip_file_without_prompt(path, default_source=default_source.strip())
                        imported += zip_imported
                        skipped_items.extend(zip_skipped)
                    elif self.import_one_book_without_prompt(path, default_source=default_source.strip()):
                        imported += 1
                    else:
                        skipped_items.append(f"{path} -- unsupported or not imported")
                except sqlite3.IntegrityError as exc:
                    skipped_items.append(f"{path} -- already imported or duplicate stored path")
                    self.log_error(f"Duplicate during folder import: {path}", exc)
                except Exception as exc:
                    skipped_items.append(f"{path} -- {exc}")
                    self.log_error(f"Importing file: {path}", exc)

            report_path = self.write_import_report(imported, skipped_items)

            try:
                self.refresh_books()
                if imported:
                    self.focus_books_list()
            except Exception as exc:
                self.log_error("Refreshing after folder import", exc)

            skipped = len(skipped_items)
            messagebox.showinfo(
                "Folder import complete",
                f"Imported {imported} book{'s' if imported != 1 else ''}. "
                f"Skipped {skipped} file{'s' if skipped != 1 else ''}.\n\n"
                f"Import report saved at:\n{report_path}\n\n"
                f"If the app crashed before, check:\n{self.crash_log_path()}"
            )
            self.status_var.set(f"Folder import complete. Imported {imported}. Skipped {skipped}.")
        except Exception as exc:
            self.log_error("Fatal folder import error", exc)
            messagebox.showerror(
                "Folder import crashed",
                f"Folder import hit an unexpected error, but the app caught it this time.\n\n{exc}\n\n"
                f"Crash details saved at:\n{self.crash_log_path()}"
            )

    def add_book(self):
        paths = filedialog.askopenfilenames(
            title="Choose books to add",
            filetypes=[
                ("Ebook and document files", "*.epub *.pdf *.docx *.doc *.txt *.rtf *.mobi *.azw *.azw3 *.html *.htm *.zip"),
                ("All files", "*.*"),
            ],
        )
        if not paths:
            return

        added = 0
        for path in paths:
            source_path = Path(path)

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

            default_title = clean_filename_title(source_path)
            initial_metadata = {"title": default_title, "author": "", "source": "", "tags": "", "notes": ""}

            if source_path.suffix.lower() == ".epub":
                epub_metadata = read_epub_metadata(source_path)
                initial_metadata.update({key: value for key, value in epub_metadata.items() if value})

            initial_metadata = detect_metadata_from_text(source_path, existing=initial_metadata)

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
                str(source_path),
                str(destination),
            )
            self.db.update_extra_fields(
                new_book_id,
                metadata.get("edition", ""),
                metadata.get("year", ""),
                metadata.get("isbn", ""),
                metadata.get("publisher", ""),
            )
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

    def row_to_metadata(self, row):
        return {
            "title": row[1],
            "author": row[2],
            "edition": row[10],
            "year": row[11],
            "isbn": row[12],
            "publisher": row[13],
            "source": row[3],
            "tags": row[4],
            "notes": row[5],
            "cover_url": row[14] if len(row) > 14 else "",
        }

    def merge_online_metadata(self, existing, online, replace_existing=False):
        merged = dict(existing)
        for key in ["title", "author", "edition", "year", "isbn", "publisher", "tags", "notes", "cover_url"]:
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
        if metadata.get("cover_url"):
            self.db.update_cover_url(book_id, metadata.get("cover_url", ""))

        stored_path = Path(row[8])
        if stored_path.suffix.lower() == ".epub" and stored_path.exists():
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
            messagebox.showerror("Not found", "The selected book was not found.")
            return

        stored_path = Path(row[8])
        if not stored_path.exists():
            messagebox.showerror("File missing", "The stored book file could not be found.")
            return

        existing = self.row_to_metadata(row)

        detected = detect_metadata_from_text(stored_path, existing=existing)

        summary = (
            "Detected metadata:\n\n"
            f"Title: {detected.get('title', '')}\n"
            f"Author: {detected.get('author', '')}\n"
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
        self.refresh_books(selected_book_id=book_id)
        self.focus_books_list()

    def lookup_selected_metadata_online(self):
        book_id = self.selected_book_id()
        if book_id is None:
            return

        row = self.db.get_book(book_id)
        if not row:
            messagebox.showerror("Not found", "The selected book was not found.")
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
            messagebox.showerror("Not found", "The selected book was not found.")
            return

        cover_url = row[14] if len(row) > 14 else ""
        if not cover_url:
            if not messagebox.askyesno(
                "No cover saved",
                "No cover image is saved for this book yet. Look online now?"
            ):
                return
            try:
                online, service = lookup_online_metadata(self.row_to_metadata(row))
            except Exception as exc:
                messagebox.showerror("Cover lookup failed", f"Could not look up a cover image online.\n\n{exc}")
                return
            cover_url = online.get("cover_url", "")
            if not cover_url:
                messagebox.showinfo("No cover found", "No cover image was found online for this book.")
                return
            self.db.update_cover_url(book_id, cover_url)
            self.status_var.set(f"Cover image found from {service}.")

        try:
            import webbrowser
            webbrowser.open(cover_url)
            self.status_var.set("Opened cover image.")
        except Exception as exc:
            messagebox.showerror("Open cover failed", f"Could not open the cover image.\n\n{exc}")

    def edit_book(self, initial_focus_field="title"):
        book_id = self.selected_book_id()
        if book_id is None:
            return

        row = self.db.get_book(book_id)
        if not row:
            messagebox.showerror("Not found", "The selected book was not found.")
            return

        try:
            metadata = TkMetadataDialog.ask(
                self.root,
                "Edit Book Metadata",
                {
                    "title": row[1],
                    "author": row[2],
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

        stored_path = Path(row[8])
        if stored_path.suffix.lower() == ".epub" and stored_path.exists():
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
            messagebox.showerror("Not found", "The selected book was not found in the library database.")
            return

        path = str(row[8])
        if not os.path.exists(path):
            messagebox.showerror(
                "File missing",
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
                    messagebox.showerror("Open Kindle failed", str(exc))
                    return

        try:
            subprocess.Popen("start kindle:", shell=True)
            self.status_var.set("Tried to open Kindle for PC.")
        except Exception:
            messagebox.showerror(
                "Kindle not found",
                "I could not find Kindle for PC. Install Kindle for PC, then try again."
            )

    def find_ebook_convert(self):
        found = shutil.which("ebook-convert")
        if found:
            return found

        possible_paths = []
        if sys.platform.startswith("win"):
            possible_paths.extend([
                Path(os.environ.get("PROGRAMFILES", "")) / "Calibre2" / "ebook-convert.exe",
                Path(os.environ.get("PROGRAMFILES(X86)", "")) / "Calibre2" / "ebook-convert.exe",
                Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "Calibre2" / "ebook-convert.exe",
            ])
        elif sys.platform == "darwin":
            possible_paths.append(Path("/Applications/calibre.app/Contents/MacOS/ebook-convert"))

        for candidate in possible_paths:
            if candidate.exists():
                return str(candidate)
        return None

    def convert_selected_to_epub(self):
        book_id = self.selected_book_id()
        if book_id is None:
            return

        row = self.db.get_book(book_id)
        if not row:
            messagebox.showerror("Not found", "The selected book was not found.")
            return

        source_file = Path(row[8])
        if not source_file.exists():
            messagebox.showerror("File missing", "The stored book file could not be found.")
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
            messagebox.showerror("Conversion failed", str(exc))
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
            messagebox.showerror("File missing", "The stored book file could not be found.")
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
            messagebox.showerror("Send to Kindle failed", str(exc))

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
            messagebox.showerror("Open library folder failed", str(exc))

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
            messagebox.showerror("Not found", "The selected book was not found.")
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
            messagebox.showerror("Not found", "The selected book was not found.")
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
        book_id = self.selected_book_id()
        if book_id is None:
            return

        row = self.db.get_book(book_id)
        if not row:
            messagebox.showerror("Not found", "The selected book was not found.")
            return

        source = Path(row[8])
        if not source.exists():
            messagebox.showerror(
                "File missing",
                "The stored book file could not be found, so it cannot be sent to Voice Dream."
            )
            return

        loader_folder = self.get_voice_dream_folder()
        if not loader_folder:
            return

        destination_folder = Path(loader_folder)
        if not destination_folder.exists():
            messagebox.showerror(
                "Folder missing",
                "The Voice Dream Loader folder does not exist. Please choose it again."
            )
            self.db.set_setting("voice_dream_loader_folder", "")
            return

        destination = destination_folder / source.name
        if destination.exists():
            replace = messagebox.askyesno(
                "Replace existing file",
                f"{destination.name} already exists in the Voice Dream Loader folder. Replace it?"
            )
            if not replace:
                return

        try:
            shutil.copy2(source, destination)
        except Exception as exc:
            messagebox.showerror(
                "Send to Voice Dream failed",
                f"Could not copy the book to the Voice Dream Loader folder.\n\n{exc}"
            )
            return

        self.status_var.set(f"Sent to Voice Dream: {source.name}")
        messagebox.showinfo(
            "Sent to Voice Dream",
            f"Copied this book to the Voice Dream Loader folder:\n\n{destination}"
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

        exported = 0
        skipped = []
        for row in rows:
            source = Path(row[8])
            if not source.exists():
                skipped.append(f"{row[1]}: stored file missing")
                continue
            destination = Path(folder) / source.name
            if destination.exists():
                if not messagebox.askyesno("Replace file", f"{destination.name} already exists. Replace it?"):
                    skipped.append(f"{row[1]}: skipped because the export file already exists")
                    continue
            try:
                shutil.copy2(source, destination)
                exported += 1
            except Exception as exc:
                skipped.append(f"{row[1]}: {exc}")

        self.status_var.set(f"Exported {exported} book{'s' if exported != 1 else ''}.")
        if skipped:
            messagebox.showwarning(
                "Export finished with warnings",
                f"Exported {exported} book{'s' if exported != 1 else ''}.\n\n"
                + "\n".join(skipped[:20])
            )
        else:
            messagebox.showinfo("Export complete", f"Exported {exported} book{'s' if exported != 1 else ''}.")

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

    def show_help(self):
        messagebox.showinfo(
            "Help",
            "Accessible Ebook Library Manager keyboard commands:\n\n"
            "Alt: Open the menu bar.\n"
            "Control+N: Add book.\n"
            "Control+Shift+N: Import a folder of books, including Bookshare ZIP files.\n"
            "F2: Edit selected book metadata. On Windows, the metadata editor uses native edit boxes so screen readers can read field names, contents, and typed text. Use Tab and Shift+Tab to move between fields.\n"
            "Control+D: Auto-detect metadata from the selected book.\n"
            "Use the Book menu, Look Up Book Metadata from Internet, to search Open Library and Google Books for metadata.\n"
            "Use the Book menu, View Cover Image, to open a visual cover image when one is available or look one up online.\n"
            "Enter or Control+O: Open selected book.\n"
            "Control+E: Export selected book.\n"
            "Control+R: Convert selected book to EPUB.\n"
            "Control+Shift+K: Send selected book to Kindle.\n"
            "Control+Shift+V: Send selected book to Voice Dream Loader folder.\n"
            "Control+Shift+E: Send selected book to an NLS eReader if it is connected.\n"
            "Use File, Send To, HumanWare Braille eReader MTP, for HumanWare devices that appear under This PC but do not have a normal drive letter.\n"
            "Control+Space: Select or unselect the current book for batch actions. Kindle, Export, and Delete use selected books when any are selected.\n"
            "Control+Shift+A: Deselect all books selected for batch actions.\n"
            "Control+K: Open Kindle for PC.\n"
            "Delete: Remove selected book from library.\n"
            "Control+F: Search metadata.\n"
            "Escape: Clear the current search and return to the full book list.\n"
            "Control+I: Show selected book information.\n"
            "Control+Shift+I: Read the current book list item.\n"
            "In the books list, Alt+1 reads title, Alt+2 reads author, Alt+3 reads edition, Alt+4 reads year, Alt+5 reads ISBN, Alt+6 reads publisher, Alt+7 reads source, Alt+8 reads tags, Alt+9 reads format, and Alt+0 reads date added. Press the same Alt+number twice quickly to edit that field when it is editable.\n"
            "Use the Organize menu to sort by title, author, published year, or date added, and to filter by source, tag, or format.\n"
            "Use Organize, Remove Duplicates Prefer EPUB, to remove likely duplicate library entries while keeping an EPUB version when one exists.\n"
            "Use Settings, Book List Speech, to choose title only, title and author, title author and edition, or full details.\n"
            "Use Settings, Missing Metadata Sound, to choose whether the alert means missing author only, missing useful textbook details, or more complete metadata.\n"
            "Use Settings, Library Backup, to choose a Google Drive, OneDrive, iCloud Drive, or other synced folder for database backups. You can back up on demand, daily, weekly, or monthly, and restore from the cloud backup if the local database is lost.\n"
            "Use Settings, Toggle NVDA Book List Announcements, if NVDA does not automatically read book list rows.\n"
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


def main():
    try:
        root = Tk()
        LibraryApp(root)
        root.mainloop()
    except Exception:
        try:
            folder = app_data_folder()
            log_path = folder / "crash_log.txt"
            with open(log_path, "a", encoding="utf-8") as log:
                log.write("\n" + "=" * 60 + "\n")
                log.write("Fatal application crash\n")
                log.write(traceback.format_exc())
                log.write("\n")
            messagebox.showerror(
                "Application error",
                f"The app hit an unexpected error.\n\nCrash details saved at:\n{log_path}"
            )
        except Exception:
            raise


if __name__ == "__main__":
    main()
