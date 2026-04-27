"""
Accessible Ebook Library Manager
A JAWS-friendly starter ebook manager for Windows.

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
import zipfile
import re
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

OPF_NS = "http://www.idpf.org/2007/opf"
DC_NS = "http://purl.org/dc/elements/1.1/"
CONTAINER_NS = "urn:oasis:names:tc:opendocument:xmlns:container"

ET.register_namespace("opf", OPF_NS)
ET.register_namespace("dc", DC_NS)


def app_data_folder() -> Path:
    base = os.environ.get("APPDATA")
    if base:
        folder = Path(base) / "AccessibleEbookLibraryManager"
    else:
        folder = Path.home() / "AccessibleEbookLibraryManager"
    folder.mkdir(parents=True, exist_ok=True)
    (folder / BOOKS_FOLDER).mkdir(parents=True, exist_ok=True)
    return folder


class LibraryDatabase:
    def __init__(self):
        self.folder = app_data_folder()
        self.db_path = self.folder / DB_NAME
        self.books_path = self.folder / BOOKS_FOLDER
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

    def get_book(self, book_id):
        cursor = self.connection.execute(
            """
            SELECT id, title, author, source, tags, notes, format, original_path, stored_path, added_at,
                   edition, year, isbn, publisher
            FROM books WHERE id = ?
            """,
            (book_id,),
        )
        return cursor.fetchone()

    def search_books(self, query=""):
        query = query.strip()
        if not query:
            cursor = self.connection.execute(
                """
                SELECT id, title, author
                FROM books
                ORDER BY author COLLATE NOCASE, title COLLATE NOCASE
                """
            )
        else:
            like = f"%{query}%"
            cursor = self.connection.execute(
                """
                SELECT id, title, author
                FROM books
                WHERE title LIKE ?
                   OR author LIKE ?
                   OR source LIKE ?
                   OR tags LIKE ?
                   OR notes LIKE ?
                   OR format LIKE ?
                   OR edition LIKE ?
                   OR year LIKE ?
                   OR isbn LIKE ?
                   OR publisher LIKE ?
                ORDER BY author COLLATE NOCASE, title COLLATE NOCASE
                """,
                (like, like, like, like, like, like, like, like, like, like),
            )
        return cursor.fetchall()


def safe_filename(text: str) -> str:
    bad = '<>:"/\\|?*'
    cleaned = "".join("_" if ch in bad else ch for ch in text).strip()
    return cleaned or "Untitled"


def first_or_empty(root, xpath, namespaces):
    item = root.find(xpath, namespaces)
    if item is None or item.text is None:
        return ""
    return item.text.strip()


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


def read_text_from_epub(epub_path: Path, max_chars: int = 12000) -> str:
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

            for name in text_names[:12]:
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


def read_text_from_docx(docx_path: Path, max_chars: int = 12000) -> str:
    try:
        with zipfile.ZipFile(docx_path, "r") as archive:
            raw = archive.read("word/document.xml")
        decoded = raw.decode("utf-8", errors="ignore")
        return strip_xml_html_tags(decoded)[:max_chars]
    except Exception:
        return ""


def read_text_from_plain_file(path: Path, max_chars: int = 12000) -> str:
    try:
        raw = path.read_bytes()[:max_chars * 2]
        return raw.decode("utf-8", errors="ignore")[:max_chars]
    except Exception:
        return ""


def read_text_for_metadata_detection(path: Path, max_chars: int = 12000) -> str:
    suffix = path.suffix.lower()
    if suffix == ".epub":
        return read_text_from_epub(path, max_chars=max_chars)
    if suffix == ".docx":
        return read_text_from_docx(path, max_chars=max_chars)
    if suffix in {".txt", ".rtf", ".html", ".htm"}:
        return read_text_from_plain_file(path, max_chars=max_chars)
    return ""


def clean_metadata_line(line: str) -> str:
    line = line.strip()
    line = re.sub(r"\s+", " ", line)
    line = re.sub(r"^[#*\\-–—: ]+", "", line)
    return line.strip()


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
    lines = [line for line in lines if line and len(line) < 200]

    # Look for common author patterns.
    author_patterns = [
        r"^by\s+(.+)$",
        r"^author[:\s]+(.+)$",
        r"^written by\s+(.+)$",
    ]
    for line in lines[:80]:
        for pattern in author_patterns:
            match = re.match(pattern, line, flags=re.IGNORECASE)
            if match and not result["author"]:
                result["author"] = clean_metadata_line(match.group(1))
                break
        if result["author"]:
            break

    # Look for common title labels.
    for line in lines[:80]:
        match = re.match(r"^title[:\s]+(.+)$", line, flags=re.IGNORECASE)
        if match:
            result["title"] = clean_metadata_line(match.group(1))
            break

    # If there is still no good title, use the first substantial line that is not a boilerplate line.
    boilerplate_words = [
        "bookshare", "copyright", "all rights reserved", "dedication",
        "contents", "table of contents", "chapter", "isbn", "published"
    ]
    if not result["title"] or result["title"].lower() == path.stem.lower():
        for line in lines[:60]:
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
        result["title"] = path.stem.replace("_", " ").replace("-", " ")

    # ISBN, year, publisher, and edition guesses.
    if text and not result.get("isbn"):
        isbn_match = re.search(r"\b(?:ISBN(?:-1[03])?:?\s*)?((?:97[89][-\s]?)?\d[-\s]?\d{2,5}[-\s]?\d{2,7}[-\s]?\d{1,7}[-\s]?[\dXx])\b", text)
        if isbn_match:
            result["isbn"] = re.sub(r"[^0-9Xx]", "", isbn_match.group(1))

    if text and not result.get("year"):
        year_match = re.search(r"\b(19[5-9]\d|20[0-4]\d)\b", text[:4000])
        if year_match:
            result["year"] = year_match.group(1)

    if text and not result.get("publisher"):
        for line in lines[:120]:
            match = re.match(r"^publisher[:\s]+(.+)$", line, flags=re.IGNORECASE)
            if match:
                result["publisher"] = clean_metadata_line(match.group(1))
                break
        if not result.get("publisher"):
            for line in lines[:120]:
                if re.search(r"\b(press|publishing|publishers|pearson|mcgraw|cengage|wiley|openstax)\b", line, flags=re.IGNORECASE):
                    result["publisher"] = line
                    break

    if text and not result.get("edition"):
        for line in lines[:120]:
            match = re.search(r"\b(\d+(?:st|nd|rd|th)\s+edition|first edition|second edition|third edition|fourth edition|fifth edition|sixth edition|seventh edition|eighth edition|ninth edition|tenth edition|revised edition|international edition)\b", line, flags=re.IGNORECASE)
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
        if "bookshare" in combined:
            tags.append("Bookshare")
        if path.suffix.lower() == ".epub":
            tags.append("EPUB")
        elif path.suffix.lower() == ".docx":
            tags.append("DOCX")
        elif path.suffix.lower() == ".pdf":
            tags.append("PDF")
        result["tags"] = ", ".join(tags)

    # Do not auto-fill notes. Notes remain available for manual editing only.
    return result



class AccessibleSingleFieldDialog:
    """One accessible edit dialog.

    The edit box itself contains a prefix such as "Title: current value".
    This works around cases where JAWS does not announce the separate label.
    """

    def __init__(self, parent, field_label: str, instructions: str, current_value: str):
        import tkinter as tk

        self.result = None
        self.field_label = field_label
        self.window = tk.Toplevel(parent)
        self.window.title(field_label)
        self.window.transient(parent)
        self.window.grab_set()

        main = ttk.Frame(self.window, padding=12)
        main.pack(fill=BOTH, expand=True)

        prompt_text = (
            f"{field_label}. {instructions} "
            "The edit box starts with the field name. "
            "Press Enter to accept. Press Escape to cancel."
        )
        ttk.Label(main, text=prompt_text, wraplength=600).pack(anchor="w", pady=(0, 8))

        self.value_var = StringVar(value=f"{field_label}: {current_value}")
        self.entry = ttk.Entry(main, textvariable=self.value_var, width=70)
        self.entry.pack(fill=X, expand=True, pady=(0, 8))

        button_frame = ttk.Frame(main)
        button_frame.pack(anchor="e")

        ok_button = ttk.Button(button_frame, text="OK", command=self.ok)
        ok_button.pack(side=LEFT, padx=(0, 8))
        cancel_button = ttk.Button(button_frame, text="Cancel", command=self.cancel)
        cancel_button.pack(side=LEFT)

        self.window.bind("<Return>", lambda event: self.ok())
        self.window.bind("<Escape>", lambda event: self.cancel())
        self.window.protocol("WM_DELETE_WINDOW", self.cancel)

        self.window.after(100, self.focus_entry)

    def focus_entry(self):
        self.entry.focus_force()
        prefix = f"{self.field_label}: "
        text = self.value_var.get()
        start = len(prefix) if text.startswith(prefix) else 0
        self.entry.selection_range(start, END)
        self.entry.icursor(END)

    def ok(self):
        value = self.value_var.get().strip()
        prefix = f"{self.field_label}:"
        if value.lower().startswith(prefix.lower()):
            value = value[len(prefix):].strip()
        self.result = value
        self.window.destroy()

    def cancel(self):
        self.result = None
        self.window.destroy()

    @staticmethod
    def ask(parent, field_label: str, instructions: str, current_value: str):
        dialog = AccessibleSingleFieldDialog(parent, field_label, instructions, current_value)
        parent.wait_window(dialog.window)
        return dialog.result


class TkMetadataDialog:
    """Accessible step-by-step metadata editor."""

    @staticmethod
    def ask(parent, heading="Book Metadata", initial=None):
        initial = initial or {}

        fields = [
            ("title", "Title", "Enter the book title."),
            ("author", "Author", "Enter the author."),
            ("edition", "Edition", "Enter the edition, for example 3rd edition or Revised edition."),
            ("year", "Year", "Enter the publication year."),
            ("isbn", "ISBN", "Enter the ISBN."),
            ("publisher", "Publisher", "Enter the publisher."),
            ("source", "Source", "Enter the source, for example Bookshare, Kindle, Personal, or Web."),
            ("tags", "Tags", "Enter tags, separated by commas."),
            ("notes", "Notes", "Enter notes or description."),
        ]

        result = {}

        messagebox.showinfo(
            heading,
            "You will now edit metadata one field at a time. "
            "Each edit box begins with the field name, such as Title or Author. "
            "Press Enter to accept the current value and move to the next field."
        )

        for key, label, instructions in fields:
            current_value = initial.get(key, "")
            while True:
                value = AccessibleSingleFieldDialog.ask(parent, label, instructions, current_value)

                if value is None:
                    if messagebox.askyesno(
                        "Cancel metadata editing",
                        "Cancel metadata editing and discard changes?"
                    ):
                        return None
                    continue

                value = value.strip()

                if key == "title" and not value:
                    messagebox.showerror("Missing title", "Title is required.")
                    current_value = value
                    continue

                result[key] = value
                break

        return result

class LibraryApp:
    def __init__(self, root):
        self.root = root
        self.db = LibraryDatabase()
        self.root.title(APP_NAME)
        self.root.geometry("1000x600")

        self.search_var = StringVar()
        self.status_var = StringVar(value="Ready")
        self.book_list_ids = []

        self.build_menu()
        self.build_ui()
        self.refresh_books()

    def build_menu(self):
        menu_bar = Menu(self.root)

        file_menu = Menu(menu_bar, tearoff=False)
        file_menu.add_command(label="Open Book\tCtrl+O", command=self.open_book)
        file_menu.add_command(label="Add Book...\tCtrl+N", command=self.add_book)
        file_menu.add_command(label="Import Folder...\tCtrl+Shift+N", command=self.import_folder)
        file_menu.add_command(label="Export Copy...\tCtrl+E", command=self.export_book)
        file_menu.add_command(label="Send to Voice Dream...\tCtrl+Shift+V", command=self.send_to_voice_dream)
        file_menu.add_command(label="Send to Kindle...\tCtrl+Shift+K", command=self.send_to_kindle)
        file_menu.add_separator()
        file_menu.add_command(label="Exit\tAlt+F4", command=self.root.quit)
        menu_bar.add_cascade(label="File", menu=file_menu, underline=0)

        book_menu = Menu(menu_bar, tearoff=False)
        book_menu.add_command(label="Edit Metadata...\tF2", command=self.edit_book)
        book_menu.add_command(label="Auto-Detect Metadata...\tCtrl+D", command=self.auto_detect_selected_metadata)
        book_menu.add_command(label="Convert to EPUB...\tCtrl+R", command=self.convert_selected_to_epub)
        book_menu.add_command(label="Show Selected Book Information\tCtrl+I", command=self.show_selected_book_info)
        book_menu.add_command(label="Focus Books List\tCtrl+L", command=self.focus_books_list)
        book_menu.add_command(label="Delete from Library\tDelete", command=self.delete_book)
        menu_bar.add_cascade(label="Book", menu=book_menu, underline=0)

        search_menu = Menu(menu_bar, tearoff=False)
        search_menu.add_command(label="Move to Search Box\tCtrl+F", command=self.focus_search)
        search_menu.add_command(label="Move to Books List\tCtrl+L", command=self.focus_books_list)
        search_menu.add_command(label="Clear Search", command=self.clear_search)
        menu_bar.add_cascade(label="Search", menu=search_menu, underline=0)

        settings_menu = Menu(menu_bar, tearoff=False)
        settings_menu.add_command(label="Set Voice Dream Loader Folder...", command=self.choose_voice_dream_folder)
        settings_menu.add_command(label="Set Kindle Email Address...", command=self.set_kindle_email)
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

        ttk.Label(search_frame, text="Search").pack(side=LEFT, padx=(0, 8))
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_var)
        self.search_entry.pack(side=LEFT, fill=X, expand=True, padx=(0, 8))
        self.search_entry.bind("<Return>", lambda event: self.search_and_focus())

        list_frame = ttk.Frame(main)
        list_frame.pack(fill=BOTH, expand=True)

        ttk.Label(
            list_frame,
            text="Books list. Each item includes title, author, source, tags, format, and date added."
        ).pack(anchor="w")

        self.book_list = Listbox(list_frame, selectmode=SINGLE, height=20)
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
        self.book_list.bind("<Control-n>", lambda event: self.add_book())
        self.book_list.bind("<Control-f>", lambda event: self.focus_search())
        self.book_list.bind("<Control-i>", lambda event: self.show_selected_book_info())
        self.book_list.bind("<Control-d>", lambda event: self.auto_detect_selected_metadata())
        self.book_list.bind("<Up>", self.on_book_list_arrow)
        self.book_list.bind("<Down>", self.on_book_list_arrow)
        self.book_list.bind("<Prior>", self.on_book_list_arrow)
        self.book_list.bind("<Next>", self.on_book_list_arrow)
        self.book_list.bind("<Home>", self.on_book_list_arrow)
        self.book_list.bind("<End>", self.on_book_list_arrow)

        self.root.bind("<Control-n>", lambda event: self.add_book())
        self.root.bind("<Control-N>", lambda event: self.import_folder())
        self.root.bind("<Control-o>", lambda event: self.open_book())
        self.root.bind("<Control-e>", lambda event: self.export_book())
        self.root.bind("<Control-V>", lambda event: self.send_to_voice_dream())
        self.root.bind("<Control-k>", lambda event: self.open_kindle())
        self.root.bind("<Control-r>", lambda event: self.convert_selected_to_epub())
        self.root.bind("<Control-K>", lambda event: self.send_to_kindle())
        self.root.bind("<Control-f>", lambda event: self.focus_search())
        self.root.bind("<Control-i>", lambda event: self.show_selected_book_info())
        self.root.bind("<Control-d>", lambda event: self.auto_detect_selected_metadata())
        self.root.bind("<Control-l>", lambda event: self.focus_books_list())
        self.root.bind("<F1>", lambda event: self.show_help())

        status = ttk.Label(main, textvariable=self.status_var, relief="sunken", anchor="w")
        status.pack(fill=X, pady=(8, 0))

        self.book_list.focus_set()

    def focus_search(self):
        self.book_list.focus_set()
        self.search_entry.select_range(0, END)
        self.status_var.set("Search box focused.")

    def focus_books_list(self):
        self.book_list.focus_set()
        if self.book_list.size() > 0 and not self.book_list.curselection():
            self.book_list.selection_set(0)
            self.book_list.activate(0)
        self.status_var.set("Books list focused. Use up and down arrow to choose a book.")

    def on_book_list_arrow(self, event):
        # Let Tk move the active item first, then synchronize selection to it.
        self.root.after(1, self.sync_selection_to_active)
        return None

    def sync_selection_to_active(self):
        if self.book_list.size() == 0:
            return
        active = self.book_list.index("active")
        self.book_list.selection_clear(0, END)
        self.book_list.selection_set(active)
        self.book_list.activate(active)
        self.book_list.see(active)
        self.status_var.set(self.book_list.get(active))

    def clear_search(self):
        self.search_var.set("")
        self.refresh_books()
        self.book_list.focus_set()

    def refresh_books(self):
        self.book_list.delete(0, END)
        self.book_list_ids = []

        rows = self.db.search_books(self.search_var.get())
        for row in rows:
            book_id, title, author = row
            spoken_row = f"Title: {title}. Author: {author or 'Unknown'}."
            self.book_list.insert(END, spoken_row)
            self.book_list_ids.append(book_id)

        count = len(rows)
        self.status_var.set(f"{count} book{'s' if count != 1 else ''} shown.")

        if count:
            self.book_list.selection_clear(0, END)
            self.book_list.selection_set(0)
            self.book_list.activate(0)
            self.book_list.see(0)
            self.book_list.focus_set()

    def current_book_index(self):
        """Return the selected or active list index, and make it the selection."""
        if self.book_list.size() == 0:
            return None

        selected = self.book_list.curselection()
        if selected:
            index = selected[0]
        else:
            try:
                index = self.book_list.index("active")
            except Exception:
                index = 0

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
                f"Source: {row[3] or 'Not specified'}\n"
                f"Tags: {row[4] or 'None'}\n"
                f"Format: {row[6] or 'Unknown'}\n"
                f"Stored path: {row[8]}"
            )
        self.status_var.set(text.replace("\n", " "))
        messagebox.showinfo("Selected book information", text)

    def extract_supported_files_from_zip(self, zip_path: Path) -> list[Path]:
        """Extract a ZIP file and return supported ebook/document files inside it.

        Bookshare ZIP files often contain an EPUB. EPUB files are imported as
        whole files, which preserves images, navigation, and structure.
        """
        extract_root = self.db.folder / "Extracted_Zips" / safe_filename(zip_path.stem)
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
        default_title = source_path.stem.replace("_", " ").replace("-", " ")
        metadata = {"title": default_title, "author": "", "edition": "", "year": "", "isbn": "", "publisher": "", "source": "", "tags": "", "notes": ""}

        if source_path.suffix.lower() == ".epub":
            epub_metadata = read_epub_metadata(source_path)
            metadata.update({key: value for key, value in epub_metadata.items() if value})

        return metadata

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
        if include_subfolders:
            candidates = [path for path in root_folder.rglob("*") if path.is_file()]
        else:
            candidates = [path for path in root_folder.iterdir() if path.is_file()]

        supported = [path for path in candidates if path.suffix.lower() in SUPPORTED_EXTENSIONS]

        if not supported:
            messagebox.showinfo("No supported books found", "No supported ebook or document files were found in that folder.")
            return

        if not messagebox.askyesno(
            "Confirm folder import",
            f"Found {len(supported)} supported file{'s' if len(supported) != 1 else ''}. Import them now?"
        ):
            return

        imported = 0
        skipped_items = []

        for path in supported:
            try:
                if path.suffix.lower() == ".zip":
                    zip_imported, zip_skipped = self.import_zip_file_without_prompt(path, default_source=default_source.strip())
                    imported += zip_imported
                    skipped_items.extend(zip_skipped)
                elif self.import_one_book_without_prompt(path, default_source=default_source.strip()):
                    imported += 1
                else:
                    skipped_items.append(f"{path} -- unsupported or not imported")
            except sqlite3.IntegrityError:
                skipped_items.append(f"{path} -- already imported or duplicate stored path")
            except Exception as exc:
                skipped_items.append(f"{path} -- {exc}")

        report_path = self.write_import_report(imported, skipped_items)

        self.refresh_books()
        if imported:
            self.focus_books_list()

        skipped = len(skipped_items)
        messagebox.showinfo(
            "Folder import complete",
            f"Imported {imported} book{'s' if imported != 1 else ''}. "
            f"Skipped {skipped} file{'s' if skipped != 1 else ''}.\n\n"
            f"Import report saved at:\n{report_path}"
        )
        self.status_var.set(f"Folder import complete. Imported {imported}. Skipped {skipped}.")

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

            default_title = source_path.stem.replace("_", " ").replace("-", " ")
            initial_metadata = {"title": default_title, "author": "", "source": "", "tags": "", "notes": ""}

            if source_path.suffix.lower() == ".epub":
                epub_metadata = read_epub_metadata(source_path)
                initial_metadata.update({key: value for key, value in epub_metadata.items() if value})

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

        existing = {
            "title": row[1],
            "author": row[2],
            "edition": row[10],
            "year": row[11],
            "isbn": row[12],
            "publisher": row[13],
            "source": row[3],
            "tags": row[4],
            "notes": row[5],
        }

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

        if stored_path.suffix.lower() == ".epub":
            try:
                write_epub_metadata(
                    stored_path,
                    metadata["title"],
                    metadata["author"],
                    metadata["source"],
                    metadata["tags"],
                    metadata["notes"],
                )
                self.status_var.set("Detected metadata saved in the database and EPUB file.")
            except Exception as exc:
                messagebox.showwarning(
                    "EPUB metadata not written",
                    f"The database was updated, but I could not write metadata into the EPUB file itself.\n\n{exc}"
                )
                self.status_var.set("Detected metadata saved in the database only.")
        else:
            self.status_var.set("Detected metadata saved in the database only.")

        self.refresh_books()
        self.focus_books_list()

    def edit_book(self):
        book_id = self.selected_book_id()
        if book_id is None:
            return

        row = self.db.get_book(book_id)
        if not row:
            messagebox.showerror("Not found", "The selected book was not found.")
            return

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
        )

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

        self.refresh_books()

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
            "Kindle Email",
            "Enter your Send to Kindle email address, for example name@kindle.com.",
            current,
        )
        if value is None:
            return
        value = value.strip()
        if value and "@" not in value:
            messagebox.showerror("Invalid email", "Please enter a valid Kindle email address.")
            return
        self.db.set_setting("kindle_email", value)
        self.status_var.set("Kindle email address saved.")

    def send_to_kindle(self):
        book_id = self.selected_book_id()
        if book_id is None:
            return

        row = self.db.get_book(book_id)
        if not row:
            messagebox.showerror("Not found", "The selected book was not found.")
            return

        source = Path(row[8])
        if not source.exists():
            messagebox.showerror("File missing", "The stored book file could not be found.")
            return

        kindle_email = self.db.get_setting("kindle_email", "")
        if not kindle_email:
            self.set_kindle_email()
            kindle_email = self.db.get_setting("kindle_email", "")
            if not kindle_email:
                return

        try:
            # Standard mailto cannot reliably attach files. This opens a draft and shows the file path to attach.
            import webbrowser
            subject = f"Send to Kindle: {row[1]}"
            body = (
                "Attach this file before sending to Kindle:%0D%0A%0D%0A"
                + str(source).replace(" ", "%20")
            )
            webbrowser.open(f"mailto:{kindle_email}?subject={subject.replace(' ', '%20')}&body={body}")
            messagebox.showinfo(
                "Send to Kindle",
                "Your email app should open. Attach the selected book file if it is not already attached, then send it.\\n\\n"
                f"File path:\\n{source}"
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
        book_id = self.selected_book_id()
        if book_id is None:
            return

        row = self.db.get_book(book_id)
        if not row:
            return

        source = Path(row[8])
        if not source.exists():
            messagebox.showerror("File missing", "The stored book file could not be found.")
            return

        folder = filedialog.askdirectory(title="Choose export folder")
        if not folder:
            return

        destination = Path(folder) / source.name
        if destination.exists():
            if not messagebox.askyesno("Replace file", f"{destination.name} already exists. Replace it?"):
                return

        shutil.copy2(source, destination)
        self.status_var.set(f"Exported copy to {destination}.")

    def delete_book(self):
        book_id = self.selected_book_id()
        if book_id is None:
            return

        row = self.db.get_book(book_id)
        if not row:
            return

        answer = messagebox.askyesnocancel(
            "Delete book",
            "Remove this book from the library database?\n\n"
            "Choose Yes to continue. Choose No or Cancel to stop.",
        )
        if answer is not True:
            return

        delete_file = messagebox.askyesno(
            "Delete stored file",
            "Also delete the stored copy of the book file? Choose No to keep the file but remove it from the library list.",
        )

        self.db.delete_book(book_id, delete_file=delete_file)
        self.refresh_books()
        self.status_var.set("Book removed from library.")

    def show_help(self):
        messagebox.showinfo(
            "Help",
            "Accessible Ebook Library Manager keyboard commands:\n\n"
            "Alt: Open the menu bar.\n"
            "Control+N: Add book.\n"
            "Control+Shift+N: Import a folder of books, including Bookshare ZIP files.\n"
            "F2: Edit selected book metadata.\n"
            "Control+D: Auto-detect metadata from the selected book.\n"
            "Enter or Control+O: Open selected book.\n"
            "Control+E: Export selected book.\n"
            "Control+R: Convert selected book to EPUB.\n"
            "Control+Shift+K: Send selected book to Kindle.\n"
            "Control+Shift+V: Send selected book to Voice Dream Loader folder.\n"
            "Control+K: Open Kindle for PC.\n"
            "Delete: Remove selected book from library.\n"
            "Control+F: Move to search box.\n"
            "Control+I: Show selected book information.\n"
            "Control+L: Move to the books list.\n"
            "Use Up and Down Arrow in the books list to choose a book.\n"
            "F1: Help.\n\n"
            "The book list is not a table. Each book is one list item with labels for title, author, source, tags, format, and date added. The app automatically focuses the books list when it opens and after searches/imports.\n\n"
            "EPUB metadata changes are written into the EPUB file itself, with a .bak backup made first. "
            "Other file formats are updated in the library database only. This app does not remove DRM."
        )


def main():
    root = Tk()
    LibraryApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
