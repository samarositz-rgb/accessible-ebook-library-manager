"""SQLite-backed library storage.

Holds LibraryDatabase plus the small filesystem/time helpers it depends on,
so the rest of the app does not need to know how the on-disk layout works.
"""

import os
import re
import sqlite3
import threading
from datetime import datetime
from pathlib import Path


DB_NAME = "library.db"
BOOKS_FOLDER = "Books"
SCHEMA_VERSION = 3

ACCESSIBILITY_METADATA_KEYS = [
    "accessibility_summary",
    "accessibility_features",
    "accessibility_hazards",
    "accessibility_access_modes",
    "accessibility_access_modes_sufficient",
    "accessibility_certified_by",
]


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


class _LockedConnection:
    """Wraps a sqlite3.Connection so execute/commit/backup acquire a lock.

    LibraryDatabase is touched from the Tk thread and from the background
    content-indexer thread. sqlite3.Connection itself is not safe for
    concurrent use, so we serialize the calls that we actually make.
    """

    def __init__(self, connection: sqlite3.Connection, lock: threading.RLock):
        self._connection = connection
        self._lock = lock

    def execute(self, *args, **kwargs):
        with self._lock:
            return self._connection.execute(*args, **kwargs)

    def commit(self):
        with self._lock:
            self._connection.commit()

    def close(self):
        with self._lock:
            self._connection.close()

    def backup(self, target):
        with self._lock:
            self._connection.backup(target)


class LibraryDatabase:
    def __init__(self):
        self.folder = app_data_folder()
        self.db_path = self.folder / DB_NAME
        self.books_path = managed_books_folder(self.folder)
        self._lock = threading.RLock()
        raw_connection = sqlite3.connect(self.db_path, check_same_thread=False)
        # The background content indexer touches this connection from a
        # worker thread, so all access is serialized through ``self._lock``.
        self.connection = _LockedConnection(raw_connection, self._lock)
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
        # ensure_book_columns must run before any index that references a
        # dynamically-added column (e.g. year) is created, because on a fresh
        # database those columns do not exist until this call adds them.
        if int(self.get_setting("schema_version", "0")) < SCHEMA_VERSION:
            self.ensure_book_columns()
            self.set_setting("schema_version", str(SCHEMA_VERSION))
        self.connection.execute("CREATE INDEX IF NOT EXISTS idx_books_year ON books(year)")
        self.connection.commit()

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
            "content_flags": "TEXT DEFAULT ''",
            "source_book_id": "TEXT DEFAULT ''",
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
        tokens = re.findall(r"\S+", query)
        fts_query = " ".join(f'"{token}"' for token in tokens)
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

    def set_content_flags(self, book_id: int, flags: str) -> None:
        """Store the comma-separated content_flags string for a book."""
        self.connection.execute(
            "UPDATE books SET content_flags = ? WHERE id = ?",
            (flags or "", book_id),
        )
        self.connection.commit()

    def get_content_flags(self, book_id: int) -> list[str]:
        cursor = self.connection.execute(
            "SELECT content_flags FROM books WHERE id = ?", (book_id,)
        )
        row = cursor.fetchone()
        if not row or not row[0]:
            return []
        return [flag.strip() for flag in row[0].split(",") if flag.strip()]

    def set_source_book_id(self, book_id: int, source_book_id: str) -> None:
        self.connection.execute(
            "UPDATE books SET source_book_id = ? WHERE id = ?",
            (source_book_id or "", book_id),
        )
        self.connection.commit()

    def get_book_by_source_book_id(self, source_book_id: str):
        if not source_book_id:
            return None
        cursor = self.connection.execute(
            "SELECT id, title, author, stored_path FROM books WHERE source_book_id = ? LIMIT 1",
            (source_book_id,),
        )
        return cursor.fetchone()

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
                if os.path.exists(stored_path):
                    os.remove(stored_path)
            except OSError:
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
                   accessibility_access_modes, accessibility_access_modes_sufficient, accessibility_certified_by
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
                   accessibility_access_modes, accessibility_access_modes_sufficient, accessibility_certified_by
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
                   accessibility_access_modes, accessibility_access_modes_sufficient, accessibility_certified_by
            FROM books WHERE original_path = ?
            """,
            (original_path,),
        )
        return cursor.fetchone()

    def search_books(
        self,
        query="",
        sort_by="title",
        source_filter="",
        tag_filter="",
        format_filter="",
        accessibility_feature_filter="",
        accessibility_hazard_filter="",
        content_flag_filter="",
        extra_ids=None,
    ):
        query = query.strip()
        source_filter = source_filter.strip()
        tag_filter = tag_filter.strip()
        format_filter = format_filter.strip().lower()
        accessibility_feature_filter = accessibility_feature_filter.strip()
        accessibility_hazard_filter = accessibility_hazard_filter.strip()
        content_flag_filter = content_flag_filter.strip()

        order_map = {
            "title": "title COLLATE NOCASE ASC, author COLLATE NOCASE ASC",
            "title_desc": "title COLLATE NOCASE DESC, author COLLATE NOCASE ASC",
            "author": "author COLLATE NOCASE ASC, title COLLATE NOCASE ASC",
            "author_desc": "author COLLATE NOCASE DESC, title COLLATE NOCASE ASC",
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
                params.extend([like] * 16)
                params.extend(extra_ids)
            else:
                conditions.append(metadata_condition)
                params.extend([like] * 16)

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

        if accessibility_hazard_filter:
            conditions.append("accessibility_hazards LIKE ?")
            params.append(f"%{accessibility_hazard_filter}%")

        if content_flag_filter:
            conditions.append("content_flags LIKE ?")
            params.append(f"%{content_flag_filter}%")

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
