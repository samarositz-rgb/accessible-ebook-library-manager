"""Small Qt accessibility prototype; reads the production library without writing to it."""

import os
import re
import sqlite3
import sys
from contextlib import closing
from pathlib import Path
from urllib.parse import quote

from PySide6.QtCore import Qt
from PySide6.QtGui import QShortcut, QKeySequence
from PySide6.QtWidgets import (
    QApplication,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QMainWindow,
    QMenuBar,
    QScrollArea,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)


GROUPS = (
    ("A-F", "abcdef"),
    ("G-L", "ghijkl"),
    ("M-R", "mnopqr"),
    ("S-Z", "stuvwxyz"),
    ("Other", ""),
)


def library_database_path():
    base = Path(os.environ.get("APPDATA", Path.home()))
    return base / "AccessibleEbookLibraryManager" / "library.db"


def read_library():
    database = library_database_path()
    if not database.is_file():
        return [], ["id", "title", "author", "edition"], "The library database was not found. Showing sample books."

    uri = f"file:{quote(database.resolve().as_posix(), safe='/:')}?mode=ro"
    try:
        with closing(sqlite3.connect(uri, uri=True, timeout=2)) as connection:
            available = {row[1] for row in connection.execute("PRAGMA table_info(books)")}
            selected = [name for name in ("id", "title", "author", "edition") if name in available]
            if "title" not in selected:
                return [], ["id", "title", "author", "edition"], "The library database has no title field. Showing sample books."
            query = "SELECT " + ", ".join(selected) + " FROM books ORDER BY title COLLATE NOCASE"
            return connection.execute(query).fetchall(), selected, f"Read-only library: {database}"
    except sqlite3.Error as exc:
        return [], ["id", "title", "author", "edition"], f"Could not read the library database: {exc}"


def title_group(title):
    value = (title or "").strip()
    for article in ("The ", "A "):
        if value.casefold().startswith(article.casefold()):
            value = value[len(article):]
            break
    value = re.sub(r"^[^A-Za-z0-9]+", "", value)
    initial = value[:1].casefold()
    for index, (_label, letters) in enumerate(GROUPS):
        if letters and initial in letters:
            return index
    return len(GROUPS) - 1


def book_row_to_dict(row, columns):
    return dict(zip(columns, row))


class BookList(QListWidget):
    def __init__(self, parent=None, column_index=None, on_horizontal_move=None):
        super().__init__(parent)
        self.setAccessibleName("Books")
        self.setAccessibleDescription(
            "Use Up and Down Arrow to move through books, and Left and Right Arrow to move between alphabetical columns."
        )
        self.setSelectionMode(QListWidget.SelectionMode.SingleSelection)
        self.setUniformItemSizes(True)
        self.column_index = column_index
        self.on_horizontal_move = on_horizontal_move

    def keyPressEvent(self, event):
        if self.column_index is not None and event.key() in (Qt.Key.Key_Left, Qt.Key.Key_Right):
            direction = -1 if event.key() == Qt.Key.Key_Left else 1
            if self.on_horizontal_move and self.on_horizontal_move(self.column_index + direction):
                event.accept()
                return
        super().keyPressEvent(event)


class AccessibilityTestWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Accessible Ebook Library Manager")
        self.resize(980, 650)
        self.books, self.columns, status = read_library()
        if not self.books:
            self.books = [
                (1, "A Sample Book", "Example Author", "First edition"),
                (2, "The Narrator Test", "Accessibility Tester", "Second edition"),
                (3, "Zebra Crossing", "Another Author", ""),
            ]
            self.columns = ["id", "title", "author", "edition"]
            if status.startswith("Read-only library:"):
                status = "No books found. Showing sample books."

        self.book_records = [book_row_to_dict(row, self.columns) for row in self.books]
        self.speech_fields = ["title", "author"]
        self._load_speech_settings()
        self.current_mode = "single"
        self.single_list = BookList(self)
        self.columns_page = QWidget(self)
        columns_layout = QHBoxLayout(self.columns_page)
        columns_layout.setContentsMargins(0, 0, 0, 0)
        columns_layout.setSpacing(8)
        self.group_lists = []
        self.group_headers = []
        for index, (label, _letters) in enumerate(GROUPS):
            column = QWidget(self.columns_page)
            column_layout = QVBoxLayout(column)
            column_layout.setContentsMargins(0, 0, 0, 0)
            header = QLabel(label, column)
            book_list = BookList(column, index, self._move_to_column)
            book_list.setAccessibleName(f"{label} books")
            book_list.setWordWrap(True)
            column_layout.addWidget(header)
            column_layout.addWidget(book_list)
            columns_layout.addWidget(column, 1)
            self.group_headers.append(header)
            self.group_lists.append(book_list)

        self.columns_scroll = QScrollArea(self)
        self.columns_scroll.setWidgetResizable(True)
        self.columns_scroll.setWidget(self.columns_page)
        self.stack = QStackedWidget(self)
        self.stack.addWidget(self.single_list)
        self.stack.addWidget(self.columns_scroll)
        self.stack.setAccessibleName("Book list view")
        self.setCentralWidget(self.stack)
        self.statusBar().showMessage(status)
        self._populate_lists()
        self._create_menus()
        self._create_view_shortcuts()
        self._connect_selection_events()

        first = self.single_list.item(0)
        if first is not None:
            self.single_list.setCurrentItem(first)
        self.single_list.setFocus(Qt.FocusReason.OtherFocusReason)
        self.statusBar().showMessage(f"{len(self.book_records)} books loaded. {status}")

    def _create_view_shortcuts(self):
        self.single_shortcut = QShortcut(QKeySequence("Ctrl+1"), self)
        self.single_shortcut.setContext(Qt.ShortcutContext.WindowShortcut)
        self.single_shortcut.activated.connect(lambda: self.set_view_mode("single"))
        self.multiple_shortcut = QShortcut(QKeySequence("Ctrl+2"), self)
        self.multiple_shortcut.setContext(Qt.ShortcutContext.WindowShortcut)
        self.multiple_shortcut.activated.connect(lambda: self.set_view_mode("multiple"))

    def _load_speech_settings(self):
        database = library_database_path()
        if not database.is_file():
            return
        uri = f"file:{quote(database.resolve().as_posix(), safe='/:')}?mode=ro"
        try:
            with closing(sqlite3.connect(uri, uri=True, timeout=2)) as connection:
                row = connection.execute(
                    "SELECT value FROM settings WHERE key = ?", ("book_list_speech_fields",)
                ).fetchone()
            if row and row[0]:
                choices = [field for field in row[0].split(",") if field in {"title", "author", "edition"}]
                if choices:
                    self.speech_fields = choices
        except sqlite3.Error:
            pass

    def _create_menus(self):
        menu_bar = QMenuBar(self)
        menu_bar.setNativeMenuBar(False)
        self.setMenuBar(menu_bar)

        file_menu = menu_bar.addMenu("&File")
        exit_action = file_menu.addAction("E&xit")
        exit_action.setShortcut("Alt+F4")
        exit_action.triggered.connect(self.close)

        view_menu = menu_bar.addMenu("&View")
        self.single_action = view_menu.addAction("&Single list")
        self.single_action.setCheckable(True)
        self.single_action.setChecked(True)
        self.multiple_action = view_menu.addAction("&Multiple alphabetical lists")
        self.multiple_action.setCheckable(True)
        self.single_action.triggered.connect(lambda: self.set_view_mode("single"))
        self.multiple_action.triggered.connect(lambda: self.set_view_mode("multiple"))

        settings_menu = menu_bar.addMenu("&Settings")
        speech_menu = settings_menu.addMenu("Book list &reading")
        self.speech_actions = {}
        for label, fields in (
            ("Title only", ["title"]),
            ("Title and author", ["title", "author"]),
            ("Title, author, and edition", ["title", "author", "edition"]),
        ):
            action = speech_menu.addAction(label)
            action.setCheckable(True)
            action.triggered.connect(lambda _checked=False, fields=fields: self.set_speech_fields(fields))
            self.speech_actions[tuple(fields)] = action
        self._update_speech_menu_checks()

        help_menu = menu_bar.addMenu("&Help")
        test_action = help_menu.addAction("About this accessibility test")
        test_action.triggered.connect(self.show_test_info)

    def _populate_lists(self):
        lists = [self.single_list, *self.group_lists]
        for book_list in lists:
            book_list.clear()
        for book in self.book_records:
            title = book.get("title") or "Untitled"
            author = book.get("author") or "Unknown author"
            details = []
            for field in self.speech_fields:
                if field == "title":
                    details.append(f"{title}.")
                elif field == "author":
                    details.append(f"{author}.")
                elif field == "edition":
                    details.append(f"{book.get('edition') or 'Edition not specified'}.")
            item = ". ".join(details)
            self.single_list.addItem(item)
            group_list = self.group_lists[title_group(title)]
            group_list.addItem(item)

    def _connect_selection_events(self):
        self.single_list.currentItemChanged.connect(
            lambda current, _previous: self._announce_status(current, self.single_list)
        )
        for index, book_list in enumerate(self.group_lists):
            book_list.currentItemChanged.connect(
                lambda current, _previous, index=index, book_list=book_list: self._announce_status(
                    current, book_list, index
                )
            )

    def _announce_status(self, item, book_list, group_index=None):
        if item is None:
            return
        prefix = f"{GROUPS[group_index][0]} range, " if group_index is not None else ""
        self.statusBar().showMessage(
            f"{prefix}book {book_list.currentRow() + 1} of {book_list.count()}: {item.text()}"
        )

    def _move_to_column(self, index):
        if not 0 <= index < len(self.group_lists):
            return False
        book_list = self.group_lists[index]
        if book_list.count() == 0:
            self.statusBar().showMessage(f"{GROUPS[index][0]} column, no books")
            return True
        row = min(max(book_list.currentRow(), 0), book_list.count() - 1)
        book_list.setCurrentRow(row)
        book_list.setFocus(Qt.FocusReason.OtherFocusReason)
        self.statusBar().showMessage(f"{GROUPS[index][0]} column, {book_list.count()} books")
        return True

    def set_view_mode(self, mode):
        self.current_mode = mode
        self.stack.setCurrentIndex(0 if mode == "single" else 1)
        self.single_action.setChecked(mode == "single")
        self.multiple_action.setChecked(mode == "multiple")
        if mode == "single":
            self.single_list.setFocus(Qt.FocusReason.OtherFocusReason)
            self.statusBar().showMessage(f"Single list, {len(self.book_records)} books")
        else:
            selected_title = self.book_records[self.single_list.currentRow()].get("title", "") if self.single_list.currentRow() >= 0 else ""
            column_index = title_group(selected_title)
            if self.group_lists[column_index].count() == 0:
                column_index = next((i for i, view in enumerate(self.group_lists) if view.count()), 0)
            self._move_to_column(column_index)
            counts = ", ".join(f"{label}: {self.group_lists[i].count()}" for i, (label, _) in enumerate(GROUPS))
            self.statusBar().showMessage(f"Multiple alphabetical columns. {counts}. Use Left and Right Arrow between columns.")

    def set_speech_fields(self, fields):
        self.speech_fields = fields
        self._populate_lists()
        self._update_speech_menu_checks()
        self.statusBar().showMessage("Book list reading: " + ", ".join(fields))
        current = self.single_list.item(0)
        if current:
            self.single_list.setCurrentItem(current)

    def _update_speech_menu_checks(self):
        for fields, action in self.speech_actions.items():
            action.setChecked(tuple(self.speech_fields) == fields)

    def show_test_info(self):
        self.statusBar().showMessage(
            "Prototype only. Use Alt to reach the menu, Ctrl+1 or Ctrl+2 to change list view, and arrow keys in a list."
        )


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("Accessible Ebook Library Manager")
    window = AccessibilityTestWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
