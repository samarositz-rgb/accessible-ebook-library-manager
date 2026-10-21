"""Qt-based accessible main-window migration, sharing the existing library database."""

import os
import json
import re
import string
import shutil
import sqlite3
import subprocess
import sys
import time
import urllib.request
from datetime import datetime, timezone
from importlib.util import find_spec
from pathlib import Path


_QT_DLL_DIRECTORY_HANDLES = []


def prepare_qt_dll_search_path():
    if sys.platform != "win32":
        return
    bundle_root = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    qt_directory = bundle_root / "PySide6"
    shiboken_directory = bundle_root / "shiboken6"
    if not qt_directory.is_dir():
        spec = find_spec("PySide6")
        if spec and spec.origin:
            qt_directory = Path(spec.origin).resolve().parent
            shiboken_directory = qt_directory.parent / "shiboken6"
    directories = [path for path in (qt_directory, shiboken_directory) if path.is_dir()]
    for directory in directories:
        _QT_DLL_DIRECTORY_HANDLES.append(os.add_dll_directory(str(directory)))
    if directories:
        os.environ["PATH"] = os.pathsep.join(map(str, directories)) + os.pathsep + os.environ.get("PATH", "")


prepare_qt_dll_search_path()

from PySide6.QtCore import Qt, QUrl, QThread, Signal, QTimer, QRect, QPoint
from PySide6.QtGui import QAction, QAccessible, QAccessibleEvent, QAccessibleAnnouncementEvent, QKeySequence, QShortcut, QDesktopServices, QPalette, QPainter, QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QDialog,
    QDialogButtonBox,
    QCheckBox,
    QBoxLayout,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMenu,
    QMessageBox,
    QProgressDialog as QtProgressDialog,
    QPlainTextEdit,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QSpinBox,
    QStackedWidget,
    QStatusBar,
    QTabWidget,
    QToolButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from library_manager import (
    APP_NAME,
    app_data_folder,
    BACKUP_SCHEDULES,
    BOOKS_FOLDER,
    BOOK_LIST_ALPHABET_GROUPS,
    BOOK_LIST_SHORTCUT_FIELDS,
    BOOK_LIST_SPEECH_FIELDS,
    EXPLORER_CONTEXT_MENU_KEY,
    EXPLORER_CONTEXT_MENU_LABEL,
    book_part_sort_key,
    DEFAULT_BACKUP_SCHEDULE,
    DEFAULT_MISSING_METADATA_SOUND_MODE,
    MISSING_METADATA_SOUND_MODES,
    LibraryDatabase,
    LibraryApp,
    clean_filename_title,
    detect_metadata_from_text,
    read_epub_metadata,
    write_epub_metadata,
    SUPPORTED_EXTENSIONS,
    detect_metadata_from_text_content,
    read_epub_accessibility_metadata,
    read_text_for_metadata_detection,
    parse_utc_text,
    utc_now_text,
    cloud_backup_subfolder,
)
from document_text import extract_text_for_indexing
from book_flags import IMAGE_ONLY_PDF, describe_flags, index_book_file
from audiobook_scan import first_audio_file, import_audiobooks, scan_audiobook_folder
from goodreads_export import export_goodreads_csv
from find_elsewhere import GROUP_TITLES as FIND_GROUP_TITLES, LIBRARY_KEY_SETTING as FIND_LIBRARY_SETTING, SERVICES as FIND_SERVICES, search_gutenberg, search_words, service_target, services_in_group
from goodreads_import import add_wishlist_entries, apply_goodreads_entries, is_wishlist_path, parse_goodreads_csv, rename_legacy_wishlist_tags, write_unmatched_file
from help_file import ensure_user_help, load_help_text
from library_utils import (
    metadata_score_from_row,
    normalize_duplicate_key,
    normalize_isbn_key,
    safe_filename,
    title_keys_look_same,
    folder_file_stats,
    sync_folder_contents,
)
from calibre_tools import find_calibre_tool
from online_metadata import lookup_online_metadata
from braille_export import BrailleEngine, BrailleOptions, ExportCancelled, publish_files
from conversion_warnings import inspect_conversion_risks
from book_categories import COMMON_TAGS

VIEW_LABELS = {"single": "Single List", "multiple": "Multiple Alphabetical Lists", "letters": "One Column per Letter", "tags": "Columns by Tag"}
SPEECH_PRESETS = (
    ("Title Only", ["title"]),
    ("Title and Author", ["title", "author"]),
    ("Title, Author, and Edition", ["title", "author", "edition"]),
    ("All Details", [key for key, _ in BOOK_LIST_SPEECH_FIELDS]),
)


def title_navigation_key(title):
    text = str(title or "").strip().lstrip("\"'([{ ").casefold()
    return re.sub(r"^(?:the|a)\s+", "", text).lstrip("\"'([{ ")


def book_tags(value):
    return list(dict.fromkeys(tag.strip().casefold() for tag in re.split(r"[,;]", value or "") if tag.strip()))


def tag_categories(value):
    aliases = {"biography": "biography and memoir", "textbook": "textbooks",
               "nonfiction": "non-fiction", "non fiction": "non-fiction"}
    known = {tag.casefold() for tag in COMMON_TAGS}
    return list(dict.fromkeys(
        tag if tag in known else "other"
        for tag in (aliases.get(tag, tag) for tag in book_tags(value))
    ))


class TagMenu(QMenu):
    def keyPressEvent(self, event):
        action = self.activeAction()
        if event.key() == Qt.Key.Key_Space and action and action.isCheckable():
            action.trigger()
            event.accept()
            return
        super().keyPressEvent(event)

    def mouseReleaseEvent(self, event):
        action = self.actionAt(event.position().toPoint())
        if event.button() == Qt.MouseButton.LeftButton and action and action.isCheckable():
            action.trigger()
            event.accept()
            return
        super().mouseReleaseEvent(event)


class TagInput(QToolButton):
    def __init__(self, text="", parent=None):
        super().__init__(parent)
        self.field_name = "Tags"
        self.tag_actions = []
        self.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.setText(text)

    def text(self):
        return ", ".join(action.text() for action in self.tag_actions if action.isChecked())

    def setText(self, text):
        old_menu = self.menu()
        menu = TagMenu(self)
        self.tag_actions = []
        tags = book_tags(text)
        known = {tag.casefold() for tag in COMMON_TAGS}
        for tag in COMMON_TAGS:
            self.add_tag_action(menu, tag, tag.casefold() in tags)
        legacy = [tag.strip() for tag in re.split(r"[,;]", text or "")
                  if tag.strip() and tag.strip().casefold() not in known]
        if legacy:
            legacy_menu = TagMenu("Existing Tags", menu)
            menu.addMenu(legacy_menu)
            for tag in dict.fromkeys(legacy):
                self.add_tag_action(legacy_menu, tag, True)
        self.setMenu(menu)
        if old_menu:
            old_menu.deleteLater()
        self.update_summary()

    def add_tag_action(self, menu, tag, checked):
        action = QAction(tag, menu)
        action.setCheckable(True)
        action.setChecked(checked)
        action.toggled.connect(self.update_summary)
        menu.addAction(action)
        self.tag_actions.append(action)

    def setAccessibleName(self, name):
        self.field_name = name
        self.update_summary()

    def update_summary(self):
        value = self.text() or "No tags"
        count = sum(action.isChecked() for action in self.tag_actions)
        QToolButton.setText(self, f"{count} tags selected" if count > 1 else value)
        super().setAccessibleName(f"{self.field_name}: {value}")
        self.setToolTip(value)

try:
    import winreg
except ImportError:
    winreg = None


class BookColumnHeading(QLabel):
    def __init__(self, text, book_list, parent):
        super().__init__("", parent)
        self.caption = text
        self.setMinimumHeight(self.fontMetrics().height() + 8)
        self.book_list = book_list
        self.setAccessibleName("")
        self.setFocusPolicy(Qt.FocusPolicy.ClickFocus)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setPen(self.palette().color(QPalette.ColorRole.WindowText))
        painter.drawText(self.rect(), Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter, self.caption)
        painter.end()

    def keyPressEvent(self, event):
        view = self.book_list
        previous, following = view.category_keys()
        if event.key() in (previous, following):
            delta = -1 if event.key() == previous else 1
            view.move_column(view.column + delta)
            event.accept()
            return
        if view.currentItem() is not None:
            if event.key() in view.book_keys():
                view.setFocus()
                event.accept()
                return
            if view.handle_key(event, view.accessible_row):
                event.accept()
                return
        super().keyPressEvent(event)


class AccessibleBookRow(QLabel):
    def __init__(self, text, book_list):
        super().__init__(text, book_list.content)
        self.book_list = book_list
        self.setAccessibleName(text)
        self.setWordWrap(True)
        self.setFocusPolicy(Qt.FocusPolicy.ClickFocus)
        self.setAutoFillBackground(True)
        self.setMargin(4)
        self.normal_palette = self.palette()
        self.setStyleSheet("padding: 3px 5px;")

    def focusInEvent(self, event):
        if not self.book_list.column_canvas and getattr(self, "item", None):
            name = self.item.navigation_text()
            if self.accessibleName() != name:
                self.setAccessibleName(name)
            if self.accessibleDescription():
                self.setAccessibleDescription("")
        super().focusInEvent(event)

    def keyPressEvent(self, event):
        if self.book_list.handle_key(event, self):
            event.accept()
            return
        super().keyPressEvent(event)

    def mousePressEvent(self, event):
        self.book_list.setCurrentRow(self.book_list._items.index(self.item))
        self.setFocus(Qt.FocusReason.MouseFocusReason)
        super().mousePressEvent(event)


class AccessibleBookItem:
    def __init__(self, text, book_id, title):
        self._text = text
        self._book_id = book_id
        self._title = title
        self._selected = False
        self.row_widget = None

    def text(self):
        return self._text

    def navigation_text(self):
        return f"Selected. {self._text}" if self._selected else self._text

    def data(self, role):
        if role == Qt.ItemDataRole.UserRole:
            return self._book_id
        if role == Qt.ItemDataRole.UserRole + 1:
            return self._title
        return None

    def setSelected(self, selected):
        self._selected = bool(selected)
        if self.row_widget and hasattr(self.row_widget, "normal_palette"):
            palette = QPalette(self.row_widget.normal_palette)
            if self._selected:
                palette.setColor(QPalette.ColorRole.Window, self.row_widget.palette().color(QPalette.ColorRole.Highlight))
                palette.setColor(QPalette.ColorRole.WindowText, self.row_widget.palette().color(QPalette.ColorRole.HighlightedText))
            self.row_widget.setPalette(palette if self._selected else self.row_widget.normal_palette)
            self.row_widget.setAccessibleDescription("Selected" if self._selected else "")
            # Explicit command feedback announces the change while focused.
            # Include the state in the name when this book is next entered.
            if not self.row_widget.hasFocus():
                self.row_widget.setAccessibleName(self.navigation_text())
                self.row_widget.setAccessibleDescription("")
        if getattr(self, "column_view", None):
            self.column_view.column_canvas.update()

    def isSelected(self):
        return self._selected


class BookColumnCanvas(QWidget):
    def __init__(self, book_list):
        super().__init__()
        self.book_list = book_list
        self.setMouseTracking(True)
        self.setMinimumWidth(120)

    def item_rects(self):
        horizontal = self.book_list.horizontal_books
        width = max(240, self.fontMetrics().averageCharWidth() * 36) if horizontal else max(40, self.width() - 18)
        rects = []
        y = 0
        for item in self.book_list._items:
            text_rect = self.fontMetrics().boundingRect(
                0, 0, width - 12, 10000, Qt.TextFlag.TextWordWrap, item.text()
            )
            height = max(self.fontMetrics().height() + 16, text_rect.height() + 16)
            rects.append((item, QRect(0, y, width, height)))
            y += height
        if horizontal:
            height = max((rect.height() for _, rect in rects), default=self.fontMetrics().height() + 16)
            return [(item, QRect(index * width, 0, width, height)) for index, (item, _) in enumerate(rects)]
        return rects

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.update_content_height()
        self.book_list.update_accessible_row()

    def update_content_height(self):
        rects = self.item_rects()
        height = max((rect.bottom() + 1 for _, rect in rects), default=0)
        width = max((rect.right() + 1 for _, rect in rects), default=120) if self.book_list.horizontal_books else 120
        if self.minimumWidth() != width:
            self.setMinimumWidth(width)
        if self.minimumHeight() != height:
            self.setMinimumHeight(height)
        if self.book_list.horizontal_books:
            self.book_list.setMinimumHeight(height + self.book_list.scroll_area.horizontalScrollBar().sizeHint().height() + 8)
        else:
            self.book_list.setMinimumHeight(0)

    def paintEvent(self, _event):
        painter = QPainter(self)
        palette = self.palette()
        current_index = self.book_list.currentRow()
        for index, (item, rect) in enumerate(self.item_rects()):
            selected = item.isSelected()
            background = palette.color(QPalette.ColorRole.Highlight if selected else QPalette.ColorRole.Base)
            foreground = palette.color(QPalette.ColorRole.HighlightedText if selected else QPalette.ColorRole.Text)
            painter.fillRect(rect, background)
            painter.setPen(foreground)
            painter.drawText(
                rect.adjusted(8, 3, -6, -3),
                int(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter | Qt.TextFlag.TextWordWrap),
                item.text(),
            )
        painter.end()

    def mousePressEvent(self, event):
        for index, (_item, rect) in enumerate(self.item_rects()):
            if rect.contains(event.position().toPoint()):
                self.book_list.setCurrentRow(index)
                self.book_list.setFocus(Qt.FocusReason.MouseFocusReason)
                break
        super().mousePressEvent(event)


class AccessibleBookList(QWidget):
    currentItemChanged = Signal(object, object)

    def __init__(self, parent=None, column=None, move_column=None, shared_row=None):
        super().__init__(parent)
        self.column = column
        self.horizontal_books = False
        self.column_label = ""
        self.move_column = move_column
        self._items = []
        self._current_row = -1
        self.accessibility_active = column is None
        self.include_column_name = False
        self.setAccessibleName("Books" if column is None else "")
        self.setAccessibleDescription(
            "Use Up and Down Arrow to move between books. Use Left and Right Arrow to move between alphabetical columns."
        )
        self.setFocusPolicy(Qt.FocusPolicy.ClickFocus)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self.column_canvas = None
        self.accessible_row = None
        self.shared_row = column is not None
        if column is not None:
            self.scroll_area = QScrollArea(self)
            self.scroll_area.setWidgetResizable(True)
            self.column_canvas = BookColumnCanvas(self)
            self.scroll_area.setWidget(self.column_canvas)
            self.content = self.column_canvas
            self.accessible_row = shared_row if shared_row is not None else AccessibleBookRow("", self)
            self.accessible_row.setAutoFillBackground(False)
            if shared_row is None:
                self.accessible_row.setParent(self.column_canvas)
            self.accessible_row.setStyleSheet("background: transparent; padding: 3px 5px; border: 1px solid transparent; QLabel:focus { border-color: palette(highlight); }")
            self.accessible_row.setMinimumHeight(0)
            layout.addWidget(self.scroll_area)
        else:
            self.scroll_area = QScrollArea(self)
            self.scroll_area.setWidgetResizable(True)
            self.content = QWidget(self.scroll_area)
            self.rows_layout = QVBoxLayout(self.content)
            self.rows_layout.setContentsMargins(0, 0, 0, 0)
            self.rows_layout.setSpacing(0)
            self.rows_layout.addStretch(1)
            self.scroll_area.setWidget(self.content)
            layout.addWidget(self.scroll_area)

    def count(self):
        return len(self._items)

    def item(self, index):
        return self._items[index] if 0 <= index < len(self._items) else None

    def addItem(self, item):
        self._items.append(item)
        if self.column_canvas:
            item.column_view = self
            if not getattr(self, "bulk_loading", False):
                self.column_canvas.updateGeometry()
                self.column_canvas.update_content_height()
                self.column_canvas.update()
        else:
            item.row_widget = AccessibleBookRow(item.text(), self)
            item.row_widget.item = item
            self.rows_layout.insertWidget(self.rows_layout.count() - 1, item.row_widget)
        item.setSelected(item.isSelected())

    def clear(self):
        previous = self.currentItem()
        self._current_row = -1
        self._items.clear()
        if self.column_canvas:
            self.set_accessibility_active(False)
            self.column_canvas.updateGeometry()
            self.column_canvas.update_content_height()
            self.column_canvas.update()
        else:
            while self.rows_layout.count() > 1:
                layout_item = self.rows_layout.takeAt(0)
                widget = layout_item.widget()
                if widget:
                    widget.hide()
                    widget.setParent(None)
                    widget.deleteLater()
        if previous:
            self.currentItemChanged.emit(None, previous)

    def currentItem(self):
        return self.item(self._current_row)

    def currentRow(self):
        return self._current_row

    def setCurrentRow(self, index, _selection_flag=None):
        self.include_column_name = False
        previous = self.currentItem()
        self._current_row = index if 0 <= index < self.count() else -1
        current = self.currentItem()
        if self.column_canvas:
            self.update_accessible_row()
            self.column_canvas.updateGeometry()
            self.column_canvas.update_content_height()
            self.column_canvas.update()
        if current is not previous:
            self.currentItemChanged.emit(current, previous)

    def update_accessible_row(self):
        if not self.column_canvas or not self.accessible_row:
            return
        if self.shared_row and not self.accessibility_active:
            return
        current = self.currentItem()
        previous = getattr(self, "_accessible_item", None)
        if previous and previous.row_widget is self.accessible_row:
            previous.row_widget = None
        self._accessible_item = current
        old_view = self.accessible_row.book_list
        if old_view is not self and old_view.currentItem() is not None:
            old_view.currentItem().row_widget = None
        self.accessible_row.book_list = self
        if current is None:
            if previous:
                previous.row_widget = None
            self.accessible_row.item = None
            # Only touch the shared row when something actually changes; every
            # unconditional rename is another event a screen reader can speak.
            if self.accessible_row.accessibleName():
                self.accessible_row.setAccessibleName("")
            if not self.accessible_row.isHidden():
                self.accessible_row.hide()
            return
        self.accessible_row.item = current
        text = current.navigation_text()
        if self.include_column_name:
            text = f"{self.column_label}. {text}"
        # The canvas paints the title; expose speech only through the name, not
        # a second QLabel text value containing the same words.
        # Qt emits the name-change event itself; layout refreshes must stay silent.
        if self.accessible_row.accessibleName() != text:
            self.accessible_row.setAccessibleName(text)
        description = ""
        if self.accessible_row.accessibleDescription() != description:
            self.accessible_row.setAccessibleDescription(description)
        if not self.accessibility_active:
            current.row_widget = None
            self.accessible_row.hide()
            return
        current.row_widget = self.accessible_row
        self.accessible_row.show()
        for item, rect in self.column_canvas.item_rects():
            if item is current:
                position = self.column_canvas.mapTo(self.accessible_row.parentWidget(), rect.topLeft())
                self.accessible_row.setGeometry(QRect(position, rect.size()))
                break
        self.accessible_row.raise_()
        if self.accessible_row.hasFocus():
            QTimer.singleShot(0, self.ensure_current_visible)

    def ensure_current_visible(self):
        if not self.column_canvas or self.accessible_row.book_list is not self:
            return
        for item, rect in self.column_canvas.item_rects():
            if item is self.currentItem():
                self.scroll_area.ensureVisible(rect.center().x(), rect.center().y(), rect.width() // 2, rect.height() // 2)
                position = self.column_canvas.mapTo(self.accessible_row.parentWidget(), rect.topLeft())
                self.accessible_row.move(position)
                break

    def set_accessibility_active(self, active):
        if not self.column_canvas:
            return
        self.accessibility_active = bool(active)
        if not active and self.accessible_row.book_list is self:
            self.accessible_row.hide()
        self.update_accessible_row()
        self.column_canvas.update()

    def selectedItems(self):
        return [item for item in self._items if item.isSelected()]

    def clearSelection(self):
        for item in self._items:
            item.setSelected(False)

    def setFocus(self, reason=Qt.FocusReason.OtherFocusReason):
        if self.column_canvas and not self._items:
            # An empty category must never take focus or be spoken. Hand focus to
            # the nearest category that has books, or leave focus where it is.
            window = self.window()
            finder = getattr(window, "nearest_populated_view", None)
            target = finder(self) if finder else None
            if target is not None and target is not self:
                window.move_to_column(target.column)
            return
        if self.column_canvas:
            for view in getattr(self.window(), "groups", []):
                view.accessibility_active = view is self
            self.window().active_column = self.column
            self.accessibility_active = True
            self.update_accessible_row()
            self.accessible_row.setFocus(reason)
            QTimer.singleShot(0, self.ensure_current_visible)
            return
        item = self.currentItem()
        if item and item.row_widget:
            item.row_widget.setFocus(reason)
        else:
            super().setFocus(reason)

    def book_keys(self):
        return (Qt.Key.Key_Left, Qt.Key.Key_Right) if self.horizontal_books else (Qt.Key.Key_Up, Qt.Key.Key_Down)

    def category_keys(self):
        return (Qt.Key.Key_Up, Qt.Key.Key_Down) if self.horizontal_books else (Qt.Key.Key_Left, Qt.Key.Key_Right)

    def handle_key(self, event, row):
        previous, following = self.category_keys()
        if self.column is not None and event.key() in (previous, following):
            delta = -1 if event.key() == previous else 1
            if self.move_column:
                self.move_column(self.column + delta)
            return True
        if self.column is not None and self.currentItem() is None:
            return event.key() in (*self.book_keys(), previous, following, Qt.Key.Key_Home, Qt.Key.Key_End)
        try:
            self._current_row = self._items.index(row.item)
        except (ValueError, AttributeError):
            return False
        if event.key() == Qt.Key.Key_A and event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            if event.modifiers() & Qt.KeyboardModifier.ShiftModifier:
                self.window().deselect_all()
            else:
                self.window().select_all()
            return True
        if event.modifiers() & Qt.KeyboardModifier.AltModifier and Qt.Key.Key_0 <= event.key() <= Qt.Key.Key_9:
            if event.isAutoRepeat():
                return True
            digit = "0" if event.key() == Qt.Key.Key_0 else str(event.key() - Qt.Key.Key_0)
            field_index = 9 if digit == "0" else int(digit) - 1
            if 0 <= field_index < len(BOOK_LIST_SHORTCUT_FIELDS):
                self.window().read_or_edit_field(BOOK_LIST_SHORTCUT_FIELDS[field_index], digit)
                return True
        if event.key() == Qt.Key.Key_Space and event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            if event.isAutoRepeat():
                return True
            item = self.currentItem()
            if item is not None:
                self.window().toggle_marked(item)
            return True
        if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            self.window().open_book()
            return True
        previous, following = self.book_keys()
        if event.key() in (previous, following, Qt.Key.Key_Home, Qt.Key.Key_End):
            if event.key() == Qt.Key.Key_Home:
                target = 0
            elif event.key() == Qt.Key.Key_End:
                target = self.count() - 1
            else:
                target = self.currentRow() + (-1 if event.key() == previous else 1)
            if 0 <= target < self.count():
                self.setCurrentRow(target)
                self.setFocus()
            return True
        if not event.modifiers() and len(event.text()) == 1 and event.text().isalpha():
            target = event.text().casefold()
            start = self.currentRow()
            for offset in range(1, self.count() + 1):
                index = (start + offset) % self.count()
                title = title_navigation_key(self.item(index).data(Qt.ItemDataRole.UserRole + 1))
                if title.startswith(target):
                    self.setCurrentRow(index)
                    self.setFocus()
                    return True
        return False


class DiscardChangesDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Unsaved Metadata Changes")
        layout = QHBoxLayout(self)
        self.keep_editing = QPushButton("&Keep Editing", self)
        self.keep_editing.setDefault(True)
        self.keep_editing.clicked.connect(self.reject)
        self.discard = QPushButton("&Discard Changes", self)
        self.discard.clicked.connect(self.accept)
        layout.addWidget(self.keep_editing)
        layout.addWidget(self.discard)
        self.keep_editing.setFocus()


class RemoveBooksDialog(QDialog):
    def __init__(self, count, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Remove {count} books from library; stored files will be kept")
        layout = QHBoxLayout(self)
        cancel = QPushButton("Cancel", self)
        cancel.setDefault(True)
        cancel.clicked.connect(self.reject)
        remove = QPushButton("Remove from Library", self)
        remove.clicked.connect(self.accept)
        layout.addWidget(cancel)
        layout.addWidget(remove)
        cancel.setFocus()


class FolderImportOptionsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Import Folder Options")
        layout = QVBoxLayout(self)
        form = QFormLayout()
        self.include = QComboBox(self)
        self.include.addItems(["This folder and its subfolders", "This folder only"])
        self.include.setAccessibleName("Folders to search")
        self.source = QLineEdit(self)
        self.source.setAccessibleName("Source, optional")
        form.addRow("Folders to search", self.include)
        form.addRow("Source, optional", self.source)
        layout.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel, self)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.include.setFocus()


class MetadataDialog(QDialog):
    FIELDS = (
        ("title", "Title"), ("author", "Author"), ("series", "Series"),
        ("series_number", "Series number"), ("edition", "Edition"),
        ("year", "Year"), ("isbn", "ISBN"), ("publisher", "Publisher"),
        ("source", "Source"), ("tags", "Tags"), ("notes", "Notes"),
    )

    def __init__(self, metadata, parent=None, initial_focus_field="title"):
        super().__init__(parent)
        self.setWindowTitle("Edit Book Metadata")
        self.resize(560, 680)
        layout = QVBoxLayout(self)
        form = QFormLayout()
        form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.AllNonFixedFieldsGrow)
        self.inputs = {}
        for key, label_text in self.FIELDS:
            label = QLabel(label_text, self)
            if key == "notes":
                field = QTextEdit(self)
                field.setTabChangesFocus(True)
                field.setPlainText(str(metadata.get(key, "") or ""))
                field.setMinimumHeight(100)
            elif key == "tags":
                field = TagInput(str(metadata.get(key, "") or ""), self)
            else:
                field = QLineEdit(str(metadata.get(key, "") or ""), self)
            field.setAccessibleName(label_text)
            label.setBuddy(field)
            form.addRow(label, field)
            self.inputs[key] = field
        layout.addLayout(form)
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel,
            parent=self,
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.save_button = buttons.button(QDialogButtonBox.StandardButton.Save)
        self.cancel_button = buttons.button(QDialogButtonBox.StandardButton.Cancel)
        self.save_button.setDefault(True)
        tab_order = [*self.inputs.values(), self.save_button, self.cancel_button]
        for first, second in zip(tab_order, tab_order[1:] + tab_order[:1]):
            self.setTabOrder(first, second)
        self.save_shortcut = QShortcut(QKeySequence("Alt+S"), self)
        self.save_shortcut.setContext(Qt.ShortcutContext.WidgetWithChildrenShortcut)
        self.save_shortcut.activated.connect(self.accept)
        self.control_save_shortcut = QShortcut(QKeySequence("Ctrl+S"), self)
        self.control_save_shortcut.setContext(Qt.ShortcutContext.WidgetWithChildrenShortcut)
        self.control_save_shortcut.activated.connect(self.accept)
        self.inputs.get(initial_focus_field, self.inputs["title"]).setFocus(Qt.FocusReason.OtherFocusReason)
        self.initial_values = self.values()

    def accept(self):
        if not self.values()["title"]:
            self.inputs["title"].setFocus(Qt.FocusReason.OtherFocusReason)
            self.inputs["title"].setAccessibleDescription("Required. Enter a title before saving.")
            event = QAccessibleAnnouncementEvent(self.inputs["title"], "Enter a title before saving.")
            QAccessible.updateAccessibility(event)
            return
        super().accept()

    def reject(self):
        if self.values() != self.initial_values:
            if DiscardChangesDialog(self).exec() != QDialog.DialogCode.Accepted:
                return
        super().reject()

    def values(self):
        return {
            key: (field.toPlainText() if isinstance(field, QTextEdit) else field.text()).strip()
            for key, field in self.inputs.items()
        }


class OnlineMetadataDialog(MetadataDialog):
    def __init__(self, existing, online, service, parent=None):
        self.existing = dict(existing)
        self.online = dict(online)
        super().__init__(self.proposed_values(False), parent)
        self.setWindowTitle("Review Online Metadata")
        options = QHBoxLayout()
        label = QLabel("Apply online values", self)
        self.merge_mode = QComboBox(self)
        self.merge_mode.addItems(["Fill Blank Fields", "Replace Existing"])
        self.merge_mode.setAccessibleName("Apply online values")
        label.setBuddy(self.merge_mode)
        options.addWidget(label)
        options.addWidget(self.merge_mode)
        self.layout().insertWidget(0, QLabel(f"Metadata source: {service}", self))
        self.layout().insertLayout(1, options)
        self.setTabOrder(self.cancel_button, self.merge_mode)
        self.setTabOrder(self.merge_mode, self.inputs["title"])
        self.suggested_values = self.values()
        self.merge_mode.currentIndexChanged.connect(self.change_merge_mode)
        self.merge_mode.setFocus()

    def proposed_values(self, replace):
        values = dict(self.existing)
        for key, _label in self.FIELDS:
            if key == "source":
                continue
            incoming = str(self.online.get(key, "") or "").strip()
            if incoming and (replace or not str(values.get(key, "") or "").strip()):
                values[key] = incoming
        return values

    def change_merge_mode(self, index):
        current = self.values()
        proposed = self.proposed_values(index == 1)
        for key, field in self.inputs.items():
            # Only replace untouched suggestions, not the user's corrections.
            if current[key] != self.suggested_values[key]:
                continue
            value = str(proposed.get(key, "") or "")
            if isinstance(field, QTextEdit):
                field.setPlainText(value)
            else:
                field.setText(value)
            self.suggested_values[key] = self.values()[key]


class ExportFormatDialog(QDialog):
    def __init__(self, parent=None, braille_defaults=None, send_to=None):
        super().__init__(parent)
        self.setWindowTitle("Export Format")
        self.setAccessibleName("Export Format")
        if send_to:
            self.setWindowTitle(f"Send to {send_to}")
            self.setAccessibleName(f"Send to {send_to}")
        layout = QVBoxLayout(self)
        instruction = QLabel(
            "Choose the format for exported copies. Books in your library will not be changed.",
            self,
        )
        instruction.setWordWrap(True)
        layout.addWidget(instruction)
        self.original_format = QRadioButton("Keep original format", self)
        self.epub_format = QRadioButton("Convert exported copies to EPUB", self)
        self.brf_format = QRadioButton("Convert exported copies to &BRF", self)
        self.ebraille_format = QRadioButton("Convert exported copies to eBraille (.ebrl)", self)
        self.original_format.setChecked(True)
        layout.addWidget(self.original_format)
        layout.addWidget(self.epub_format)
        layout.addWidget(self.brf_format)
        layout.addWidget(self.ebraille_format)
        defaults = braille_defaults or BrailleOptions()
        self.braille_settings = QWidget(self)
        form = QFormLayout(self.braille_settings)
        self.braille_grade = QComboBox(self.braille_settings)
        self.braille_grade.addItems(["English UEB contracted (grade 2)", "English UEB uncontracted (grade 1)"])
        self.braille_grade.setCurrentIndex(0 if defaults.contracted else 1)
        self.cells_per_line = QSpinBox(self.braille_settings)
        self.cells_per_line.setRange(10, 120)
        self.cells_per_line.setValue(defaults.cells_per_line)
        self.lines_per_page = QSpinBox(self.braille_settings)
        self.lines_per_page.setRange(5, 100)
        self.lines_per_page.setValue(defaults.lines_per_page)
        self.braille_labels = []
        for title, field in (("Braille code", self.braille_grade), ("Cells per line", self.cells_per_line), ("Lines per page", self.lines_per_page)):
            label = QLabel(title, self.braille_settings)
            label.setBuddy(field)
            field.setAccessibleName(title)
            form.addRow(label, field)
            self.braille_labels.append(label)
        layout.addWidget(self.braille_settings)
        for radio in (self.original_format, self.epub_format, self.brf_format, self.ebraille_format):
            radio.toggled.connect(self.update_braille_settings)
        self.update_braille_settings()
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Cancel, parent=self)
        export_button = buttons.addButton("Send" if send_to else "Export", QDialogButtonBox.ButtonRole.AcceptRole)
        export_button.setDefault(True)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.original_format.setFocus(Qt.FocusReason.OtherFocusReason)

    def export_as_epub(self):
        return self.epub_format.isChecked()

    def update_braille_settings(self):
        self.braille_settings.setVisible(self.brf_format.isChecked() or self.ebraille_format.isChecked())
        for field, label in zip((self.cells_per_line, self.lines_per_page), self.braille_labels[1:]):
            field.setVisible(self.brf_format.isChecked())
            label.setVisible(self.brf_format.isChecked())

    def braille_options(self):
        if not (self.brf_format.isChecked() or self.ebraille_format.isChecked()):
            return None
        return BrailleOptions(
            "brf" if self.brf_format.isChecked() else "ebraille",
            self.braille_grade.currentIndex() == 0,
            self.cells_per_line.value(), self.lines_per_page.value(),
        )


class ConversionReportDialog(QDialog):
    def __init__(self, title, report, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(650, 450)
        layout = QVBoxLayout(self)
        label = QLabel("Conversion report", self)
        self.report = QPlainTextEdit(self)
        self.report.setReadOnly(True)
        self.report.setTabChangesFocus(True)
        self.report.setAccessibleName("Conversion report")
        self.report.setPlainText(report)
        label.setBuddy(self.report)
        layout.addWidget(label)
        layout.addWidget(self.report)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close, parent=self)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.report.setFocus()


class BulkMetadataDialog(QDialog):
    FIELDS = (
        ("author", "Author"), ("source", "Source"), ("tags", "Tags"),
        ("notes", "Notes"), ("edition", "Edition"), ("year", "Year"),
        ("publisher", "Publisher"), ("series", "Series"),
    )

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Bulk Edit Metadata")
        self.resize(560, 660)
        layout = QVBoxLayout(self)
        form = QFormLayout()
        self.inputs = {}
        for key, label_text in self.FIELDS:
            label = QLabel(label_text, self)
            field = TagInput(parent=self) if key == "tags" else QLineEdit(self)
            label.setBuddy(field)
            field.setAccessibleName(label_text)
            form.addRow(label, field)
            self.inputs[key] = field
        layout.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Apply | QDialogButtonBox.StandardButton.Cancel, parent=self)
        # Apply has ApplyRole, so the button box does not emit accepted for it.
        self.apply_button = buttons.button(QDialogButtonBox.StandardButton.Apply)
        self.apply_button.setAccessibleName("Apply")
        self.apply_button.clicked.connect(self.accept)
        self.apply_button.setDefault(True)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.cancel_button = buttons.button(QDialogButtonBox.StandardButton.Cancel)
        tab_order = [*self.inputs.values(), self.apply_button, self.cancel_button]
        for first, second in zip(tab_order, tab_order[1:] + tab_order[:1]):
            self.setTabOrder(first, second)
        self.save_shortcuts = []
        for sequence in ("Ctrl+S", "Alt+S"):
            shortcut = QShortcut(QKeySequence(sequence), self)
            shortcut.setContext(Qt.ShortcutContext.WidgetWithChildrenShortcut)
            shortcut.activated.connect(self.accept)
            self.save_shortcuts.append(shortcut)
        self.inputs["author"].setFocus(Qt.FocusReason.OtherFocusReason)

    def values(self):
        return {key: field.text().strip() for key, field in self.inputs.items() if field.text().strip()}

    def reject(self):
        if self.values():
            if DiscardChangesDialog(self).exec() != QDialog.DialogCode.Accepted:
                return
        super().reject()


class MetadataDetectionWorker(QThread):
    completed = Signal(dict, str)

    def __init__(self, path, existing, parent=None, preserve_source=True):
        super().__init__(parent)
        self.path = Path(path)
        self.existing = dict(existing)
        self.preserve_source = preserve_source

    def run(self):
        try:
            found = detect_metadata_from_text(self.path, existing=self.existing)
            if self.preserve_source:
                found["source"] = self.existing.get("source", "")
            self.completed.emit(found, "")
        except Exception as exc:
            self.completed.emit({}, str(exc))


class ReindexWorker(QThread):
    progress = Signal(str)
    completed = Signal(int, list)

    def run(self):
        database = None
        indexed = 0
        issues = []
        try:
            database = LibraryDatabase()
            rows = database.connection.execute("SELECT id, stored_path FROM books ORDER BY id").fetchall()
            for index, (book_id, stored_path) in enumerate(rows, start=1):
                if self.isInterruptionRequested():
                    break
                path = Path(stored_path)
                self.progress.emit(f"Indexing {index} of {len(rows)}: {path.name}")
                try:
                    if not path.is_file():
                        raise OSError("stored book file is missing")
                    index_book_file(database, book_id, path)
                    if path.suffix.lower() == ".epub":
                        # Re-indexing also refreshes accessibility details so that filters
                        # by feature or hazard work for books imported before they existed.
                        database.update_accessibility_metadata(book_id, read_epub_accessibility_metadata(path))
                    indexed += 1
                except Exception as exc:
                    issues.append(f"{path.name}: {exc}")
        except Exception as exc:
            issues.append(str(exc))
        finally:
            if database:
                database.close()
        self.completed.emit(indexed, issues)


class LibraryBackupWorker(QThread):
    completed = Signal(str, str)

    def __init__(self, destination, automatic=False, parent=None):
        super().__init__(parent)
        self.destination = Path(destination) if destination else None
        self.automatic = automatic

    def run(self):
        database = None
        try:
            database = LibraryDatabase()
            if self.destination is None:
                self.completed.emit("", "No backup folder is configured.")
                return
            books_root = database.books_path.resolve()
            backup_root = self.destination.resolve()
            try:
                backup_root.relative_to(books_root)
                overlaps_library = True
            except ValueError:
                try:
                    books_root.relative_to(backup_root)
                    overlaps_library = True
                except ValueError:
                    overlaps_library = False
            if overlaps_library:
                self.completed.emit("", "The backup folder overlaps the managed Books folder. Choose a separate folder outside the library.")
                return
            if self.automatic:
                schedule = database.get_setting("backup_schedule", DEFAULT_BACKUP_SCHEDULE)
                interval = BACKUP_SCHEDULES.get(schedule, BACKUP_SCHEDULES[DEFAULT_BACKUP_SCHEDULE])[1]
                if interval is None:
                    self.completed.emit("", "")
                    return
                backup_db = self.destination / "library_backup.db"
                backup_books = self.destination / BOOKS_FOLDER
                if backup_db.exists():
                    last_backup = parse_utc_text(database.get_setting("last_backup_at", ""))
                    db_mtime = str(database.db_path.stat().st_mtime)
                    file_count, byte_count = folder_file_stats(database.books_path)
                    signature = f"{file_count}:{byte_count}"
                    changed = (
                        db_mtime != database.get_setting("last_backup_db_mtime", "")
                        or signature != database.get_setting("last_backup_books_signature", "")
                    )
                    if not changed or (last_backup and datetime.now(timezone.utc).replace(tzinfo=None) - last_backup < interval):
                        self.completed.emit("", "")
                        return
            self.destination.mkdir(parents=True, exist_ok=True)
            backup_db = self.destination / "library_backup.db"
            backup_books = self.destination / BOOKS_FOLDER
            manifest_path = self.destination / "library_backup_manifest.json"
            database.connection.commit()
            database.backup_to(backup_db)
            sync_folder_contents(database.books_path, backup_books)
            db_mtime = str(database.db_path.stat().st_mtime)
            book_count = database.connection.execute("SELECT COUNT(*) FROM books").fetchone()[0]
            source_count, source_bytes = folder_file_stats(database.books_path)
            backup_count, backup_bytes = folder_file_stats(backup_books)
            created_at = utc_now_text()
            manifest = {
                "app": APP_NAME,
                "created_at": created_at,
                "source_database": str(database.db_path),
                "backup_database": str(backup_db),
                "source_books_folder": str(database.books_path),
                "backup_books_folder": str(backup_books),
                "source_database_mtime": db_mtime,
                "book_count": book_count,
                "book_file_count": source_count,
                "book_byte_count": source_bytes,
                "backup_book_file_count": backup_count,
                "backup_book_byte_count": backup_bytes,
            }
            manifest_path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")
            database.set_setting("last_backup_at", created_at)
            database.set_setting("last_backup_db_mtime", db_mtime)
            database.set_setting("last_backup_books_signature", f"{source_count}:{source_bytes}")
            database.set_setting("last_seen_backup_file_mtime", str(backup_db.stat().st_mtime))
            self.completed.emit(f"Library and {backup_count} book file(s) backed up to {self.destination}.", "")
        except Exception as exc:
            self.completed.emit("", str(exc))
        finally:
            if database:
                database.close()


class LibraryRestoreWorker(QThread):
    completed = Signal(str, str)

    def __init__(self, backup_folder, preserved_settings, parent=None):
        super().__init__(parent)
        self.backup_folder = Path(backup_folder)
        self.preserved_settings = dict(preserved_settings)

    def run(self):
        backup_db = self.backup_folder / "library_backup.db"
        backup_books = self.backup_folder / BOOKS_FOLDER
        manifest_path = self.backup_folder / "library_backup_manifest.json"
        database = None
        restored_database = None
        safety_db = None
        safety_books = None
        books_path = None
        db_path = None
        try:
            if not backup_db.is_file():
                raise OSError(f"No library database backup was found at:\n{backup_db}")
            database = LibraryDatabase()
            db_path = database.db_path
            books_path = database.books_path
            folder_root = self.backup_folder.resolve()
            books_root = books_path.resolve()
            if folder_root == books_root or folder_root in books_root.parents or books_root in folder_root.parents:
                raise OSError("The backup folder overlaps the live Books folder. Restore was stopped.")
            check_db = sqlite3.connect(backup_db)
            try:
                integrity = check_db.execute("PRAGMA integrity_check").fetchone()
                tables = {row[0] for row in check_db.execute("SELECT name FROM sqlite_master WHERE type='table'")}
                if not integrity or integrity[0] != "ok" or "books" not in tables:
                    raise OSError("The backup database failed its integrity check.")
                book_count = check_db.execute("SELECT COUNT(*) FROM books").fetchone()[0]
                stored_names = {Path(row[0] or "").name for row in check_db.execute("SELECT stored_path FROM books") if Path(row[0] or "").name}
            finally:
                check_db.close()
            manifest = {}
            if manifest_path.is_file():
                manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
                if manifest.get("app") not in {None, APP_NAME}:
                    raise OSError("The backup belongs to a different application.")
                if "book_count" in manifest and int(manifest["book_count"]) != book_count:
                    raise OSError("The backup database does not match its manifest. It may still be syncing.")
            has_book_backup = backup_books.is_dir()
            if has_book_backup:
                file_count, byte_count = folder_file_stats(backup_books)
                if manifest.get("backup_book_file_count") is not None and int(manifest["backup_book_file_count"]) != file_count:
                    raise OSError("The backed-up book files do not match the manifest count.")
                if manifest.get("backup_book_byte_count") is not None and int(manifest["backup_book_byte_count"]) != byte_count:
                    raise OSError("The backed-up book files do not match the manifest size.")
                available = {path.name for path in backup_books.rglob("*") if path.is_file()}
                missing = stored_names - available
                if missing:
                    raise OSError(f"The backup is missing {len(missing)} book file(s) referenced by its database.")
                if book_count and file_count == 0:
                    raise OSError("The backup database has books, but its Books folder is empty.")
            elif book_count and manifest.get("backup_book_file_count") is not None:
                raise OSError("The backup manifest lists book files, but its Books folder is missing.")

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            safety_db = database.folder / f"library_before_restore_{timestamp}.db"
            safety_books = database.folder / f"{BOOKS_FOLDER}_before_restore_{timestamp}"
            database.connection.commit()
            shutil.copy2(db_path, safety_db)
            sync_folder_contents(books_path, safety_books)
            database.close()
            database = None
            shutil.copy2(backup_db, db_path)
            if has_book_backup:
                sync_folder_contents(backup_books, books_path)
            restored_database = LibraryDatabase()
            rows = restored_database.connection.execute("SELECT id, stored_path FROM books").fetchall()
            for book_id, stored_path in rows:
                restored_path = books_path / Path(stored_path or "").name
                if restored_path.is_file():
                    restored_database.connection.execute(
                        "UPDATE books SET stored_path = ?, format = ? WHERE id = ?",
                        (str(restored_path), restored_path.suffix.casefold().lstrip("."), book_id),
                    )
            for key, value in self.preserved_settings.items():
                restored_database.set_setting(key, value)
            restored_database.connection.commit()
            restored_database.close()
            restored_database = None
            suffix = " Database and book files restored." if has_book_backup else " Database restored; no book-file backup was included."
            self.completed.emit(f"Library restored. Safety copies are at {safety_db} and {safety_books}.{suffix}", "")
        except Exception as exc:
            if restored_database:
                restored_database.close()
                restored_database = None
            rollback_errors = []
            try:
                if safety_db and safety_db.is_file() and db_path:
                    shutil.copy2(safety_db, db_path)
                if safety_books and safety_books.is_dir() and books_path:
                    sync_folder_contents(safety_books, books_path)
            except Exception as rollback_exc:
                rollback_errors.append(f"Safety-copy recovery also failed: {rollback_exc}")
            message = str(exc)
            if rollback_errors:
                message += "\n\n" + "\n".join(rollback_errors)
            self.completed.emit("", message)
        finally:
            if database:
                database.close()


class FolderScanWorker(QThread):
    completed = Signal(list, list)

    def __init__(self, folder, include_subfolders, managed_folder, parent=None):
        super().__init__(parent)
        self.folder = Path(folder)
        self.include_subfolders = include_subfolders
        self.managed_folder = Path(managed_folder)

    def run(self):
        files = []
        skipped = []
        iterator = self.folder.rglob("*") if self.include_subfolders else self.folder.glob("*")
        for path in iterator:
            if self.isInterruptionRequested():
                break
            if not path.is_file() or path.suffix.casefold() not in SUPPORTED_EXTENSIONS:
                continue
            try:
                path.resolve().relative_to(self.managed_folder.resolve())
                skipped.append(f"{path} -- already inside the managed library")
                continue
            except ValueError:
                pass
            if LibraryApp.is_ignored_import_file(None, path):
                skipped.append(f"{path} -- ignored helper/notice file")
                continue
            files.append(path)
        self.completed.emit(files, skipped)


class FolderImportWorker(QThread):
    progress = Signal(str)
    completed = Signal(int, list)

    def __init__(self, paths, source, parent=None):
        super().__init__(parent)
        self.paths = list(paths)
        self.source = source

    def run(self):
        imported = 0
        skipped = []
        database = None
        try:
            database = LibraryDatabase()
            operations = LibraryOperations(database)
            for index, source_path in enumerate(self.paths, start=1):
                if self.isInterruptionRequested():
                    break
                self.progress.emit(f"Importing {index} of {len(self.paths)}: {source_path.name}")
                try:
                    original_path = operations.canonical_path_text(source_path)
                    if database.get_book_by_original_path(original_path):
                        skipped.append(f"{source_path} -- already in the library")
                        continue
                    if operations.watched_file_matches_existing_book(source_path, database=database):
                        skipped.append(f"{source_path} -- matching book is already in the library")
                        continue
                    metadata_match, reason = operations.watched_file_matches_existing_metadata(source_path, database=database)
                    if metadata_match:
                        skipped.append(f"{source_path} -- already in library as {metadata_match[1]} ({reason})")
                        continue
                    if source_path.suffix.casefold() == ".zip":
                        added, notes = operations.import_zip_file_without_prompt(source_path, default_source=self.source, database=database)
                        imported += added
                        skipped.extend(notes)
                    elif operations.import_one_book_without_prompt(source_path, default_source=self.source, database=database):
                        imported += 1
                except Exception as exc:
                    skipped.append(f"{source_path} -- {exc}")
            report = database.folder / "last_import_report.txt"
            report.write_text("Accessible Ebook Library Manager import report\n\n" + f"Imported: {imported}\nSkipped: {len(skipped)}\n\n" + "\n".join(skipped), encoding="utf-8")
        except Exception as exc:
            skipped.append(f"Import failed: {exc}")
        finally:
            if database:
                database.close()
        self.completed.emit(imported, skipped)


class CombinedPartsWorker(QThread):
    progress = Signal(str)
    completed = Signal(list, list, str)

    def __init__(self, folder, paths, parent=None):
        super().__init__(parent)
        self.folder = Path(folder)
        self.paths = list(paths)

    def run(self):
        readable = []
        sample = []
        skipped = []
        for index, path in enumerate(self.paths, start=1):
            if self.isInterruptionRequested():
                break
            self.progress.emit(f"Reading part {index} of {len(self.paths)}: {path.name}")
            if path.suffix.casefold() == ".zip":
                skipped.append(f"{path} -- ZIP files cannot be combined into one EPUB")
                continue
            try:
                text = read_text_for_metadata_detection(path, max_chars=250000)
            except Exception as exc:
                skipped.append(f"{path} -- {exc}")
                continue
            if not text.strip():
                skipped.append(f"{path} -- no readable text found")
                continue
            readable.append((path, text))
            if len(" ".join(sample)) < 50000:
                sample.append(text[:10000])
        self.completed.emit(readable, sample, "\n".join(skipped))


class CombinedEpubBuilder:
    def xhtml_body_from_text(self, text):
        return LibraryApp.xhtml_body_from_text(self, text)

    def epub_pagebreak_span(self, page_label, used_page_ids):
        return LibraryApp.epub_pagebreak_span(self, page_label, used_page_ids)

    def create(self, destination, parts, metadata, root):
        LibraryApp.create_combined_epub(self, destination, parts, metadata, root_folder=root)


class CombinedSaveWorker(QThread):
    completed = Signal(int, str)

    def __init__(self, folder, parts, metadata, parent=None):
        super().__init__(parent)
        self.folder = Path(folder)
        self.parts = list(parts)
        self.metadata = dict(metadata)

    def run(self):
        database = None
        try:
            database = LibraryDatabase()
            base = safe_filename(f"{self.metadata['author']} - {self.metadata['title']}").strip(" -") or safe_filename(self.metadata["title"])
            destination = database.books_path / f"{base}.epub"
            number = 2
            while destination.exists():
                destination = database.books_path / f"{base} ({number}).epub"
                number += 1
            CombinedEpubBuilder().create(destination, self.parts, self.metadata, self.folder)
            book_id = database.add_book(
                self.metadata["title"], self.metadata["author"], self.metadata["source"],
                self.metadata["tags"], self.metadata["notes"], str(self.folder.resolve()), str(destination),
            )
            database.update_extra_fields(book_id, self.metadata.get("edition", ""), self.metadata.get("year", ""), self.metadata.get("isbn", ""), self.metadata.get("publisher", ""))
            database.update_series_fields(book_id, self.metadata.get("series", ""), self.metadata.get("series_number", ""))
            database.update_accessibility_metadata(book_id, read_epub_accessibility_metadata(destination))
            index_book_file(database, book_id, destination)
            self.completed.emit(book_id, "")
        except Exception as exc:
            self.completed.emit(-1, str(exc))
        finally:
            if database:
                database.close()


class CopyBooksWorker(QThread):
    progress = Signal(str)
    completed = Signal(int, list)

    def __init__(self, transfers, parent=None):
        super().__init__(parent)
        self.transfers = list(transfers)

    def run(self):
        copied = 0
        skipped = []
        for index, (title, source, destination, replace) in enumerate(self.transfers, start=1):
            if self.isInterruptionRequested():
                break
            self.progress.emit(f"Copying {index} of {len(self.transfers)}: {title}")
            if not source.is_file():
                skipped.append(f"{title}: stored file is missing")
                continue
            if destination.exists() and not replace:
                skipped.append(f"{title}: destination already exists")
                continue
            try:
                shutil.copy2(source, destination)
                copied += 1
            except OSError as exc:
                skipped.append(f"{title}: {exc}")
        self.completed.emit(copied, skipped)


class ExportEpubWorker(QThread):
    progress = Signal(str)
    completed = Signal(int, list)

    def __init__(self, rows, destination_folder, converter, parent=None):
        super().__init__(parent)
        self.rows = list(rows)
        self.destination_folder = Path(destination_folder)
        self.converter = converter

    def run(self):
        exported = 0
        skipped = []
        planned = set()
        try:
            self.destination_folder.mkdir(parents=True, exist_ok=True)
            for index, row in enumerate(self.rows, start=1):
                if self.isInterruptionRequested():
                    break
                title, source = row[1], Path(row[8])
                self.progress.emit(f"Exporting {index} of {len(self.rows)}: {title}")
                if not source.is_file():
                    skipped.append(f"{title}: stored book file is missing")
                    continue
                base = safe_filename(f"{row[2]} - {title}").strip(" -") or safe_filename(title)
                destination = self.destination_folder / f"{base}.epub"
                suffix = 2
                while destination.exists() or destination.name.casefold() in planned:
                    destination = self.destination_folder / f"{base} ({suffix}).epub"
                    suffix += 1
                planned.add(destination.name.casefold())
                try:
                    if source.suffix.casefold() == ".epub":
                        shutil.copy2(source, destination)
                    else:
                        command = [self.converter, str(source), str(destination), "--title", title]
                        if row[2]:
                            command.extend(["--authors", row[2]])
                        if row[4]:
                            command.extend(["--tags", row[4]])
                        result = subprocess.run(command, capture_output=True, text=True, check=False, timeout=300)
                        if result.returncode != 0 or not destination.is_file():
                            details = result.stderr.strip() or result.stdout.strip() or "Conversion did not create an EPUB file."
                            raise RuntimeError(details[:500])
                        try:
                            write_epub_metadata(destination, title, row[2], row[3], row[4], row[5])
                        except Exception:
                            pass
                    exported += 1
                except subprocess.TimeoutExpired:
                    skipped.append(f"{title}: conversion timed out")
                    destination.unlink(missing_ok=True)
                except Exception as exc:
                    skipped.append(f"{title}: {exc}")
                    destination.unlink(missing_ok=True)
        except Exception as exc:
            skipped.append(str(exc))
        self.completed.emit(exported, skipped)


class BrailleExportWorker(QThread):
    progress = Signal(str)
    completed = Signal(int, list, bool, list)

    def __init__(self, rows, destination, options, parent=None):
        super().__init__(parent)
        self.rows = list(rows)
        self.destination = Path(destination)
        self.options = options

    def run(self):
        exported, skipped, cancelled, warnings = 0, [], False, []
        try:
            self.progress.emit("Preparing built-in braille converter...")
            with BrailleEngine(self.isInterruptionRequested) as engine:
                for index, row in enumerate(self.rows, 1):
                    engine.check_cancelled()
                    self.progress.emit(f"Converting {index} of {len(self.rows)}: {row[1]}")
                    try:
                        source = Path(row[8])
                        risks = inspect_conversion_risks(source, engine.check_cancelled) if source.suffix.lower() == ".epub" else []
                        files = engine.convert(Path(row[8]), self.options)
                        risks.extend("Converter warning: " + message for message in engine.warnings)
                        base = safe_filename(f"{row[2]} - {row[1]}").strip(" -") or "Book"
                        publish_files(files, self.destination, base, self.isInterruptionRequested)
                        exported += 1
                        if risks:
                            warnings.append(f"{row[1]}:\n" + "\n".join(risks))
                    except ExportCancelled:
                        raise
                    except Exception as exc:
                        skipped.append(f"{row[1]}: {exc}")
        except ExportCancelled:
            cancelled = True
        except Exception as exc:
            skipped.append(str(exc))
        self.completed.emit(exported, skipped, cancelled, warnings)


class HumanWareMtpWorker(QThread):
    progress = Signal(str)
    completed = Signal(list, list)

    def __init__(self, books, device_name="", parent=None):
        super().__init__(parent)
        self.books = list(books)
        self.device_name = device_name

    def run(self):
        outcomes = []
        devices = []
        for index, (title, source) in enumerate(self.books, start=1):
            self.progress.emit(f"Sending {index} of {len(self.books)}: {title}")
            try:
                result = LibraryApp.copy_to_humanware_mtp(None, source, self.device_name)
                if result.get("code") == "choose_device":
                    devices = result.get("devices") or []
                    self.completed.emit(outcomes, devices)
                    return
                outcomes.append((title, result.get("ok", False), result.get("message", "")))
            except Exception as exc:
                outcomes.append((title, False, str(exc)))
        self.completed.emit(outcomes, devices)


class ConvertBooksWorker(QThread):
    progress = Signal(str)
    completed = Signal(int, list)

    def __init__(self, books, converter, parent=None):
        super().__init__(parent)
        self.books = list(books)
        self.converter = converter

    def run(self):
        database = None
        converted = 0
        issues = []
        try:
            database = LibraryDatabase()
            for index, (row, source) in enumerate(self.books, start=1):
                if self.isInterruptionRequested():
                    break
                self.progress.emit(f"Converting {index} of {len(self.books)}: {row[1]}")
                try:
                    base = safe_filename(f"{row[2]} - {row[1]}").strip(" -") or safe_filename(row[1])
                    destination = database.books_path / f"{base}.epub"
                    suffix = 2
                    while destination.exists():
                        destination = database.books_path / f"{base} ({suffix}).epub"
                        suffix += 1
                    command = [self.converter, str(source), str(destination), "--title", row[1]]
                    if row[2]:
                        command.extend(["--authors", row[2]])
                    if row[4]:
                        command.extend(["--tags", row[4]])
                    result = subprocess.run(command, capture_output=True, text=True, check=False, timeout=300)
                    if result.returncode != 0 or not destination.is_file():
                        raise RuntimeError((result.stderr.strip() or result.stdout.strip() or "Conversion did not produce an EPUB")[:2500])
                    try:
                        write_epub_metadata(destination, row[1], row[2], row[3], row[4], row[5])
                    except Exception:
                        pass
                    new_id = database.add_book(row[1], row[2], row[3], row[4], row[5], str(source), str(destination))
                    database.update_extra_fields(new_id, row[10], row[11], row[12], row[13])
                    database.update_series_fields(new_id, row[21] if len(row) > 21 else "", row[22] if len(row) > 22 else "")
                    database.update_accessibility_metadata(new_id, read_epub_accessibility_metadata(destination))
                    try:
                        index_book_file(database, new_id, destination)
                    except Exception:
                        pass
                    converted += 1
                except Exception as exc:
                    issues.append(f"{row[1]}: {exc}")
        except Exception as exc:
            issues.append(str(exc))
        finally:
            if database:
                database.close()
        self.completed.emit(converted, issues)


class LibraryOperations:
    """Adapt existing non-UI import and watch-folder routines for Qt workers."""

    def __init__(self, database):
        self.db = database

    is_ignored_import_file = LibraryApp.is_ignored_import_file
    is_file_ready_for_import = LibraryApp.is_file_ready_for_import
    guess_metadata_from_file = LibraryApp.guess_metadata_from_file
    unique_destination = LibraryApp.unique_destination
    canonical_path_text = LibraryApp.canonical_path_text
    update_accessibility_from_epub = LibraryApp.update_accessibility_from_epub
    extract_supported_files_from_zip = LibraryApp.extract_supported_files_from_zip
    import_one_book_without_prompt = LibraryApp.import_one_book_without_prompt
    import_zip_file_without_prompt = LibraryApp.import_zip_file_without_prompt
    calibre_metadata_reading_enabled = LibraryApp.calibre_metadata_reading_enabled
    file_signature = LibraryApp.file_signature
    path_is_inside = LibraryApp.path_is_inside
    watched_file_matches_existing_book = LibraryApp.watched_file_matches_existing_book
    watched_file_matches_existing_metadata = LibraryApp.watched_file_matches_existing_metadata
    update_existing_watched_book = LibraryApp.update_existing_watched_book
    removable_windows_drives = LibraryApp.removable_windows_drives
    likely_nls_ereader_drives = LibraryApp.likely_nls_ereader_drives

    def queue_book_for_indexing(self, book_id, stored_path):
        try:
            index_book_file(self.db, book_id, Path(stored_path))
        except Exception:
            pass


class ReaderDeviceDetection:
    removable_windows_drives = LibraryApp.removable_windows_drives
    likely_nls_ereader_drives = LibraryApp.likely_nls_ereader_drives


def duplicate_removal_ids(database):
    rows = database.all_books_for_duplicate_check()
    library_ids = database.source_book_ids_by_book()
    parents = {row[0]: row[0] for row in rows}

    def find(book_id):
        while parents[book_id] != book_id:
            parents[book_id] = parents[parents[book_id]]
            book_id = parents[book_id]
        return book_id

    for index, left in enumerate(rows):
        for right in rows[index + 1:]:
            left_isbn = normalize_isbn_key(left[12] or "")
            right_isbn = normalize_isbn_key(right[12] or "")
            same_isbn = bool(left_isbn and right_isbn and left_isbn == right_isbn)
            same_library_id = bool(library_ids.get(left[0]) and library_ids.get(left[0]) == library_ids.get(right[0]))
            left_author = normalize_duplicate_key(left[2] or "")
            right_author = normalize_duplicate_key(right[2] or "")
            same_title = title_keys_look_same(left[1] or "", right[1] or "", allow_richer_metadata_match=True)
            if same_isbn or same_library_id or (same_title and (not left_author or not right_author or left_author == right_author)):
                parents[find(right[0])] = find(left[0])

    grouped = {}
    for row in rows:
        grouped.setdefault(find(row[0]), []).append(row)
    remove_ids = []
    for group in grouped.values():
        if len(group) < 2:
            continue
        keep = min(
            group,
            key=lambda row: (
                0 if (row[6] or Path(row[8]).suffix).casefold().lstrip(".") == "epub" else 1,
                0 if Path(row[8]).exists() else 1,
                -metadata_score_from_row(row),
                row[9] or "",
                row[0],
            ),
        )
        remove_ids.extend(row[0] for row in group if row[0] != keep[0])
    return remove_ids


class DuplicateCleanupWorker(QThread):
    completed = Signal(int, str)

    def run(self):
        database = None
        try:
            database = LibraryDatabase()
            if database.get_setting("automatic_duplicate_cleanup", "0") != "1":
                self.completed.emit(0, "")
                return
            ids = duplicate_removal_ids(database)
            if ids:
                database.connection.executemany("DELETE FROM books_fts WHERE rowid = ?", ((book_id,) for book_id in ids))
                database.connection.executemany("DELETE FROM books WHERE id = ?", ((book_id,) for book_id in ids))
                database.connection.commit()
            self.completed.emit(len(ids), "")
        except Exception as exc:
            self.completed.emit(0, str(exc))
        finally:
            if database:
                database.close()


class OnlineMetadataWorker(QThread):
    completed = Signal(dict, str, str)

    def __init__(self, metadata, parent=None):
        super().__init__(parent)
        self.metadata = dict(metadata)

    def run(self):
        try:
            found, service = lookup_online_metadata(self.metadata, cancel_check=self.isInterruptionRequested)
            self.completed.emit(found or {}, service or "", "")
        except Exception as exc:
            self.completed.emit({}, "", str(exc))


class CoverImageWorker(QThread):
    completed = Signal(int, str, bytes, str)

    def __init__(self, book_id, url, parent=None):
        super().__init__(parent)
        self.book_id = book_id
        self.url = url

    def run(self):
        try:
            request = urllib.request.Request(self.url, headers={"User-Agent": APP_NAME})
            with urllib.request.urlopen(request, timeout=8) as response:
                data = response.read(2_000_000)
            self.completed.emit(self.book_id, self.url, data, "")
        except Exception as exc:
            self.completed.emit(self.book_id, self.url, b"", str(exc))


class EpubRepairOperations:
    epub_pagebreak_span = LibraryApp.epub_pagebreak_span
    ensure_xhtml_epub_namespace = LibraryApp.ensure_xhtml_epub_namespace
    insert_pagebreaks_in_xhtml_text = LibraryApp.insert_pagebreaks_in_xhtml_text
    add_page_breaks_to_epub_file = LibraryApp.add_page_breaks_to_epub_file
    build_nav_xhtml = LibraryApp.build_nav_xhtml
    rebuild_epub_toc_from_text = LibraryApp.rebuild_epub_toc_from_text
    paragraph_texts_from_xhtml = LibraryApp.paragraph_texts_from_xhtml
    repeated_paragraphs_for_cleanup = LibraryApp.repeated_paragraphs_for_cleanup
    clean_xhtml_text = LibraryApp.clean_xhtml_text
    clean_epub_text_file = LibraryApp.clean_epub_text_file


class EpubRepairWorker(QThread):
    completed = Signal(str, object, str)

    def __init__(self, operation, path, parent=None):
        super().__init__(parent)
        self.operation = operation
        self.path = Path(path)

    def run(self):
        try:
            tools = EpubRepairOperations()
            if self.operation == "pages":
                result = tools.add_page_breaks_to_epub_file(self.path)
            elif self.operation == "toc":
                pages = tools.add_page_breaks_to_epub_file(self.path)
                added, skipped = tools.rebuild_epub_toc_from_text(self.path)
                result = (pages, added, skipped)
            else:
                result = tools.clean_epub_text_file(self.path)
            self.completed.emit(self.operation, result, "")
        except Exception as exc:
            self.completed.emit(self.operation, None, str(exc))


class CleanupWizardWorker(QThread):
    completed = Signal(list, str)

    def __init__(self, book_id, path, choices, parent=None):
        super().__init__(parent)
        self.book_id = book_id
        self.path = Path(path)
        self.choices = dict(choices)

    def run(self):
        database = None
        results = []
        try:
            database = LibraryDatabase()
            operations = LibraryOperations(database)
            repairs = EpubRepairOperations()
            if self.choices.get("accessibility"):
                metadata = operations.update_accessibility_from_epub(self.book_id, self.path)
                results.append("Accessibility metadata checked." if metadata else "No accessibility metadata was found.")
            if self.choices.get("clean"):
                changed = repairs.clean_epub_text_file(self.path)
                results.append(f"Text cleanup made {changed} change(s).")
            if self.choices.get("pages") and not self.choices.get("toc") and not self.choices.get("clean"):
                inserted = repairs.add_page_breaks_to_epub_file(self.path)
                results.append(f"Added {inserted} page-break marker(s).")
            if self.choices.get("toc"):
                inserted = repairs.add_page_breaks_to_epub_file(self.path)
                added, skipped = repairs.rebuild_epub_toc_from_text(self.path)
                results.append(f"Added {inserted} page-break marker(s) before rebuilding navigation. Built {added} TOC item(s); skipped {skipped} without matching anchors.")
            if any(self.choices.get(key) for key in ("clean", "pages", "toc")):
                operations.update_accessibility_from_epub(self.book_id, self.path)
        except Exception as exc:
            self.completed.emit(results, str(exc))
            return
        finally:
            if database:
                database.close()
        self.completed.emit(results, "")


class WatchedFolderWorker(QThread):
    progress = Signal(str)
    completed = Signal(int, int, list, str)

    def run(self):
        database = None
        imported = updated = 0
        skipped = []
        error = ""
        try:
            database = LibraryDatabase()
            try:
                folders = json.loads(database.get_setting("watched_folders", "[]"))
            except (TypeError, ValueError):
                folders = []
            try:
                signatures = json.loads(database.get_setting("watched_file_signatures", "{}"))
            except (TypeError, ValueError):
                signatures = {}
            if not isinstance(signatures, dict):
                signatures = {}
            operations = LibraryOperations(database)
            for folder_text in folders:
                if self.isInterruptionRequested():
                    break
                folder = Path(folder_text)
                if not folder.is_dir():
                    skipped.append(f"{folder} -- watched folder not found")
                    continue
                for path in folder.rglob("*"):
                    if self.isInterruptionRequested():
                        break
                    if not path.is_file() or path.suffix.casefold() not in SUPPORTED_EXTENSIONS:
                        continue
                    if operations.path_is_inside(path, database.books_path) or operations.is_ignored_import_file(path):
                        continue
                    canonical = operations.canonical_path_text(path)
                    signature = operations.file_signature(path)
                    if signature and signatures.get(canonical) == signature:
                        continue
                    self.progress.emit(path.name)
                    try:
                        existing = database.get_book_by_original_path(canonical)
                        if existing and path.suffix.casefold() != ".zip":
                            operations.update_existing_watched_book(existing, path, database=database)
                            updated += 1
                        elif operations.watched_file_matches_existing_book(path, database=database):
                            skipped.append(f"{path} -- already in library")
                        else:
                            match, reason = operations.watched_file_matches_existing_metadata(path, database=database)
                            if match:
                                skipped.append(f"{path} -- already in library as {match[1]} ({reason})")
                            elif path.suffix.casefold() == ".zip":
                                added, notes = operations.import_zip_file_without_prompt(path, default_source="Watched Folder", database=database)
                                imported += added
                                skipped.extend(notes)
                            elif operations.import_one_book_without_prompt(path, default_source="Watched Folder", database=database):
                                imported += 1
                        if signature:
                            signatures[canonical] = signature
                    except Exception as exc:
                        skipped.append(f"{path} -- {exc}")
            database.set_setting("watched_file_signatures", json.dumps(signatures, sort_keys=True))
            if imported or updated or skipped:
                report = database.folder / "last_import_report.txt"
                report.write_text(
                    "Accessible Ebook Library Manager watched-folder report\n\n"
                    f"Imported: {imported}\nUpdated: {updated}\nSkipped: {len(skipped)}\n\n"
                    + "\n".join(skipped),
                    encoding="utf-8",
                )
        except Exception as exc:
            error = str(exc)
        finally:
            if database:
                database.close()
        self.completed.emit(imported, updated, skipped, error)


class QProgressDialog(QtProgressDialog):
    def request_safe_cancel(self):
        worker = getattr(self, "operation_worker", None)
        if not worker or not worker.isRunning():
            return False
        if hasattr(worker, "cancel_callback"):
            worker.cancel_callback()
        elif isinstance(worker, BrailleExportWorker):
            self.parent().cancel_braille_export(worker)
        return True

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Escape and self.request_safe_cancel():
            event.accept()
            return
        super().keyPressEvent(event)

    def closeEvent(self, event):
        if self.request_safe_cancel():
            event.ignore()
            return
        super().closeEvent(event)


class ReviewableStatusBar(QStatusBar):
    def __init__(self, parent):
        super().__init__(parent)
        self.messages = []
        self.pending_announcement = ""
        self.announcement_timer = QTimer(self)
        self.announcement_timer.setSingleShot(True)
        self.announcement_timer.timeout.connect(self.announce_pending)

    def showMessage(self, message, timeout=0):
        super().showMessage(message, timeout)
        # Per-file progress remains in its progress dialog; do not flood history
        # or interrupt screen-reader speech as each file is processed.
        if not message or message.startswith(("Checking:", "Selected:", "Deselected:")):
            return
        if not self.messages or self.messages[-1][1] != message:
            self.messages.append((datetime.now().strftime("%H:%M:%S"), message))
            self.messages = self.messages[-10:]
        database = getattr(self.parent(), "db", None)
        if database and database.get_setting("supplementary_announcements", "0") == "1":
            self.pending_announcement = message
            self.announcement_timer.start(700)

    def announce_pending(self):
        window = self.parent()
        if (not window.db or window.db.get_setting("supplementary_announcements", "0") != "1"
                or QApplication.activeModalWidget() or not window.isActiveWindow()):
            return
        self.setAccessibleName(self.pending_announcement)
        QAccessible.updateAccessibility(QAccessibleEvent(self, QAccessible.Event.Alert))


class RecentMessagesDialog(QDialog):
    def __init__(self, messages, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Recent Messages")
        self.resize(720, 420)
        layout = QVBoxLayout(self)
        self.history = QListWidget(self)
        self.history.setAccessibleName("Recent messages, newest first")
        self.history.setWordWrap(True)
        for stamp, message in reversed(messages[-10:]):
            self.history.addItem(f"{stamp}  {message}")
        if not self.history.count():
            self.history.addItem("No messages in this session.")
        self.history.setCurrentRow(0)
        layout.addWidget(self.history)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close, self)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.history.setFocus()


class PreferencesDialog(QDialog):
    def __init__(self, window):
        super().__init__(window)
        self.setWindowTitle("Settings")
        self.resize(740, 560)
        self.command = None
        layout = QVBoxLayout(self)
        self.tabs = QTabWidget(self)
        self.tabs.setAccessibleName("Settings categories")
        layout.addWidget(self.tabs)
        categories = {"General": []}
        for action in window.settings_commands.actions():
            if action.isSeparator():
                continue
            if action.menu():
                commands = []
                self.collect_commands(action.menu(), commands)
                if action.text() == "File Explorer Integration":
                    categories["General"].extend(commands)
                elif action.text() == "Missing Metadata Alert Sound":
                    continue
                elif action.text() == "Book List Reading":
                    categories.setdefault("Book List Reading", [])
                else:
                    categories.setdefault(action.text(), []).extend(commands)
            else:
                categories["General"].append((action.text(), action))
        for name, commands in categories.items():
            page = QWidget(self)
            page_layout = QVBoxLayout(page)
            if name == "Book List Reading":
                speech_label = QLabel("Book list speech", page)
                self.speech_mode = QComboBox(page)
                self.speech_mode.setAccessibleName("Book list speech")
                speech_label.setBuddy(self.speech_mode)
                current_fields = window.speech_fields()
                selected = -1
                for index, (label, fields) in enumerate(SPEECH_PRESETS):
                    self.speech_mode.addItem(label, fields)
                    if fields == current_fields:
                        selected = index
                if selected < 0:
                    self.speech_mode.addItem("Current custom selection", current_fields)
                    selected = self.speech_mode.count() - 1
                self.speech_mode.setCurrentIndex(selected)
                page_layout.addWidget(speech_label)
                page_layout.addWidget(self.speech_mode)
                sound_label = QLabel("Missing metadata sound", page)
                self.sound_mode = QComboBox(page)
                self.sound_mode.setAccessibleName("Missing metadata sound")
                sound_label.setBuddy(self.sound_mode)
                for mode, label in [("off", "Off"), *((key, value[0]) for key, value in MISSING_METADATA_SOUND_MODES.items())]:
                    self.sound_mode.addItem(label, mode)
                self.sound_mode.setCurrentIndex(self.sound_mode.findData(window.missing_metadata_sound_mode()))
                page_layout.addWidget(sound_label)
                page_layout.addWidget(self.sound_mode)
                test_sound = QPushButton("Test Sound", page)
                test_sound.clicked.connect(lambda: LibraryApp.play_missing_metadata_sound(None))
                page_layout.addWidget(test_sound)
            for label, action in commands:
                button = QPushButton(label, page)
                button.clicked.connect(lambda _checked=False, action=action: self.choose_command(action))
                page_layout.addWidget(button)
            page_layout.addStretch()
            self.tabs.addTab(page, name)
        feedback = QWidget(self)
        feedback_layout = QVBoxLayout(feedback)
        self.announcements = QCheckBox("Announce supplementary status messages", feedback)
        self.announcements.setChecked(window.db.get_setting("supplementary_announcements", "0") == "1")
        feedback_layout.addWidget(self.announcements)
        history = QPushButton("Review Recent Messages", feedback)
        history.clicked.connect(lambda: self.choose_command(window.history_action))
        feedback_layout.addWidget(history)
        feedback_layout.addStretch()
        self.tabs.addTab(feedback, "Feedback")
        for button in self.tabs.tabBar().findChildren(QToolButton):
            button.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Apply | QDialogButtonBox.StandardButton.Cancel, self)
        buttons.accepted.connect(self.accept)
        buttons.button(QDialogButtonBox.StandardButton.Apply).clicked.connect(self.apply_settings)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.tabs.setFocus()

    def collect_commands(self, menu, result, prefix=""):
        for action in menu.actions():
            if action.isSeparator():
                continue
            if action.menu():
                self.collect_commands(action.menu(), result, prefix + action.text() + ": ")
            else:
                result.append((prefix + action.text(), action))

    def choose_command(self, action):
        focused = self.focusWidget()
        action.trigger()
        self.activateWindow()
        if focused:
            focused.setFocus()

    def apply_settings(self):
        window = self.parent()
        if self.speech_mode.currentData() != window.speech_fields():
            window.set_speech_fields(self.speech_mode.currentData())
        window.set_missing_metadata_sound_mode(self.sound_mode.currentData())
        window.db.set_setting("supplementary_announcements", "1" if self.announcements.isChecked() else "0")

    def accept(self):
        self.apply_settings()
        super().accept()


class QtLibraryWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.db = LibraryDatabase()
        try:
            rename_legacy_wishlist_tags(self.db)
        except Exception:
            pass
        self.setStatusBar(ReviewableStatusBar(self))
        self.setWindowTitle(APP_NAME)
        self.resize(1120, 760)
        self.search = QLineEdit(self)
        self.search.setAccessibleName("Search books")
        self.search.setPlaceholderText("Search book metadata and indexed text")
        self.search.returnPressed.connect(self.search_and_focus)
        self.single_list = AccessibleBookList(self)
        self.single_list.currentItemChanged.connect(self.update_position)
        self.groups = []
        self.group_columns = []
        columns = QWidget(self)
        columns_layout = QHBoxLayout(columns)
        columns_layout.setContentsMargins(0, 0, 0, 0)
        columns_layout.setSpacing(8)
        self.columns_container = columns
        self.columns_layout = columns_layout
        self.column_pool = []
        self.column_specs = []
        self.configure_columns(list(BOOK_LIST_ALPHABET_GROUPS))
        self.column_scroll = QScrollArea(self)
        self.column_scroll.setWidgetResizable(True)
        self.column_scroll.setWidget(columns)
        self.views = QStackedWidget(self)
        self.views.addWidget(self.single_list)
        self.views.addWidget(self.column_scroll)

        self.cover_cache = {}
        self.cover_preview_book_id = None
        self.pending_cover_preview_book_id = None
        self.cover_preview_workers = []
        self.cover_lookup_worker = None
        self.cover_preview_timer = QTimer(self)
        self.cover_preview_timer.setSingleShot(True)
        self.cover_preview_timer.timeout.connect(self.load_cover_preview)
        cover_panel = QWidget(self)
        cover_panel.setFixedWidth(205)
        cover_layout = QVBoxLayout(cover_panel)
        cover_heading = QLabel("Cover Preview", cover_panel)
        self.cover_image = QLabel("No cover saved", cover_panel)
        self.cover_image.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.cover_image.setMinimumSize(175, 245)
        self.cover_image.setMaximumSize(175, 245)
        self.cover_image.setStyleSheet("border: 1px solid palette(mid); padding: 4px;")
        self.cover_image.setAccessibleName("")
        self.cover_title = QLabel("", cover_panel)
        self.cover_title.setWordWrap(True)
        self.cover_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.view_cover_button = QPushButton("View Cover", cover_panel)
        self.view_cover_button.clicked.connect(self.view_cover_image)
        self.find_cover_button = QPushButton("Find Cover", cover_panel)
        self.find_cover_button.clicked.connect(self.find_cover_image)
        cover_layout.addWidget(cover_heading)
        cover_layout.addWidget(self.cover_image, 0, Qt.AlignmentFlag.AlignHCenter)
        cover_layout.addWidget(self.cover_title)
        cover_layout.addWidget(self.view_cover_button)
        cover_layout.addWidget(self.find_cover_button)
        cover_layout.addStretch(1)

        center = QWidget(self)
        layout = QVBoxLayout(center)
        layout.addWidget(self.search)
        content_layout = QHBoxLayout()
        content_layout.addWidget(self.views, 1)
        content_layout.addWidget(cover_panel)
        layout.addLayout(content_layout, 1)
        self.setCentralWidget(center)
        self.statusBar().showMessage("Ready")
        self.current_rows = []
        self.marked_ids = set()
        self.active_column = 0
        self.active_worker = None
        self.progress_dialog = None
        self.sort_by = "title"
        self.filter_source = ""
        self.filter_tag = ""
        self.filter_format = ""
        self.filter_accessibility_feature = ""
        self.filter_hide_hazard = ""
        self.filter_content_flag = ""
        self.filter_file_access = ""
        self.active_saved_view = ""
        self.watch_scan_running = False
        self.duplicate_worker = None
        self.metadata_worker = None
        self.pending_lookup_row = None
        self.last_missing_metadata_sound_book_id = None
        self.backup_worker = None
        self._closing = False
        self.build_menus()
        self.bind_shortcuts()
        self.set_view_mode(self.db.get_setting("book_list_display_mode", "single"), initial=True)
        self.refresh_books()
        QTimer.singleShot(15000, self.check_watched_folders)
        QTimer.singleShot(60 * 60 * 1000, self.check_library_backup)
        if self.db.get_setting("automatic_duplicate_cleanup", "0") == "1":
            QTimer.singleShot(10000, self.automatic_duplicate_cleanup)

    def build_menus(self):
        file_menu = self.menuBar().addMenu("&File")
        self.add_action(file_menu, "Add Book...", self.add_book, "Ctrl+N")
        self.add_action(file_menu, "Import Folder...", self.import_folder, "Ctrl+Shift+N")
        self.add_action(file_menu, "Import Goodreads Export...", self.import_goodreads_export)
        self.add_action(file_menu, "Export Library for Goodreads...", self.export_library_for_goodreads)
        self.add_action(file_menu, "Import Audiobook Folder or Drive...", self.import_audiobook_folder)
        self.add_action(file_menu, "Open Book", self.open_book, "Ctrl+O")
        self.add_action(file_menu, "Export Book Copy...", self.export_book, "Ctrl+E")
        self.add_action(file_menu, "Open Kindle for PC", self.open_kindle, "Ctrl+K")
        send_menu = file_menu.addMenu("Send To")
        self.add_action(send_menu, "Voice Dream...", lambda: self.send_to_folder("Voice Dream", "voice_dream_loader_folder"), "Ctrl+Shift+V")
        self.add_action(send_menu, "Dolphin EasyReader...", lambda: self.send_to_folder("Dolphin EasyReader", "dolphin_easyreader_folder"))
        self.add_action(send_menu, "NLS eReader...", self.send_to_nls_ereader, "Ctrl+Shift+E")
        self.add_action(send_menu, "Kindle...", self.send_to_kindle, "Ctrl+Shift+K")
        self.add_action(send_menu, "HumanWare Braille eReader (MTP)...", self.send_to_humanware)
        self.add_action(file_menu, "Exit", self.close, "Alt+F4")

        book_menu = self.menuBar().addMenu("&Book")
        self.add_action(book_menu, "Edit Metadata...", self.edit_book, "F2")
        self.add_action(book_menu, "Bulk Edit Selected Metadata...", self.bulk_edit_metadata)
        self.add_action(book_menu, "Find Missing Metadata Wizard...", self.auto_detect_metadata, "Ctrl+D")
        self.add_action(book_menu, "Look Up Metadata Online...", self.lookup_metadata_online)
        self.add_action(book_menu, "Find Book on Other Services...", self.find_book_elsewhere)
        repair_menu = book_menu.addMenu("Repair")
        self.add_action(repair_menu, "Cleanup and Repair Wizard...", self.cleanup_repair_wizard)
        self.add_action(repair_menu, "Check EPUB Accessibility Metadata", self.check_epub_accessibility)
        repair_menu.addSeparator()
        self.add_action(repair_menu, "Add EPUB Page Breaks from Page Labels...", lambda: self.run_epub_repair("pages"))
        self.add_action(repair_menu, "Rebuild EPUB Table of Contents...", lambda: self.run_epub_repair("toc"))
        self.add_action(repair_menu, "Clean Repaired EPUB Text...", lambda: self.run_epub_repair("clean"))
        self.add_action(repair_menu, "Combine Selected Books into One EPUB...", self.combine_selected_books)
        self.add_action(repair_menu, "Remove Repair Backup Files...", self.remove_repair_backups)
        self.add_action(book_menu, "Show Book Details", self.show_details, "Ctrl+I")
        self.add_action(book_menu, "Focus Book List", self.focus_book_list, "Ctrl+L")
        self.add_action(book_menu, "View Cover Image...", self.view_cover_image)
        self.add_action(book_menu, "Find Cover Image Online...", self.find_cover_image)
        self.add_action(book_menu, "Convert Selected Books to EPUB...", self.convert_selected_to_epub, "Ctrl+R")
        self.add_action(book_menu, "Select All Books (Ctrl+A)", self.select_all)
        self.add_action(book_menu, "Deselect All Books (Ctrl+Shift+A)", self.deselect_all)
        self.add_action(book_menu, "Delete from Library...", self.delete_books, "Delete")

        search_menu = self.menuBar().addMenu("&Search")
        self.add_action(search_menu, "Focus Search", self.focus_search, "Ctrl+F")
        self.add_action(search_menu, "Clear Search", self.clear_search, "Escape")
        self.add_action(search_menu, "Focus Book List", self.focus_book_list)
        self.add_action(search_menu, "Explain Search", self.explain_search)
        self.add_action(search_menu, "Re-index Book Content...", self.reindex_content)

        organize_menu = self.menuBar().addMenu("&Organize")
        display_menu = organize_menu.addMenu("Book List Display")
        self.view_actions = {}
        for mode, label in VIEW_LABELS.items():
            self.view_actions[mode] = self.add_action(display_menu, label, lambda _checked=False, mode=mode: self.set_view_mode(mode))
            self.view_actions[mode].setCheckable(True)
        self.single_action = self.view_actions["single"]
        self.multiple_action = self.view_actions["multiple"]
        self.add_action(display_menu, "Show Current Display", self.show_view_mode)
        orientation_menu = organize_menu.addMenu("Category Layout")
        self.orientation_actions = {}
        for orientation, label in (("columns", "Columns"), ("rows", "Rows")):
            action = self.add_action(orientation_menu, label, lambda _checked=False, orientation=orientation: self.set_category_layout(orientation))
            action.setCheckable(True)
            action.setChecked(orientation == self.category_layout())
            self.orientation_actions[orientation] = action
        sort_menu = organize_menu.addMenu("Sort")
        for label, sort_key in (
            ("Title A to Z", "title"), ("Title Z to A", "title_desc"),
            ("Author A to Z", "author"), ("Author Z to A", "author_desc"),
            ("Series A to Z", "series"), ("Series Z to A", "series_desc"),
            ("Published Year Newest to Oldest", "date"),
            ("Published Year Oldest to Newest", "date_oldest"),
            ("Date Added Newest to Oldest", "date_added"),
            ("Date Added Oldest to Newest", "date_added_oldest"),
        ):
            self.add_action(sort_menu, label, lambda sort_key=sort_key: self.set_sort(sort_key))
        filter_menu = organize_menu.addMenu("Filter")
        self.add_action(filter_menu, "By Source...", lambda: self.set_filter("source"))
        self.add_action(filter_menu, "By Tag...", lambda: self.set_filter("tag"))
        self.add_action(filter_menu, "By Format...", lambda: self.set_filter("format"))
        self.add_action(filter_menu, "By Accessibility Feature...", lambda: self.set_accessibility_filter("feature"))
        self.add_action(filter_menu, "Hide Books With Accessibility Hazard...", lambda: self.set_accessibility_filter("hazard"))
        self.add_action(filter_menu, "Show Only Books I Have", lambda: self.set_file_access_filter("have"))
        self.add_action(filter_menu, "Show Only Books I Do Not Have", lambda: self.set_file_access_filter("missing"))
        self.add_action(filter_menu, "Show Books Whether or Not I Have Them", lambda: self.set_file_access_filter(""))
        self.add_action(filter_menu, "Clear Filters", self.clear_filters)
        saved_menu = organize_menu.addMenu("Saved Views")
        for label, view in (("Missing Metadata", "missing_metadata"), ("Recently Added", "recently_added"), ("EPUB Books", "epub"), ("Bookshare Books", "bookshare"), ("Image-Only PDFs That Need OCR", "image_only_pdf")):
            self.add_action(saved_menu, label, lambda view=view: self.apply_saved_view(view))
        self.add_action(saved_menu, "Save Current View...", self.save_current_view)
        self.add_action(saved_menu, "Choose Saved View...", self.choose_saved_view)
        self.add_action(saved_menu, "Clear Saved View", self.clear_saved_view)
        organize_menu.addSeparator()
        self.add_action(organize_menu, "Remove Duplicate Books, Prefer EPUB...", self.remove_duplicates)
        self.add_action(organize_menu, "Toggle Automatic Duplicate Cleanup", self.toggle_automatic_duplicate_cleanup)
        self.add_action(organize_menu, "Show Sort and Filter Settings", self.show_organize_settings)

        settings_menu = self.menuBar().addMenu("&Settings")
        self.add_action(settings_menu, "Set Voice Dream Loader Folder...", lambda: self.configure_folder("Voice Dream Loader", "voice_dream_loader_folder"))
        self.add_action(settings_menu, "Set Dolphin EasyReader Folder...", lambda: self.configure_folder("Dolphin EasyReader", "dolphin_easyreader_folder"))
        self.add_action(settings_menu, "Set NLS eReader Folder...", lambda: self.configure_folder("NLS eReader", "nls_ereader_folder"))
        self.add_action(settings_menu, "Set Kindle Email Addresses...", self.set_kindle_emails)
        reader_menu = settings_menu.addMenu("Ebook Reader")
        self.add_action(reader_menu, "Choose Default Reader...", self.choose_default_reader)
        self.add_action(reader_menu, "Use System Default Reader", self.clear_default_reader)
        self.add_action(settings_menu, "Open Library Folder", self.open_library_folder)
        calibre_menu = settings_menu.addMenu("Kindle and Calibre")
        self.add_action(calibre_menu, "Show Calibre Tools Status", self.show_calibre_status)
        self.add_action(calibre_menu, "Toggle Calibre Metadata Reading", self.toggle_calibre_metadata)
        settings_menu.addSeparator()
        speech_menu = settings_menu.addMenu("Book List Reading")
        for label, fields in SPEECH_PRESETS:
            self.add_action(speech_menu, label, lambda fields=fields: self.set_speech_fields(fields))
        self.add_action(speech_menu, "Show Current Reading", self.show_speech_fields)
        sound_menu = settings_menu.addMenu("Missing Metadata Alert Sound")
        for mode, label in [("off", "Off"), *((key, value[0]) for key, value in MISSING_METADATA_SOUND_MODES.items())]:
            self.add_action(sound_menu, label, lambda mode=mode: self.set_missing_metadata_sound_mode(mode))
        sound_menu.addSeparator()
        self.add_action(sound_menu, "Show Current Setting", self.show_missing_metadata_sound_mode)
        self.add_action(sound_menu, "Test Sound", lambda: LibraryApp.play_missing_metadata_sound(None))
        settings_menu.addSeparator()
        backup_menu = settings_menu.addMenu("Library Backup")
        self.add_action(backup_menu, "Choose Backup Folder...", self.choose_backup_folder)
        schedule_menu = backup_menu.addMenu("Backup Schedule")
        for key, (label, _interval) in BACKUP_SCHEDULES.items():
            self.add_action(schedule_menu, label, lambda key=key: self.set_backup_schedule(key))
        self.add_action(backup_menu, "Back Up Now", self.backup_library_now)
        self.add_action(backup_menu, "Restore From Backup...", self.restore_library_backup)
        self.add_action(backup_menu, "Show Backup Status", self.show_backup_status)
        settings_menu.addSeparator()
        watch_menu = settings_menu.addMenu("Watched Folders")
        self.add_action(watch_menu, "Add Watched Folder...", self.add_watched_folder)
        self.add_action(watch_menu, "Remove Watched Folder...", self.remove_watched_folder)
        self.add_action(watch_menu, "Scan Watched Folders Now", lambda: self.scan_watched_folders(manual=True))
        self.add_action(watch_menu, "Toggle Automatic Scanning", self.toggle_watched_auto_scan)
        self.add_action(watch_menu, "Show Watched Folder Status", self.show_watched_folder_status)
        explorer_menu = settings_menu.addMenu("File Explorer Integration")
        self.add_action(explorer_menu, "Add Right-Click Command...", self.install_file_explorer_context_menu)
        self.add_action(explorer_menu, "Remove Right-Click Command...", self.remove_file_explorer_context_menu)
        self.add_action(explorer_menu, "Show Integration Status", self.show_file_explorer_context_menu_status)
        self.settings_commands = settings_menu
        self.menuBar().removeAction(settings_menu.menuAction())
        settings_menu = self.menuBar().addMenu("&Settings")
        self.add_action(settings_menu, "Settings...", self.show_preferences)
        help_menu = self.menuBar().addMenu("&Help")
        self.add_action(help_menu, "Help", self.show_help, "F1")
        self.history_action = self.add_action(help_menu, "Review Recent Messages...", self.show_message_history)

    def show_preferences(self):
        dialog = PreferencesDialog(self)
        dialog.exec()
        if dialog.command:
            dialog.command.trigger()

    def show_message_history(self):
        RecentMessagesDialog(self.statusBar().messages, self).exec()

    @staticmethod
    def add_action(menu, label, callback, shortcut=None):
        action = QAction(label, menu)
        if shortcut:
            action.setShortcut(QKeySequence(shortcut))
        action.triggered.connect(lambda _checked=False: callback())
        menu.addAction(action)
        return action

    def bind_shortcuts(self):
        for key, mode in (("Ctrl+1", "single"), ("Ctrl+2", "multiple"), ("Ctrl+3", "letters"), ("Ctrl+4", "tags")):
            shortcut = QShortcut(QKeySequence(key), self)
            shortcut.setContext(Qt.ShortcutContext.WindowShortcut)
            shortcut.activated.connect(lambda mode=mode: self.set_view_mode(mode))
    def set_view_mode(self, mode, initial=False):
        current = self.selected_row() if not initial else None
        current_id = current[0] if current else None
        mode = mode if mode in VIEW_LABELS else "single"
        self.views.setCurrentIndex(0 if mode == "single" else 1)
        for key, action in self.view_actions.items():
            action.setChecked(key == mode)
            action.setText(self.view_label(key))
        self.db.set_setting("book_list_display_mode", mode)
        if not initial:
            self.refresh_books(current_id)
            self.statusBar().showMessage(self.view_label(mode))

    def category_layout(self):
        return "rows" if self.db.get_setting("book_list_category_layout", "columns") == "rows" else "columns"

    def view_label(self, mode):
        label = VIEW_LABELS.get(mode, VIEW_LABELS["single"])
        if self.category_layout() == "rows":
            label = label.replace("Column", "Row")
        return label

    def set_category_layout(self, orientation):
        current = self.selected_row()
        orientation = "rows" if orientation == "rows" else "columns"
        self.db.set_setting("book_list_category_layout", orientation)
        for key, action in self.orientation_actions.items():
            action.setChecked(key == orientation)
        for key, action in self.view_actions.items():
            action.setText(self.view_label(key))
        self.refresh_books(current[0] if current else None)
        self.setFocusToCurrentBook()
        self.statusBar().showMessage(f"Category layout: {orientation}")

    def set_speech_fields(self, fields):
        current = self.selected_row()
        self.db.set_setting("book_list_speech_fields", ",".join(fields))
        self.refresh_books(current[0] if current else None)
        self.statusBar().showMessage("Book list reading updated")

    def show_speech_fields(self):
        labels = dict(BOOK_LIST_SPEECH_FIELDS)
        current = ", ".join(labels[key] for key in self.speech_fields())
        QMessageBox.information(self, "Book List Reading", f"The book list currently reads: {current}.")

    def show_view_mode(self):
        mode = self.db.get_setting("book_list_display_mode", "single")
        label = self.view_label(mode)
        QMessageBox.information(self, "Book List Display", f"The book list is currently shown as: {label}.")

    def missing_metadata_sound_mode(self):
        mode = self.db.get_setting("missing_metadata_sound_mode", "")
        if mode in MISSING_METADATA_SOUND_MODES or mode == "off":
            return mode
        if self.db.get_setting("missing_metadata_sound", "0") == "1":
            return "complete"
        return DEFAULT_MISSING_METADATA_SOUND_MODE

    def set_missing_metadata_sound_mode(self, mode):
        self.db.set_setting("missing_metadata_sound_mode", mode)
        self.db.set_setting("missing_metadata_sound", "0" if mode == "off" else "1")
        self.last_missing_metadata_sound_book_id = None
        label = "Off" if mode == "off" else MISSING_METADATA_SOUND_MODES[mode][0]
        self.routine_feedback(f"Missing metadata alert sound: {label}.")

    def show_missing_metadata_sound_mode(self):
        mode = self.missing_metadata_sound_mode()
        label = "Off" if mode == "off" else MISSING_METADATA_SOUND_MODES[mode][0]
        QMessageBox.information(self, "Missing Metadata Alert Sound", f"Current setting: {label}.")

    def announce_missing_metadata_if_needed(self, item):
        book_id = item.data(Qt.ItemDataRole.UserRole)
        mode = self.missing_metadata_sound_mode()
        if mode == "off":
            self.last_missing_metadata_sound_book_id = None
            return
        row = self.db.get_book(book_id)
        if not row:
            return
        indexes = {"author": 2, "tags": 4, "edition": 10, "year": 11, "isbn": 12, "publisher": 13}
        required = MISSING_METADATA_SOUND_MODES[mode][1]
        missing = any(not (book_tags(row[4]) if key == "tags" else str(row[indexes[key]] or "").strip()) for key in required)
        if missing and book_id != self.last_missing_metadata_sound_book_id:
            LibraryApp.play_missing_metadata_sound(None)
            self.last_missing_metadata_sound_book_id = book_id
        elif not missing:
            self.last_missing_metadata_sound_book_id = None

    def explain_search(self):
        QMessageBox.information(
            self, "Search Library",
            "Search checks book metadata and indexed full text. Metadata includes title, author, series, source, tags, format, notes, edition, year, ISBN, and publisher. Use Re-index Book Content to rebuild text search for the existing library.",
        )

    def reindex_content(self):
        answer = QMessageBox.question(
            self, "Re-index Book Content",
            "Rebuild the full-text search index for every book? This may take several minutes, but you can continue using the app.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        self.progress_dialog = QProgressDialog("Preparing full-text indexing...", "", 0, 0, self)
        self.progress_dialog.setWindowTitle("Re-index Book Content")
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.setCancelButton(None)
        self.progress_dialog.show()
        worker = ReindexWorker(self)
        self.active_worker = worker
        worker.progress.connect(self.progress_dialog.setLabelText)
        self.connect_worker_result(worker, self.reindex_finished)
        worker.start()

    def reindex_finished(self, indexed, issues):
        if self.progress_dialog:
            self.progress_dialog.close()
            self.progress_dialog = None
        message = f"Indexed text from {indexed} book(s)."
        if issues:
            message += "\n\nSome files could not be indexed:\n" + "\n".join(issues[:15])
            QMessageBox.warning(self, "Re-indexing Finished with Warnings", message)
        else:
            self.routine_feedback(message)

    def set_sort(self, sort_by):
        self.sort_by = sort_by
        self.refresh_books()
        self.statusBar().showMessage(f"Sorted by {sort_by.replace('_', ' ')}")

    def set_filter(self, field):
        current = {"source": self.filter_source, "tag": self.filter_tag, "format": self.filter_format}[field]
        value, accepted = QInputDialog.getText(self, f"Filter by {field.title()}", f"Enter {field} text. Leave blank to clear this filter:", text=current)
        if not accepted:
            return
        if field == "source":
            self.filter_source = value.strip()
        elif field == "tag":
            self.filter_tag = value.strip()
        else:
            self.filter_format = value.strip().lower()
        self.active_saved_view = ""
        self.refresh_books()

    ACCESSIBILITY_FEATURE_CHOICES = (
        "structuralNavigation", "tableOfContents", "pageNavigation", "alternativeText",
        "longDescription", "readingOrder", "displayTransformability", "MathML",
        "synchronizedAudioText", "index", "printPageNumbers", "taggedPDF",
    )
    HAZARD_CHOICES = ("flashing", "motionSimulation", "sound")

    def clear_accessibility_filters(self):
        self.filter_accessibility_feature = ""
        self.filter_hide_hazard = ""
        self.filter_content_flag = ""
        self.filter_file_access = ""

    @staticmethod
    def row_has_book_file(row):
        """True when the book's file is in the library. No File Available entries never are."""
        stored = row[13] if len(row) > 13 else ""
        return bool(stored) and not is_wishlist_path(stored) and Path(stored).exists()

    def set_file_access_filter(self, mode):
        self.filter_file_access = mode
        self.active_saved_view = ""
        self.refresh_books()
        message = {
            "have": "Showing only books you have.",
            "missing": "Showing only books you do not have: No File Available entries and books whose file is missing.",
        }.get(mode, "Showing all books, whether or not you have them.")
        self.routine_feedback(message)

    def set_accessibility_filter(self, kind):
        """Ask which accessibility feature to require, or which hazard to hide."""
        if kind == "feature":
            title = "Filter by Accessibility Feature"
            label = "Show only books that declare this feature. Leave blank to clear this filter:"
            choices, current = self.ACCESSIBILITY_FEATURE_CHOICES, self.filter_accessibility_feature
        else:
            title = "Hide Books With Accessibility Hazard"
            label = "Hide books that declare this hazard. Leave blank to clear this filter:"
            choices, current = self.HAZARD_CHOICES, self.filter_hide_hazard
        items = ["", *choices]
        start = items.index(current) if current in items else 0
        value, accepted = QInputDialog.getItem(self, title, label, items, start, True)
        if not accepted:
            return
        if kind == "feature":
            self.filter_accessibility_feature = value.strip()
        else:
            self.filter_hide_hazard = value.strip()
        self.active_saved_view = ""
        self.refresh_books()

    def clear_filters(self):
        self.filter_source = self.filter_tag = self.filter_format = ""
        self.clear_accessibility_filters()
        self.active_saved_view = ""
        self.refresh_books()
        self.statusBar().showMessage("Filters cleared")

    def show_organize_settings(self):
        sort_labels = {
            "title": "Title A to Z", "title_desc": "Title Z to A",
            "author": "Author A to Z", "author_desc": "Author Z to A",
            "series": "Series A to Z", "series_desc": "Series Z to A",
            "date": "Published Year Newest to Oldest", "date_oldest": "Published Year Oldest to Newest",
            "date_added": "Date Added Newest to Oldest", "date_added_oldest": "Date Added Oldest to Newest",
        }
        filters = []
        if self.filter_source:
            filters.append(f"Source contains {self.filter_source}")
        if self.filter_tag:
            filters.append(f"Tag contains {self.filter_tag}")
        if self.filter_format:
            filters.append(f"Format is {self.filter_format}")
        if self.filter_accessibility_feature:
            filters.append(f"Accessibility feature: {self.filter_accessibility_feature}")
        if self.filter_hide_hazard:
            filters.append(f"Hiding books with hazard: {self.filter_hide_hazard}")
        if self.filter_content_flag:
            filters.append(f"Content note: {self.filter_content_flag.replace('_', ' ')}")
        if self.filter_file_access:
            filters.append("Only books I have" if self.filter_file_access == "have" else "Only books I do not have")
        if self.active_saved_view:
            filters.append(f"Saved view: {self.active_saved_view.replace('_', ' ')}")
        QMessageBox.information(
            self, "Sort and Filter Settings",
            f"Sort: {sort_labels.get(self.sort_by, 'Title A to Z')}\nFilters: {'; '.join(filters) if filters else 'None'}",
        )

    def apply_saved_view(self, name):
        self.active_saved_view = name
        self.filter_source = self.filter_tag = self.filter_format = ""
        self.clear_accessibility_filters()
        if name == "recently_added":
            self.sort_by = "date_added"
        elif name == "epub":
            self.filter_format = "epub"
        elif name == "bookshare":
            self.filter_source = "Bookshare"
        elif name == "image_only_pdf":
            self.filter_content_flag = IMAGE_ONLY_PDF
        self.search.clear()
        self.refresh_books()
        self.statusBar().showMessage(f"Saved view: {name.replace('_', ' ')}")

    def saved_views(self):
        import json
        try:
            value = json.loads(self.db.get_setting("saved_views", "{}"))
            return value if isinstance(value, dict) else {}
        except (TypeError, ValueError):
            return {}

    def save_current_view(self):
        name, accepted = QInputDialog.getText(self, "Save Current View", "View name:")
        if not accepted or not name.strip():
            return
        views = self.saved_views()
        views[name.strip()] = {
            "query": self.search.text(), "sort_by": self.sort_by,
            "source": self.filter_source, "tag": self.filter_tag,
            "format": self.filter_format, "view": self.active_saved_view,
            "accessibility_feature": self.filter_accessibility_feature,
            "hide_hazard": self.filter_hide_hazard, "content_flag": self.filter_content_flag,
            "file_access": self.filter_file_access,
        }
        import json
        self.db.set_setting("saved_views", json.dumps(views, sort_keys=True))
        self.statusBar().showMessage(f"Saved view: {name.strip()}")

    def choose_saved_view(self):
        views = self.saved_views()
        if not views:
            self.routine_feedback("There are no custom saved views yet.")
            return
        name, accepted = QInputDialog.getItem(self, "Choose Saved View", "Saved view:", sorted(views), 0, False)
        if accepted:
            view = views[name]
            self.search.setText(view.get("query", ""))
            self.sort_by = view.get("sort_by", "title")
            self.filter_source = view.get("source", "")
            self.filter_tag = view.get("tag", "")
            self.filter_format = view.get("format", "")
            self.filter_accessibility_feature = view.get("accessibility_feature", "")
            self.filter_hide_hazard = view.get("hide_hazard", "")
            self.filter_content_flag = view.get("content_flag", "")
            self.filter_file_access = view.get("file_access", "")
            self.active_saved_view = view.get("view", "")
            self.refresh_books()

    def clear_saved_view(self):
        self.active_saved_view = ""
        self.refresh_books()

    def speech_fields(self):
        configured = self.db.get_setting("book_list_speech_fields", "title,author")
        return [key for key in configured.split(",") if key in {field for field, _ in BOOK_LIST_SPEECH_FIELDS}] or ["title", "author"]

    @staticmethod
    def group_index(title):
        initial = title_navigation_key(title)[:1]
        for index, (_label, letters) in enumerate(BOOK_LIST_ALPHABET_GROUPS):
            if letters and initial in letters:
                return index
        return len(BOOK_LIST_ALPHABET_GROUPS) - 1

    def configure_columns(self, specs):
        orientation = self.category_layout()
        if specs == self.column_specs and getattr(self, "configured_orientation", None) == orientation:
            return
        self.configured_orientation = orientation
        self.columns_layout.setDirection(QBoxLayout.Direction.TopToBottom if orientation == "rows" else QBoxLayout.Direction.LeftToRight)
        shared_row = self.column_pool[0][1].accessible_row if self.column_pool else None
        for view_group, view in self.column_pool:
            view.set_accessibility_active(False)
            view.clear()
            view_group.hide()
        while len(self.column_pool) < len(specs):
            index = len(self.column_pool)
            group = QWidget(self.columns_container)
            group.setMinimumWidth(160)
            layout = QVBoxLayout(group)
            layout.setContentsMargins(0, 0, 0, 0)
            view = AccessibleBookList(group, index, self.move_to_column, shared_row)
            if shared_row is None:
                shared_row = view.accessible_row
                shared_row.setParent(self.columns_container)
            heading = BookColumnHeading("", view, group)
            view.heading = heading
            layout.addWidget(heading)
            layout.addWidget(view)
            self.columns_layout.addWidget(group, 1)
            view.scroll_area.verticalScrollBar().valueChanged.connect(view.update_accessible_row)
            view.scroll_area.horizontalScrollBar().valueChanged.connect(view.update_accessible_row)
            view.currentItemChanged.connect(self.update_position)
            self.column_pool.append((group, view))
        self.column_specs = specs
        self.group_columns = []
        self.groups = []
        for index, (label, _key) in enumerate(specs):
            group, view = self.column_pool[index]
            view.horizontal_books = orientation == "rows"
            view.setAccessibleDescription(
                "Use Left and Right Arrow to move between books. Use Up and Down Arrow to move between categories."
                if view.horizontal_books else
                "Use Up and Down Arrow to move between books. Use Left and Right Arrow to move between categories."
            )
            view.column_canvas.update_content_height()
            view.column_label = label
            view.heading.caption = label
            view.heading.update()
            view.heading.setAccessibleName("")
            group.show()
            self.group_columns.append(group)
            self.groups.append(view)
        self.active_column = min(getattr(self, "active_column", 0), len(self.groups) - 1)

    def row_text(self, row):
        fields = {
            "title": row[1], "author": row[2], "source": row[3], "tags": row[4],
            "format": row[5], "added_at": row[6], "edition": row[7], "year": row[8],
            "isbn": row[9], "publisher": row[10], "series": row[11], "series_number": row[12],
        }
        text = ". ".join(str(fields[key] or f"No {key.replace('_', ' ')}") for key in self.speech_fields())
        if len(row) > 13 and is_wishlist_path(row[13]):
            text += ". No file available"
        return text

    def refresh_books(self, selected_id=None, focus_books=True):
        modal = QApplication.activeModalWidget()
        focus_books = focus_books and not (modal and self.isAncestorOf(modal))
        query = self.search.text().strip()
        content_ids = self.db.search_content(query) if query else None
        rows = self.db.search_books(
            query, sort_by=self.sort_by, source_filter=self.filter_source,
            tag_filter=self.filter_tag, format_filter=self.filter_format,
            accessibility_feature_filter=self.filter_accessibility_feature,
            exclude_hazard_filter=self.filter_hide_hazard,
            content_flag_filter=self.filter_content_flag,
            extra_ids=content_ids,
        )
        if self.filter_file_access:
            wanted = self.filter_file_access == "have"
            rows = [row for row in rows if self.row_has_book_file(row) == wanted]
        if self.active_saved_view == "missing_metadata":
            rows = [row for row in rows if any(not str(row[index] or "").strip() for index in (2, 7, 8, 9, 10))]
        if self.sort_by in {"title", "title_desc"}:
            rows = sorted(rows, key=lambda row: (title_navigation_key(row[1]), str(row[2] or "").casefold(), row[0]), reverse=self.sort_by == "title_desc")
        if query:
            metadata_ids = {row[0] for row in self.db.search_books(query)}
            needle = query.casefold()
            def relevance(row):
                title = str(row[1] or "").casefold()
                if title == needle or title_navigation_key(title) == title_navigation_key(needle):
                    return 0
                if needle in title:
                    return 1
                if needle in str(row[2] or "").casefold():
                    return 2
                return 3 if row[0] in metadata_ids else 4
            rows.sort(key=relevance)
        mode = self.db.get_setting("book_list_display_mode", "single")
        if mode == "letters":
            self.configure_columns([(letter, letter.lower()) for letter in string.ascii_uppercase] + [("Other", "")])
        elif mode == "tags":
            self.configure_columns(sorted(
                [(tag, tag.casefold()) for tag in COMMON_TAGS] + [("Untagged", "")],
                key=lambda spec: spec[0].casefold()))
        else:
            self.configure_columns(list(BOOK_LIST_ALPHABET_GROUPS))
        self.current_rows = rows
        visible_ids = {row[0] for row in rows}
        self.marked_ids.intersection_update(visible_ids)
        for view in [self.single_list, *self.groups]:
            view.clear()
            view.bulk_loading = True
        id_to_group_row = {}
        multiple_mode = self.views.currentIndex() == 1
        for row_index, row in enumerate(rows):
            text = self.row_text(row)
            if multiple_mode:
                if mode == "tags":
                    tags = tag_categories(row[4])
                    indices = [i for i, (_label, key) in enumerate(self.column_specs) if key in tags or (not tags and not key)]
                elif mode == "letters":
                    initial = title_navigation_key(row[1])[:1]
                    indices = [string.ascii_lowercase.index(initial) if initial and initial in string.ascii_lowercase else 26]
                else:
                    indices = [self.group_index(row[1])]
                for group_index in indices:
                    grouped_item = AccessibleBookItem(text, row[0], row[1])
                    self.groups[group_index].addItem(grouped_item)
                    grouped_item.setSelected(row[0] in self.marked_ids)
                    if row[0] not in id_to_group_row or group_index == self.active_column:
                        id_to_group_row[row[0]] = (group_index, self.groups[group_index].count() - 1)
            else:
                item = AccessibleBookItem(text, row[0], row[1])
                self.single_list.addItem(item)
                item.setSelected(row[0] in self.marked_ids)
        for view in [self.single_list, *self.groups]:
            view.bulk_loading = False
            if view.column_canvas:
                view.column_canvas.updateGeometry()
                view.column_canvas.update_content_height()
                view.column_canvas.update()
        self.apply_column_visibility()
        if rows:
            if not multiple_mode:
                row_index = next((i for i, row in enumerate(rows) if row[0] == selected_id), 0)
                self.single_list.setCurrentRow(row_index)
                if focus_books:
                    self.single_list.setFocus(Qt.FocusReason.OtherFocusReason)
            else:
                group_index, local_index = id_to_group_row.get(
                    selected_id,
                    next(value for value in id_to_group_row.values()),
                )
                self.active_column = group_index
                self.groups[group_index].setCurrentRow(local_index)
                self.groups[group_index].set_accessibility_active(True)
                if focus_books:
                    self.groups[group_index].setFocus(Qt.FocusReason.OtherFocusReason)
            active_list = self.groups[self.active_column] if multiple_mode else self.single_list
            current_item = active_list.currentItem()
            self.schedule_cover_preview_update(current_item.data(Qt.ItemDataRole.UserRole) if current_item else None)
        else:
            self.clear_cover_preview()
            for view in self.groups:
                view.accessibility_active = False
            if self.column_pool:
                shared = self.column_pool[0][1].accessible_row
                shared.item = None
                shared.setAccessibleName("")
                shared.hide()
        self.statusBar().clearMessage()

    def apply_column_visibility(self):
        """Show only categories that contain books.

        An empty category has nothing to read, so it is removed from the window
        entirely. Hidden widgets are not exposed to screen readers, so review
        commands and focus can never land on an empty column or row.
        """
        for group, view in zip(self.group_columns, self.groups):
            group.setVisible(view.count() > 0)

    def nearest_populated_view(self, view):
        """Return the category with books closest to ``view``, or None."""
        populated = [candidate for candidate in self.groups if candidate.count()]
        if not populated:
            return None
        start = view.column if view.column is not None else self.active_column
        return min(populated, key=lambda candidate: (abs(candidate.column - start), candidate.column < start))

    def move_to_column(self, index):
        direction = 1 if index > self.active_column else -1
        while 0 <= index < len(self.groups) and not self.groups[index].count():
            index += direction
        if not 0 <= index < len(self.groups):
            return False
        if index == self.active_column:
            # Already in this category: speak its name once with the book and
            # drop any stale status text so nothing is announced a second time.
            current = self.groups[index]
            self.statusBar().clearMessage()
            if current.currentItem() is None:
                return True
            current.include_column_name = True
            if self.focusWidget() is current.accessible_row:
                current.update_accessible_row()
            else:
                current.setFocus(Qt.FocusReason.OtherFocusReason)
            return True
        target = self.groups[index]
        previous = self.groups[self.active_column]
        self.active_column = index
        target.accessibility_active = False
        row = min(max(target.currentRow(), 0), target.count() - 1)
        # Establish the current book without announcing it until the user enters.
        blocked = target.blockSignals(True)
        try:
            target.setCurrentRow(row)
        finally:
            target.blockSignals(blocked)
        target.include_column_name = True
        target.setFocus(Qt.FocusReason.OtherFocusReason)
        if previous is not target:
            previous.set_accessibility_active(False)
        self.column_scroll.ensureWidgetVisible(target.heading)
        self.statusBar().clearMessage()
        item = target.currentItem()
        if item:
            self.announce_missing_metadata_if_needed(item)
        else:
            self.last_missing_metadata_sound_book_id = None
        self.schedule_cover_preview_update(item.data(Qt.ItemDataRole.UserRole) if item else None)
        return True

    def update_position(self, item, _previous=None):
        if item is not None:
            self.announce_selection(item)
            self.announce_missing_metadata_if_needed(item)
            self.schedule_cover_preview_update(item.data(Qt.ItemDataRole.UserRole))

    def announce_selection(self, item):
        if item is None:
            return
        if item.data(Qt.ItemDataRole.UserRole) in self.marked_ids:
            self.statusBar().showMessage(f"Selected: {item.text()}")

    def book_list_for_focus(self):
        focused = self.focusWidget()
        if self.views.currentIndex() == 0:
            return self.single_list
        if isinstance(focused, AccessibleBookRow):
            if focused.book_list in self.groups:
                return focused.book_list
        for view in self.groups:
            if focused is view or focused is getattr(view, "heading", None) or view.isAncestorOf(focused):
                return view
        return self.single_list if self.views.currentIndex() == 0 else self.groups[self.active_column]

    def current_ids(self):
        if self.marked_ids:
            return [row[0] for row in self.current_rows if row[0] in self.marked_ids]
        view = self.book_list_for_focus()
        selected = [item for item in view.selectedItems()]
        if not selected and view.currentItem():
            selected = [view.currentItem()]
        return list(dict.fromkeys(item.data(Qt.ItemDataRole.UserRole) for item in selected))

    def toggle_marked(self, item):
        book_id = item.data(Qt.ItemDataRole.UserRole)
        if book_id in self.marked_ids:
            self.marked_ids.remove(book_id)
            message = "Deselected"
        else:
            self.marked_ids.add(book_id)
            message = "Selected"
        for view in [self.single_list, *self.groups]:
            for row_index in range(view.count()):
                row_item = view.item(row_index)
                if row_item.data(Qt.ItemDataRole.UserRole) == book_id:
                    row_item.setSelected(book_id in self.marked_ids)
        self.statusBar().showMessage(f"{message}: {item.text()}")
        self.announce_requested_feedback(message)

    def select_all(self):
        self.marked_ids = {row[0] for row in self.current_rows}
        for book_view in [self.single_list, *self.groups]:
            for row_index in range(book_view.count()):
                item = book_view.item(row_index)
                item.setSelected(item.data(Qt.ItemDataRole.UserRole) in self.marked_ids)
        self.statusBar().showMessage(f"Selected {len(self.marked_ids)} books")

    def deselect_all(self):
        self.marked_ids.clear()
        for view in [self.single_list, *self.groups]:
            view.clearSelection()
        self.statusBar().showMessage("All books deselected")

    def focus_search(self):
        self.search.setFocus(Qt.FocusReason.ShortcutFocusReason)
        self.search.selectAll()

    def clear_search(self):
        self.search.clear()
        self.filter_source = self.filter_tag = self.filter_format = ""
        self.clear_accessibility_filters()
        self.active_saved_view = ""
        self.refresh_books()

    def search_and_focus(self):
        if not self.search.text().strip():
            self.clear_search()
            return
        self.filter_source = self.filter_tag = self.filter_format = ""
        self.clear_accessibility_filters()
        self.active_saved_view = ""
        self.refresh_books()
        if not self.current_rows:
            self.search.setFocus()
            self.routine_feedback("No books matched the search. Press Escape to clear the search; use Organize to clear any filters.")
        else:
            self.routine_feedback(
                f"{len(self.current_rows)} books found for {self.search.text().strip()}.",
                f"{len(self.current_rows)} search results for {self.search.text().strip()}. Matching titles first, then other metadata and indexed text.")

    def focus_book_list(self):
        self.setFocusToCurrentBook()

    def selected_row(self):
        view = self.book_list_for_focus()
        item = view.currentItem()
        if item is not None:
            return self.db.get_book(item.data(Qt.ItemDataRole.UserRole))
        ids = self.current_ids()
        return self.db.get_book(ids[0]) if ids else None

    def read_or_edit_field(self, field, digit):
        row = self.selected_row()
        if not row:
            return
        values = {
            "title": ("Title", row[1]), "author": ("Author", row[2]),
            "series": ("Series", row[21] if len(row) > 21 else ""),
            "series_number": ("Series number", row[22] if len(row) > 22 else ""),
            "edition": ("Edition", row[10]), "year": ("Year", row[11]),
            "isbn": ("ISBN", row[12]), "publisher": ("Publisher", row[13]),
            "source": ("Source", row[3]), "tags": ("Tags", row[4]),
        }
        label, value = values[field]
        now = time.monotonic()
        repeated = (digit == getattr(self, "last_alt_number_key", None)
                    and row[0] == getattr(self, "last_alt_number_book_id", None)
                    and now - getattr(self, "last_alt_number_time", 0) <= 0.7)
        self.last_alt_number_key = digit
        self.last_alt_number_time = now
        self.last_alt_number_book_id = row[0]
        if repeated:
            self.last_alt_number_key = None
            self.edit_book(initial_focus_field=field)
            return
        message = f"{label}: {value or 'Not specified'}"
        status_bar = self.statusBar()
        status_bar.showMessage(message)
        status_bar.announcement_timer.stop()
        self.announce_requested_feedback(message)

    def announce_requested_feedback(self, message):
        # A notification carries its own text; an Alert on a status bar can be
        # ignored because it is not a native alert control. Never move focus.
        self.statusBar().announcement_timer.stop()
        event = QAccessibleAnnouncementEvent(self, message)
        event.setPoliteness(QAccessible.AnnouncementPoliteness.Assertive)
        QAccessible.updateAccessibility(event)

    def routine_feedback(self, message, details=None, speak=True):
        self.statusBar().showMessage(details or message)
        self.statusBar().announcement_timer.stop()
        if speak and not self._closing and self.isActiveWindow() and QApplication.activeModalWidget() is None:
            event = QAccessibleAnnouncementEvent(self, message)
            event.setPoliteness(QAccessible.AnnouncementPoliteness.Polite)
            QAccessible.updateAccessibility(event)

    def open_book(self):
        row = self.selected_row()
        if not row:
            return
        if is_wishlist_path(row[8]):
            QMessageBox.information(self, "No File Available", "This book is in your library without a book file yet. Add the file to read it.")
            return
        path = Path(row[8])
        if not path.exists():
            QMessageBox.warning(self, "Book file missing", f"The stored book file could not be found:\n{path}\n\nIf it is on a drive or network location, check that it is connected.")
            return
        if path.is_dir():
            first = first_audio_file(path)
            if first is None:
                QMessageBox.warning(self, "Audiobook folder empty", f"No audio files were found in:\n{path}")
                return
            path = first
        notes = describe_flags(self.db.get_content_flags(row[0]))
        if notes:
            # Say it before the reader takes focus, so a silent scanned PDF is no surprise.
            self.routine_feedback(notes[0])
        reader = self.db.get_setting("default_reader_path", "")
        try:
            if reader and Path(reader).exists():
                subprocess.Popen([reader, str(path)])
            else:
                QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))
        except OSError as exc:
            QMessageBox.critical(self, "Open failed", str(exc))

    def open_kindle(self):
        if not sys.platform.startswith("win"):
            QMessageBox.information(self, "Windows only", "Kindle for PC launching is only supported on Windows.")
            return

        possible_paths = [
            Path(os.environ.get("LOCALAPPDATA", "")) / "Amazon" / "Kindle" / "application" / "Kindle.exe",
            Path(os.environ.get("PROGRAMFILES", "")) / "Amazon" / "Kindle" / "Kindle.exe",
            Path(os.environ.get("PROGRAMFILES(X86)", "")) / "Amazon" / "Kindle" / "Kindle.exe",
        ]
        for kindle_path in possible_paths:
            if kindle_path.is_file():
                try:
                    subprocess.Popen([str(kindle_path)], shell=False)
                    self.statusBar().showMessage("Opened Kindle for PC.")
                except OSError as exc:
                    QMessageBox.critical(self, "Open Kindle failed", f"Could not open Kindle for PC.\n\n{exc}")
                return

        try:
            subprocess.Popen("start kindle:", shell=True)
            self.statusBar().showMessage("Tried to open Kindle for PC.")
        except OSError:
            QMessageBox.warning(self, "Kindle not found", "I could not find Kindle for PC. Install it, then try again.")

    def edit_book(self, initial_focus_field="title"):
        row = self.selected_row()
        if not row:
            return
        values = self.row_metadata(row)
        dialog = MetadataDialog(values, self, initial_focus_field)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        values = dialog.values()
        if not values["title"]:
            QMessageBox.warning(self, "Title required", "Enter a title before saving.")
            return
        book_id = row[0]
        self.db.update_book(book_id, values["title"], values["author"], values["source"], values["tags"], values["notes"])
        self.db.update_extra_fields(book_id, values["edition"], values["year"], values["isbn"], values["publisher"])
        self.db.update_series_fields(book_id, values["series"], values["series_number"])
        stored = Path(row[8])
        if stored.suffix.casefold() == ".epub" and stored.exists():
            try:
                write_epub_metadata(stored, values["title"], values["author"], values["source"], values["tags"], values["notes"])
            except Exception as exc:
                QMessageBox.warning(self, "EPUB metadata not written", f"Library metadata was saved, but the EPUB file could not be updated.\n\n{exc}")
        self.refresh_books(book_id)
        self.setFocusToCurrentBook()

    def bulk_edit_metadata(self):
        book_ids = self.current_ids()
        if not book_ids:
            return
        dialog = BulkMetadataDialog(self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        updates = dialog.values()
        if not updates:
            self.routine_feedback("No changes. No fields were entered.")
            return
        fields = ", ".join(f"{key.replace('_', ' ').title()}: {value}" for key, value in updates.items())
        answer = QMessageBox.question(
            self, "Apply Bulk Metadata",
            f"Apply these values to {len(book_ids)} selected book(s)? Blank fields will remain unchanged.\n\n{fields}",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        changed = self.db.update_bulk_metadata_fields(book_ids, updates)
        current = self.selected_row()
        self.refresh_books(current[0] if current else book_ids[0])
        self.statusBar().showMessage(f"Bulk metadata updated for {changed} book(s)")
        self.setFocusToCurrentBook()

    def auto_detect_metadata(self):
        row = self.selected_row()
        if not row:
            return
        path = Path(row[8])
        if not path.is_file():
            QMessageBox.warning(self, "Book File Missing", f"The stored book file could not be found:\n{path}")
            return
        from metadata_wizard import MetadataWizard
        wizard = MetadataWizard(path, self.row_metadata(row), self)
        def review_metadata(active_wizard):
            dialog = MetadataDialog(active_wizard.metadata, active_wizard)
            dialog.setWindowTitle("Review Detected Metadata")
            if dialog.exec() != QDialog.DialogCode.Accepted:
                return False
            active_wizard.metadata = dialog.values()
            return True
        wizard.review_callback = review_metadata
        if wizard.exec() == QDialog.DialogCode.Accepted:
            self.pending_detection_row = row
            self.metadata_detection_finished(wizard.metadata, "", reviewed=True)

    def metadata_detection_finished(self, detected, error, reviewed=False):
        if self.progress_dialog:
            self.progress_dialog.close()
            self.progress_dialog = None
        row = self.pending_detection_row
        self.pending_detection_row = None
        if error:
            QMessageBox.warning(self, "Metadata Detection Failed", error)
            return
        if not row:
            return
        if reviewed:
            values = detected
        else:
            dialog = MetadataDialog(detected, self)
            dialog.setWindowTitle("Review Detected Metadata")
            if dialog.exec() != QDialog.DialogCode.Accepted:
                return
            values = dialog.values()
        if not values["title"]:
            QMessageBox.warning(self, "Title Required", "Enter a title before saving.")
            return
        book_id = row[0]
        self.db.update_book(book_id, values["title"], values["author"], row[3], values["tags"], values["notes"])
        self.db.update_extra_fields(book_id, values["edition"], values["year"], values["isbn"], values["publisher"])
        self.db.update_series_fields(book_id, values["series"], values["series_number"])
        stored = Path(row[8])
        if stored.suffix.casefold() == ".epub" and stored.is_file():
            try:
                write_epub_metadata(stored, values["title"], values["author"], row[3], values["tags"], values["notes"])
            except Exception as exc:
                QMessageBox.warning(self, "EPUB Metadata Not Written", f"Library metadata was saved, but the EPUB could not be changed.\n\n{exc}")
        self.refresh_books(book_id)
        self.setFocusToCurrentBook()

    def view_cover_image(self):
        row = self.selected_row()
        if not row:
            return
        cover_url = row[14] if len(row) > 14 else ""
        if cover_url:
            QDesktopServices.openUrl(QUrl(cover_url))
            return
        self.find_cover_image()

    def find_cover_image(self):
        row = self.selected_row()
        if not row:
            return
        if row[14]:
            self.routine_feedback("A cover is already saved. Choose View Cover to open it.")
            return
        if self.cover_lookup_worker and self.cover_lookup_worker.isRunning():
            self.routine_feedback("A cover lookup is already running.")
            return
        self.pending_cover_lookup_book_id = row[0]
        self.progress_dialog = QProgressDialog("Looking for a cover image...", "", 0, 0, self)
        self.progress_dialog.setWindowTitle("Find Cover Image")
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.setCancelButton(None)
        self.progress_dialog.show()
        worker = OnlineMetadataWorker(self.row_metadata(row), self)
        self.cover_lookup_worker = worker
        self.connect_worker_result(worker, self.cover_lookup_finished)
        worker.start()

    def cover_lookup_finished(self, online, service, error):
        if self.progress_dialog:
            self.progress_dialog.close()
            self.progress_dialog = None
        book_id = self.pending_cover_lookup_book_id
        self.pending_cover_lookup_book_id = None
        if error:
            QMessageBox.warning(self, "Cover Lookup Failed", error)
            return
        cover_url = str(online.get("cover_url", "") or "").strip()
        if not cover_url:
            self.routine_feedback("No cover image found.")
            return
        if not book_id or not self.db.get_book(book_id):
            return
        self.db.update_cover_url(book_id, cover_url)
        current = self.selected_row()
        if current and current[0] == book_id:
            self.schedule_cover_preview_update(book_id)
        self.routine_feedback("Cover saved.", f"Cover image saved from {service}.")

    def schedule_cover_preview_update(self, book_id):
        self.pending_cover_preview_book_id = book_id
        self.cover_preview_timer.start(160)

    def load_cover_preview(self):
        if self._closing or self.db is None:
            return
        book_id = self.pending_cover_preview_book_id
        self.cover_preview_book_id = book_id
        row = self.db.get_book(book_id) if book_id else None
        if not row:
            self.clear_cover_preview()
            return
        self.cover_title.setText(row[1] or "")
        url = row[14] or ""
        if not url:
            self.cover_image.setPixmap(QPixmap())
            self.cover_image.setText("No cover saved")
            return
        cached = self.cover_cache.get(url)
        if cached is not None:
            self.cover_cache.pop(url)
            self.cover_cache[url] = cached
            self.show_cover_pixmap(cached)
            return
        self.cover_image.setPixmap(QPixmap())
        self.cover_image.setText("Loading cover...")
        worker = CoverImageWorker(book_id, url, self)
        self.cover_preview_workers.append(worker)
        self.connect_worker_result(worker, self.cover_preview_loaded)
        worker.finished.connect(lambda worker=worker: self.release_cover_worker(worker))
        worker.start()

    def show_cover_pixmap(self, pixmap):
        if pixmap.width() > self.cover_image.width() or pixmap.height() > self.cover_image.height():
            pixmap = pixmap.scaled(
                self.cover_image.size(), Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation,
            )
        self.cover_image.setText("")
        self.cover_image.setPixmap(pixmap)

    def cover_preview_loaded(self, book_id, url, data, error):
        if self._closing:
            return
        if book_id != self.cover_preview_book_id:
            return
        pixmap = QPixmap()
        if error or not data or not pixmap.loadFromData(data):
            self.cover_image.setPixmap(QPixmap())
            self.cover_image.setText("Cover unavailable")
            return
        thumbnail = pixmap.scaled(
            self.cover_image.size(), Qt.AspectRatioMode.KeepAspectRatio,
            Qt.TransformationMode.SmoothTransformation,
        )
        self.cover_cache[url] = thumbnail
        while len(self.cover_cache) > 64:
            self.cover_cache.pop(next(iter(self.cover_cache)))
        self.show_cover_pixmap(thumbnail)

    def release_cover_worker(self, worker):
        if worker in self.cover_preview_workers:
            self.cover_preview_workers.remove(worker)

    def clear_cover_preview(self):
        self.cover_preview_timer.stop()
        self.pending_cover_preview_book_id = None
        self.cover_preview_book_id = None
        self.cover_title.clear()
        self.cover_image.setPixmap(QPixmap())
        self.cover_image.setText("No cover saved")

    def lookup_metadata_online(self):
        row = self.selected_row()
        if not row:
            return
        self.pending_lookup_row = row
        self.progress_dialog = QProgressDialog("Searching online metadata sources...", "", 0, 0, self)
        self.progress_dialog.setWindowTitle("Metadata Lookup")
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.setCancelButton(None)
        self.progress_dialog.show()
        worker = OnlineMetadataWorker(self.row_metadata(row), self)
        self.metadata_worker = worker
        self.connect_worker_result(worker, self.metadata_lookup_finished)
        worker.start()

    def metadata_lookup_finished(self, online, service, error):
        if self.progress_dialog:
            self.progress_dialog.close()
            self.progress_dialog = None
        row = self.pending_lookup_row
        self.pending_lookup_row = None
        if error:
            QMessageBox.warning(self, "Metadata Lookup Failed", error)
            return
        if not online:
            self.routine_feedback("No online metadata match found.")
            return
        if not row:
            return
        metadata = self.row_metadata(row)
        if online.get("cover_url"):
            metadata["cover_url"] = online["cover_url"]
        dialog = OnlineMetadataDialog(metadata, online, service, self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        values = dialog.values()
        if not values["title"]:
            QMessageBox.warning(self, "Title Required", "Enter a title before saving.")
            return
        book_id = row[0]
        self.db.update_book(book_id, values["title"], values["author"], row[3], values["tags"], values["notes"])
        self.db.update_extra_fields(book_id, values["edition"], values["year"], values["isbn"], values["publisher"])
        self.db.update_series_fields(book_id, values["series"], values["series_number"])
        self.db.update_cover_url(book_id, metadata.get("cover_url", ""))
        stored = Path(row[8])
        if stored.suffix.casefold() == ".epub" and stored.exists():
            try:
                write_epub_metadata(stored, values["title"], values["author"], row[3], values["tags"], values["notes"])
            except Exception as exc:
                QMessageBox.warning(self, "EPUB Metadata Not Written", f"The database was updated, but the EPUB could not be changed.\n\n{exc}")
        self.refresh_books(book_id)
        self.setFocusToCurrentBook()

    @staticmethod
    def row_metadata(row):
        return {
            "title": row[1], "author": row[2], "source": row[3], "tags": row[4], "notes": row[5],
            "edition": row[10], "year": row[11], "isbn": row[12], "publisher": row[13],
            "cover_url": row[14], "series": row[21], "series_number": row[22],
        }

    def setFocusToCurrentBook(self):
        view = self.single_list if self.views.currentIndex() == 0 else self.groups[self.active_column]
        item = view.currentItem()
        if item and item.row_widget and item.row_widget.hasFocus():
            return
        view.setFocus(Qt.FocusReason.OtherFocusReason)

    def show_details(self):
        row = self.selected_row()
        if not row:
            return
        accessibility = self.accessibility_text_from_row(row)
        details = [
            f"Title: {row[1] or 'Not specified'}",
            f"Author: {row[2] or 'Unknown'}",
            f"Series: {row[21] or 'Not specified'}",
            f"Series number: {row[22] or 'Not specified'}",
            f"Edition: {row[10] or 'Not specified'}",
            f"Year: {row[11] or 'Not specified'}",
            f"ISBN: {row[12] or 'Not specified'}",
            f"Publisher: {row[13] or 'Not specified'}",
            f"Cover image: {'Available' if row[14] else 'Not saved'}",
            f"EPUB accessibility:\n{accessibility}",
            f"Source: {row[3] or 'Not specified'}",
            f"Tags: {row[4] or 'None'}",
            f"Format: {row[6] or 'Unknown'}",
            f"Stored path: {row[8]}",
        ]
        library_id = self.db.get_source_book_id(row[0])
        if library_id:
            details.insert(-1, f"Library book ID: {library_id}")
        content_notes = describe_flags(self.db.get_content_flags(row[0]))
        if content_notes:
            details.insert(-1, "Content notes: " + " ".join(content_notes))
        QMessageBox.information(self, "Book Details", "\n".join(details))

    @staticmethod
    def accessibility_text_from_row(row):
        if not row or len(row) <= 20:
            return "Not checked"
        labels = (
            ("Accessibility summary", row[15]),
            ("Accessibility features", row[16]),
            ("Accessibility hazards", row[17]),
            ("Access modes", row[18]),
            ("Access modes sufficient", row[19]),
            ("Certified by", row[20]),
        )
        parts = [f"{label}: {value}" for label, value in labels if value]
        return "\n".join(parts) if parts else "Not checked"

    def run_epub_repair(self, operation):
        row = self.selected_row()
        if not row:
            return
        path = Path(row[8])
        if path.suffix.casefold() != ".epub":
            QMessageBox.information(self, "EPUB Only", "This repair is available for EPUB books.")
            return
        if not path.is_file():
            QMessageBox.warning(self, "Book File Missing", f"The stored EPUB could not be found:\n{path}")
            return
        messages = {
            "pages": "Find written page labels and add EPUB page-break markers? A backup copy will be kept.",
            "toc": "Check for written page labels first, then rebuild EPUB navigation from a text table of contents? A backup copy will be kept.",
            "clean": "Remove repeated short headers/footers and blank-page notices, detect language, and add page breaks? A backup copy will be kept.",
        }
        answer = QMessageBox.question(
            self, "Repair EPUB", messages[operation],
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        self.pending_repair_book_id = row[0]
        self.progress_dialog = QProgressDialog("Repairing EPUB...", "", 0, 0, self)
        self.progress_dialog.setWindowTitle("EPUB Repair")
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.setCancelButton(None)
        self.progress_dialog.show()
        worker = EpubRepairWorker(operation, path, self)
        self.active_worker = worker
        self.connect_worker_result(worker, self.epub_repair_finished)
        worker.start()

    def check_epub_accessibility(self):
        row = self.selected_row()
        if not row:
            return
        path = Path(row[8])
        if path.suffix.casefold() != ".epub" or not path.is_file():
            QMessageBox.information(self, "EPUB Only", "Accessibility metadata checking is available for EPUB books with an available file.")
            return
        self.start_cleanup_wizard(row, {"accessibility": True})

    def cleanup_repair_wizard(self):
        row = self.selected_row()
        if not row:
            return
        path = Path(row[8])
        if path.suffix.casefold() != ".epub" or not path.is_file():
            QMessageBox.information(self, "EPUB Only", "The cleanup and repair wizard is available for EPUB books with an available file.")
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("Cleanup and Repair Wizard")
        dialog.resize(540, 360)
        layout = QVBoxLayout(dialog)
        prompt = QLabel(f"Choose the checks and repairs to run for {row[1]}. File repairs keep a backup copy.", dialog)
        prompt.setWordWrap(True)
        layout.addWidget(prompt)
        choices = {
            "accessibility": QCheckBox("Check EPUB accessibility metadata", dialog),
            "clean": QCheckBox("Clean repeated text, detect language, and add page breaks", dialog),
            "pages": QCheckBox("Add page breaks from written page labels", dialog),
            "toc": QCheckBox("Rebuild navigation from the text table of contents (checks page labels first)", dialog),
        }
        for checkbox in choices.values():
            checkbox.setChecked(True)
            layout.addWidget(checkbox)
        choices["pages"].setDisabled(True)
        choices["clean"].toggled.connect(lambda enabled: choices["pages"].setDisabled(enabled))
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel, parent=dialog)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        selected = {key: checkbox.isChecked() for key, checkbox in choices.items()}
        if not any(selected.values()):
            self.routine_feedback("No steps selected. No repairs were run.")
            return
        self.start_cleanup_wizard(row, selected)

    def start_cleanup_wizard(self, row, choices):
        self.pending_repair_book_id = row[0]
        self.progress_dialog = QProgressDialog("Checking and repairing the EPUB...", "", 0, 0, self)
        self.progress_dialog.setWindowTitle("Cleanup and Repair")
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.setCancelButton(None)
        self.progress_dialog.show()
        worker = CleanupWizardWorker(row[0], Path(row[8]), choices, self)
        self.active_worker = worker
        self.connect_worker_result(worker, self.cleanup_wizard_finished)
        worker.start()

    def cleanup_wizard_finished(self, results, error):
        if self.progress_dialog:
            self.progress_dialog.close()
            self.progress_dialog = None
        book_id = self.pending_repair_book_id
        self.pending_repair_book_id = None
        if error:
            QMessageBox.critical(self, "Cleanup and Repair Failed", error)
            return
        row = self.db.get_book(book_id) if book_id else None
        if row:
            try:
                self.db.update_accessibility_metadata(book_id, read_epub_accessibility_metadata(Path(row[8])))
            except Exception:
                pass
            self.refresh_books(book_id)
            self.setFocusToCurrentBook()
        message = "\n".join(results) if results else "No wizard steps were selected."
        self.statusBar().showMessage("Cleanup and repair complete.")
        QMessageBox.information(self, "Cleanup and Repair Complete", message)

    def epub_repair_finished(self, operation, result, error):
        if self.progress_dialog:
            self.progress_dialog.close()
            self.progress_dialog = None
        book_id = self.pending_repair_book_id
        self.pending_repair_book_id = None
        if error:
            QMessageBox.critical(self, "EPUB Repair Failed", error)
            return
        row = self.db.get_book(book_id) if book_id else None
        if row:
            try:
                self.db.update_accessibility_metadata(book_id, read_epub_accessibility_metadata(Path(row[8])))
            except Exception:
                pass
            self.refresh_books(book_id)
            self.setFocusToCurrentBook()
        if operation == "pages":
            message = f"Added {result} page-break markers." if result else "No written page labels were found."
        elif operation == "toc":
            pages, added, skipped = result
            message = f"Added {pages} page-break markers before rebuilding navigation. Built {added} table-of-contents entries; skipped {skipped} without matching page anchors."
        else:
            message = f"EPUB cleanup made {result} change(s)."
        self.statusBar().showMessage(message)
        QMessageBox.information(self, "EPUB Repair Complete", message)

    def remove_repair_backups(self):
        suffixes = (".bak", ".before_pagebreaks.bak", ".before_toc_rebuild.bak", ".before_text_cleanup.bak")
        backups = [path for path in self.db.books_path.rglob("*") if path.is_file() and path.name.casefold().endswith(suffixes)]
        if not backups:
            self.routine_feedback("No repair backup files found.")
            return
        total = sum(path.stat().st_size for path in backups if path.exists())
        answer = QMessageBox.question(
            self, "Remove Repair Backup Files",
            f"Delete {len(backups)} repair backup file(s) from the managed Books folder?\n\nTotal size: {total:,} bytes.\nOnly files with known repair-backup suffixes will be removed.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        removed = 0
        errors = []
        for path in backups:
            try:
                path.unlink()
                removed += 1
            except OSError as exc:
                errors.append(f"{path}: {exc}")
        self.statusBar().showMessage(f"Removed {removed} repair backup file(s)")
        if errors:
            QMessageBox.warning(self, "Backup Cleanup", f"Removed {removed}; could not remove {len(errors)}:\n\n" + "\n".join(errors[:10]))

    def delete_books(self):
        ids = self.current_ids()
        if not ids:
            return
        if RemoveBooksDialog(len(ids), self).exec() != QDialog.DialogCode.Accepted:
            return
        for book_id in ids:
            self.db.delete_book(book_id, delete_file=False)
        self.refresh_books()

    def remove_duplicates(self):
        remove_ids = duplicate_removal_ids(self.db)
        if not remove_ids:
            self.routine_feedback("No likely duplicate books found.")
            return
        answer = QMessageBox.question(
            self, "Remove Duplicate Books",
            f"Found {len(remove_ids)} duplicate entries. Remove them from the library? Files will be kept.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        delete_files = QMessageBox.question(
            self, "Delete Duplicate Files",
            "Also delete the duplicate book files from disk? Choose No to keep the files and only remove their library entries.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        ) == QMessageBox.StandardButton.Yes
        for book_id in remove_ids:
            self.db.delete_book(book_id, delete_file=delete_files)
        self.refresh_books()
        self.statusBar().showMessage(f"Removed {len(remove_ids)} duplicate entries")

    def toggle_automatic_duplicate_cleanup(self):
        enabled = self.db.get_setting("automatic_duplicate_cleanup", "0") == "1"
        if not enabled:
            answer = QMessageBox.question(
                self, "Enable Automatic Duplicate Cleanup",
                "The manager will periodically remove likely duplicate entries from the library. Book files are kept. Enable this?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return
        self.db.set_setting("automatic_duplicate_cleanup", "0" if enabled else "1")
        state = "off" if enabled else "on"
        self.statusBar().showMessage(f"Automatic duplicate cleanup is {state}")
        if not enabled:
            QTimer.singleShot(1000, self.automatic_duplicate_cleanup)

    def automatic_duplicate_cleanup(self):
        if self._closing or self.db is None:
            return
        if self.db.get_setting("automatic_duplicate_cleanup", "0") != "1":
            return
        if self.duplicate_worker and self.duplicate_worker.isRunning():
            QTimer.singleShot(60 * 1000, self.automatic_duplicate_cleanup)
            return
        worker = DuplicateCleanupWorker(self)
        self.duplicate_worker = worker
        self.connect_worker_result(worker, self.automatic_duplicate_cleanup_finished)
        worker.start()

    def automatic_duplicate_cleanup_finished(self, removed, error):
        if self._closing or self.db is None:
            return
        if error:
            self.statusBar().showMessage("Automatic duplicate cleanup failed")
        elif removed:
            self.refresh_books()
            self.statusBar().showMessage(f"Automatic cleanup removed {removed} duplicate entries; files were kept")
        if self.db.get_setting("automatic_duplicate_cleanup", "0") == "1":
            QTimer.singleShot(30 * 60 * 1000, self.automatic_duplicate_cleanup)

    def add_book(self):
        if getattr(self, "pending_add_paths", None):
            return
        paths, _ = QFileDialog.getOpenFileNames(self, "Add Books", "", "Books (*.epub *.pdf *.docx *.doc *.txt *.rtf *.mobi *.azw *.azw3 *.kfx *.prc *.html *.htm);;All files (*.*)")
        if not paths:
            return
        self.pending_add_paths = list(paths)
        self.added_count = 0
        self.add_next_book()

    def add_next_book(self):
        if self._closing:
            return
        if not self.pending_add_paths:
            self.statusBar().showMessage(f"Added {self.added_count} books")
            return
        source = Path(self.pending_add_paths.pop(0))
        if not source.is_file() or source.stat().st_size == 0:
            QMessageBox.warning(self, "Book File Unavailable", f"The file is missing or empty:\n{source}")
            QTimer.singleShot(0, self.add_next_book)
            return
        metadata = {"title": clean_filename_title(source)}
        self.progress_dialog = QProgressDialog("Reading book metadata...", "", 0, 0, self)
        self.progress_dialog.setWindowTitle("Add Book")
        self.progress_dialog.setCancelButton(None)
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.show()
        worker = MetadataDetectionWorker(source, metadata, self, preserve_source=False)
        self.active_worker = worker
        self.connect_worker_result(worker, lambda found, error: self.add_book_ready(source, found, error))
        worker.start()

    def add_book_ready(self, source, metadata, error):
        if self.progress_dialog:
            self.progress_dialog.close()
            self.progress_dialog = None
        if error:
            QMessageBox.warning(self, "Import Failed", error)
            QTimer.singleShot(0, self.add_next_book)
            return
        dialog = MetadataDialog(metadata, self)
        dialog.setWindowTitle(f"Add Book: {source.name}")
        if dialog.exec() != QDialog.DialogCode.Accepted:
            QTimer.singleShot(0, self.add_next_book)
            return
        try:
            values = dialog.values()
            base = safe_filename(f"{values['author']} - {values['title']}").strip(" -") or source.stem
            destination = self.db.books_path / f"{base}{source.suffix.lower()}"
            suffix = 2
            while destination.exists():
                destination = self.db.books_path / f"{base} ({suffix}){source.suffix.lower()}"
                suffix += 1
            shutil.copy2(source, destination)
            if destination.suffix.casefold() == ".epub":
                try:
                    write_epub_metadata(destination, values["title"], values["author"], values["source"], values["tags"], values["notes"])
                except Exception:
                    pass
            book_id = self.db.add_book(values["title"], values["author"], values["source"], values["tags"], values["notes"], str(source.resolve()), str(destination))
            self.db.update_extra_fields(book_id, values["edition"], values["year"], values["isbn"], values["publisher"])
            self.db.update_series_fields(book_id, values["series"], values["series_number"])
            self.added_count += 1
            self.search.clear()
            self.filter_source = self.filter_tag = self.filter_format = ""
            self.clear_accessibility_filters()
            self.active_saved_view = ""
            self.refresh_books(book_id)
        except (OSError, ValueError) as exc:
            QMessageBox.warning(self, "Import Failed", str(exc))
        QTimer.singleShot(0, self.add_next_book)

    def import_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Choose Folder of Books")
        if not folder:
            return
        options = FolderImportOptionsDialog(self)
        if options.exec() != QDialog.DialogCode.Accepted:
            return
        source = options.source.text().strip()
        self.progress_dialog = QProgressDialog("Searching the selected folder...", "", 0, 0, self)
        self.progress_dialog.setWindowTitle("Import Folder")
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.setCancelButton(None)
        self.progress_dialog.show()
        worker = FolderScanWorker(
            folder,
            options.include.currentIndex() == 0,
            self.db.books_path,
            self,
        )
        self.active_worker = worker
        self.connect_worker_result(worker, lambda paths, skipped, folder=Path(folder), source=source: self.folder_scan_finished(folder, source, paths, skipped))
        worker.start()

    def folder_scan_finished(self, folder, source, paths, skipped):
        if self.progress_dialog:
            self.progress_dialog.close()
            self.progress_dialog = None
        if not paths:
            report = self.db.folder / "last_import_report.txt"
            report.write_text("No supported books found.\n\n" + "\n".join(skipped), encoding="utf-8")
            if skipped:
                QMessageBox.warning(self, "No Books Found", f"No supported book files were found.\n\nImport report: {report}")
            else:
                self.routine_feedback("No supported book files found.", f"No supported book files found. Import report: {report}")
            return
        if len(paths) > 1:
            answer = QMessageBox(self)
            answer.setWindowTitle("One Book in Multiple Parts?")
            answer.setText(f"Found {len(paths)} supported files. Do these parts belong to one book?")
            answer.setInformativeText("Yes combines them into one EPUB. No imports each file separately.")
            yes = answer.addButton("Combine into One EPUB", QMessageBox.ButtonRole.YesRole)
            no = answer.addButton("Import Separately", QMessageBox.ButtonRole.NoRole)
            answer.addButton(QMessageBox.StandardButton.Cancel)
            answer.setDefaultButton(no)
            answer.exec()
            clicked = answer.clickedButton()
            if clicked is yes:
                self.start_combined_import(folder, sorted(paths, key=book_part_sort_key), source, skipped)
                return
            if clicked is not no:
                return
        answer = QMessageBox.question(
            self, "Confirm Folder Import",
            f"Import {len(paths)} book file(s) from this folder?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        self.progress_dialog = QProgressDialog("Preparing folder import...", "", 0, 0, self)
        self.progress_dialog.setWindowTitle("Importing Books")
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.setCancelButton(None)
        self.progress_dialog.show()
        worker = FolderImportWorker(paths, source, self)
        self.active_worker = worker
        worker.progress.connect(self.progress_dialog.setLabelText)
        self.connect_worker_result(worker, self.folder_import_finished)
        worker.start()

    def start_combined_import(self, folder, paths, source, skipped):
        self.pending_combined_source = source
        self.pending_combined_skipped = list(skipped)
        self.progress_dialog = QProgressDialog("Reading book parts...", "", 0, 0, self)
        self.progress_dialog.setWindowTitle("Combine Book Parts")
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.setCancelButton(None)
        self.progress_dialog.show()
        worker = CombinedPartsWorker(folder, paths, self)
        self.active_worker = worker
        worker.progress.connect(self.progress_dialog.setLabelText)
        self.connect_worker_result(worker, lambda parts, sample, errors, folder=folder: self.combined_parts_ready(folder, parts, sample, errors))
        worker.start()

    def combined_parts_ready(self, folder, parts, sample, errors):
        if self.progress_dialog:
            self.progress_dialog.close()
            self.progress_dialog = None
        skipped = [line for line in errors.splitlines() if line]
        if not parts:
            QMessageBox.warning(self, "Cannot Combine Book", "No readable text was found in those parts.\n\n" + "\n".join(skipped[:12]))
            return
        initial = {
            "title": clean_filename_title(folder), "author": "", "series": "", "series_number": "",
            "edition": "", "year": "", "isbn": "", "publisher": "",
            "source": self.pending_combined_source, "tags": "", "notes": "",
        }
        initial.update({key: value for key, value in detect_metadata_from_text_content("\n".join(sample), existing=initial).items() if value})
        dialog = MetadataDialog(initial, self)
        dialog.setWindowTitle("Combined Book Metadata")
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        metadata = dialog.values()
        if not metadata["title"]:
            QMessageBox.warning(self, "Title Required", "Enter a title before creating the EPUB.")
            return
        self.progress_dialog = QProgressDialog("Creating the combined EPUB...", "", 0, 0, self)
        self.progress_dialog.setWindowTitle("Combining Book Parts")
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.setCancelButton(None)
        self.progress_dialog.show()
        worker = CombinedSaveWorker(folder, parts, metadata, self)
        self.active_worker = worker
        self.connect_worker_result(worker, lambda book_id, error, skipped=skipped: self.combined_import_finished(book_id, error, skipped))
        worker.start()

    def combined_import_finished(self, book_id, error, skipped):
        if self.progress_dialog:
            self.progress_dialog.close()
            self.progress_dialog = None
        if error:
            QMessageBox.critical(self, "Combined EPUB Import Failed", error)
            return
        self.refresh_books(book_id)
        self.statusBar().showMessage("Combined EPUB imported")
        if skipped:
            QMessageBox.information(self, "Combined EPUB Imported", f"The combined book was imported. {len(skipped)} item(s) were skipped:\n\n" + "\n".join(skipped[:12]))

    def combine_selected_books(self):
        ids = self.current_ids()
        if len(ids) < 2:
            QMessageBox.information(self, "Select Multiple Books", "Mark at least two book parts first with Ctrl+Space, then choose this command.")
            return
        rows = [self.db.get_book(book_id) for book_id in ids]
        rows = [row for row in rows if row]
        answer = QMessageBox.question(
            self, "Combine Selected Books",
            f"Combine {len(rows)} selected books into one EPUB in the current list order? The original library entries and files will be kept.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        paths = [Path(row[8]) for row in rows if Path(row[8]).is_file() and Path(row[8]).suffix.casefold() != ".zip"]
        if len(paths) < 2:
            QMessageBox.warning(self, "Could Not Combine Books", "At least two selected books must have readable, available files.")
            return
        self.pending_combined_rows = rows
        self.pending_combined_source = next(iter({row[3].strip() for row in rows if row[3].strip()}), "")
        root = paths[0].parent
        self.progress_dialog = QProgressDialog("Reading selected book parts...", "", 0, 0, self)
        self.progress_dialog.setWindowTitle("Combine Selected Books")
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.setCancelButton(None)
        self.progress_dialog.show()
        worker = CombinedPartsWorker(root, paths, self)
        self.active_worker = worker
        worker.progress.connect(self.progress_dialog.setLabelText)
        self.connect_worker_result(worker, lambda parts, sample, errors, root=root: self.combined_selected_parts_ready(root, parts, sample, errors))
        worker.start()

    def combined_selected_parts_ready(self, root, parts, sample, errors):
        if self.progress_dialog:
            self.progress_dialog.close()
            self.progress_dialog = None
        skipped = [line for line in errors.splitlines() if line]
        if len(parts) < 2:
            QMessageBox.warning(self, "Could Not Combine Books", "Readable text was found in fewer than two selected files.\n\n" + "\n".join(skipped[:12]))
            return
        rows = getattr(self, "pending_combined_rows", [])
        authors = {row[2].strip() for row in rows if row[2].strip()}
        sources = {row[3].strip() for row in rows if row[3].strip()}
        first = rows[0] if rows else None
        initial = {
            "title": first[1] if first else "Combined Book",
            "author": next(iter(authors)) if len(authors) == 1 else "",
            "series": first[21] if first and len(first) > 21 else "",
            "series_number": first[22] if first and len(first) > 22 else "",
            "edition": "", "year": "", "isbn": "", "publisher": "",
            "source": next(iter(sources)) if len(sources) == 1 else "",
            "tags": "", "notes": "Combined from selected library books.",
        }
        initial.update({key: value for key, value in detect_metadata_from_text_content("\n".join(sample), existing=initial).items() if value})
        dialog = MetadataDialog(initial, self)
        dialog.setWindowTitle("Combined Book Metadata")
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        metadata = dialog.values()
        if not metadata["title"]:
            QMessageBox.warning(self, "Title Required", "Enter a title before creating the combined EPUB.")
            return
        self.progress_dialog = QProgressDialog("Creating the combined EPUB...", "", 0, 0, self)
        self.progress_dialog.setWindowTitle("Combine Selected Books")
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.setCancelButton(None)
        self.progress_dialog.show()
        worker = CombinedSaveWorker(root, parts, metadata, self)
        self.active_worker = worker
        self.connect_worker_result(worker, lambda book_id, error: self.combined_import_finished(book_id, error, skipped))
        worker.start()

    def folder_import_finished(self, imported, skipped):
        if self.progress_dialog:
            self.progress_dialog.close()
            self.progress_dialog = None
        self.refresh_books()
        report = self.db.folder / "last_import_report.txt"
        if skipped:
            QMessageBox.information(self, "Folder Import Complete", f"Imported {imported}. Skipped {len(skipped)}.\n\nReport: {report}")
        self.statusBar().showMessage(f"Folder import complete. Imported {imported}; skipped {len(skipped)}.")

    def export_book(self):
        if self.active_worker and self.active_worker.isRunning():
            QMessageBox.information(self, "Task in Progress", "Please wait for the current task to finish before exporting.")
            return
        ids = self.current_ids()
        rows = [self.db.get_book(book_id) for book_id in ids]
        rows = [row for row in rows if row]
        if not rows:
            return
        destination = QFileDialog.getExistingDirectory(self, "Choose Export Folder")
        if not destination:
            return
        self.export_rows(rows, destination)

    def export_rows(self, rows, destination, send_to=None):
        if self.active_worker and self.active_worker.isRunning():
            QMessageBox.information(self, "Task in Progress", "Please wait for the current task to finish before sending or exporting.")
            return
        try:
            defaults = BrailleOptions(**json.loads(self.db.get_setting("braille_export_options", "{}")))
        except (ValueError, TypeError):
            defaults = BrailleOptions()
        format_dialog = ExportFormatDialog(self, braille_defaults=defaults, send_to=send_to)
        if format_dialog.exec() != QDialog.DialogCode.Accepted:
            return
        self._conversion_destination = send_to
        braille = format_dialog.braille_options()
        if braille:
            self.db.set_setting("braille_export_options", json.dumps({
                "contracted": braille.contracted,
                "cells_per_line": braille.cells_per_line,
                "lines_per_page": braille.lines_per_page,
            }))
            self.start_braille_export(rows, Path(destination), braille)
            return
        if format_dialog.export_as_epub():
            needs_conversion = any(
                Path(row[8]).is_file() and Path(row[8]).suffix.casefold() != ".epub"
                for row in rows
            )
            converter = find_calibre_tool("ebook-convert") if needs_conversion else None
            if needs_conversion and not converter:
                QMessageBox.warning(
                    self,
                    "Calibre Not Found",
                    "Calibre's ebook-convert tool was not found. Install Calibre to export these books as EPUB.",
                )
                return
            self.start_epub_export(rows, Path(destination), converter)
            return
        transfers = self.prepare_transfers(rows, Path(destination))
        self.start_transfers(transfers, f"Sending to {send_to}" if send_to else "Exporting Books")

    def start_epub_export(self, rows, destination, converter):
        self.progress_dialog = QProgressDialog("Preparing EPUB export...", "", 0, 0, self)
        self.progress_dialog.setWindowTitle("Exporting EPUB Copies")
        if getattr(self, "_conversion_destination", None):
            self.progress_dialog.setWindowTitle(f"Sending EPUB to {self._conversion_destination}")
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.setCancelButton(None)
        self.progress_dialog.show()
        worker = ExportEpubWorker(rows, destination, converter, self)
        self.active_worker = worker
        worker.progress.connect(self.progress_dialog.setLabelText)
        self.connect_worker_result(worker, self.epub_export_finished)
        worker.start()

    def start_braille_export(self, rows, destination, options):
        self.progress_dialog = QProgressDialog("Preparing built-in braille converter...", "Cancel", 0, 0, self)
        self.progress_dialog.setWindowTitle("Exporting BRF" if options.format == "brf" else "Exporting eBraille")
        if getattr(self, "_conversion_destination", None):
            self.progress_dialog.setWindowTitle(f"Converting and Sending to {self._conversion_destination}")
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.show()
        worker = BrailleExportWorker(rows, destination, options, self)
        self.active_worker = worker
        self.progress_dialog.canceled.connect(lambda: self.cancel_braille_export(worker))
        worker.progress.connect(self.progress_dialog.setLabelText)
        self.connect_worker_result(worker, self.braille_export_finished)
        worker.start()

    def cancel_braille_export(self, worker):
        worker.requestInterruption()
        dialog = self.progress_dialog
        if dialog:
            dialog.setLabelText("Cancelling braille export...")
            dialog.setCancelButton(None)
            # Keep other operations blocked while the converter shuts down.
            QTimer.singleShot(0, lambda: dialog.show() if worker.isRunning() and self.progress_dialog is dialog and not self._closing else None)

    def braille_export_finished(self, exported, skipped, cancelled, warnings):
        if self._closing:
            return
        if self.progress_dialog:
            self.progress_dialog.close()
            self.progress_dialog = None
        destination = getattr(self, "_conversion_destination", None)
        message = f"Sent {exported} book(s) to {destination}." if destination else f"Exported {exported} book(s)."
        if cancelled:
            message += " Export cancelled. Completed copies have been kept."
        if skipped:
            message += "\n\nCould not export:\n" + "\n".join(skipped)
        if warnings:
            message += "\n\nExported books needing review:\n\n" + "\n\n".join(warnings)
        if skipped or warnings:
            message += "\n\nThese automated checks are not a certification of braille accuracy."
            ConversionReportDialog("Braille Conversion Results", message, self).exec()
        else:
            self.routine_feedback(message.split("\n")[0], message)
        if skipped or warnings:
            self.statusBar().showMessage(message.split("\n")[0])

    def epub_export_finished(self, exported, skipped):
        if self.progress_dialog:
            self.progress_dialog.close()
            self.progress_dialog = None
        destination = getattr(self, "_conversion_destination", None)
        summary = f"Sent {exported} EPUB book(s) to {destination}." if destination else f"Exported {exported} EPUB book(s)."
        if skipped:
            self.statusBar().showMessage(summary + f" Skipped {len(skipped)}.")
            QMessageBox.warning(
                self,
                "Export Finished with Warnings",
                summary + f" Skipped {len(skipped)}.\n\n" + "\n".join(skipped[:20]),
            )
        else:
            self.routine_feedback(summary)

    def convert_selected_to_epub(self):
        converter = find_calibre_tool("ebook-convert")
        if not converter:
            QMessageBox.warning(self, "Calibre Not Found", "Calibre's ebook-convert tool was not found. Install Calibre, then try again.")
            return
        books = []
        for book_id in self.current_ids():
            row = self.db.get_book(book_id)
            if not row:
                continue
            source = Path(row[8])
            if source.is_file() and source.suffix.casefold() != ".epub":
                books.append((row, source))
        if not books:
            self.routine_feedback("No books to convert. Select books with available files that are not already EPUB.")
            return
        answer = QMessageBox.question(
            self, "Convert Books to EPUB",
            f"Convert {len(books)} selected book(s) to EPUB and add the converted copies to the library? The original files will be kept.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        self.progress_dialog = QProgressDialog("Converting selected books...", "", 0, 0, self)
        self.progress_dialog.setWindowTitle("Convert to EPUB")
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.setCancelButton(None)
        self.progress_dialog.show()
        worker = ConvertBooksWorker(books, converter, self)
        self.active_worker = worker
        worker.progress.connect(self.progress_dialog.setLabelText)
        self.connect_worker_result(worker, self.conversion_finished)
        worker.start()

    def conversion_finished(self, converted, issues):
        if self.progress_dialog:
            self.progress_dialog.close()
            self.progress_dialog = None
        self.refresh_books()
        self.setFocusToCurrentBook()
        text = f"Converted {converted} book(s) to EPUB. Original files were kept."
        if issues:
            text += "\n\nNot converted:\n" + "\n".join(issues[:20])
            if len(issues) > 20:
                text += f"\nAnd {len(issues) - 20} more."
        if issues:
            self.statusBar().showMessage(f"EPUB conversion complete: {converted} converted, {len(issues)} not converted.")
            QMessageBox.warning(self, "EPUB Conversion Results", text)
        else:
            self.routine_feedback(f"Converted {converted} books to EPUB.", text)

    def configure_folder(self, label, setting_key):
        current = self.db.get_setting(setting_key, "")
        folder = QFileDialog.getExistingDirectory(self, f"Choose {label} Folder", current)
        if folder:
            self.db.set_setting(setting_key, folder)
            self.statusBar().showMessage(f"{label} folder saved")
        return folder

    def choose_default_reader(self):
        current = self.db.get_setting("default_reader_path", "")
        path, _filter = QFileDialog.getOpenFileName(self, "Choose Default Ebook Reader", current, "Applications (*.exe);;All files (*)")
        if path:
            self.db.set_setting("default_reader_path", path)
            self.routine_feedback("Default ebook reader saved.", f"Default ebook reader: {path}")

    def clear_default_reader(self):
        self.db.set_setting("default_reader_path", "")
        self.statusBar().showMessage("Books will open with the system default reader")

    def open_library_folder(self):
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.db.folder)))

    @staticmethod
    def explorer_context_menu_registry_paths():
        return [
            rf"Software\Classes\SystemFileAssociations\{extension}\shell\{EXPLORER_CONTEXT_MENU_KEY}"
            for extension in sorted(SUPPORTED_EXTENSIONS)
        ]

    @staticmethod
    def app_launch_command_for_file_argument():
        if getattr(sys, "frozen", False):
            return f'"{sys.executable}" --import "%1"'
        executable = Path(sys.executable)
        pythonw = executable.with_name("pythonw.exe")
        launcher = pythonw if pythonw.exists() else executable
        return f'"{launcher}" "{Path(__file__).resolve()}" --import "%1"'

    @staticmethod
    def delete_registry_tree(subkey):
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, subkey) as key:
                children = []
                index = 0
                while True:
                    try:
                        children.append(winreg.EnumKey(key, index))
                        index += 1
                    except OSError:
                        break
            for child in children:
                QtLibraryWindow.delete_registry_tree(subkey + "\\" + child)
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, subkey)
        except FileNotFoundError:
            pass

    def find_book_elsewhere(self):
        row = self.selected_row()
        if not row:
            return
        words = search_words(row[1], row[2])
        entries = []
        problem = ""
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            for hit in search_gutenberg(row[1], row[2]):
                entries.append((f"Project Gutenberg, free download: {hit['title']} by {hit['author']}", hit["url"], ""))
        except Exception as exc:
            problem = f"Project Gutenberg could not be searched: {exc}"
        finally:
            QApplication.restoreOverrideCursor()
        for group, group_title in FIND_GROUP_TITLES.items():
            for key in services_in_group(group):
                entries.append((f"{group_title}, {FIND_SERVICES[key][0]}", key, words))
        dialog = QDialog(self)
        dialog.setWindowTitle(f"Find {row[1]} Elsewhere")
        dialog.resize(700, 460)
        layout = QVBoxLayout(dialog)
        results = QListWidget(dialog)
        results.setAccessibleName("Places to find this book. Press Enter to open only the one you are on.")
        for text, target, paste in entries:
            item = QListWidgetItem(text)
            item.setData(Qt.ItemDataRole.UserRole, (target, paste))
            results.addItem(item)
        results.setCurrentRow(0)
        layout.addWidget(results)
        if problem:
            note = QLabel(problem, dialog)
            note.setWordWrap(True)
            layout.addWidget(note)

        def open_current():
            item = results.currentItem()
            if item is None:
                return
            target, paste = item.data(Qt.ItemDataRole.UserRole)
            if target in FIND_SERVICES:
                library_key = self.db.get_setting(FIND_LIBRARY_SETTING, "")
                if target == "libby" and not library_key:
                    name, accepted = QInputDialog.getText(
                        dialog, "Your Library",
                        "Your library's Libby or OverDrive name, such as nypl for nypl.overdrive.com. Leave blank to open the Libby site instead:",
                    )
                    if accepted and name.strip():
                        library_key = name.strip()
                        self.db.set_setting(FIND_LIBRARY_SETTING, library_key)
                url, _label, needs_paste = service_target(target, paste, library_key)
                if needs_paste:
                    QApplication.clipboard().setText(paste)
                QDesktopServices.openUrl(QUrl(url))
            else:
                QDesktopServices.openUrl(QUrl(target))

        results.itemActivated.connect(lambda _item: open_current())
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close, parent=dialog)
        open_button = buttons.addButton("Open", QDialogButtonBox.ButtonRole.ActionRole)
        open_button.clicked.connect(open_current)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        results.setFocus(Qt.FocusReason.OtherFocusReason)
        dialog.exec()

    def export_library_for_goodreads(self):
        include_all = QMessageBox.question(
            self, "Export for Goodreads",
            "Include books that have no reading status yet? Goodreads will place them on its To Read shelf. "
            "Choose No to export only No File Available books and books with a status from an earlier Goodreads import.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        ) == QMessageBox.StandardButton.Yes
        target, _ = QFileDialog.getSaveFileName(
            self, "Save Goodreads Import File", "library for goodreads.csv", "CSV file (*.csv)",
        )
        if not target:
            return
        try:
            count = export_goodreads_csv(target, self.db.all_books_for_duplicate_check(), include_all)
        except OSError as exc:
            QMessageBox.warning(self, "Export for Goodreads", f"The file could not be saved: {exc}")
            return
        message = f"Saved {count} books to {target}."
        QMessageBox.information(
            self, "Export for Goodreads",
            message + "\n\nOn Goodreads, choose My Books, then Import and export, then Import, and select this file.",
        )
        self.routine_feedback(message)

    def import_audiobook_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Choose Audiobook Folder or Drive, such as your Libation library")
        if not folder:
            return
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            items = scan_audiobook_folder(folder)
            result = import_audiobooks(self.db, items)
        finally:
            QApplication.restoreOverrideCursor()
        self.refresh_books()
        message = (
            f"Found {len(items)} audiobooks. {result['added']} added, {result['linked']} matched to books that had no file, "
            f"{result['already_linked']} already in the library, {result['already_have']} skipped because you already have that book."
        )
        QMessageBox.information(
            self, "Audiobook Import",
            message + "\n\nAudiobook files stay where they are. Removing a book from the library never deletes them.",
        )
        self.routine_feedback(message)

    def import_goodreads_export(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Choose Goodreads Library Export", "", "Goodreads export (*.csv);;All files (*.*)",
        )
        if not path:
            return
        try:
            entries = parse_goodreads_csv(path)
        except (OSError, ValueError) as exc:
            QMessageBox.warning(self, "Goodreads Import", str(exc))
            return
        result = apply_goodreads_entries(self.db, entries)
        wishlisted = add_wishlist_entries(self.db, result["unmatched_entries"])
        self.refresh_books()
        message = (
            f"Goodreads import finished. {result['updated']} updated, {result['unchanged']} already current, "
            f"{wishlisted} added to your library as No File Available entries."
        )
        saved = ""
        if result["unmatched_entries"]:
            target, _ = QFileDialog.getSaveFileName(
                self, "Save List of Books Not in Library", "Goodreads books not in library.txt", "Text file (*.txt)",
            )
            if target:
                try:
                    write_unmatched_file(target, result["unmatched_entries"])
                    saved = f"\n\nThe list of those books was saved to {target}."
                except OSError as exc:
                    saved = f"\n\nThe list could not be saved: {exc}"
        QMessageBox.information(self, "Goodreads Import", message + saved)
        self.routine_feedback(message)

    def install_file_explorer_context_menu(self):
        if not sys.platform.startswith("win") or winreg is None:
            QMessageBox.information(self, "Windows Only", "File Explorer integration is available on Windows only.")
            return
        command = self.app_launch_command_for_file_argument()
        try:
            for menu_path in self.explorer_context_menu_registry_paths():
                with winreg.CreateKey(winreg.HKEY_CURRENT_USER, menu_path) as key:
                    winreg.SetValueEx(key, "", 0, winreg.REG_SZ, EXPLORER_CONTEXT_MENU_LABEL)
                    winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, sys.executable)
                with winreg.CreateKey(winreg.HKEY_CURRENT_USER, menu_path + r"\command") as key:
                    winreg.SetValueEx(key, "", 0, winreg.REG_SZ, command)
            self.routine_feedback("File Explorer command added.", f"Added File Explorer command: {EXPLORER_CONTEXT_MENU_LABEL}")
        except OSError as exc:
            QMessageBox.critical(self, "File Explorer Integration Failed", str(exc))

    def remove_file_explorer_context_menu(self):
        if not sys.platform.startswith("win") or winreg is None:
            QMessageBox.information(self, "Windows Only", "File Explorer integration is available on Windows only.")
            return
        try:
            for menu_path in self.explorer_context_menu_registry_paths():
                self.delete_registry_tree(menu_path)
            self.routine_feedback("File Explorer command removed.")
        except OSError as exc:
            QMessageBox.critical(self, "File Explorer Integration Failed", str(exc))

    def show_file_explorer_context_menu_status(self):
        installed = 0
        if sys.platform.startswith("win") and winreg is not None:
            for menu_path in self.explorer_context_menu_registry_paths():
                try:
                    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, menu_path):
                        installed += 1
                except FileNotFoundError:
                    pass
        QMessageBox.information(
            self, "File Explorer Integration Status",
            f"Right-click entries installed: {installed} of {len(SUPPORTED_EXTENSIONS)}\n\n"
            f"Command name: {EXPLORER_CONTEXT_MENU_LABEL}\n\n"
            f"Launch command: {self.app_launch_command_for_file_argument()}",
        )

    def import_external_files(self, paths):
        paths = [Path(path) for path in paths if Path(path).suffix.casefold() in SUPPORTED_EXTENSIONS]
        paths = [path for path in paths if not LibraryApp.is_ignored_import_file(None, path)]
        if not paths:
            return
        self.progress_dialog = QProgressDialog("Adding selected files to the library...", "", 0, 0, self)
        self.progress_dialog.setWindowTitle("Add to Library")
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.setCancelButton(None)
        self.progress_dialog.show()
        worker = FolderImportWorker(paths, "File Explorer", self)
        self.active_worker = worker
        worker.progress.connect(self.progress_dialog.setLabelText)
        self.connect_worker_result(worker, self.folder_import_finished)
        worker.start()

    def show_help(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Accessible Ebook Library Manager Help")
        dialog.resize(700, 560)
        layout = QVBoxLayout(dialog)
        help_text = QTextEdit(dialog)
        help_text.setAccessibleName("Help topics")
        help_text.setReadOnly(True)
        help_text.setPlainText(load_help_text(app_data_folder()))
        layout.addWidget(help_text)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close, parent=dialog)
        edit_button = buttons.addButton("Edit Help File", QDialogButtonBox.ButtonRole.ActionRole)
        edit_button.clicked.connect(lambda: self.open_help_file_for_editing(dialog))
        buttons.rejected.connect(dialog.reject)
        buttons.accepted.connect(dialog.accept)
        layout.addWidget(buttons)
        help_text.setFocus(Qt.FocusReason.OtherFocusReason)
        dialog.exec()

    def open_help_file_for_editing(self, dialog):
        path = ensure_user_help(app_data_folder())
        try:
            os.startfile(path)  # type: ignore[attr-defined]
        except Exception as exc:
            QMessageBox.warning(dialog, "Help File", f"Could not open the help file:\n{path}\n\n{exc}")
            return
        dialog.accept()

    def choose_backup_folder(self):
        self.choose_cloud_backup_folder("other")

    def choose_cloud_backup_folder(self, service="other"):
        labels = {
            "onedrive": "OneDrive", "google_drive": "Google Drive",
            "icloud": "iCloud Drive", "other": "backup",
        }
        label = labels.get(service, "backup")
        detected = LibraryApp.detected_cloud_folder(None, service) if service != "other" else None
        folder = None
        if detected and QMessageBox.question(
            self, f"Use {label} Folder?",
            f"A {label} folder was found here:\n{detected}\n\nUse it for library backups?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        ) == QMessageBox.StandardButton.Yes:
            folder = detected
        if folder is None:
            if service != "other" and not detected:
                QMessageBox.information(
                    self, "Cloud Folder Not Found",
                    f"No {label} folder was detected. Choose the synced folder to use for backups.",
                )
            current = self.db.get_setting("backup_folder", "")
            selected = QFileDialog.getExistingDirectory(
                self, f"Choose {label} Backup Folder", current,
            )
            if not selected:
                return
            folder = Path(selected)
        backup_folder = Path(folder)
        if backup_folder.name != "Accessible Ebook Library Manager Backups":
            backup_folder = cloud_backup_subfolder(backup_folder)
        books_root = self.db.books_path.resolve()
        backup_root = backup_folder.resolve()
        try:
            backup_root.relative_to(books_root)
            overlaps_library = True
        except ValueError:
            try:
                books_root.relative_to(backup_root)
                overlaps_library = True
            except ValueError:
                overlaps_library = False
        if overlaps_library:
            QMessageBox.warning(self, "Invalid Backup Folder", "Choose a synced folder outside the managed Books folder.")
            return
        try:
            backup_folder.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            QMessageBox.critical(self, "Backup Folder Failed", f"Could not create the backup folder.\n\n{exc}")
            return
        self.db.set_setting("backup_folder", str(backup_folder))
        self.routine_feedback("Backup folder saved.", f"Backup folder: {backup_folder}. Includes the database and imported book files.")
        if self.db.get_setting("backup_schedule", DEFAULT_BACKUP_SCHEDULE) != "on_demand":
            QTimer.singleShot(1000, self.check_library_backup)

    def backup_library_now(self):
        folder = self.db.get_setting("backup_folder", "")
        if not folder:
            self.choose_backup_folder()
            folder = self.db.get_setting("backup_folder", "")
        if not folder:
            return
        self.start_library_backup(folder, automatic=False)

    def restore_library_backup(self):
        if any(worker and worker.isRunning() for worker in (self.active_worker, self.duplicate_worker, self.metadata_worker, self.backup_worker)):
            QMessageBox.information(self, "Library Busy", "Wait for the current library operation to finish before restoring a backup.")
            return
        folder = self.db.get_setting("backup_folder", "")
        if not folder:
            QMessageBox.information(self, "Library Backup", "Choose a backup folder before restoring.")
            self.choose_backup_folder()
            folder = self.db.get_setting("backup_folder", "")
        if not folder:
            return
        backup_db = Path(folder) / "library_backup.db"
        if not backup_db.is_file():
            QMessageBox.warning(self, "Backup Not Found", f"No library database backup was found here:\n{backup_db}")
            return
        answer = QMessageBox.question(
            self, "Restore Library Backup",
            "Replace the current library database and imported book files with the backup?\n\n"
            "The current database and Books folder will first be saved as safety copies. An incomplete backup will be rejected.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        preserved = {
            "backup_folder": self.db.get_setting("backup_folder", ""),
            "backup_schedule": self.db.get_setting("backup_schedule", DEFAULT_BACKUP_SCHEDULE),
        }
        self.db.connection.commit()
        self.db.close()
        self.db = None
        self.progress_dialog = QProgressDialog("Validating and restoring the library backup...", "", 0, 0, self)
        self.progress_dialog.setWindowTitle("Restore Library Backup")
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.setCancelButton(None)
        self.progress_dialog.show()
        worker = LibraryRestoreWorker(folder, preserved, self)
        self.active_worker = worker
        self.connect_worker_result(worker, self.library_restore_finished)
        worker.start()

    def library_restore_finished(self, message, error):
        if self.progress_dialog:
            self.progress_dialog.close()
            self.progress_dialog = None
        try:
            self.db = LibraryDatabase()
        except Exception as reopen_error:
            QMessageBox.critical(self, "Library Reopen Failed", f"{error or message}\n\nThe library could not be reopened:\n{reopen_error}")
            return
        if error:
            QMessageBox.critical(self, "Library Restore Failed", error)
            return
        self.marked_ids.clear()
        self.refresh_books()
        self.setFocusToCurrentBook()
        self.statusBar().showMessage("Library restore complete")
        self.routine_feedback("Library restored.", message)

    def start_library_backup(self, folder, automatic=False):
        if self.backup_worker and self.backup_worker.isRunning():
            return
        if not automatic:
            self.progress_dialog = QProgressDialog("Backing up the library and book files...", "", 0, 0, self)
            self.progress_dialog.setWindowTitle("Library Backup")
            self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
            self.progress_dialog.setCancelButton(None)
            self.progress_dialog.show()
        worker = LibraryBackupWorker(folder, automatic, self)
        self.backup_worker = worker
        self.connect_worker_result(worker, lambda message, error, automatic=automatic: self.library_backup_finished(message, error, automatic))
        worker.start()

    def library_backup_finished(self, message, error, automatic):
        if self.progress_dialog:
            self.progress_dialog.close()
            self.progress_dialog = None
        if error:
            self.statusBar().showMessage("Library backup failed")
            if not automatic:
                QMessageBox.critical(self, "Library Backup Failed", error)
        elif message:
            self.statusBar().showMessage(message)
            if not automatic:
                self.routine_feedback("Library backup complete.", message)
        if not self.isVisible():
            return

    def set_backup_schedule(self, schedule):
        if schedule not in BACKUP_SCHEDULES:
            schedule = DEFAULT_BACKUP_SCHEDULE
        self.db.set_setting("backup_schedule", schedule)
        label = BACKUP_SCHEDULES[schedule][0]
        self.routine_feedback(f"Backup schedule set to {label}.")
        if schedule != "on_demand":
            QTimer.singleShot(1000, self.check_library_backup)

    def check_library_backup(self):
        if self._closing or self.db is None:
            return
        if any(worker and worker.isRunning() for worker in (self.active_worker, self.duplicate_worker, self.metadata_worker, self.backup_worker)):
            QTimer.singleShot(60 * 60 * 1000, self.check_library_backup)
            return
        folder = self.db.get_setting("backup_folder", "")
        if folder and self.db.get_setting("backup_schedule", DEFAULT_BACKUP_SCHEDULE) != "on_demand":
            self.start_library_backup(folder, automatic=True)
        QTimer.singleShot(60 * 60 * 1000, self.check_library_backup)

    def show_backup_status(self):
        folder = self.db.get_setting("backup_folder", "")
        schedule = self.db.get_setting("backup_schedule", DEFAULT_BACKUP_SCHEDULE)
        label = BACKUP_SCHEDULES.get(schedule, BACKUP_SCHEDULES[DEFAULT_BACKUP_SCHEDULE])[0]
        last = self.db.get_setting("last_backup_at", "Never") or "Never"
        backup_db = Path(folder) / "library_backup.db" if folder else None
        manifest = {}
        manifest_path = Path(folder) / "library_backup_manifest.json" if folder else None
        if manifest_path and manifest_path.is_file():
            try:
                manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            except (OSError, ValueError):
                manifest = {}
        count = manifest.get("backup_book_file_count", 0)
        size = manifest.get("backup_book_byte_count", 0)
        text = (
            f"Backup folder: {folder or 'Not set'}\n"
            f"Schedule: {label}\n"
            f"Last backup: {last}\n"
            f"Database backup: {'Found' if backup_db and backup_db.is_file() else 'Not found'}\n"
            f"Book files backed up: {count}\n"
            f"Book backup size: {size} bytes"
        )
        QMessageBox.information(self, "Library Backup Status", text)

    def show_calibre_status(self):
        metadata_tool = find_calibre_tool("ebook-meta")
        converter = find_calibre_tool("ebook-convert")
        reading = self.db.get_setting("calibre_metadata_reading", "1") != "0"
        QMessageBox.information(
            self, "Kindle and Calibre",
            f"Calibre metadata reading: {'On' if reading else 'Off'}\n\nMetadata tool: {metadata_tool or 'Not found'}\nConversion tool: {converter or 'Not found'}\n\nCalibre runs only as a background helper; it does not remove DRM.",
        )

    def toggle_calibre_metadata(self):
        enabled = self.db.get_setting("calibre_metadata_reading", "1") == "0"
        self.db.set_setting("calibre_metadata_reading", "1" if enabled else "0")
        state = "on" if enabled else "off"
        self.routine_feedback(f"Calibre metadata reading {state}.")

    def send_to_folder(self, app_name, setting_key):
        ids = self.current_ids()
        rows = [self.db.get_book(book_id) for book_id in ids]
        rows = [row for row in rows if row]
        if not rows:
            return
        folder = self.db.get_setting(setting_key, "")
        if not folder or not Path(folder).is_dir():
            folder = self.configure_folder(app_name, setting_key)
        if not folder:
            return
        target = Path(folder)
        self.export_rows(rows, target, send_to=app_name)

    def send_to_nls_ereader(self):
        rows = [self.db.get_book(book_id) for book_id in self.current_ids()]
        rows = [row for row in rows if row and Path(row[8]).is_file()]
        if not rows:
            QMessageBox.warning(self, "No Available Book Files", "Select at least one book with an available stored file.")
            return

        folder_text = self.db.get_setting("nls_ereader_folder", "")
        folder = Path(folder_text) if folder_text else None
        if not folder or not folder.is_dir():
            folder = None
            candidates = ReaderDeviceDetection().likely_nls_ereader_drives() if sys.platform.startswith("win") else []
            if len(candidates) == 1:
                root, label, _score = candidates[0]
                display = f"{root} {label}".strip()
                if QMessageBox.question(
                    self, "NLS eReader Found", f"Use this likely NLS eReader?\n{display}",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.Yes,
                ) == QMessageBox.StandardButton.Yes:
                    folder = Path(root)
            elif candidates:
                choices = [f"{root} {label}".strip() for root, label, _score in candidates]
                choice, accepted = QInputDialog.getItem(self, "Choose NLS eReader", "Likely eReader drives:", choices, 0, False)
                if accepted:
                    folder = Path(candidates[choices.index(choice)][0])

            if folder is None:
                QMessageBox.information(
                    self, "Choose NLS eReader Folder",
                    "No NLS eReader drive was selected. Choose its drive or the folder where books should be copied.",
                )
                selected = QFileDialog.getExistingDirectory(self, "Choose NLS eReader Folder")
                if not selected:
                    return
                folder = Path(selected)
            if not folder.is_dir():
                QMessageBox.warning(self, "NLS eReader Folder Missing", f"This folder is not available:\n{folder}")
                return
            self.db.set_setting("nls_ereader_folder", str(folder))

        self.export_rows(rows, folder, send_to="NLS eReader")

    def prepare_transfers(self, rows, target_folder):
        planned = set()
        transfers = []
        for row in rows:
            source = Path(row[8])
            if not source.exists():
                transfers.append((row[1], source, target_folder / source.name, False))
                continue
            name = source.name
            if name.casefold() in planned:
                stem = safe_filename(f"{row[2]} - {row[1]}").strip(" -") or safe_filename(row[1])
                name = f"{stem}{source.suffix}"
                serial = 2
                while name.casefold() in planned:
                    name = f"{stem} ({serial}){source.suffix}"
                    serial += 1
            planned.add(name.casefold())
            try:
                same_file = source.resolve() == (target_folder / name).resolve()
            except OSError:
                same_file = False
            if same_file:
                continue
            transfers.append((row[1], source, target_folder / name, False))
        existing = [path.name for _title, _source, path, _replace in transfers if path.exists()]
        replace = False
        if existing:
            answer = QMessageBox.question(
                self, "Replace Existing Files?",
                f"{len(existing)} destination file(s) already exist. Replace them?\n\n" + "\n".join(existing[:8]),
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            replace = answer == QMessageBox.StandardButton.Yes
        return [(title, source, path, replace) for title, source, path, _ in transfers]

    def start_transfers(self, transfers, title):
        if not transfers:
            self.statusBar().showMessage("No files to copy")
            return
        self.progress_dialog = QProgressDialog("Preparing file transfer...", "", 0, 0, self)
        self.progress_dialog.setWindowTitle(title)
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.setCancelButton(None)
        self.progress_dialog.show()
        worker = CopyBooksWorker(transfers, self)
        self.active_worker = worker
        worker.progress.connect(self.progress_dialog.setLabelText)
        self.connect_worker_result(worker, self.transfers_finished)
        worker.start()

    def transfers_finished(self, copied, skipped):
        if self.progress_dialog:
            self.progress_dialog.close()
            self.progress_dialog = None
        if skipped:
            self.statusBar().showMessage(f"Copied {copied} book files. Skipped {len(skipped)}.")
            QMessageBox.warning(self, "Transfer Complete", f"Copied {copied}. Skipped {len(skipped)}.\n\n" + "\n".join(skipped[:16]))
        else:
            self.routine_feedback(f"Copied {copied} book files.")

    def set_kindle_emails(self):
        current = self.db.get_setting("kindle_email", "")
        value, accepted = QInputDialog.getText(
            self, "Kindle Email Addresses",
            "Enter Send to Kindle email addresses separated by commas, semicolons, or spaces:",
            text=current,
        )
        if not accepted:
            return
        import re
        parts = [part.strip() for part in re.split(r"[,;\s]+", value) if part.strip()]
        invalid = [part for part in parts if "@" not in part or "." not in part.rsplit("@", 1)[-1]]
        if invalid:
            QMessageBox.warning(self, "Invalid Email Address", f"Check this address:\n{invalid[0]}")
            return
        preserved = []
        seen = set()
        for part in parts:
            key = part.casefold()
            if key not in seen:
                preserved.append(part)
                seen.add(key)
        self.db.set_setting("kindle_email", ", ".join(preserved))
        self.statusBar().showMessage(f"Saved {len(preserved)} Kindle email address(es)")

    def send_to_kindle(self):
        import re
        import urllib.parse
        import webbrowser
        emails = [part.strip() for part in re.split(r"[,;\s]+", self.db.get_setting("kindle_email", "")) if part.strip()]
        if not emails:
            self.set_kindle_emails()
            emails = [part.strip() for part in re.split(r"[,;\s]+", self.db.get_setting("kindle_email", "")) if part.strip()]
        if not emails:
            return
        if len(emails) > 1:
            dialog = QDialog(self)
            dialog.setWindowTitle("Kindle Recipients")
            dialog.resize(480, min(560, 120 + len(emails) * 34))
            layout = QVBoxLayout(dialog)
            layout.addWidget(QLabel("Choose the Kindle addresses that should receive this book:", dialog))
            checks = []
            for email in emails:
                check = QCheckBox(email, dialog)
                check.setAccessibleName(email)
                check.setChecked(True)
                layout.addWidget(check)
                checks.append(check)
            buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel, parent=dialog)
            buttons.accepted.connect(dialog.accept)
            buttons.rejected.connect(dialog.reject)
            layout.addWidget(buttons)
            if dialog.exec() != QDialog.DialogCode.Accepted:
                return
            recipients = [check.text() for check in checks if check.isChecked()]
            if not recipients:
                QMessageBox.information(self, "Kindle Recipients", "Select at least one recipient.")
                return
        else:
            recipients = emails
        rows = []
        missing = []
        for book_id in self.current_ids():
            row = self.db.get_book(book_id)
            if not row:
                continue
            if Path(row[8]).is_file():
                rows.append(row)
            else:
                missing.append(row[1])
        if not rows:
            message = "Select at least one book with an available stored file."
            if missing:
                message += "\n\nMissing files:\n" + "\n".join(missing[:20])
            QMessageBox.warning(self, "No Available Book Files", message)
            return
        subject = "Send to Kindle" if len(rows) != 1 else f"Send to Kindle: {rows[0][1]}"
        body_lines = ["Attach these book files before sending:", ""]
        body_lines.extend(row[8] for row in rows)
        if missing:
            body_lines.extend(["", "These selected books were skipped because their stored files were missing:"])
            body_lines.extend(missing)
        body = "\r\n".join(body_lines)
        mailto = "mailto:" + ",".join(urllib.parse.quote(email) for email in recipients)
        mailto += "?subject=" + urllib.parse.quote(subject) + "&body=" + urllib.parse.quote(body)
        try:
            webbrowser.open(mailto)
        except Exception as exc:
            QMessageBox.warning(self, "Could Not Open Email", f"Could not open your email app.\n\n{exc}")
            return
        message = "Your email app should open. Attach the listed book file(s) and send the message:\n\n" + "\n".join(row[8] for row in rows)
        if missing:
            message += "\n\nSkipped books with missing files:\n" + "\n".join(missing[:20])
        QMessageBox.information(self, "Send to Kindle", message)

    def send_to_humanware(self):
        if not sys.platform.startswith("win"):
            QMessageBox.information(self, "HumanWare MTP Transfer", "HumanWare MTP transfer is available on Windows only.")
            return
        books = []
        for book_id in self.current_ids():
            row = self.db.get_book(book_id)
            if not row:
                continue
            source = Path(row[8])
            if source.is_file():
                books.append((row[1], source))
        if not books:
            QMessageBox.warning(self, "No Available Book Files", "Select at least one book with an available stored file.")
            return
        self.start_humanware_transfer(books)

    def start_humanware_transfer(self, books, device_name=""):
        worker = HumanWareMtpWorker(books, device_name, self)
        self.active_worker = worker
        worker.progress.connect(self.statusBar().showMessage)
        self.connect_worker_result(worker, lambda results, devices: self.humanware_transfer_finished(books, results, devices))
        worker.start()

    def humanware_transfer_finished(self, books, results, devices):
        if devices:
            device, accepted = QInputDialog.getItem(
                self,
                "Choose HumanWare eReader",
                "Windows did not identify the eReader by name. Choose the connected device:",
                devices,
                0,
                False,
            )
            if accepted and device:
                self.start_humanware_transfer(books, device)
            else:
                self.statusBar().showMessage("HumanWare transfer canceled.")
            return
        succeeded = [message for _title, ok, message in results if ok]
        failed = [f"{title}: {message}" for title, ok, message in results if not ok]
        if succeeded and not failed:
            self.routine_feedback(f"Sent {len(succeeded)} books to the HumanWare eReader.", "\n".join(succeeded))
        elif succeeded or failed:
            summary = []
            if succeeded:
                summary.append(f"Sent {len(succeeded)} book(s).")
            if failed:
                summary.append("Some books could not be sent:\n" + "\n".join(failed))
            QMessageBox.warning(self, "HumanWare Transfer Results", "\n\n".join(summary))
            self.statusBar().showMessage(f"HumanWare transfer finished: {len(succeeded)} sent, {len(failed)} failed.")
        else:
            QMessageBox.warning(self, "HumanWare eReader Not Found", "No book was sent. Connect and unlock the reader, select MTP or File Transfer mode, then try again.")

    def watched_folders(self):
        try:
            folders = json.loads(self.db.get_setting("watched_folders", "[]"))
        except (TypeError, ValueError):
            folders = []
        result = []
        seen = set()
        for folder in folders:
            path = str(Path(folder).expanduser())
            if path.casefold() not in seen:
                result.append(path)
                seen.add(path.casefold())
        return result

    def add_watched_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Choose Folder to Watch")
        if not folder:
            return
        folders = self.watched_folders()
        if folder.casefold() not in {path.casefold() for path in folders}:
            folders.append(folder)
        self.db.set_setting("watched_folders", json.dumps(folders))
        self.db.set_setting("watched_folder_auto_scan", "1")
        self.statusBar().showMessage(f"Watched folder added: {folder}")
        self.show_watched_folder_status()

    def remove_watched_folder(self):
        folders = self.watched_folders()
        if not folders:
            self.routine_feedback("No watched folders are configured.")
            return
        folder, accepted = QInputDialog.getItem(self, "Remove Watched Folder", "Choose the folder to stop watching:", folders, 0, False)
        if accepted:
            folders.remove(folder)
            self.db.set_setting("watched_folders", json.dumps(folders))
            self.statusBar().showMessage(f"Stopped watching: {folder}")

    def toggle_watched_auto_scan(self):
        enabled = self.db.get_setting("watched_folder_auto_scan", "1") == "1"
        self.db.set_setting("watched_folder_auto_scan", "0" if enabled else "1")
        state = "off" if enabled else "on"
        self.statusBar().showMessage(f"Automatic watched-folder scanning is {state}")
        if not enabled:
            QTimer.singleShot(1000, self.check_watched_folders)

    def show_watched_folder_status(self):
        folders = self.watched_folders()
        scanning = self.db.get_setting("watched_folder_auto_scan", "1") == "1"
        dialog = QDialog(self)
        dialog.setWindowTitle("Watched Folder Status")
        dialog.resize(720, 420)
        layout = QVBoxLayout(dialog)
        layout.addWidget(QLabel(f"Automatic scanning: {'On' if scanning else 'Off'}\nFolders watched: {len(folders)}", dialog))
        paths = QPlainTextEdit(dialog)
        paths.setAccessibleName("Watched folder paths")
        paths.setReadOnly(True)
        paths.setPlainText("\n".join(f"{index}. {path}" for index, path in enumerate(folders, 1)) or "No watched folders configured.")
        layout.addWidget(paths, 1)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close, parent=dialog)
        buttons.rejected.connect(dialog.reject)
        buttons.accepted.connect(dialog.accept)
        layout.addWidget(buttons)
        paths.setFocus(Qt.FocusReason.OtherFocusReason)
        dialog.exec()

    def check_watched_folders(self):
        if self._closing or self.db is None:
            return
        if self.db.get_setting("watched_folder_auto_scan", "1") == "1" and self.watched_folders():
            self.scan_watched_folders(manual=False)
        QTimer.singleShot(5 * 60 * 1000, self.check_watched_folders)

    def scan_watched_folders(self, manual=True):
        if self.watch_scan_running:
            if manual:
                QMessageBox.information(self, "Watched Folders", "A watched-folder scan is already running.")
            return
        if not self.watched_folders():
            if manual:
                QMessageBox.information(self, "Watched Folders", "Add a watched folder in Settings before scanning.")
            return
        self.watch_scan_running = True
        if manual:
            self.progress_dialog = QProgressDialog("Scanning watched folders...", "", 0, 0, self)
            self.progress_dialog.setWindowTitle("Scanning Watched Folders")
            self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
            self.progress_dialog.setCancelButton(None)
            self.progress_dialog.show()
        worker = WatchedFolderWorker(self)
        self.active_worker = worker
        worker.progress.connect(lambda name: self.statusBar().showMessage(f"Checking: {name}"))
        self.connect_worker_result(worker, lambda imported, updated, skipped, error, manual=manual: self.watched_scan_finished(imported, updated, skipped, error, manual))
        worker.start()

    def watched_scan_finished(self, imported, updated, skipped, error, manual):
        self.watch_scan_running = False
        if self.progress_dialog:
            self.progress_dialog.close()
            self.progress_dialog = None
        if error:
            self.statusBar().showMessage("Watched-folder scan failed")
            if manual:
                QMessageBox.critical(self, "Watched Folder Scan Failed", error)
            return
        if imported or updated:
            current = self.selected_row()
            self.refresh_books(current[0] if current else None, focus_books=isinstance(self.focusWidget(), AccessibleBookRow))
        if not manual or skipped:
            self.routine_feedback(f"Watched-folder scan complete. Imported {imported}; updated {updated}; skipped {len(skipped)}.", speak=False)
        if manual:
            message = f"Imported {imported}. Updated {updated}. Skipped {len(skipped)}."
            if skipped:
                report = self.db.folder / "last_import_report.txt"
                message += f"\n\nReport: {report}"
            if skipped:
                QMessageBox.warning(self, "Watched Folder Scan Results", message)
            else:
                self.routine_feedback(f"Folder scan complete. Imported {imported}; updated {updated}.")

    def connect_worker_result(self, worker, callback):
        worker.result_args = None
        worker.cancel_requested = False
        dialog = self.progress_dialog
        if isinstance(worker, (CoverImageWorker, DuplicateCleanupWorker)):
            dialog = None
        if dialog and getattr(dialog, "operation_worker", None) is not None:
            dialog = None
        worker.operation_dialog = dialog
        worker.operation_title = dialog.windowTitle() if dialog else "Background operation"
        if dialog:
            dialog.operation_worker = worker
            cancellable = isinstance(worker, (ReindexWorker, FolderScanWorker, FolderImportWorker,
                CombinedPartsWorker, CopyBooksWorker, ExportEpubWorker, ConvertBooksWorker,
                OnlineMetadataWorker, WatchedFolderWorker))
            if cancellable:
                dialog.setAutoClose(False)
                dialog.setAutoReset(False)
                dialog.setCancelButtonText("Cancel")
                worker.cancel_callback = lambda: self.cancel_operation(worker)
                dialog.canceled.connect(worker.cancel_callback)
            elif not isinstance(worker, BrailleExportWorker):
                dialog.setLabelText(dialog.labelText() + "\nThis step must finish before closing.")
        worker.completed.connect(lambda *args: setattr(worker, "result_args", args))

        def deliver_result():
            if dialog is None and self.progress_dialog is not None and not self._closing:
                # Background completion must not close another job's progress
                # dialog or open a report over its cancellation controls.
                QTimer.singleShot(100, deliver_result)
                return
            if dialog and hasattr(worker, "cancel_callback"):
                dialog.canceled.disconnect(worker.cancel_callback)
            if self._closing:
                return
            if worker.cancel_requested:
                self.cancelled_operation_finished(worker)
            elif worker.result_args is not None:
                callback(*worker.result_args)

        # Delivery after finished prevents chained dialogs/operations or closing
        # the window from destroying a thread still unwinding its final work.
        worker.finished.connect(deliver_result)

    def cancel_operation(self, worker):
        if not worker.isRunning():
            return
        worker.cancel_requested = True
        worker.requestInterruption()
        dialog = worker.operation_dialog
        if dialog:
            if hasattr(worker, "progress"):
                try:
                    worker.progress.disconnect(dialog.setLabelText)
                except (RuntimeError, TypeError):
                    pass
            dialog.setLabelText("Cancelling after the current step. Completed work will be kept.")
            dialog.setCancelButton(None)
            QTimer.singleShot(0, lambda: dialog.show() if worker.isRunning() and not self._closing else None)

    def cancelled_operation_finished(self, worker):
        dialog = worker.operation_dialog
        if dialog:
            dialog.close()
            if self.progress_dialog is dialog:
                self.progress_dialog = None
        args = worker.result_args or ()
        message = "Cancelled. Completed work has been kept."
        if args and type(args[0]) is int:
            message += f"\nCompleted items: {args[0]}."
        if isinstance(worker, WatchedFolderWorker):
            self.watch_scan_running = False
            message += f" Updated items: {args[1] if len(args) > 1 else 0}."
        if isinstance(worker, (FolderImportWorker, ConvertBooksWorker, WatchedFolderWorker)):
            row = self.selected_row()
            self.refresh_books(row[0] if row else None)
        if isinstance(worker, OnlineMetadataWorker):
            self.pending_lookup_row = None
            self.pending_cover_lookup_book_id = None
            message = "Cancelled. No online metadata was applied."
        issues = next((value for value in args if isinstance(value, list) and all(isinstance(item, str) for item in value)), [])
        if issues:
            message += "\n\n" + "\n".join(issues)
        self.statusBar().showMessage(worker.operation_title + ": " + message)
        if dialog:
            if issues:
                QMessageBox.warning(self, worker.operation_title + " Cancelled", message)
            else:
                self.routine_feedback(message.split("\n")[0], worker.operation_title + ": " + message)

    def closeEvent(self, event):
        running = [worker for worker in self.findChildren(QThread) if worker.isRunning()]
        if running:
            event.ignore()
            if not self._closing:
                answer = QMessageBox.question(
                    self, "Work in Progress", "Stop remaining work and close when the current steps finish? Completed work will be kept.",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No,
                )
                if answer != QMessageBox.StandardButton.Yes:
                    return
                self._closing = True
                self.cover_preview_timer.stop()
                for worker in running:
                    worker.requestInterruption()
                self.shutdown_dialog = QProgressDialog("Finishing current steps before closing...", "", 0, 0, self)
                self.shutdown_dialog.setWindowTitle("Closing Library Manager")
                self.shutdown_dialog.setCancelButton(None)
                self.shutdown_dialog.setWindowModality(Qt.WindowModality.WindowModal)
                self.shutdown_dialog.show()
            QTimer.singleShot(100, self.close)
            return
        self._closing = True
        self.cover_preview_timer.stop()
        self.statusBar().announcement_timer.stop()
        if self.db is not None:
            self.db.close()
            self.db = None
        super().closeEvent(event)


def command_line_import_paths(arguments):
    paths = []
    index = 0
    while index < len(arguments):
        argument = arguments[index]
        if argument == "--import" and index + 1 < len(arguments):
            index += 1
            paths.append(Path(arguments[index]))
        elif argument.startswith("--import="):
            paths.append(Path(argument.split("=", 1)[1]))
        elif not argument.startswith("-") and Path(argument).suffix.casefold() in SUPPORTED_EXTENSIONS:
            paths.append(Path(argument))
        index += 1
    return paths


def main():
    import_paths = command_line_import_paths(sys.argv[1:])
    app = QApplication([sys.argv[0]])
    app.setApplicationName(APP_NAME)
    window = QtLibraryWindow()
    window.show()
    if import_paths:
        QTimer.singleShot(0, lambda: window.import_external_files(import_paths))
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
