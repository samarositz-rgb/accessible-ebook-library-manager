import os
import tempfile
import unittest
from pathlib import Path


try:
    from PySide6.QtCore import Qt, QThread, Signal
    from PySide6.QtGui import QAccessible
    from PySide6.QtTest import QTest
    from PySide6.QtWidgets import QApplication, QDialog, QMessageBox, QProgressDialog, QPushButton, QScrollArea, QLabel
    from qt_library_manager import BrailleExportWorker, ExportEpubWorker, ExportFormatDialog, FolderImportWorker, MetadataDialog, PreferencesDialog, CopyBooksWorker, ReindexWorker, QtLibraryWindow, command_line_import_paths, DiscardChangesDialog, RecentMessagesDialog
except Exception as import_error:
    QApplication = None
    QT_IMPORT_ERROR = import_error
else:
    QT_IMPORT_ERROR = None

from unittest.mock import patch


@unittest.skipIf(QApplication is None, f"PySide6 unavailable: {QT_IMPORT_ERROR}")
class QtLibraryManagerTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.profile = tempfile.TemporaryDirectory(prefix="aelm-qt-test-")
        os.environ["APPDATA"] = self.profile.name
        self.window = QtLibraryWindow()

    def tearDown(self):
        self.window.close()
        self.profile.cleanup()

    def add_book(self, title, author=""):
        path = os.path.join(self.profile.name, title.replace(" ", "_") + ".txt")
        return self.window.db.add_book(title, author, "", "", "", path, path)

    def test_first_letter_skips_leading_articles(self):
        self.add_book("The Zebra")
        self.add_book("A Apple")
        self.window.refresh_books()
        self.window.single_list.setFocus()
        QTest.keyClick(self.window.single_list.currentItem().row_widget, Qt.Key.Key_Z)
        self.assertEqual(self.window.single_list.currentItem().data(Qt.ItemDataRole.UserRole + 1), "The Zebra")

    def test_new_library_records_only_store_dropdown_categories(self):
        book = self.window.db.add_book("Example", "", "", "EPUB, Fiction, Unknown, nonfiction", "", "example.epub", "example.epub")
        row = self.window.db.get_book(book)
        self.assertEqual(row[4], "Fiction, Non-fiction")
        self.assertEqual(row[6], "epub")

    def test_keyboard_search_ranks_title_before_content_in_every_view(self):
        unrelated = self.add_book("Alpha unrelated")
        mention = self.add_book("Another book mentioning the bridge")
        target = self.add_book("The Golden Gate Bridge")
        self.window.db.index_book_content(unrelated, "A golden crown. Many pages later, a garden gate.")
        self.window.db.index_book_content(mention, "We visited the Golden Gate on a trip.")
        self.window.show()
        self.window.activateWindow()
        self.app.processEvents()
        for mode in ("single", "multiple", "letters", "tags"):
            with self.subTest(mode=mode):
                self.window.set_view_mode(mode)
                QTest.keyClick(self.window.focusWidget(), Qt.Key.Key_F, Qt.KeyboardModifier.ControlModifier)
                self.assertIs(self.window.focusWidget(), self.window.search)
                QTest.keyClicks(self.window.search, "Golden Gate")
                QTest.keyClick(self.window.search, Qt.Key.Key_Return)
                self.assertEqual([row[0] for row in self.window.current_rows], [target, mention])
                self.assertEqual(self.window.selected_row()[0], target)
                QTest.keyClick(self.window.focusWidget(), Qt.Key.Key_F, Qt.KeyboardModifier.ControlModifier)
                self.assertEqual(self.window.search.selectedText(), "Golden Gate")
                QTest.keyClicks(self.window.search, "no matching phrase")
                QTest.keyClick(self.window.search, Qt.Key.Key_Return)
                self.assertEqual(self.window.current_rows, [])
                self.assertIsNone(self.window.selected_row())
                QTest.keyClick(self.window.search, Qt.Key.Key_Escape)
                self.assertEqual(len(self.window.current_rows), 3)

    def test_search_text_handles_quoted_phrases_as_literal_text(self):
        book = self.add_book("Example")
        self.window.db.index_book_content(book, 'She said "Golden Gate" clearly.')
        self.assertIn(book, self.window.db.search_content('"Golden Gate"'))

    def test_search_clears_hidden_filters_and_empty_results_have_no_stale_book(self):
        book = self.add_book("The Golden Gate Bridge")
        self.window.set_view_mode("tags")
        self.window.filter_source = "not this source"
        self.window.search.setText("Golden Gate")
        self.window.search_and_focus()
        self.assertEqual(self.window.selected_row()[0], book)
        self.window.search.setText("no possible match")
        self.window.search_and_focus()
        self.assertEqual(self.window.current_rows, [])
        self.assertIsNone(self.window.selected_row())
        self.assertIs(self.window.focusWidget(), self.window.search)
        self.assertFalse(self.window.column_pool[0][1].accessible_row.isVisible())
        self.window.clear_search()
        self.assertEqual(self.window.selected_row()[0], book)

    def test_sound_settings_apply_cancel_and_keep_dialog_open(self):
        self.window.set_missing_metadata_sound_mode("off")
        dialog = PreferencesDialog(self.window)
        dialog.show()
        dialog.sound_mode.setCurrentIndex(dialog.sound_mode.findData("tags"))
        self.assertEqual(self.window.missing_metadata_sound_mode(), "off")
        dialog.apply_settings()
        self.assertTrue(dialog.isVisible())
        self.assertEqual(self.window.missing_metadata_sound_mode(), "tags")
        dialog.sound_mode.setCurrentIndex(dialog.sound_mode.findData("off"))
        dialog.reject()
        self.assertEqual(self.window.missing_metadata_sound_mode(), "tags")

    def test_speech_dropdown_replaces_buttons_and_saves_only_on_apply_or_save(self):
        self.window.db.set_setting("book_list_speech_fields", "title,author")
        dialog = PreferencesDialog(self.window)
        dialog.show()
        labels = [button.text() for button in dialog.findChildren(QPushButton)]
        self.assertNotIn("Title Only", labels)
        self.assertNotIn("All Details", labels)
        self.assertNotIn("Show Current Reading", labels)
        self.assertEqual(dialog.speech_mode.currentData(), ["title", "author"])
        dialog.speech_mode.setCurrentIndex(0)
        self.assertEqual(self.window.speech_fields(), ["title", "author"])
        dialog.apply_settings()
        self.assertTrue(dialog.isVisible())
        self.assertEqual(self.window.speech_fields(), ["title"])
        dialog.speech_mode.setCurrentIndex(3)
        dialog.reject()
        self.assertEqual(self.window.speech_fields(), ["title"])
        reopened = PreferencesDialog(self.window)
        reopened.speech_mode.setCurrentIndex(2)
        reopened.accept()
        self.assertEqual(self.window.speech_fields(), ["title", "author", "edition"])

    def test_speech_dropdown_preserves_existing_custom_fields(self):
        self.window.db.set_setting("book_list_speech_fields", "title,publisher")
        dialog = PreferencesDialog(self.window)
        self.assertEqual(dialog.speech_mode.currentData(), ["title", "publisher"])
        dialog.accept()
        self.assertEqual(self.window.speech_fields(), ["title", "publisher"])

    def test_entering_category_checks_missing_metadata_sound(self):
        first = self.add_book("Alpha")
        self.window.db.update_book(first, "Alpha", "", "", "Fiction", "")
        self.add_book("Zulu")
        self.window.set_view_mode("multiple")
        self.window.set_missing_metadata_sound_mode("tags")
        with patch("qt_library_manager.LibraryApp.play_missing_metadata_sound") as sound:
            self.window.move_to_column(3)
            sound.assert_called_once()
            self.window.move_to_column(2)
            self.window.move_to_column(3)
            self.assertEqual(sound.call_count, 2)

    def test_import_completion_reveals_book_hidden_by_search(self):
        source = Path(self.profile.name) / "chapter.txt"
        source.write_text("Example text", encoding="utf-8")
        self.window.search.setText("unrelated search")
        self.window.filter_source = "Bookshare"
        self.window.added_count = 0
        self.window.pending_add_paths = []
        def accept(dialog):
            dialog.accept()
            return dialog.result()
        with patch.object(MetadataDialog, "exec", accept):
            self.window.add_book_ready(source, {"title": "Imported chapter"}, "")
        self.assertEqual(self.window.selected_row()[1], "Imported chapter")
        self.assertEqual(self.window.search.text(), "")
        self.assertEqual(self.window.filter_source, "")

    def test_settings_command_keeps_settings_and_focus(self):
        from PySide6.QtGui import QAction
        dialog = PreferencesDialog(self.window)
        dialog.show()
        button = dialog.findChildren(QPushButton)[0]
        button.setFocus()
        action = QAction("Example", dialog)
        dialog.choose_command(action)
        self.assertTrue(dialog.isVisible())
        self.assertIs(dialog.focusWidget(), button)
        dialog.reject()

    def test_wizard_reviews_before_closing(self):
        from metadata_wizard import MetadataWizard
        from unittest.mock import Mock
        wizard = MetadataWizard(Path("test.txt"), {"title": "Test"}, self.window)
        wizard.review_callback = Mock(return_value=False)
        wizard.accept()
        self.assertEqual(wizard.result(), QDialog.DialogCode.Rejected)
        wizard.review_callback.return_value = True
        wizard.accept()
        self.assertEqual(wizard.result(), QDialog.DialogCode.Accepted)

    def test_removal_dialog_uses_safe_default_without_repeated_message_label(self):
        from qt_library_manager import RemoveBooksDialog
        dialog = RemoveBooksDialog(2, self.window)
        dialog.show()
        self.assertFalse(dialog.findChildren(QLabel))
        self.assertEqual(dialog.focusWidget().text(), "Cancel")
        self.assertIn("stored files will be kept", dialog.windowTitle())
        dialog.reject()

    def test_book_rows_are_static_text_not_counted_list_items(self):
        self.add_book("Alpha")
        self.window.refresh_books()
        row = self.window.single_list.currentItem().row_widget
        interface = QAccessible.queryAccessibleInterface(row)
        self.assertEqual(interface.role(), QAccessible.Role.StaticText)
        self.assertNotEqual(interface.role(), QAccessible.Role.ListItem)
        self.assertTrue(interface.state().focusable)

    def test_arrow_keys_move_focus_between_book_rows(self):
        self.add_book("Alpha")
        self.add_book("Beta")
        self.window.refresh_books()
        self.window.single_list.setCurrentRow(0)
        self.window.single_list.setFocus()
        first_row = self.window.single_list.currentItem().row_widget
        QTest.keyClick(first_row, Qt.Key.Key_Down)
        self.assertEqual(self.window.single_list.currentItem().data(Qt.ItemDataRole.UserRole + 1), "Beta")
        self.assertIs(self.window.focusWidget(), self.window.single_list.currentItem().row_widget)

    def test_multiple_list_exposes_only_the_current_book_to_accessibility(self):
        self.add_book("Alpha")
        self.add_book("Beta")
        self.window.refresh_books()
        self.window.set_view_mode("multiple")
        view = self.window.groups[0]
        self.assertEqual(self.window.single_list.count(), 0)
        pending_names = []

        def collect_names(interface):
            for index in range(interface.childCount()):
                child = interface.child(index)
                name = child.text(QAccessible.Text.Name)
                if name:
                    pending_names.append(name)
                if child.object():
                    collect_names(QAccessible.queryAccessibleInterface(child.object()))

        collect_names(QAccessible.queryAccessibleInterface(self.window.single_list))
        self.assertNotIn("Alpha. No author", pending_names)
        self.assertEqual(QAccessible.queryAccessibleInterface(view.accessible_row).role(), QAccessible.Role.StaticText)
        self.assertEqual(QAccessible.queryAccessibleInterface(view.column_canvas).childCount(), 0)
        self.assertIs(view.accessible_row.parentWidget(), self.window.column_scroll.widget())
        QTest.keyClick(view.accessible_row, Qt.Key.Key_Down)
        self.assertEqual(view.currentItem().data(Qt.ItemDataRole.UserRole + 1), "Beta")
        self.assertIs(self.window.focusWidget(), view.accessible_row)
        self.assertEqual(view.accessible_row.accessibleName(), "Beta. No author")

    def test_multiple_list_exposes_only_the_active_column(self):
        self.add_book("Alpha")
        self.add_book("Lima")
        self.window.refresh_books()
        self.window.set_view_mode("multiple")
        old_column = self.window.groups[self.window.active_column]
        QTest.keyClick(old_column.accessible_row, Qt.Key.Key_Right)
        active_column = self.window.groups[self.window.active_column]
        self.assertNotEqual(old_column, active_column)
        old_canvas = QAccessible.queryAccessibleInterface(old_column.column_canvas)
        self.assertEqual(old_canvas.childCount(), 0)
        self.assertIsNone(old_column.currentItem().row_widget)
        self.assertEqual(QAccessible.queryAccessibleInterface(active_column.column_canvas).childCount(), 0)
        self.assertIs(old_column.accessible_row, active_column.accessible_row)
        self.assertIs(active_column.accessible_row.book_list, active_column)

    def test_switching_populated_columns_keeps_native_focus_on_one_control(self):
        self.add_book("Alpha")
        self.add_book("Lima")
        self.add_book("Maple")
        self.window.set_view_mode("multiple")
        self.window.show()
        self.app.processEvents()
        self.window.groups[0].setFocus()
        self.app.processEvents()
        row = self.window.groups[0].accessible_row
        changes = []

        def focused(old, new):
            changes.append((old, new))

        self.app.focusChanged.connect(focused)
        try:
            for key, name in ((Qt.Key.Key_Right, "G-L. Lima. No author"),
                              (Qt.Key.Key_Right, "M-R. Maple. No author"),
                              (Qt.Key.Key_Left, "G-L. Lima. No author")):
                QTest.keyClick(row, key)
                self.app.processEvents()
                self.assertIs(self.app.focusWidget(), row)
                self.assertEqual(row.accessibleName(), name)
                self.assertEqual(row.text(), "")
                interface = QAccessible.queryAccessibleInterface(row)
                self.assertEqual(interface.text(QAccessible.Text.Name), name)
                self.assertTrue(interface.state().focusable)
                self.assertTrue(interface.state().focused)
            self.assertEqual(changes, [])
        finally:
            self.app.focusChanged.disconnect(focused)

    def test_shared_focus_routes_metadata_and_selection_to_destination_book(self):
        self.add_book("Alpha", "First Author")
        destination_id = self.add_book("Lima", "Second Author")
        self.window.set_view_mode("multiple")
        row = self.window.groups[0].accessible_row
        QTest.keyClick(row, Qt.Key.Key_Right)
        self.assertEqual(self.window.selected_row()[0], destination_id)
        QTest.keyClick(row, Qt.Key.Key_2, Qt.KeyboardModifier.AltModifier)
        self.assertEqual(self.window.statusBar().currentMessage(), "Author: Second Author")
        QTest.keyClick(row, Qt.Key.Key_Space, Qt.KeyboardModifier.ControlModifier)
        self.assertEqual(self.window.marked_ids, {destination_id})
        interface = QAccessible.queryAccessibleInterface(row)
        self.assertIn("Selected", interface.text(QAccessible.Text.Name) + interface.text(QAccessible.Text.Description))

    def test_column_refresh_does_not_repeat_name_notifications(self):
        self.add_book("Alpha")
        self.add_book("Beta")
        self.add_book("Lima")
        self.window.set_view_mode("multiple")
        view = self.window.groups[0]
        with patch.object(view.accessible_row, "setAccessibleName", wraps=view.accessible_row.setAccessibleName) as name:
            with patch("qt_library_manager.QAccessible.updateAccessibility") as notify:
                view.update_accessible_row()
                view.setFocus()
                name.assert_not_called()
                view.setCurrentRow(1)
                view.setFocus()
                name.assert_called_once_with("Beta. No author")
                notify.assert_not_called()

    def test_column_switch_moves_focus_before_hiding_previous_row(self):
        self.add_book("Alpha")
        self.add_book("Lima")
        self.window.set_view_mode("multiple")
        previous = self.window.groups[0]
        target = self.window.groups[1]
        original = previous.set_accessibility_active

        def deactivate(active):
            self.assertIs(self.window.focusWidget(), target.accessible_row)
            original(active)

        with patch.object(previous, "set_accessibility_active", side_effect=deactivate):
            self.window.move_to_column(1)
        self.assertFalse(previous.accessibility_active)
        self.assertEqual(target.accessible_row.accessibleName(), "G-L. Lima. No author")

    def test_horizontal_arrows_expose_one_combined_name_and_no_duplicate_text(self):
        self.add_book("Alpha")
        self.add_book("Lima")
        self.add_book("Maple")
        self.add_book("Moon")
        self.window.set_view_mode("multiple")
        with patch.object(self.window, "announce_selection") as announce:
            QTest.keyClick(self.window.groups[0].accessible_row, Qt.Key.Key_Right)
            view = self.window.groups[1]
            self.assertIs(self.window.focusWidget(), view.accessible_row)
            self.assertEqual(view.heading.accessibleName(), "")
            self.assertEqual(view.accessible_row.accessibleName(), "G-L. Lima. No author")
            self.assertEqual(view.accessible_row.text(), "")
            self.assertEqual(view.heading.text(), "")
            self.assertFalse(view.accessible_row.isHidden())
            self.assertIs(self.window.book_list_for_focus(), view)
            self.assertEqual(self.window.statusBar().currentMessage(), "")
            announce.assert_not_called()
        QTest.keyClick(view.accessible_row, Qt.Key.Key_Right)
        view = self.window.groups[2]
        self.assertIs(self.window.focusWidget(), view.accessible_row)
        self.assertEqual(view.accessible_row.accessibleName(), "M-R. Maple. No author")
        QTest.keyClick(view.accessible_row, Qt.Key.Key_Down)
        self.assertIs(self.window.focusWidget(), view.accessible_row)
        self.assertFalse(view.accessible_row.isHidden())
        self.assertEqual(view.currentItem().data(Qt.ItemDataRole.UserRole + 1), "Moon")
        self.assertEqual(view.accessible_row.accessibleName(), "Moon. No author")

    def test_empty_column_is_skipped_with_arrows(self):
        self.add_book("Alpha")
        self.add_book("Maple")
        self.window.set_view_mode("multiple")
        QTest.keyClick(self.window.groups[0].accessible_row, Qt.Key.Key_Right)
        row = self.window.groups[2].accessible_row
        self.assertIs(self.window.focusWidget(), row)
        self.assertEqual(self.window.active_column, 2)
        self.assertIn("Maple", row.accessibleName())
        QTest.keyClick(row, Qt.Key.Key_Right)
        self.assertIs(self.window.focusWidget(), self.window.groups[2].accessible_row)
        self.assertEqual(self.window.active_column, 2)
        QTest.keyClick(row, Qt.Key.Key_Left)
        self.assertEqual(self.window.active_column, 0)

    def test_empty_categories_are_removed_from_the_window_in_every_grouped_view(self):
        self.add_book("Alpha")
        self.add_book("Maple")
        container = self.window.columns_container
        for mode in ("multiple", "letters", "tags"):
            for layout in ("columns", "rows"):
                with self.subTest(mode=mode, layout=layout):
                    self.window.set_category_layout(layout)
                    self.window.set_view_mode(mode)
                    for group, view in zip(self.window.group_columns, self.window.groups):
                        self.assertEqual(group.isVisibleTo(container), view.count() > 0)
                    shown = [view for group, view in zip(self.window.group_columns, self.window.groups)
                             if group.isVisibleTo(container)]
                    self.assertTrue(shown)
                    self.assertTrue(all(view.count() for view in shown))

    def test_empty_category_never_takes_focus_or_speaks(self):
        self.add_book("Alpha")
        self.add_book("Maple")
        self.window.set_view_mode("multiple")
        row = self.window.groups[0].accessible_row
        empty = self.window.groups[1]
        self.assertEqual(empty.count(), 0)
        empty.setFocus()
        self.app.processEvents()
        # Focus went to the nearest category that has books; the shared row never went blank.
        self.assertIs(self.window.groups[self.window.active_column], self.window.nearest_populated_view(empty))
        self.assertGreater(self.window.groups[self.window.active_column].count(), 0)
        self.assertNotEqual(row.accessibleName(), "")
        self.assertIs(self.window.focusWidget(), row)

    def test_category_appears_when_its_first_book_is_added_and_disappears_when_emptied(self):
        self.add_book("Alpha")
        self.window.set_view_mode("multiple")
        container = self.window.columns_container
        self.assertFalse(self.window.group_columns[2].isVisibleTo(container))
        maple = self.add_book("Maple")
        self.window.refresh_books(maple)
        self.assertTrue(self.window.group_columns[2].isVisibleTo(container))
        self.window.db.delete_book(maple)
        self.window.refresh_books()
        self.assertFalse(self.window.group_columns[2].isVisibleTo(container))
        self.assertEqual(self.window.groups[self.window.active_column].count(), 1)

    def test_empty_library_shows_no_categories_and_keeps_focus_out_of_them(self):
        self.window.set_view_mode("letters")
        container = self.window.columns_container
        self.assertFalse(any(group.isVisibleTo(container) for group in self.window.group_columns))
        for view in self.window.groups:
            view.setFocus()
        self.assertFalse(any(view.accessible_row.isVisible() for view in self.window.groups[:1]))

    def test_category_moves_keep_one_focus_target_and_one_name_update(self):
        from PySide6.QtTest import QSignalSpy
        self.add_book("Alpha")
        self.add_book("Maple")
        self.add_book("Zulu")
        self.window.show()
        self.window.activateWindow()
        for orientation in ("columns", "rows"):
            self.window.set_view_mode("multiple")
            self.window.set_category_layout(orientation)
            self.window.move_to_column(0)
            self.app.processEvents()
            row = self.window.groups[0].accessible_row
            focus_changes = QSignalSpy(self.app.focusChanged)
            with patch.object(row, "setAccessibleName", wraps=row.setAccessibleName) as names, \
                    patch.object(row, "hide", wraps=row.hide) as hide:
                for index in (2, 3, 2, 0):
                    self.window.move_to_column(index)
                    self.app.processEvents()
                    self.assertIs(self.window.focusWidget(), row)
                    self.assertIs(row.book_list, self.window.groups[index])
                    self.assertEqual(row.text(), "")
                    self.assertEqual(self.window.groups[index].heading.accessibleName(), "")
                    self.window.groups[index].update_accessible_row()
                self.assertEqual(names.call_count, 4)
                hide.assert_not_called()
            self.assertEqual(focus_changes.count(), 0)

    def test_marked_book_is_announced_when_focus_returns(self):
        self.add_book("Alpha")
        self.add_book("Beta")
        self.window.refresh_books()
        self.window.single_list.setCurrentRow(0)
        self.window.toggle_marked(self.window.single_list.currentItem())
        self.window.single_list.setCurrentRow(1)
        self.window.single_list.setCurrentRow(0)
        self.assertTrue(self.window.statusBar().currentMessage().startswith("Selected:"))

    def test_alt_number_reads_metadata_once_and_double_press_edits(self):
        self.add_book("Alpha", "Author A")
        self.window.refresh_books()
        row = self.window.single_list.currentItem().row_widget
        self.window.single_list.setFocus()

        with patch.object(self.window, "edit_book") as edit_book:
            QTest.keyClick(row, Qt.Key.Key_1, Qt.KeyboardModifier.AltModifier)
            self.assertEqual(self.window.statusBar().currentMessage(), "Title: Alpha")
            edit_book.assert_not_called()

            QTest.keyClick(row, Qt.Key.Key_1, Qt.KeyboardModifier.AltModifier)
            edit_book.assert_called_once_with(initial_focus_field="title")

    def test_alt_number_shortcuts_match_metadata_field_order(self):
        book_id = self.add_book("Alpha", "Author A")
        self.window.db.update_extra_fields(book_id, "2nd edition", "2010", "9780000000000", "Example Press")
        self.window.refresh_books(book_id)
        row = self.window.single_list.currentItem().row_widget

        for key, label, value in ((Qt.Key.Key_2, "Author", "Author A"), (Qt.Key.Key_5, "Edition", "2nd edition"), (Qt.Key.Key_0, "Tags", "Not specified")):
            with self.subTest(key=key):
                QTest.keyClick(row, key, Qt.KeyboardModifier.AltModifier)
                self.assertEqual(self.window.statusBar().currentMessage(), f"{label}: {value}")

    def test_single_book_action_uses_current_item_not_marked_selection(self):
        marked_id = self.add_book("Alpha", "Author A")
        current_id = self.add_book("Zulu", "Author Z")
        self.window.refresh_books()
        self.window.single_list.setCurrentRow(0)
        self.window.single_list.setFocus()
        QTest.keyClick(self.window.single_list.currentItem().row_widget, Qt.Key.Key_Space, Qt.KeyboardModifier.ControlModifier)
        self.window.single_list.setCurrentRow(1)
        self.window.single_list.setFocus()
        self.assertEqual(self.window.current_ids(), [marked_id])
        self.assertEqual(self.window.selected_row()[0], current_id)

    def test_metadata_tab_order_reaches_save_and_cancel_in_both_directions(self):
        dialog = MetadataDialog({"title": "Example", "notes": "Existing notes"})
        try:
            dialog.show()
            self.app.processEvents()
            dialog.inputs["title"].setFocus()
            order = [*dialog.inputs.values(), dialog.save_button, dialog.cancel_button]
            for expected in order[1:] + order[:1]:
                QTest.keyClick(dialog.focusWidget(), Qt.Key.Key_Tab)
                self.assertIs(dialog.focusWidget(), expected)
            for expected in list(reversed(order[1:])) + order[:1]:
                QTest.keyClick(dialog.focusWidget(), Qt.Key.Key_Backtab)
                self.assertIs(dialog.focusWidget(), expected)
            self.assertEqual(dialog.values()["notes"], "Existing notes")
            dialog.inputs["notes"].setFocus()
            QTest.keyClick(dialog.focusWidget(), Qt.Key.Key_Tab)
            QTest.keyClick(dialog.focusWidget(), Qt.Key.Key_Space)
            self.assertEqual(dialog.result(), QDialog.DialogCode.Accepted)
        finally:
            dialog.close()

    def test_metadata_save_shortcuts_work_from_every_field(self):
        for modifier in (Qt.KeyboardModifier.ControlModifier, Qt.KeyboardModifier.AltModifier):
            for field, _label in MetadataDialog.FIELDS:
                with self.subTest(modifier=modifier, field=field):
                    dialog = MetadataDialog({"title": "Example"}, initial_focus_field=field)
                    try:
                        dialog.show()
                        self.app.processEvents()
                        widget = dialog.inputs[field]
                        widget.setFocus()
                        if field == "tags":
                            widget.tag_actions[0].trigger()
                        else:
                            QTest.keyClicks(widget, "edited")
                        QTest.keyClick(widget, Qt.Key.Key_S, modifier)
                        self.assertEqual(dialog.result(), QDialog.DialogCode.Accepted)
                        self.assertIn("Fiction" if field == "tags" else "edited", dialog.values()[field])
                    finally:
                        dialog.close()

    def test_invalid_metadata_keeps_edits_and_focuses_title(self):
        dialog = MetadataDialog({"title": "Before", "author": "Author"})
        dialog.show()
        dialog.inputs["title"].clear()
        dialog.inputs["author"].setText("Edited author")
        with patch("qt_library_manager.QMessageBox.warning"):
            dialog.accept()
        self.assertTrue(dialog.isVisible())
        self.assertIs(dialog.focusWidget(), dialog.inputs["title"])
        self.assertEqual(dialog.inputs["author"].text(), "Edited author")
        dialog.inputs["title"].setText("After")
        dialog.accept()
        self.assertEqual(dialog.result(), QDialog.DialogCode.Accepted)

    def test_bulk_apply_is_reachable_and_accepts_dialog(self):
        from qt_library_manager import BulkMetadataDialog
        dialog = BulkMetadataDialog()
        dialog.show()
        self.app.processEvents()
        dialog.inputs["author"].setText("Updated author")
        dialog.inputs["series"].setFocus()
        QTest.keyClick(dialog.focusWidget(), Qt.Key.Key_Tab)
        self.assertIs(dialog.focusWidget(), dialog.apply_button)
        QTest.keyClick(dialog.apply_button, Qt.Key.Key_Space)
        self.assertEqual(dialog.result(), QDialog.DialogCode.Accepted)
        self.assertFalse(dialog.isVisible())
        self.assertEqual(dialog.values(), {"author": "Updated author"})
        dialog.deleteLater()

    def test_bulk_save_shortcuts_accept_from_all_fields(self):
        from qt_library_manager import BulkMetadataDialog
        for modifier in (Qt.KeyboardModifier.ControlModifier, Qt.KeyboardModifier.AltModifier):
            for key, _ in BulkMetadataDialog.FIELDS:
                with self.subTest(modifier=modifier, field=key):
                    dialog = BulkMetadataDialog()
                    dialog.inputs["author"].setText("Updated author")
                    dialog.show()
                    self.app.processEvents()
                    dialog.inputs[key].setFocus()
                    QTest.keyClick(dialog.inputs[key], Qt.Key.Key_S, modifier)
                    self.assertEqual(dialog.result(), QDialog.DialogCode.Accepted)
                    dialog.deleteLater()

    def test_bulk_editor_excludes_individual_identifiers(self):
        from qt_library_manager import BulkMetadataDialog
        bulk = BulkMetadataDialog()
        individual = MetadataDialog({"title": "Alpha"})
        for field in ("title", "isbn", "series_number"):
            self.assertNotIn(field, bulk.inputs)
            self.assertIn(field, individual.inputs)
        self.assertEqual(set(bulk.inputs), {"author", "source", "tags", "notes", "edition", "year", "publisher", "series"})
        bulk.deleteLater()
        individual.deleteLater()

    def test_bulk_update_preserves_each_books_identifiers(self):
        first = self.add_book("Alpha")
        second = self.add_book("Beta")
        for book_id, isbn, number in ((first, "9780000000001", "1"), (second, "9780000000002", "2")):
            self.window.db.update_extra_fields(book_id, "First", "2010", isbn, "Old Publisher")
            self.window.db.update_series_fields(book_id, "Old Series", number)
        self.assertEqual(self.window.db.update_bulk_metadata_fields([first, second], {
            "title": "Wrong title", "isbn": "Wrong ISBN", "series_number": "99",
        }), 0)
        self.assertEqual(self.window.db.update_bulk_metadata_fields([first, second], {
            "publisher": "Shared Publisher", "series": "Shared Series",
            "isbn": "Wrong ISBN", "series_number": "99",
        }), 2)
        for book_id, title, isbn, number in ((first, "Alpha", "9780000000001", "1"), (second, "Beta", "9780000000002", "2")):
            row = self.window.db.get_book(book_id)
            self.assertEqual((row[1], row[12], row[22]), (title, isbn, number))
            self.assertEqual((row[13], row[21]), ("Shared Publisher", "Shared Series"))

    def test_bulk_apply_saves_marked_books_and_returns_focus(self):
        from qt_library_manager import BulkMetadataDialog
        first = self.add_book("Alpha", "Original author")
        second = self.add_book("Beta", "Original author")
        other = self.add_book("Gamma", "Unchanged author")
        self.window.refresh_books(second)
        self.window.marked_ids = {first, second}

        def apply_edits(dialog):
            dialog.inputs["author"].setText("Updated author")
            dialog.inputs["tags"].tag_actions[0].trigger()
            dialog.inputs["tags"].tag_actions[2].trigger()
            dialog.apply_button.click()
            return dialog.result()

        with patch.object(BulkMetadataDialog, "exec", apply_edits), \
                patch("qt_library_manager.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes), \
                patch("qt_library_manager.QMessageBox.information") as information:
            self.window.bulk_edit_metadata()
            information.assert_not_called()
        self.assertEqual(self.window.statusBar().currentMessage(), "Bulk metadata updated for 2 book(s)")
        for book_id in (first, second):
            row = self.window.db.get_book(book_id)
            self.assertEqual(row[2], "Updated author")
            self.assertEqual(row[4], "Fiction, Textbooks")
        self.assertEqual(self.window.db.get_book(other)[2], "Unchanged author")
        self.assertEqual(self.window.selected_row()[0], second)
        self.assertIs(self.window.focusWidget(), self.window.single_list.currentItem().row_widget)

    def test_metadata_cancel_requires_confirmation_for_unsaved_changes(self):
        dialog = MetadataDialog({"title": "Before"})
        dialog.show()
        dialog.inputs["title"].setText("After")
        with patch("qt_library_manager.DiscardChangesDialog.exec", return_value=QDialog.DialogCode.Rejected):
            dialog.reject()
        self.assertTrue(dialog.isVisible())
        self.assertEqual(dialog.values()["title"], "After")
        with patch("qt_library_manager.DiscardChangesDialog.exec", return_value=QDialog.DialogCode.Accepted):
            dialog.close()
        self.assertFalse(dialog.isVisible())

    def test_preferences_exposes_existing_commands_and_announcement_setting(self):
        dialog = PreferencesDialog(self.window)
        labels = [button.text() for button in dialog.findChildren(QPushButton)]
        self.assertIn("Choose Backup Folder...", labels)
        self.assertNotIn("Use OneDrive Folder...", labels)
        self.assertIn("Set Kindle Email Addresses...", labels)
        self.assertIn("Add Watched Folder...", labels)
        self.assertFalse(dialog.announcements.isChecked())
        dialog.announcements.setChecked(True)
        self.assertNotEqual(self.window.db.get_setting("supplementary_announcements"), "1")
        dialog.apply_settings()
        self.assertEqual(self.window.db.get_setting("supplementary_announcements"), "1")
        dialog.announcements.setChecked(False)
        dialog.apply_settings()
        self.assertEqual(self.window.db.get_setting("supplementary_announcements"), "0")
        dialog.reject()

    def test_message_history_survives_status_clear_without_moving_focus(self):
        self.add_book("Example")
        self.window.refresh_books()
        before = self.window.focusWidget()
        bar = self.window.statusBar()
        bar.showMessage("Book saved")
        bar.clearMessage()
        self.assertIn("Book saved", [message for _, message in bar.messages])
        self.assertIs(self.window.focusWidget(), before)
        self.assertFalse(bar.announcement_timer.isActive())
        for number in range(220):
            bar.showMessage(str(number))
        self.assertEqual(len(bar.messages), 10)
        self.assertEqual(bar.messages[0][1], "210")

    def test_discard_confirmation_has_no_repeated_question_and_safe_default(self):
        dialog = DiscardChangesDialog(self.window)
        dialog.show()
        self.app.processEvents()
        self.assertFalse(dialog.findChildren(QLabel))
        self.assertIs(dialog.focusWidget(), dialog.keep_editing)
        self.assertTrue(dialog.keep_editing.isDefault())
        QTest.keyClick(dialog.keep_editing, Qt.Key.Key_Return)
        self.assertEqual(dialog.result(), QDialog.DialogCode.Rejected)
        dialog = DiscardChangesDialog(self.window)
        dialog.show()
        QTest.keyClick(dialog.keep_editing, Qt.Key.Key_Tab)
        self.assertIs(dialog.focusWidget(), dialog.discard)
        QTest.keyClick(dialog.discard, Qt.Key.Key_Space)
        self.assertEqual(dialog.result(), QDialog.DialogCode.Accepted)

    def test_settings_categories_have_no_scroll_area_or_extra_tabs(self):
        dialog = PreferencesDialog(self.window)
        self.assertFalse(dialog.findChildren(QScrollArea))
        tabs = {dialog.tabs.tabText(i): dialog.tabs.widget(i) for i in range(dialog.tabs.count())}
        self.assertNotIn("File Explorer Integration", tabs)
        self.assertNotIn("Missing Metadata Alert Sound", tabs)
        general = [button.text() for button in tabs["General"].findChildren(QPushButton)]
        self.assertIn("Add Right-Click Command...", general)
        reading = [button.text() for button in tabs["Book List Reading"].findChildren(QPushButton)]
        self.assertIn("Test Sound", reading)
        self.assertEqual(dialog.sound_mode.currentData(), self.window.missing_metadata_sound_mode())
        dialog.show()
        self.app.processEvents()
        dialog.tabs.setFocus()
        QTest.keyClick(dialog.tabs, Qt.Key.Key_Tab)
        self.assertIsInstance(dialog.focusWidget(), QPushButton)
        dialog.close()

    def test_recent_messages_are_ten_arrow_navigable_items(self):
        messages = [("12:00", f"Message {i}") for i in range(15)]
        dialog = RecentMessagesDialog(messages, self.window)
        dialog.show()
        self.app.processEvents()
        self.assertEqual(dialog.history.count(), 10)
        self.assertTrue(dialog.history.currentItem().text().endswith("Message 14"))
        QTest.keyClick(dialog.history, Qt.Key.Key_Down)
        self.assertTrue(dialog.history.currentItem().text().endswith("Message 13"))
        QTest.keyClick(dialog.history, Qt.Key.Key_End)
        self.assertTrue(dialog.history.currentItem().text().endswith("Message 5"))
        QTest.keyClick(dialog.history, Qt.Key.Key_Tab)
        self.assertIsInstance(dialog.focusWidget(), QPushButton)
        dialog.close()
        empty = RecentMessagesDialog([], self.window)
        self.assertEqual(empty.history.item(0).text(), "No messages in this session.")

    def test_copy_cancellation_keeps_completed_files(self):
        first = Path(self.profile.name) / "first.txt"
        second = Path(self.profile.name) / "second.txt"
        first.write_text("first", encoding="utf-8")
        second.write_text("second", encoding="utf-8")
        target1 = Path(self.profile.name) / "copy1.txt"
        target2 = Path(self.profile.name) / "copy2.txt"
        worker = CopyBooksWorker([("First", first, target1, False), ("Second", second, target2, False)])
        results = []
        worker.completed.connect(lambda *args: results.append(args))
        with patch.object(worker, "isInterruptionRequested", side_effect=[False, True]):
            worker.run()
        self.assertEqual(results[0][0], 1)
        self.assertTrue(target1.exists())
        self.assertFalse(target2.exists())

    def test_cancelled_reindex_keeps_existing_search_index(self):
        book = self.add_book("Example")
        self.window.db.index_book_content(book, "searchable phrase")
        worker = ReindexWorker()
        with patch.object(worker, "isInterruptionRequested", return_value=True):
            worker.run()
        self.assertIn(book, self.window.db.search_content("searchable"))

    def test_result_delivery_waits_for_thread_finish_and_cancellation_skips_callback(self):
        import threading
        from unittest.mock import Mock
        release = threading.Event()

        class SlowWorker(QThread):
            completed = Signal(int, list)

            def run(self):
                self.completed.emit(1, [])
                release.wait(2)

        worker = SlowWorker(self.window)
        callback = Mock()
        self.window.progress_dialog = QProgressDialog("Working", "Cancel", 0, 0, self.window)
        self.window.connect_worker_result(worker, callback)
        worker.start()
        QTest.qWait(30)
        self.assertFalse(callback.called)
        self.window.cancel_operation(worker)
        with patch.object(self.window, "routine_feedback") as report:
            release.set()
            for _ in range(100):
                QTest.qWait(10)
                if report.called:
                    break
            self.assertTrue(report.called)
        self.assertFalse(callback.called)
        worker.wait()

    def test_close_waits_responsively_and_suppresses_late_results(self):
        import threading
        from unittest.mock import Mock
        release = threading.Event()

        class SlowWorker(QThread):
            completed = Signal(int, list)

            def run(self):
                release.wait(2)
                self.completed.emit(0, [])

        worker = SlowWorker(self.window)
        callback = Mock()
        self.window.connect_worker_result(worker, callback)
        worker.start()
        QTest.qWait(10)
        with patch("qt_library_manager.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes):
            self.window.close()
        self.assertTrue(worker.isRunning())
        self.assertIsNotNone(self.window.db)
        release.set()
        for _ in range(100):
            QTest.qWait(10)
            if self.window.db is None:
                break
        self.assertIsNone(self.window.db)
        self.assertFalse(callback.called)
        worker.wait()

    def test_metadata_edit_saves_fields_and_allows_manual_source_change(self):
        book_id = self.add_book("Original Title", "Original Author")
        self.window.db.update_book(book_id, "Original Title", "Original Author", "Bookshare", "", "")
        self.window.db.update_extra_fields(book_id, "1st edition", "2010", "9780000000000", "Example Publisher")
        self.window.refresh_books(book_id)

        values = {
            "title": "Corrected Title", "author": "Corrected Author", "source": "Updated Source",
            "tags": "law, textbook", "notes": "Reviewed", "edition": "2nd edition", "year": "2012",
            "isbn": "9780000000001", "publisher": "Corrected Publisher", "series": "Casebooks",
            "series_number": "3", "cover_url": "",
        }

        class AcceptedDialog:
            def __init__(self, *_args):
                pass

            def exec(self):
                return QDialog.DialogCode.Accepted

            def values(self):
                return values

        with patch("qt_library_manager.MetadataDialog", AcceptedDialog):
            self.window.edit_book()

        row = self.window.db.get_book(book_id)
        self.assertEqual(row[1:6], ("Corrected Title", "Corrected Author", "Updated Source", "law, textbook", "Reviewed"))
        self.assertEqual(row[10:14], ("2nd edition", "2012", "9780000000001", "Corrected Publisher"))
        self.assertEqual(row[21:23], ("Casebooks", "3"))
        self.assertEqual(self.window.selected_row()[0], book_id)

    def test_online_metadata_lookup_preserves_book_source(self):
        book_id = self.add_book("Original Title", "Original Author")
        self.window.db.update_book(book_id, "Original Title", "Original Author", "Bookshare", "", "")
        self.window.refresh_books(book_id)
        self.window.pending_lookup_row = self.window.db.get_book(book_id)
        online = {
            "title": "Corrected Title", "author": "Corrected Author", "source": "Google Books",
            "year": "2010", "publisher": "Example Publisher",
        }

        from qt_library_manager import OnlineMetadataDialog

        def review(dialog):
            self.assertEqual(dialog.merge_mode.currentIndex(), 0)
            dialog.merge_mode.setCurrentIndex(1)
            dialog.accept()
            return dialog.result()

        with patch.object(OnlineMetadataDialog, "exec", review), patch("qt_library_manager.QMessageBox.question") as question:
            self.window.metadata_lookup_finished(online, "Google Books", "")
            question.assert_not_called()

        row = self.window.db.get_book(book_id)
        self.assertEqual(row[1:4], ("Corrected Title", "Corrected Author", "Bookshare"))
        self.assertEqual(row[11], "2010")
        self.assertEqual(row[13], "Example Publisher")

    def test_online_review_preserves_manual_edits_when_switching_merge_modes(self):
        from qt_library_manager import OnlineMetadataDialog
        dialog = OnlineMetadataDialog({"title": "Original", "author": "Original author", "source": "Bookshare"},
                                      {"title": "Online", "author": "Online author", "year": "2010", "source": "Service"}, "Service")
        self.assertEqual(dialog.values()["title"], "Original")
        self.assertEqual(dialog.values()["year"], "2010")
        dialog.inputs["author"].setText("My correction")
        dialog.merge_mode.setCurrentIndex(1)
        self.assertEqual(dialog.values()["title"], "Online")
        self.assertEqual(dialog.values()["author"], "My correction")
        self.assertEqual(dialog.values()["source"], "Bookshare")
        dialog.merge_mode.setCurrentIndex(0)
        self.assertEqual(dialog.values()["title"], "Original")
        self.assertEqual(dialog.values()["author"], "My correction")
        dialog.deleteLater()

    def test_online_review_cancel_does_not_update_library(self):
        from qt_library_manager import OnlineMetadataDialog
        book_id = self.add_book("Original", "Author")
        original = self.window.db.get_book(book_id)
        self.window.pending_lookup_row = original
        with patch.object(OnlineMetadataDialog, "exec", return_value=QDialog.DialogCode.Rejected):
            self.window.metadata_lookup_finished({"title": "Changed", "year": "2010"}, "Service", "")
        self.assertEqual(self.window.db.get_book(book_id), original)

    def test_routine_feedback_announces_once_and_preserves_focus_and_history(self):
        self.add_book("Alpha")
        self.window.refresh_books()
        focus = self.window.focusWidget()
        self.window.db.set_setting("supplementary_announcements", "1")
        with patch.object(self.window, "isActiveWindow", return_value=True), \
                patch("qt_library_manager.QApplication.activeModalWidget", return_value=None), \
                patch("qt_library_manager.QAccessible.updateAccessibility") as announce, \
                patch("qt_library_manager.QMessageBox.information") as info:
            self.window.routine_feedback("Backup complete.", "Backup saved to the chosen folder.")
            announce.assert_called_once()
            self.assertEqual(announce.call_args.args[0].message(), "Backup complete.")
            self.assertEqual(announce.call_args.args[0].politeness(), QAccessible.AnnouncementPoliteness.Polite)
            info.assert_not_called()
        self.assertFalse(self.window.statusBar().announcement_timer.isActive())
        self.assertIs(self.window.focusWidget(), focus)
        self.assertEqual(self.window.statusBar().messages[-1][1], "Backup saved to the chosen folder.")

    def test_routine_feedback_does_not_interrupt_another_dialog(self):
        with patch.object(self.window, "isActiveWindow", return_value=True), \
                patch("qt_library_manager.QApplication.activeModalWidget", return_value=object()), \
                patch("qt_library_manager.QAccessible.updateAccessibility") as announce:
            self.window.routine_feedback("Finished.")
            announce.assert_not_called()
        self.assertEqual(self.window.statusBar().messages[-1][1], "Finished.")

    def test_watched_scan_preserves_non_book_focus_and_stays_quiet(self):
        first = self.add_book("Alpha")
        second = self.add_book("Beta")
        self.window.refresh_books(second)
        self.window.search.setFocus()
        with patch("qt_library_manager.QAccessible.updateAccessibility") as announce:
            self.window.watched_scan_finished(1, 0, [], "", False)
            announce.assert_not_called()
        self.assertIs(self.window.focusWidget(), self.window.search)
        self.assertEqual(self.window.selected_row()[0], second)
        self.assertIn("Imported 1", self.window.statusBar().messages[-1][1])

    def test_no_result_messages_do_not_open_dialogs(self):
        with patch("qt_library_manager.QMessageBox.information") as info, \
                patch.object(self.window, "routine_feedback") as feedback:
            self.window.pending_cover_lookup_book_id = None
            self.window.cover_lookup_finished({}, "Service", "")
            self.window.pending_lookup_row = None
            self.window.metadata_lookup_finished({}, "Service", "")
            info.assert_not_called()
            self.assertEqual(feedback.call_count, 2)

    def test_settings_menu_dispatch_keeps_configured_value(self):
        from PySide6.QtWidgets import QMenu
        menu = QMenu(self.window)
        action = self.window.add_action(menu, "Off", lambda mode="off": self.window.set_missing_metadata_sound_mode(mode))
        with patch("qt_library_manager.QMessageBox.information") as info:
            action.trigger()
            info.assert_not_called()
        self.assertEqual(self.window.missing_metadata_sound_mode(), "off")

    def test_success_results_do_not_require_acknowledgement(self):
        with patch("qt_library_manager.QMessageBox.information") as info, \
                patch("qt_library_manager.QMessageBox.warning") as warning, \
                patch.object(self.window, "routine_feedback") as feedback:
            self.window.reindex_finished(2, [])
            self.window.epub_export_finished(2, [])
            self.window.braille_export_finished(2, [], False, [])
            self.window.transfers_finished(2, [])
            self.window.library_backup_finished("Saved backup.", "", False)
            self.window.watched_scan_finished(0, 0, [], "", True)
            self.window.set_missing_metadata_sound_mode("off")
            self.window.set_backup_schedule("on_demand")
            self.window.toggle_calibre_metadata()
            self.assertEqual(feedback.call_count, 9)
            info.assert_not_called()
            warning.assert_not_called()

    def test_partial_results_still_show_warnings(self):
        with patch("qt_library_manager.QMessageBox.warning") as warning:
            self.window.reindex_finished(1, ["Missing book"])
            self.window.epub_export_finished(1, ["Missing book"])
            self.window.transfers_finished(1, ["Missing book"])
            self.window.watched_scan_finished(0, 0, ["Unreadable book"], "", True)
            self.assertEqual(warning.call_count, 4)

    def test_cover_lookup_starts_without_online_confirmation(self):
        book_id = self.add_book("Original Title", "Original Author")
        self.window.refresh_books(book_id)
        with patch("qt_library_manager.OnlineMetadataWorker") as worker_class, \
                patch.object(self.window, "connect_worker_result") as connect, \
                patch("qt_library_manager.QMessageBox.question") as question:
            worker = worker_class.return_value
            worker.isRunning.return_value = False
            self.window.find_cover_image()
            question.assert_not_called()
            worker.start.assert_called_once()
            connect.assert_called_once_with(worker, self.window.cover_lookup_finished)
            self.assertEqual(self.window.pending_cover_lookup_book_id, book_id)
            self.assertEqual(worker_class.call_args.args[0]["title"], "Original Title")
        self.window.progress_dialog.close()
        self.window.progress_dialog = None
        self.window.cover_lookup_worker = None

    def test_cover_lookup_saves_only_cover_url(self):
        book_id = self.add_book("Original Title", "Original Author")
        self.window.refresh_books(book_id)
        self.window.pending_cover_lookup_book_id = book_id
        with patch.object(self.window, "refresh_books"), patch("qt_library_manager.QMessageBox.information"):
            self.window.cover_lookup_finished(
                {"cover_url": "https://covers.example/book.jpg", "title": "Changed Title", "source": "Online"},
                "Open Library", "",
            )
        row = self.window.db.get_book(book_id)
        self.assertEqual(row[1], "Original Title")
        self.assertEqual(row[2], "Original Author")
        self.assertEqual(row[3], "")
        self.assertEqual(row[14], "https://covers.example/book.jpg")

    def test_cloud_backup_choice_sets_safe_backup_subfolder(self):
        with patch("qt_library_manager.QFileDialog.getExistingDirectory", return_value=self.profile.name), \
                patch("qt_library_manager.QMessageBox.information"):
            self.window.choose_cloud_backup_folder("other")
        backup = Path(self.window.db.get_setting("backup_folder"))
        self.assertEqual(backup.name, "Accessible Ebook Library Manager Backups")
        self.assertTrue(backup.is_dir())
        self.assertNotEqual(backup, self.window.db.books_path)

    def test_detected_onedrive_folder_is_offered_for_backup(self):
        with patch.dict(os.environ, {"OneDrive": self.profile.name}), \
                patch("qt_library_manager.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes), \
                patch("qt_library_manager.QMessageBox.information"):
            self.window.choose_cloud_backup_folder("onedrive")
        backup = Path(self.window.db.get_setting("backup_folder"))
        self.assertEqual(backup.parent, Path(self.profile.name))
        self.assertEqual(backup.name, "Accessible Ebook Library Manager Backups")

    def test_cloud_backup_rejects_managed_books_folder(self):
        with patch("qt_library_manager.QFileDialog.getExistingDirectory", return_value=str(self.window.db.books_path)), \
                patch("qt_library_manager.QMessageBox.warning") as warning:
            self.window.choose_cloud_backup_folder("other")
        warning.assert_called_once()
        self.assertEqual(self.window.db.get_setting("backup_folder", ""), "")

    def test_nls_send_detects_drive_and_sends_selected_books(self):
        book_id = self.add_book("NLS Test Book")
        row = self.window.db.get_book(book_id)
        Path(row[8]).write_text("book data", encoding="utf-8")
        self.window.refresh_books(book_id)
        target = Path(self.profile.name) / "NLS eReader"
        target.mkdir()
        with patch("qt_library_manager.ReaderDeviceDetection") as detection, \
                patch("qt_library_manager.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes), \
                patch.object(self.window, "export_rows") as send:
            detection.return_value.likely_nls_ereader_drives.return_value = [(str(target), "NLS eReader", 15)]
            self.window.send_to_nls_ereader()
        self.assertEqual(self.window.db.get_setting("nls_ereader_folder"), str(target))
        send.assert_called_once_with([row], target, send_to="NLS eReader")

    def test_main_menu_exposes_parity_actions_without_duplicate_reading_menu(self):
        def walk(menu):
            names = []
            for action in menu.actions():
                names.append(action.text().replace("&", ""))
                if action.menu():
                    names.extend(walk(action.menu()))
            return names

        names = walk(self.window.menuBar())
        self.assertIn("Open Kindle for PC", names)
        self.assertIn("Find Cover Image Online...", names)
        self.assertIn("Settings...", names)
        settings = walk(self.window.settings_commands)
        self.assertIn("Choose Backup Folder...", settings)
        self.assertNotIn("Use OneDrive Folder...", settings)
        self.assertNotIn("Use Google Drive Folder...", settings)
        self.assertNotIn("Use iCloud Drive Folder...", settings)
        self.assertEqual(settings.count("Book List Reading"), 1)

    def test_open_kindle_has_original_ctrl_k_shortcut(self):
        file_menu = self.window.menuBar().actions()[0].menu()
        action = next(action for action in file_menu.actions() if action.text() == "Open Kindle for PC")
        self.assertEqual(action.shortcut().toString(), "Ctrl+K")

    def test_export_format_dialog_has_labeled_radio_choices(self):
        dialog = ExportFormatDialog(self.window)
        self.assertTrue(dialog.original_format.isChecked())
        self.assertEqual(dialog.original_format.text(), "Keep original format")
        self.assertEqual(dialog.epub_format.text(), "Convert exported copies to EPUB")
        dialog.epub_format.setChecked(True)
        self.assertTrue(dialog.export_as_epub())
        dialog.close()

    def test_braille_export_settings_are_labeled_and_keyboard_reachable(self):
        from braille_export import BrailleOptions
        dialog = ExportFormatDialog(self.window, BrailleOptions(contracted=False, cells_per_line=32))
        dialog.show()
        dialog.brf_format.setChecked(True)
        dialog.brf_format.setFocus()
        self.app.processEvents()
        QTest.keyClick(dialog.brf_format, Qt.Key.Key_Tab)
        self.assertIs(dialog.focusWidget(), dialog.braille_grade)
        for field, name in ((dialog.braille_grade, "Braille code"), (dialog.cells_per_line, "Cells per line"), (dialog.lines_per_page, "Lines per page")):
            self.assertEqual(QAccessible.queryAccessibleInterface(field).text(QAccessible.Text.Name), name)
        QTest.keyClick(dialog.braille_grade, Qt.Key.Key_Tab)
        self.assertTrue(dialog.cells_per_line.hasFocus())
        QTest.keyClick(dialog.cells_per_line, Qt.Key.Key_Tab)
        self.assertTrue(dialog.lines_per_page.hasFocus())
        QTest.keyClick(dialog.lines_per_page, Qt.Key.Key_Tab)
        self.assertIn(dialog.focusWidget().text(), ("Export", "Cancel"))
        self.assertEqual(dialog.braille_options(), BrailleOptions("brf", False, 32, 25))
        dialog.ebraille_format.setChecked(True)
        self.assertEqual(dialog.braille_options().format, "ebraille")
        self.assertFalse(dialog.cells_per_line.isVisible())
        self.assertTrue(dialog.braille_grade.isVisible())
        dialog.original_format.setChecked(True)
        self.assertIsNone(dialog.braille_options())
        self.assertFalse(dialog.braille_grade.isVisible())
        dialog.close()

    def test_braille_worker_exports_multiple_books_and_reports_failed_book(self):
        from braille_export import BrailleOptions
        from unittest.mock import MagicMock
        rows = [self.window.db.get_book(self.add_book(title)) for title in ("Alpha", "Bad Book", "Beta")]
        translated = Path(self.profile.name) / "sample.brf"
        translated.write_bytes(b",BOOK\r\n\f")
        destination = Path(self.profile.name) / "exports"
        engine = MagicMock()
        engine.convert.side_effect = [[translated], ValueError("Unsupported book"), [translated]]
        results = []
        worker = BrailleExportWorker(rows, destination, BrailleOptions())
        worker.completed.connect(lambda *args: results.append(args))
        with patch("qt_library_manager.BrailleEngine") as factory:
            factory.return_value.__enter__.return_value = engine
            worker.run()
        self.assertEqual(results[0][0], 2)
        self.assertIn("Bad Book: Unsupported book", results[0][1])
        self.assertFalse(results[0][2])
        self.assertEqual(sorted(p.name for p in destination.iterdir()), ["Alpha.brf", "Beta.brf"])

    def test_braille_worker_reports_cancellation(self):
        from braille_export import BrailleOptions, ExportCancelled
        results = []
        worker = BrailleExportWorker([], Path(self.profile.name), BrailleOptions())
        worker.completed.connect(lambda *args: results.append(args))
        with patch("qt_library_manager.BrailleEngine") as factory:
            factory.return_value.__enter__.side_effect = ExportCancelled()
            worker.run()
        self.assertEqual(results, [(0, [], True, [])])

    def test_braille_worker_includes_content_and_converter_warnings(self):
        from braille_export import BrailleOptions
        from unittest.mock import MagicMock
        from tests.braille_samples import sample_epub
        row = list(self.window.db.get_book(self.add_book("Review Book")))
        row[8] = str(sample_epub(Path(self.profile.name) / "review.epub"))
        translated = Path(self.profile.name) / "output.brf"
        translated.write_bytes(b"ABC\r\n\f")
        engine = MagicMock()
        engine.convert.return_value = [translated]
        engine.warnings = ["Check table layout"]
        results = []
        worker = BrailleExportWorker([row], Path(self.profile.name) / "exports", BrailleOptions())
        worker.completed.connect(lambda *args: results.append(args))
        with patch("qt_library_manager.BrailleEngine") as factory, patch("qt_library_manager.inspect_conversion_risks", return_value=["Images need review"]):
            factory.return_value.__enter__.return_value = engine
            worker.run()
        self.assertEqual(results[0][:3], (1, [], False))
        self.assertIn("Review Book", results[0][3][0])
        self.assertIn("Images need review", results[0][3][0])
        self.assertIn("Converter warning: Check table layout", results[0][3][0])

    def test_braille_warnings_open_accessible_report(self):
        with patch("qt_library_manager.ConversionReportDialog") as report:
            self.window.braille_export_finished(1, [], False, ["Alpha: check tables"])
        self.assertIn("Alpha: check tables", report.call_args.args[1])
        report.return_value.exec.assert_called_once()

    @unittest.skipUnless(os.environ.get("AELM_TEST_BRAILLE") == "1", "Real converter test is opt-in")
    def test_real_convert_send_preserves_source_and_reports_table(self):
        import zipfile
        from braille_export import BrailleOptions
        from tests.braille_samples import sample_epub
        template = sample_epub(Path(self.profile.name) / "template.epub")
        source = Path(self.profile.name) / "table.epub"
        with zipfile.ZipFile(template) as original, zipfile.ZipFile(source, "w") as output:
            for item in original.infolist():
                data = original.read(item)
                if item.filename == "chapter.xhtml":
                    data = data.replace(b"</body>", b"<table><tr><td>First</td><td>Second</td></tr></table></body>")
                output.writestr(item, data)
        before = source.read_bytes()
        row = [None, "Table Book", "Writer", "", "", "", "", "", str(source)]
        target = Path(self.profile.name) / "simulated-reader"
        target.mkdir()
        worker = BrailleExportWorker([row], target, BrailleOptions())
        results = []
        worker.completed.connect(lambda *args: results.append(args))
        worker.run()
        self.assertEqual(results[0][:3], (1, [], False))
        self.assertIn("1 table", "\n".join(results[0][3]))
        self.assertTrue((target / "Writer - Table Book.brf").is_file())
        self.assertEqual(source.read_bytes(), before)

    def test_folder_send_routes_all_selected_books_to_format_choice(self):
        ids = [self.add_book("Alpha"), self.add_book("Beta")]
        target = Path(self.profile.name) / "Reader"
        target.mkdir()
        self.window.db.set_setting("voice_dream_loader_folder", str(target))
        with patch.object(self.window, "current_ids", return_value=ids), patch.object(self.window, "export_rows") as send:
            self.window.send_to_folder("Voice Dream", "voice_dream_loader_folder")
        send.assert_called_once_with([self.window.db.get_book(i) for i in ids], target, send_to="Voice Dream")

    def test_convert_and_send_braille_uses_selected_destination(self):
        from braille_export import BrailleOptions
        rows = [self.window.db.get_book(self.add_book("Alpha"))]
        target = Path(self.profile.name)
        def choose(dialog):
            dialog.brf_format.setChecked(True)
            self.assertEqual(dialog.windowTitle(), "Send to NLS eReader")
            return QDialog.DialogCode.Accepted
        with patch.object(ExportFormatDialog, "exec", choose), patch.object(self.window, "start_braille_export") as start:
            self.window.export_rows(rows, target, send_to="NLS eReader")
        start.assert_called_once_with(rows, target, BrailleOptions())

    def test_convert_and_send_epub_does_not_need_converter_for_epub_source(self):
        row = list(self.window.db.get_book(self.add_book("Alpha")))
        source = Path(self.profile.name) / "Alpha.epub"
        source.write_bytes(b"fixture")
        row[8] = str(source)
        def choose(dialog):
            dialog.epub_format.setChecked(True)
            return QDialog.DialogCode.Accepted
        with patch.object(ExportFormatDialog, "exec", choose), patch.object(self.window, "start_epub_export") as start:
            self.window.export_rows([row], Path(self.profile.name), send_to="Dolphin EasyReader")
        start.assert_called_once_with([row], Path(self.profile.name), None)

    def test_send_original_and_cancel_still_behave_correctly(self):
        rows = [self.window.db.get_book(self.add_book("Alpha"))]
        with patch.object(ExportFormatDialog, "exec", return_value=QDialog.DialogCode.Accepted), patch.object(self.window, "prepare_transfers", return_value=["transfer"]), patch.object(self.window, "start_transfers") as start:
            self.window.export_rows(rows, Path(self.profile.name), send_to="Voice Dream")
        start.assert_called_once_with(["transfer"], "Sending to Voice Dream")
        with patch.object(ExportFormatDialog, "exec", return_value=QDialog.DialogCode.Rejected), patch.object(self.window, "start_transfers") as start:
            self.window.export_rows(rows, Path(self.profile.name), send_to="Voice Dream")
        start.assert_not_called()

    def test_conversion_report_exposes_text_and_tabs_to_close(self):
        from qt_library_manager import ConversionReportDialog
        dialog = ConversionReportDialog("Conversion Results", "Alpha: tables need review", self.window)
        dialog.show()
        dialog.report.setFocus()
        self.app.processEvents()
        self.assertEqual(dialog.report.toPlainText(), "Alpha: tables need review")
        self.assertEqual(QAccessible.queryAccessibleInterface(dialog.report).text(QAccessible.Text.Name), "Conversion report")
        QTest.keyClick(dialog.report, Qt.Key.Key_Tab)
        self.assertEqual(dialog.focusWidget().text(), "Close")
        QTest.keyClick(dialog.focusWidget(), Qt.Key.Key_Backtab)
        self.assertIs(dialog.focusWidget(), dialog.report)
        dialog.close()

    def test_epub_export_copies_existing_epub_without_overwriting(self):
        book_id = self.add_book("Export Book", "Writer")
        row = self.window.db.get_book(book_id)
        source = Path(row[8])
        source.write_bytes(b"epub-content")
        source = source.with_suffix(".epub")
        source.write_bytes(b"epub-content")
        row = list(row)
        row[8] = str(source)
        destination = Path(self.profile.name) / "export"
        destination.mkdir()
        (destination / "Writer - Export Book.epub").write_bytes(b"existing")
        worker = ExportEpubWorker([tuple(row)], destination, None)
        results = []
        worker.completed.connect(lambda count, skipped: results.append((count, skipped)))
        worker.run()
        self.assertEqual(results, [(1, [])])
        self.assertEqual((destination / "Writer - Export Book.epub").read_bytes(), b"existing")
        self.assertEqual((destination / "Writer - Export Book (2).epub").read_bytes(), b"epub-content")

    def test_book_details_include_accessibility_and_file_information(self):
        book_id = self.add_book("Detailed Book", "Writer")
        self.window.db.update_accessibility_metadata(book_id, {
            "accessibility_summary": "Screen-reader friendly",
            "accessibility_features": "structuralNavigation",
        })
        self.window.refresh_books(book_id)
        with patch("qt_library_manager.QMessageBox.information") as info:
            self.window.show_details()
        details = info.call_args.args[2]
        self.assertIn("EPUB accessibility:", details)
        self.assertIn("Accessibility summary: Screen-reader friendly", details)
        self.assertIn("Format: txt", details)
        self.assertIn("Stored path:", details)

    def test_navigation_status_does_not_include_position_count(self):
        self.add_book("Alpha")
        self.add_book("Beta")
        self.window.refresh_books()
        self.assertEqual(self.window.statusBar().currentMessage(), "")
        self.window.single_list.setCurrentRow(1)
        self.window.single_list.setCurrentRow(0)
        self.assertNotIn(" of ", self.window.statusBar().currentMessage().lower())

    def test_switching_list_layout_keeps_current_book(self):
        self.add_book("Alpha")
        current_id = self.add_book("Zulu")
        self.window.refresh_books()
        self.window.single_list.setCurrentRow(1)
        self.window.single_list.setFocus()
        self.window.set_view_mode("multiple")
        self.assertEqual(self.window.selected_row()[0], current_id)
        self.window.set_view_mode("single")
        self.assertEqual(self.window.selected_row()[0], current_id)

    def test_alphabetical_column_has_no_duplicate_accessible_name(self):
        self.add_book("123 Other Title")
        self.window.refresh_books()
        self.window.set_view_mode("multiple")
        self.assertEqual(self.window.groups[-1].accessibleName(), "")
        self.window.move_to_column(len(self.window.groups) - 1)
        self.assertEqual(self.window.statusBar().currentMessage(), "")
        self.assertEqual(self.window.focusWidget().accessibleName(), "Other. 123 Other Title. No author")

    def test_alt_number_reads_field_without_repeating_book_title(self):
        self.add_book("Alpha", "Author A")
        self.window.refresh_books()
        self.window.single_list.setFocus()
        with patch("qt_library_manager.QAccessible.updateAccessibility") as announce:
            QTest.keyClick(self.window.single_list.currentItem().row_widget, Qt.Key.Key_2, Qt.KeyboardModifier.AltModifier)
        self.assertEqual(self.window.statusBar().currentMessage(), "Author: Author A")
        event = announce.call_args.args[0]
        self.assertEqual(event.message(), "Author: Author A")
        self.assertEqual(event.type(), QAccessible.Event.Announcement)
        self.assertEqual(event.politeness(), QAccessible.AnnouncementPoliteness.Assertive)

    def test_selection_shortcut_and_announcements_in_every_layout(self):
        first = self.add_book("Alpha", "Author A")
        second = self.add_book("Alpine", "Author B")
        for mode in ("single", "multiple", "letters", "tags"):
            with self.subTest(mode=mode):
                self.window.deselect_all()
                self.window.set_view_mode(mode)
                self.window.refresh_books(first)
                view = self.window.book_list_for_focus()
                with patch("qt_library_manager.QAccessible.updateAccessibility") as announce:
                    QTest.keyClick(view.currentItem().row_widget, Qt.Key.Key_Space, Qt.KeyboardModifier.ControlModifier)
                    self.assertEqual(announce.call_args.args[0].message(), "Selected")
                self.assertEqual(self.window.marked_ids, {first})
                QTest.keyClick(view.currentItem().row_widget, Qt.Key.Key_Down)
                QTest.keyClick(view.currentItem().row_widget, Qt.Key.Key_Space, Qt.KeyboardModifier.ControlModifier)
                self.assertEqual(set(self.window.current_ids()), {first, second})
                QTest.keyClick(view.currentItem().row_widget, Qt.Key.Key_Up)
                self.assertTrue(view.currentItem().row_widget.accessibleName().startswith("Selected. "))
                with patch("qt_library_manager.QAccessible.updateAccessibility") as announce:
                    QTest.keyClick(view.currentItem().row_widget, Qt.Key.Key_Space, Qt.KeyboardModifier.ControlModifier)
                    self.assertEqual(announce.call_args.args[0].message(), "Deselected")
                self.assertEqual(self.window.marked_ids, {second})
                self.assertEqual(self.window.selected_row()[0], first)

    def test_explicit_metadata_reading_ignores_supplementary_setting(self):
        self.add_book("Alpha", "Author A")
        self.window.db.set_setting("supplementary_announcements", "0")
        self.window.refresh_books()
        row = self.window.single_list.currentItem().row_widget
        before = self.window.focusWidget()
        with patch("qt_library_manager.QAccessible.updateAccessibility") as announce:
            QTest.keyClick(row, Qt.Key.Key_2, Qt.KeyboardModifier.AltModifier)
            announce.assert_called_once()
            self.assertEqual(announce.call_args.args[0].message(), "Author: Author A")
        self.assertIs(self.window.focusWidget(), before)

    def test_alt_number_on_different_book_does_not_open_editor(self):
        self.add_book("Alpha", "Author A")
        self.add_book("Beta", "Author B")
        self.window.refresh_books()
        view = self.window.single_list
        with patch.object(self.window, "edit_book") as edit, patch("qt_library_manager.time.monotonic", return_value=10):
            QTest.keyClick(view.currentItem().row_widget, Qt.Key.Key_2, Qt.KeyboardModifier.AltModifier)
            QTest.keyClick(view.currentItem().row_widget, Qt.Key.Key_Down)
            QTest.keyClick(view.currentItem().row_widget, Qt.Key.Key_2, Qt.KeyboardModifier.AltModifier)
            edit.assert_not_called()
            self.assertEqual(self.window.statusBar().currentMessage(), "Author: Author B")

    def test_held_selection_key_does_not_toggle_repeatedly(self):
        from PySide6.QtCore import QEvent
        from PySide6.QtGui import QKeyEvent
        book_id = self.add_book("Alpha")
        self.window.refresh_books()
        row = self.window.single_list.currentItem().row_widget
        QTest.keyClick(row, Qt.Key.Key_Space, Qt.KeyboardModifier.ControlModifier)
        repeated = QKeyEvent(QEvent.Type.KeyPress, Qt.Key.Key_Space, Qt.KeyboardModifier.ControlModifier, " ", True)
        QApplication.sendEvent(row, repeated)
        self.assertEqual(self.window.marked_ids, {book_id})

    def test_folder_import_skips_duplicate_content(self):
        first = os.path.join(self.profile.name, "first.txt")
        duplicate = os.path.join(self.profile.name, "duplicate.txt")
        contents = "A test book by Avery Writer. Unique searchable contents for duplicate detection."
        for path in (first, duplicate):
            with open(path, "w", encoding="utf-8") as file:
                file.write(contents)
        worker = FolderImportWorker([Path(first), Path(duplicate)], "Test Source")
        results = []
        worker.completed.connect(lambda imported, skipped: results.append((imported, skipped)))
        worker.run()
        self.assertEqual(results[0][0], 1)
        self.assertEqual(len(results[0][1]), 1)
        self.assertIn("already in the library", results[0][1][0])

    def test_article_sorting_in_grouped_and_letter_views(self):
        for title in ("Lemon", "The Golden Gate Bridge", "Lake", "A Garden"):
            self.add_book(title)
        self.window.set_view_mode("multiple")
        titles = [self.window.groups[1].item(i).data(Qt.ItemDataRole.UserRole + 1) for i in range(4)]
        self.assertEqual(titles, ["A Garden", "The Golden Gate Bridge", "Lake", "Lemon"])
        self.window.set_view_mode("letters")
        self.assertEqual(len(self.window.groups), 27)
        self.assertEqual(self.window.groups[6].count(), 2)
        self.assertEqual(self.window.groups[11].count(), 2)
        self.window.move_to_column(6)
        QTest.keyClick(self.window.groups[6].accessible_row, Qt.Key.Key_Right)
        self.assertEqual(self.window.active_column, 11)
        self.assertIn("Lake", self.window.groups[11].accessible_row.accessibleName())
        self.assertTrue(all(view.heading.text() == "" for view in self.window.groups))

    def test_tag_columns_use_exact_tags_and_unique_selections(self):
        first = self.add_book("First")
        second = self.add_book("Second")
        third = self.add_book("Third")
        self.window.db.update_book(first, "First", "", "", "Fiction; Textbooks", "")
        self.window.db.update_book(second, "Second", "", "", "Non-fiction", "")
        self.window.set_view_mode("tags")
        from qt_library_manager import COMMON_TAGS
        expected = sorted(COMMON_TAGS + ("Untagged",), key=str.casefold)
        self.assertEqual([label for label, _key in self.window.column_specs], expected)
        self.assertEqual([view.count() for view in self.window.groups],
                         [int(tag in ("Fiction", "Non-fiction", "Textbooks", "Untagged")) for tag in expected])
        self.window.select_all()
        self.assertEqual(set(self.window.current_ids()), {first, second, third})
        self.assertEqual(len(self.window.current_ids()), 3)
        textbooks = expected.index("Textbooks")
        self.window.move_to_column(textbooks)
        self.window.refresh_books(first)
        self.assertEqual(self.window.active_column, textbooks)
        self.window.set_category_layout("rows")
        self.assertEqual([label for label, _key in self.window.column_specs], expected)
        self.assertEqual(self.window.active_column, textbooks)

    def test_layout_shortcuts_and_organize_menu(self):
        self.add_book("Alpha")
        self.window.refresh_books()
        self.window.show()
        self.window.activateWindow()
        self.app.processEvents()
        for key, mode in ((Qt.Key.Key_3, "letters"), (Qt.Key.Key_4, "tags"), (Qt.Key.Key_1, "single")):
            QTest.keyClick(self.window.book_list_for_focus().currentItem().row_widget, key, Qt.KeyboardModifier.ControlModifier)
            self.assertEqual(self.window.db.get_setting("book_list_display_mode"), mode)
        menus = [action.text().replace("&", "") for action in self.window.menuBar().actions()]
        self.assertNotIn("View", menus)
        self.assertIn("Organize", menus)

    def test_rows_swap_navigation_for_each_grouped_view(self):
        first = self.add_book("Alpha", "Author A")
        second = self.add_book("Alpine", "Author B")
        third = self.add_book("Beta", "Author C")
        fourth = self.add_book("Lima", "Author D")
        for book_id in (first, second):
            self.window.db.update_book(book_id, "Alpha" if book_id == first else "Alpine", "Author", "", "Fiction", "")
        self.window.db.update_book(third, "Beta", "Author C", "", "History", "")
        self.window.set_category_layout("rows")
        for mode, destination in (("multiple", fourth), ("letters", third), ("tags", third)):
            with self.subTest(mode=mode):
                self.window.last_alt_number_key = None
                self.window.deselect_all()
                self.window.view_actions[mode].trigger()
                self.assertEqual(self.window.db.get_setting("book_list_display_mode"), mode)
                self.window.refresh_books(first)
                view = self.window.book_list_for_focus()
                QTest.keyClick(view.currentItem().row_widget, Qt.Key.Key_Right)
                self.assertEqual(self.window.selected_row()[0], second)
                QTest.keyClick(view.currentItem().row_widget, Qt.Key.Key_Left)
                self.assertEqual(self.window.selected_row()[0], first)
                QTest.keyClick(view.currentItem().row_widget, Qt.Key.Key_Down)
                self.assertEqual(self.window.selected_row()[0], destination)
                view = self.window.book_list_for_focus()
                self.assertTrue(view.accessible_row.accessibleName().startswith(self.window.column_specs[self.window.active_column][0] + ". "))
                QTest.keyClick(view.currentItem().row_widget, Qt.Key.Key_Space, Qt.KeyboardModifier.ControlModifier)
                self.assertIn(destination, self.window.marked_ids)
                QTest.keyClick(view.currentItem().row_widget, Qt.Key.Key_2, Qt.KeyboardModifier.AltModifier)
                self.assertIn("Author:", self.window.statusBar().currentMessage())
                QTest.keyClick(view.currentItem().row_widget, Qt.Key.Key_Up)
                self.assertEqual(self.window.selected_row()[0], first)

    def test_rows_skip_empty_categories(self):
        self.add_book("Alpha")
        third = self.add_book("Charlie")
        self.window.set_view_mode("letters")
        self.window.set_category_layout("rows")
        QTest.keyClick(self.window.focusWidget(), Qt.Key.Key_Down)
        self.assertEqual(self.window.active_column, 2)
        self.assertEqual(self.window.selected_row()[0], third)
        QTest.keyClick(self.window.focusWidget(), Qt.Key.Key_Up)
        self.assertEqual(self.window.active_column, 0)

    def test_category_layout_preserves_focus_selection_and_persists(self):
        first = self.add_book("Alpha")
        second = self.add_book("Lima")
        self.window.set_view_mode("multiple")
        self.window.refresh_books(second)
        self.window.marked_ids = {first, second}
        self.window.orientation_actions["rows"].trigger()
        self.assertEqual(self.window.selected_row()[0], second)
        self.assertEqual(self.window.marked_ids, {first, second})
        self.assertTrue(self.window.orientation_actions["rows"].isChecked())
        self.assertFalse(self.window.orientation_actions["columns"].isChecked())
        reopened = QtLibraryWindow()
        try:
            self.assertEqual(reopened.category_layout(), "rows")
            self.assertTrue(all(view.horizontal_books for view in reopened.groups))
            self.assertTrue(reopened.orientation_actions["rows"].isChecked())
        finally:
            reopened.close()
        self.window.orientation_actions["columns"].trigger()
        self.assertEqual(self.window.selected_row()[0], second)
        self.assertEqual(self.window.marked_ids, {first, second})
        self.assertFalse(any(view.horizontal_books for view in self.window.groups))

    def test_rows_do_not_change_single_list_arrow_keys(self):
        self.add_book("Alpha")
        second = self.add_book("Beta")
        self.window.set_category_layout("rows")
        QTest.keyClick(self.window.focusWidget(), Qt.Key.Key_Down)
        self.assertEqual(self.window.selected_row()[0], second)
        self.assertFalse(self.window.single_list.horizontal_books)

    def test_row_geometry_and_scrolling_follow_book_focus(self):
        for number in range(8):
            self.add_book(f"Alpha {number} " + "A long book title with readable wrapping " * 3)
        self.window.resize(1000, 720)
        self.window.show()
        self.window.set_view_mode("multiple")
        self.window.set_category_layout("rows")
        self.app.processEvents()
        view = self.window.groups[0]
        rects = view.column_canvas.item_rects()
        self.assertGreater(rects[1][1].x(), rects[0][1].x())
        self.assertEqual(rects[1][1].y(), rects[0][1].y())
        for (_, first), (_, second) in zip(rects, rects[1:]):
            self.assertFalse(first.intersects(second))
        self.assertGreater(self.window.group_columns[1].y(), self.window.group_columns[0].y())
        QTest.keyClick(view.accessible_row, Qt.Key.Key_End)
        self.app.processEvents()
        self.assertEqual(view.currentRow(), 7)
        self.assertGreater(view.scroll_area.horizontalScrollBar().value(), 0)
        self.assertEqual(QAccessible.queryAccessibleInterface(view.column_canvas).childCount(), 0)

    def test_tag_dropdown_preserves_existing_and_toggles_multiple_tags(self):
        dialog = MetadataDialog({"title": "Alpha", "tags": "My course, Fiction"}, self.window)
        field = dialog.inputs["tags"]
        self.assertEqual(field.text(), "Fiction, My course")
        menu = field.menu()
        action = next(a for a in menu.actions() if a.text() == "Textbooks")
        menu.setActiveAction(action)
        QTest.keyClick(menu, Qt.Key.Key_Space)
        self.assertTrue(action.isChecked())
        self.assertEqual(dialog.values()["tags"], "Fiction, Textbooks, My course")
        self.assertIn("Textbooks", field.accessibleName())
        QTest.keyClick(menu, Qt.Key.Key_Space)
        self.assertFalse(action.isChecked())
        self.assertEqual(dialog.values()["tags"], "Fiction, My course")
        dialog.deleteLater()

    def test_tag_columns_bound_imported_tags_without_losing_books(self):
        first = self.add_book("Legacy")
        self.window.db.update_book(first, "Legacy", "", "", "Rare topic; Another topic; Fiction", "")
        self.window.set_view_mode("tags")
        self.assertEqual(len(self.window.groups), 11)
        counts = {label: view.count() for (label, _key), view in zip(self.window.column_specs, self.window.groups)}
        self.assertEqual(counts["Fiction"], 1)
        self.assertEqual(counts["Other"], 1)
        self.assertEqual(counts["Untagged"], 0)
        self.assertEqual(self.window.db.get_book(first)[4], "Rare topic; Another topic; Fiction")

    def test_tag_menu_exposes_checked_state_and_keeps_focus(self):
        from PySide6.QtCore import QPoint
        dialog = MetadataDialog({"title": "Alpha"})
        dialog.show()
        field = dialog.inputs["tags"]
        field.setFocus()
        menu = field.menu()
        menu.popup(field.mapToGlobal(QPoint(0, field.height())))
        self.app.processEvents()
        menu.setActiveAction(menu.actions()[0])
        QTest.keyClick(menu, Qt.Key.Key_Space)
        self.assertTrue(menu.isVisible())
        interface = QAccessible.queryAccessibleInterface(menu)
        item = interface.child(0)
        self.assertTrue(item.state().checkable)
        self.assertTrue(item.state().checked)
        menu.setActiveAction(menu.actions()[2])
        QTest.keyClick(menu, Qt.Key.Key_Space)
        QTest.keyClick(menu, Qt.Key.Key_Escape)
        self.assertFalse(menu.isVisible())
        self.assertEqual(field.text(), "Fiction, Textbooks")
        QTest.keyClick(field, Qt.Key.Key_Tab)
        self.assertIs(dialog.focusWidget(), dialog.inputs["notes"])
        dialog.accept()
        dialog.deleteLater()

    def test_missing_tags_alert_modes_and_repeat_suppression(self):
        first = self.add_book("Alpha", "Author")
        second = self.add_book("Beta", "Author")
        self.window.db.update_book(second, "Beta", "Author", "", "Fiction", "")
        self.window.db.set_setting("missing_metadata_sound_mode", "off")
        self.window.refresh_books()
        first_item = self.window.single_list.item(0)
        second_item = self.window.single_list.item(1)
        with patch("qt_library_manager.LibraryApp.play_missing_metadata_sound") as sound:
            self.window.db.set_setting("missing_metadata_sound_mode", "tags")
            self.window.announce_missing_metadata_if_needed(first_item)
            self.window.announce_missing_metadata_if_needed(first_item)
            self.assertEqual(sound.call_count, 1)
            self.window.announce_missing_metadata_if_needed(second_item)
            self.assertEqual(sound.call_count, 1)
            self.window.announce_missing_metadata_if_needed(first_item)
            self.assertEqual(sound.call_count, 2)
            self.window.db.set_setting("missing_metadata_sound_mode", "off")
            self.window.announce_missing_metadata_if_needed(first_item)
            self.assertEqual(sound.call_count, 2)
            self.window.db.set_setting("missing_metadata_sound_mode", "author")
            self.window.announce_missing_metadata_if_needed(first_item)
            self.assertEqual(sound.call_count, 3)

    def test_file_explorer_arguments_are_parsed(self):
        paths = command_line_import_paths(["--import", "one.epub", "two.pdf", "--import=three.zip", "--other"])
        self.assertEqual([path.name for path in paths], ["one.epub", "two.pdf", "three.zip"])

    def test_search_returns_focus_to_matching_book(self):
        self.add_book("A Different Title")
        self.add_book("Target Search Result")
        self.window.search.setText("Target Search")
        self.window.search_and_focus()
        self.assertIs(self.window.book_list_for_focus(), self.window.single_list)
        self.assertEqual(self.window.selected_row()[1], "Target Search Result")

    def test_select_all_shortcut_does_not_steal_search_editing(self):
        first = self.add_book("Alpha")
        second = self.add_book("Beta")
        self.window.refresh_books()
        self.window.single_list.setFocus()
        QTest.keyClick(self.window.single_list.currentItem().row_widget, Qt.Key.Key_A, Qt.KeyboardModifier.ControlModifier)
        self.assertEqual(self.window.marked_ids, {first, second})
        self.window.search.setText("Target")
        self.window.search.setFocus()
        QTest.keyClick(self.window.search, Qt.Key.Key_A, Qt.KeyboardModifier.ControlModifier)
        self.assertEqual(self.window.search.selectedText(), "Target")
        self.assertEqual(self.window.marked_ids, {first, second})

    def test_find_book_elsewhere_lists_gutenberg_and_login_services(self):
        from PySide6.QtWidgets import QListWidget
        book_id = self.add_book("Pride and Prejudice", "Jane Austen")
        seen = []

        def capture(dialog):
            seen.extend(dialog.findChild(QListWidget).item(i).text() for i in range(dialog.findChild(QListWidget).count()))
            return 0

        hit = [{"title": "Pride and Prejudice", "author": "Jane Austen", "url": "https://www.gutenberg.org/ebooks/1342"}]
        with patch.object(self.window, "selected_row", return_value=self.window.db.get_book(book_id)), \
                patch("qt_library_manager.search_gutenberg", return_value=hit), \
                patch.object(QDialog, "exec", new=capture):
            self.window.find_book_elsewhere("free")
        self.assertTrue(any(text.startswith("Project Gutenberg, free download") for text in seen))
        self.assertTrue(any(text.startswith("NLS BARD") and "copied" in text for text in seen))
        self.assertTrue(any(text.startswith("Bookshare") for text in seen))

    def test_file_access_filter_separates_owned_and_wishlist_books(self):
        real = Path(self.profile.name) / "Owned Book.txt"
        real.write_text("text", encoding="utf-8")
        owned = self.window.db.add_book("Owned Book", "A Author", "", "", "", str(real), str(real))
        wish = self.window.db.add_book("Wish Book", "B Author", "", "", "", "", "wishlist:goodreads:123")

        def titles():
            return {self.window.db.get_book(i)[1] for i in self.window.single_list_ids()} if hasattr(self.window, "single_list_ids") else {
                self.window.single_list.item(n).data(Qt.ItemDataRole.UserRole + 1) for n in range(self.window.single_list.count())}

        self.window.refresh_books()
        self.assertEqual(titles(), {"Owned Book", "Wish Book"})
        self.window.set_file_access_filter("have")
        self.assertEqual(titles(), {"Owned Book"})
        self.window.set_file_access_filter("missing")
        self.assertEqual(titles(), {"Wish Book"})
        self.window.set_file_access_filter("")
        self.assertEqual(titles(), {"Owned Book", "Wish Book"})
        self.assertIn("No file available", self.window.row_text(self.window.db.search_books("Wish")[0]))
        self.assertNotIn("No file available", self.window.row_text(self.window.db.search_books("Owned")[0]))
        self.assertTrue(owned and wish)

    def test_find_services_choice_is_saved(self):
        from PySide6.QtWidgets import QCheckBox
        def choose(dialog):
            for box in dialog.findChildren(QCheckBox):
                box.setChecked(box.text().startswith(("Bookshare", "Project Gutenberg")))
            return QDialog.DialogCode.Accepted

        with patch.object(QDialog, "exec", new=choose):
            self.window.choose_find_services()
        self.assertEqual(self.window.db.get_setting("find_elsewhere_services", ""), "gutenberg,bookshare")
        self.assertEqual(self.window.db.get_setting("find_commercial_services", ""), "none")


if __name__ == "__main__":
    unittest.main()
