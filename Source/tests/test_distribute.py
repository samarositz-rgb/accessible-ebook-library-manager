"""Unit tests for distribute.load/save + prepare_for_destination.

The Calibre-conversion path is exercised only through error branches so
these tests do not depend on ebook-convert being installed.
"""

import sys
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase, main


PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))
import distribute


class InMemoryDb:
    """Minimal stand-in for the SQLite settings table."""

    def __init__(self):
        self.settings = {}

    def get_setting(self, key, default=""):
        return self.settings.get(key, default)

    def set_setting(self, key, value):
        self.settings[key] = value


class LoadSaveRoundTripTests(TestCase):
    def test_empty_db_returns_empty_list(self):
        self.assertEqual(distribute.load_destinations(InMemoryDb()), [])

    def test_malformed_json_returns_empty_list(self):
        db = InMemoryDb()
        db.set_setting(distribute.SETTING_KEY, "not json {{")
        self.assertEqual(distribute.load_destinations(db), [])

    def test_round_trip_preserves_fields(self):
        db = InMemoryDb()
        original = [
            distribute.Destination(name="A", path="/tmp/a", target_format="epub", post_copy_notice="hi"),
            distribute.Destination(name="B", path="/tmp/b"),
        ]
        distribute.save_destinations(db, original)
        recovered = distribute.load_destinations(db)
        self.assertEqual(len(recovered), 2)
        self.assertEqual(recovered[0].name, "A")
        self.assertEqual(recovered[0].target_format, "epub")
        self.assertEqual(recovered[0].post_copy_notice, "hi")
        self.assertEqual(recovered[1].target_format, "")

    def test_entries_missing_required_fields_are_dropped(self):
        db = InMemoryDb()
        import json
        db.set_setting(distribute.SETTING_KEY, json.dumps([
            {"name": "", "path": "/tmp/a"},          # empty name
            {"name": "A", "path": ""},                # empty path
            {"name": "B", "path": "/tmp/b"},          # valid
            "not a dict",                             # wrong type
        ]))
        result = distribute.load_destinations(db)
        self.assertEqual([d.name for d in result], ["B"])


class NormalizedFormatTests(TestCase):
    def test_strips_dot_and_lowercases(self):
        self.assertEqual(distribute.Destination("n", "/p", ".EPUB").normalized_format(), "epub")

    def test_blank_returns_blank(self):
        self.assertEqual(distribute.Destination("n", "/p", "  ").normalized_format(), "")


class PrepareForDestinationTests(TestCase):
    def test_no_target_format_returns_source(self):
        with TemporaryDirectory() as tmp:
            source = Path(tmp) / "book.epub"
            source.write_bytes(b"epub-body")
            path, temp_dir = distribute.prepare_for_destination(source, "", None)
            self.assertEqual(path, source)
            self.assertIsNone(temp_dir)

    def test_matching_format_returns_source(self):
        with TemporaryDirectory() as tmp:
            source = Path(tmp) / "book.EPUB"
            source.write_bytes(b"epub-body")
            path, temp_dir = distribute.prepare_for_destination(source, "epub", None)
            self.assertEqual(path, source)
            self.assertIsNone(temp_dir)

    def test_conversion_without_calibre_raises(self):
        with TemporaryDirectory() as tmp:
            source = Path(tmp) / "book.pdf"
            source.write_bytes(b"pdf-body")
            with self.assertRaises(RuntimeError) as ctx:
                distribute.prepare_for_destination(source, "epub", None)
            self.assertIn("ebook-convert", str(ctx.exception).lower())


class SendBookErrorsTests(TestCase):
    def test_missing_destination_folder_raises(self):
        with TemporaryDirectory() as tmp:
            source = Path(tmp) / "book.epub"
            source.write_bytes(b"body")
            dest = distribute.Destination(name="X", path=str(Path(tmp) / "does_not_exist"))
            with self.assertRaises(FileNotFoundError):
                distribute.send_book_to_destination(source, dest)

    def test_copies_when_no_conversion_needed(self):
        with TemporaryDirectory() as tmp:
            source = Path(tmp) / "book.epub"
            source.write_bytes(b"body")
            dest_folder = Path(tmp) / "out"
            dest_folder.mkdir()
            dest = distribute.Destination(name="X", path=str(dest_folder))
            written = distribute.send_book_to_destination(source, dest)
            self.assertEqual(written.read_bytes(), b"body")

    def test_existing_file_raises_without_overwrite(self):
        with TemporaryDirectory() as tmp:
            source = Path(tmp) / "book.epub"
            source.write_bytes(b"new")
            dest_folder = Path(tmp) / "out"
            dest_folder.mkdir()
            (dest_folder / "book.epub").write_bytes(b"old")
            dest = distribute.Destination(name="X", path=str(dest_folder))
            with self.assertRaises(FileExistsError):
                distribute.send_book_to_destination(source, dest, overwrite=False)
            distribute.send_book_to_destination(source, dest, overwrite=True)
            self.assertEqual((dest_folder / "book.epub").read_bytes(), b"new")

    def test_post_copy_notice_is_written_once(self):
        with TemporaryDirectory() as tmp:
            source = Path(tmp) / "book.epub"
            source.write_bytes(b"body")
            dest_folder = Path(tmp) / "out"
            dest_folder.mkdir()
            dest = distribute.Destination(
                name="X", path=str(dest_folder), post_copy_notice="Do not delete."
            )
            distribute.send_book_to_destination(source, dest)
            notice_path = dest_folder / "IMPORTANT NOTICE.txt"
            self.assertTrue(notice_path.exists())
            notice_path.write_text("user-modified", encoding="utf-8")
            source2 = Path(tmp) / "book2.epub"
            source2.write_bytes(b"more")
            distribute.send_book_to_destination(source2, dest)
            self.assertEqual(notice_path.read_text(encoding="utf-8"), "user-modified")


if __name__ == "__main__":
    main()
