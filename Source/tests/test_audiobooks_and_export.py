import csv
import os
import shutil
import struct
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import audiobook_scan as ab
import goodreads_export as ge
import goodreads_import as gr


def atom(kind: bytes, body: bytes) -> bytes:
    return struct.pack(">I4s", 8 + len(body), kind) + body


def text_item(kind: bytes, value: str) -> bytes:
    return atom(kind, atom(b"data", struct.pack(">II", 1, 0) + value.encode("utf-8")))


def freeform(label: str, value: str) -> bytes:
    return atom(b"----", atom(b"mean", b"\0\0\0\0com.apple.iTunes") + atom(b"name", b"\0\0\0\0" + label.encode())
                + atom(b"data", struct.pack(">II", 1, 0) + value.encode()))


def make_m4b(path: Path, title="", author="", narrator="", year="", asin=""):
    items = b""
    for kind, value in ((b"\xa9nam", title), (b"\xa9ART", author), (b"\xa9nrt", narrator), (b"\xa9day", year)):
        if value:
            items += text_item(kind, value)
    if asin:
        items += freeform("ASIN", asin)
    ilst = atom(b"ilst", items)
    meta = atom(b"meta", b"\0\0\0\0" + ilst)
    moov = atom(b"moov", atom(b"udta", meta))
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(atom(b"ftyp", b"M4B \0\0\0\0") + atom(b"mdat", b"x" * 50) + moov)


class AudiobookScanTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="ab_test_")
        self.addCleanup(lambda: shutil.rmtree(self.tmp, ignore_errors=True))
        self.root = Path(self.tmp) / "Libation"

    def test_reads_embedded_tags(self):
        path = self.root / "Frank Herbert" / "Dune [B002V0QK4C]" / "Dune [B002V0QK4C].m4b"
        make_m4b(path, "Dune", "Frank Herbert", "Scott Brick", "2007-01-01", "B002V0QK4C")
        item = ab.describe_audiobook(path, self.root)
        self.assertEqual((item["title"], item["author"], item["narrator"], item["year"], item["asin"]),
                         ("Dune", "Frank Herbert", "Scott Brick", "2007", "B002V0QK4C"))

    def test_falls_back_to_folder_and_file_names(self):
        path = self.root / "Andy Weir" / "Project Hail Mary [B08G9PRS1K]" / "Project Hail Mary [B08G9PRS1K].m4b"
        path.parent.mkdir(parents=True)
        path.write_bytes(b"not an mp4")
        item = ab.describe_audiobook(path, self.root)
        self.assertEqual((item["title"], item["author"], item["asin"]), ("Project Hail Mary", "Andy Weir", "B08G9PRS1K"))

    def test_scan_finds_only_m4b(self):
        make_m4b(self.root / "A" / "One.m4b", "One", "A")
        (self.root / "A" / "cover.jpg").write_bytes(b"x")
        self.assertEqual(len(ab.scan_audiobook_folder(self.root)), 1)


class DatabaseTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="ab_db_")
        self.addCleanup(lambda: shutil.rmtree(self.tmp, ignore_errors=True))
        old = os.environ.get("APPDATA")
        os.environ["APPDATA"] = self.tmp
        self.addCleanup(lambda: os.environ.__setitem__("APPDATA", old) if old is not None else os.environ.pop("APPDATA", None))
        from library_manager import LibraryDatabase
        self.db = LibraryDatabase()
        self.addCleanup(self.db.connection.close)
        self.root = Path(self.tmp) / "Libation"

    def test_import_links_wishlist_adds_new_and_is_repeatable(self):
        csv_path = Path(self.tmp) / "gr.csv"
        csv_path.write_text(
            "Title,Author,ISBN,ISBN13,My Rating,Date Read,Bookshelves,Exclusive Shelf,My Review\n"
            'Dune,Frank Herbert,"=""""","=""9780441172719""",5,2024/03/02,favorites,read,Great\n',
            encoding="utf-8")
        entries = gr.parse_goodreads_csv(csv_path)
        gr.add_wishlist_entries(self.db, entries)
        make_m4b(self.root / "Frank Herbert" / "Dune.m4b", "Dune", "Frank Herbert", "Scott Brick")
        make_m4b(self.root / "Andy Weir" / "Hail Mary.m4b", "Project Hail Mary", "Andy Weir")
        items = ab.scan_audiobook_folder(self.root)
        first = ab.import_audiobooks(self.db, items)
        self.assertEqual((first["linked"], first["added"]), (1, 1))
        rows = {r[1]: r for r in self.db.all_books_for_duplicate_check()}
        dune = rows["Dune"] if "Dune" in rows else next(r for r in rows.values() if r[1].startswith("Dune"))
        self.assertTrue(dune[8].endswith("Dune.m4b"))
        self.assertIn("Audiobook", dune[4])
        self.assertIn("Goodreads:", dune[5])
        self.assertIn("Narrated by Scott Brick", dune[5])
        second = ab.import_audiobooks(self.db, items)
        self.assertEqual(second["already_linked"], 2)
        self.assertEqual(len(self.db.all_books_for_duplicate_check()), 2)

    def test_delete_never_removes_linked_file(self):
        path = self.root / "A" / "Linked.m4b"
        make_m4b(path, "Linked", "A")
        ab.import_audiobooks(self.db, ab.scan_audiobook_folder(self.root))
        book_id = self.db.all_books_for_duplicate_check()[0][0]
        self.db.delete_book(book_id, delete_file=True)
        self.assertTrue(path.exists())
        self.assertEqual(self.db.all_books_for_duplicate_check(), [])


class ExportTests(unittest.TestCase):
    def row(self, **kw):
        row = [None] * 23
        row[:14] = [1, kw.get("title", "Dune"), kw.get("author", "Frank Herbert"), "", kw.get("tags", ""),
                    kw.get("notes", ""), "", "", kw.get("stored", "dune.epub"), "", "", kw.get("year", "1965"),
                    kw.get("isbn", "9780441172719"), kw.get("publisher", "Ace")]
        return row

    def test_note_round_trip(self):
        entry = {"status": "read", "rating": 4, "date_read": "2024/03/02", "review": "Good book"}
        info = ge.parse_goodreads_note("My note\n" + gr.build_note(entry))
        self.assertEqual((info["status"], info["rating"], info["date_read"], info["review"]),
                         ("read", "4", "2024/03/02", "Good book"))
        self.assertEqual(info["other"], ["My note"])

    def test_wishlist_exports_as_to_read_and_tags_become_shelves(self):
        record = ge.export_row(self.row(stored="wishlist:goodreads:9780441172719", tags="Wishlist, Science Fiction, Audiobook"))
        self.assertEqual(record["Exclusive Shelf"], "to-read")
        self.assertEqual(record["Bookshelves"], "science-fiction")
        self.assertEqual(record["ISBN13"], "9780441172719")

    def test_books_without_status_can_be_left_out(self):
        self.assertIsNone(ge.export_row(self.row(), include_without_status=False))
        self.assertEqual(ge.export_row(self.row(), include_without_status=True)["Exclusive Shelf"], "")

    def test_csv_file_is_readable_by_our_importer(self):
        note = gr.build_note({"status": "read", "rating": 5, "date_read": "2024/03/02", "review": ""})
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "out.csv"
            self.assertEqual(ge.export_goodreads_csv(out, [self.row(notes=note)]), 1)
            back = gr.parse_goodreads_csv(out)
            with open(out, newline="", encoding="utf-8") as handle:
                self.assertEqual(next(csv.DictReader(handle))["My Rating"], "5")
        self.assertEqual((back[0]["rating"], back[0]["status"], back[0]["isbn"]), (5, "read", "9780441172719"))


if __name__ == "__main__":
    unittest.main()
