import os
import shutil
import struct
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import kindle_monitor as km


def make_azw(path: Path, title, author="", asin="", publisher="", isbn="", encrypted=True, cdetype="EBOK"):
    records = []
    for kind, value in ((100, author), (101, publisher), (104, isbn), (113, asin), (501, cdetype)):
        if value:
            data = value.encode("utf-8")
            records.append(struct.pack(">II", kind, 8 + len(data)) + data)
    exth_body = b"".join(records)
    exth = b"EXTH" + struct.pack(">II", 12 + len(exth_body), len(records)) + exth_body
    exth += b"\0" * (-len(exth) % 4)
    header_length = 232
    record = bytearray(16 + header_length)
    struct.pack_into(">H", record, 12, 2 if encrypted else 0)
    record[16:20] = b"MOBI"
    struct.pack_into(">I", record, 20, header_length)
    struct.pack_into(">I", record, 28, 65001)
    name_offset = len(record) + len(exth)
    name = title.encode("utf-8")
    struct.pack_into(">II", record, 84, name_offset, len(name))
    struct.pack_into(">I", record, 128, 0x40)
    record = bytes(record) + exth + name + b"\0\0"
    header = bytearray(78)
    header[60:68] = b"BOOKMOBI"
    struct.pack_into(">H", header, 76, 2)
    first = 78 + 16 + 2
    entries = struct.pack(">I4s", first, b"\0\0\0\0") + struct.pack(">I4s", first + len(record), b"\0\0\0\0") + b"\0\0"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(bytes(header) + entries + record + b"ENCRYPTED BOOK TEXT" * 10)


class FakeDb:
    def __init__(self):
        self.books, self.ids, self.next_id = [], {}, 1

    def all_books_for_duplicate_check(self):
        return [list(b) for b in self.books]

    def find_book_id_by_source_book_id(self, key):
        return self.ids.get(key)

    def set_source_book_id(self, book_id, key):
        self.ids[key] = book_id

    def add_book(self, title, author, source, tags, notes, original, stored):
        row = [None] * 23
        row[:9] = [self.next_id, title, author, source, tags, notes, "", original, stored]
        row[12] = ""
        self.books.append(row)
        self.next_id += 1
        return row[0]

    def update_book(self, book_id, title, author, source, tags, notes):
        for row in self.books:
            if row[0] == book_id:
                row[1:6] = [title, author, source, tags, notes]

    def update_extra_fields(self, book_id, edition="", year="", isbn="", publisher=""):
        for row in self.books:
            if row[0] == book_id:
                row[12], row[13] = isbn, publisher


class KindleMonitorTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="kindle_")
        self.addCleanup(lambda: shutil.rmtree(self.tmp, ignore_errors=True))
        self.root = Path(self.tmp) / "My Kindle Content"

    def test_header_gives_title_author_and_asin_without_book_text(self):
        make_azw(self.root / "B0039GL21I_EBOK" / "B0039GL21I_EBOK.azw", "Dune", "Frank Herbert", "B0039GL21I", "Ace", "9780441172719")
        details = km.read_kindle_header(self.root / "B0039GL21I_EBOK" / "B0039GL21I_EBOK.azw")
        self.assertEqual((details["title"], details["author"], details["asin"], details["protected"]),
                         ("Dune", "Frank Herbert", "B0039GL21I", True))

    def test_non_kindle_files_are_ignored(self):
        self.root.mkdir(parents=True)
        (self.root / "notes.azw").write_bytes(b"not a kindle file at all" * 20)
        self.assertEqual(km.scan_kindle_folder(self.root)["books"], [])

    def test_scan_counts_new_formats_and_skips_personal_documents(self):
        make_azw(self.root / "A_EBOK" / "A_EBOK.azw", "One", "Author", "A")
        make_azw(self.root / "B_EBOK" / "B_EBOK.azw", "A Personal Doc", "", "B", cdetype="PDOC")
        (self.root / "C_EBOK").mkdir()
        (self.root / "C_EBOK" / "C_EBOK.kfx").write_bytes(b"x")
        result = km.scan_kindle_folder(self.root)
        self.assertEqual([b["title"] for b in result["books"]], ["One"])
        self.assertEqual(result["unreadable"], 1)

    def test_online_only_files_are_skipped_not_opened(self):
        make_azw(self.root / "A_EBOK" / "A_EBOK.azw", "One", "Author", "A")
        with mock.patch.object(km, "is_online_only", return_value=True), \
                mock.patch.object(km, "read_kindle_header", side_effect=AssertionError("must not open")):
            result = km.scan_kindle_folder(self.root)
        self.assertEqual((result["books"], result["online_only"]), ([], 1))

    def test_new_books_are_added_once_as_no_file_entries(self):
        make_azw(self.root / "A_EBOK" / "A_EBOK.azw", "Emma", "Jane Austen", "B00A")
        books = km.scan_kindle_folder(self.root)["books"]
        db = FakeDb()
        first = km.add_kindle_books(db, books, lookup=lambda t, a: ["Free copy on Project Gutenberg: https://example/1."])
        self.assertEqual((first["added"], first["free_found"]), (1, 1))
        row = db.books[0]
        self.assertTrue(row[8].startswith("wishlist:kindle:"))
        for tag in ("No File Available", "On Kindle", "Free Copy Found"):
            self.assertIn(tag, row[4])
        self.assertIn("protected by Amazon", row[5])
        again = km.add_kindle_books(db, books)
        self.assertEqual((again["added"], again["already"], len(db.books)), (0, 1, 1))

    def test_existing_goodreads_entry_is_marked_not_duplicated(self):
        db = FakeDb()
        db.add_book("Emma", "Jane Austen", "Goodreads", "No File Available", "", "", "wishlist:goodreads:emma")
        books = [{"asin": "B00A", "title": "Emma (Unabridged)", "author": "Jane Austen", "publisher": "", "isbn": "", "protected": True}]
        result = km.add_kindle_books(db, books)
        self.assertEqual((result["marked"], result["added"], len(db.books)), (1, 0, 1))
        self.assertIn("On Kindle", db.books[0][4])

    def test_free_copy_lookup_needs_matching_title_and_author(self):
        def fetch(url):
            return {"docs": [{"key": "/works/OL1W", "title": "Emma", "author_name": ["Jane Austen"], "ebook_access": "public"},
                             {"key": "/works/OL2W", "title": "Emma Reborn", "author_name": ["Jane Austen"], "ebook_access": "public"}]}
        gutenberg = lambda title, author: [{"title": "Emma", "author": "Austen, Jane", "url": "https://www.gutenberg.org/ebooks/158"}]
        found = km.find_free_copies("Emma", "Jane Austen", fetch, gutenberg)
        self.assertEqual(len(found), 2)
        self.assertIn("gutenberg.org/ebooks/158", found[0])
        self.assertIn("free to read", found[1])
        self.assertEqual(km.find_free_copies("Other Book", "Someone Else", fetch, lambda t, a: []), [])


if __name__ == "__main__":
    unittest.main()
