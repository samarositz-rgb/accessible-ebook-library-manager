import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import goodreads_import as gr

CSV_TEXT = (
    'Book Id,Title,Author,ISBN,ISBN13,My Rating,Date Read,Bookshelves,Exclusive Shelf,My Review\n'
    '1,"Dune (Dune, #1)",Frank Herbert,"=""0441172717""","=""9780441172719""",5,2024/03/02,"favorites, sci-fi",read,"Great<br/>book"\n'
    '2,The Hobbit,J.R.R. Tolkien,"=""""","=""""",0,,,to-read,\n'
    '3,Unknown Book,Nobody,"=""""","=""""",3,,,read,\n'
)


class FakeDb:
    def __init__(self):
        self.updates = []

    def update_book(self, *args):
        self.updates.append(args)


def library_row(book_id, title, author, tags="", notes="", isbn=""):
    row = [None] * 23
    row[0], row[1], row[2], row[3], row[4], row[5], row[12] = book_id, title, author, "", tags, notes, isbn
    return row


class GoodreadsTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.path = Path(self.tmp.name) / "export.csv"
        self.path.write_text(CSV_TEXT, encoding="utf-8")

    def test_parse_cleans_isbn_and_reads_fields(self):
        rows = gr.parse_goodreads_csv(self.path)
        self.assertEqual(len(rows), 3)
        self.assertEqual(rows[0]["isbn"], "9780441172719")
        self.assertEqual(rows[0]["rating"], 5)
        self.assertEqual(rows[0]["shelves"], ["favorites", "sci-fi", "read"])

    def test_non_goodreads_file_rejected(self):
        bad = Path(self.tmp.name) / "bad.csv"
        bad.write_text("a,b\n1,2\n", encoding="utf-8")
        with self.assertRaises(ValueError):
            gr.parse_goodreads_csv(bad)

    def test_match_by_isbn_then_title(self):
        entries = gr.parse_goodreads_csv(self.path)
        books = [library_row(10, "Dune", "Frank Herbert", isbn="978-0-441-17271-9"),
                 library_row(11, "Hobbit, The", "Tolkien, J. R. R.")]
        self.assertEqual(gr.find_match(entries[0], books)[0], 10)
        self.assertEqual(gr.find_match(entries[1], books)[0], 11)
        self.assertIsNone(gr.find_match(entries[2], books))

    def test_apply_updates_and_keeps_user_notes(self):
        entries = gr.parse_goodreads_csv(self.path)
        books = [library_row(10, "Dune", "Frank Herbert", tags="Fiction", notes="My own note\nGoodreads: old", isbn="9780441172719")]
        db = FakeDb()
        result = gr.apply_goodreads_entries(db, entries, books)
        self.assertEqual(result["updated"], 1)
        self.assertEqual(result["unmatched"], 2)
        _, _, _, _, tags, notes = db.updates[0]
        self.assertEqual(tags, "Fiction, favorites, sci-fi, read")
        self.assertIn("My own note", notes)
        self.assertNotIn("old", notes)
        self.assertIn("rated 5 of 5", notes)
        self.assertIn("Great book", notes)

    def test_second_run_is_unchanged(self):
        entries = gr.parse_goodreads_csv(self.path)[:1]
        books = [library_row(10, "Dune", "Frank Herbert", isbn="9780441172719")]
        db = FakeDb()
        gr.apply_goodreads_entries(db, entries, books)
        books[0][4], books[0][5] = db.updates[0][4], db.updates[0][5]
        result = gr.apply_goodreads_entries(db, entries, books)
        self.assertEqual(result["unchanged"], 1)


class WishlistTests(unittest.TestCase):
    def setUp(self):
        import os
        import shutil
        self.tmp = tempfile.mkdtemp(prefix="gr_wishlist_")
        self.addCleanup(lambda: shutil.rmtree(self.tmp, ignore_errors=True))
        old = os.environ.get("APPDATA")
        os.environ["APPDATA"] = self.tmp
        self.addCleanup(lambda: os.environ.__setitem__("APPDATA", old) if old is not None else os.environ.pop("APPDATA", None))
        from library_manager import LibraryDatabase
        self.db = LibraryDatabase()
        self.addCleanup(self.db.connection.close)
        path = Path(self.tmp) / "export.csv"
        path.write_text(CSV_TEXT, encoding="utf-8")
        self.entries = gr.parse_goodreads_csv(path)

    def test_unmatched_become_wishlist_entries_once(self):
        first = gr.apply_goodreads_entries(self.db, self.entries)
        self.assertEqual(first["unmatched"], 3)
        self.assertEqual(gr.add_wishlist_entries(self.db, first["unmatched_entries"]), 3)
        rows = self.db.all_books_for_duplicate_check()
        self.assertEqual(len(rows), 3)
        dune = next(r for r in rows if r[1].startswith("Dune"))
        self.assertIn("Wishlist", dune[4])
        self.assertIn("favorites", dune[4])
        self.assertTrue(gr.is_wishlist_path(dune[8]))
        self.assertEqual(dune[12], "9780441172719")
        second = gr.apply_goodreads_entries(self.db, self.entries)
        self.assertEqual(second["unmatched"], 0)
        self.assertEqual(len(self.db.all_books_for_duplicate_check()), 3)

    def test_unmatched_file_lists_titles(self):
        out = Path(self.tmp) / "missing.txt"
        gr.write_unmatched_file(out, self.entries)
        text = out.read_text(encoding="utf-8")
        self.assertIn("not in your library: 3", text)
        self.assertIn("Dune (Dune, #1) by Frank Herbert (ISBN 9780441172719, read, rated 5 of 5)", text)


if __name__ == "__main__":
    unittest.main()
