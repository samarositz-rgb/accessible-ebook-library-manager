import csv
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import goodreads_export as ge
import goodreads_import as gr


class DuplicateMatchingTests(unittest.TestCase):
    def match(self, gr_title, gr_author, lib_title, lib_author):
        entry = {"title": gr_title, "author": gr_author, "isbn": "", "isbn10": ""}
        row = [None] * 23
        row[0], row[1], row[2], row[8], row[12] = 1, lib_title, lib_author, "x", ""
        return gr.find_match(entry, [row]) is not None

    def test_store_style_titles_match_goodreads_titles(self):
        self.assertTrue(self.match("The Hunger Games (The Hunger Games, #1)", "Suzanne Collins", "The Hunger Games, Book 1", "Suzanne Collins"))
        self.assertTrue(self.match("Dune (Dune, #1)", "Frank Herbert", "Dune (Unabridged)", "Frank Herbert"))
        self.assertTrue(self.match("Atomic Habits: Tiny Changes, Remarkable Results", "James Clear", "Atomic Habits", "James Clear"))
        self.assertTrue(self.match("Emma", "Jane Austen", "Emma, Unabridged", "Austen, Jane"))

    def test_authors_in_different_order_or_lists_still_match(self):
        self.assertTrue(self.match("Good Omens", "Terry Pratchett", "Good Omens", "Neil Gaiman, Terry Pratchett"))
        self.assertTrue(self.match("The Shining", "Stephen King", "The Shining", "King, Stephen"))
        self.assertTrue(self.match("Harry Potter and the Sorcerer's Stone", "J.K. Rowling", "Harry Potter and the Sorcerer's Stone", "J. K. Rowling"))

    def test_different_books_or_authors_do_not_match(self):
        self.assertFalse(self.match("Dune", "Frank Herbert", "Dune Messiah", "Frank Herbert"))
        self.assertFalse(self.match("It", "Stephen King", "It", "Someone Else"))
        self.assertFalse(self.match("Dune", "John Smith", "Dune", "John Jones"))


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
        record = ge.export_row(self.row(stored="wishlist:goodreads:9780441172719", tags="No File Available, Science Fiction"))
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
