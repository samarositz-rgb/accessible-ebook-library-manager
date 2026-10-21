import os
import shutil
import struct
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import audiobook_scan as ab
from tests.test_audiobooks_and_export import make_m4b


def id3_frame(frame_id: bytes, text: str) -> bytes:
    body = b"\x03" + text.encode("utf-8")
    return frame_id + struct.pack(">I", len(body)) + b"\0\0" + body


def make_mp3(path: Path, title="", author="", album="", year=""):
    frames = b""
    for frame_id, value in ((b"TIT2", title), (b"TPE1", author), (b"TALB", album), (b"TYER", year)):
        if value:
            frames += id3_frame(frame_id, value)
    size = len(frames)
    syncsafe = bytes([(size >> 21) & 0x7F, (size >> 14) & 0x7F, (size >> 7) & 0x7F, size & 0x7F])
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(b"ID3\x03\x00\x00" + syncsafe + frames + b"\xff\xfb" + b"\0" * 64)


NCC = """<html><head><title>x</title>
<meta name="dc:title" content="The Hobbit" />
<meta name="dc:creator" content="Tolkien, J. R. R." />
<meta name="dc:identifier" content="DB12345" />
<meta name="dc:date" content="1996-05-01" />
</head><body></body></html>"""

LA_NCC = """<html><head>
<meta name="dc:title" content="Silent Spring" />
<meta name="dc:creator" content="Rachel Carson" />
<meta name="dc:publisher" content="Learning Ally" />
<meta name="dc:identifier" content="BR98765" />
</head></html>"""


def make_zip(path: Path, ncc: str, audio_count=2):
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("ncc.html", ncc)
        for number in range(audio_count):
            archive.writestr(f"track{number}.mp3", b"\xff\xfb" + b"\0" * 20)


class AudiobookFormatTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="abf_")
        self.addCleanup(lambda: shutil.rmtree(self.tmp, ignore_errors=True))
        self.root = Path(self.tmp) / "Audio"

    def by_title(self, items):
        return {item["title"]: item for item in items}

    def test_mp3_folder_is_one_book_with_id3_details(self):
        for number in range(1, 4):
            make_mp3(self.root / "Frank Herbert" / "Dune" / f"{number:02d}.mp3", f"Chapter {number}", "Frank Herbert", "Dune", "1965")
        items = ab.scan_audiobook_folder(self.root)
        self.assertEqual(len(items), 1)
        item = items[0]
        self.assertEqual((item["kind"], item["title"], item["author"], item["year"]), ("folder", "Dune", "Frank Herbert", "1965"))
        self.assertTrue(item["path"].endswith("Dune"))

    def test_mp3_folder_without_tags_uses_folder_names(self):
        for number in range(2):
            make_mp3(self.root / "Andy Weir" / "The Martian" / f"{number}.mp3")
        item = ab.scan_audiobook_folder(self.root)[0]
        self.assertEqual((item["title"], item["author"]), ("The Martian", "Andy Weir"))

    def test_disc_subfolders_make_one_book(self):
        for disc in ("CD 1", "CD 2"):
            for number in range(2):
                make_mp3(self.root / "Jane Austen" / "Emma" / disc / f"{number}.mp3")
        items = ab.scan_audiobook_folder(self.root)
        self.assertEqual([(i["title"], i["author"]) for i in items], [("Emma", "Jane Austen")])

    def test_single_m4a_file_is_a_book(self):
        make_m4b(self.root / "A" / "Solo.m4a", "Solo", "A Author")
        item = ab.scan_audiobook_folder(self.root)[0]
        self.assertEqual((item["kind"], item["title"], item["author"]), ("file", "Solo", "A Author"))

    def test_bard_zip(self):
        make_zip(self.root / "Tolkien_DB12345.zip", NCC)
        item = ab.scan_audiobook_folder(self.root)[0]
        self.assertEqual((item["kind"], item["source"], item["number"], item["title"], item["year"]),
                         ("zip", "NLS BARD", "DB12345", "The Hobbit", "1996"))
        self.assertIn("NLS BARD book number DB12345", ab.audiobook_note(item))

    def test_learning_ally_zip(self):
        make_zip(self.root / "silent spring.zip", LA_NCC)
        item = ab.scan_audiobook_folder(self.root)[0]
        self.assertEqual((item["source"], item["title"], item["author"]), ("Learning Ally", "Silent Spring", "Rachel Carson"))

    def test_text_zips_and_unrelated_zips_are_ignored(self):
        self.root.mkdir(parents=True)
        with zipfile.ZipFile(self.root / "book.zip", "w") as archive:
            archive.writestr("book.epub", b"x")
        with zipfile.ZipFile(self.root / "photos.zip", "w") as archive:
            archive.writestr("a.jpg", b"x")
        self.assertEqual(ab.scan_audiobook_folder(self.root), [])

    def test_zip_is_left_untouched_and_folder_opens_at_first_part(self):
        zip_path = self.root / "x_DB555.zip"
        make_zip(zip_path, NCC)
        before = zip_path.read_bytes()
        ab.scan_audiobook_folder(self.root)
        self.assertEqual(zip_path.read_bytes(), before)
        for number in (2, 1):
            make_mp3(self.root / "Book" / f"0{number}.mp3")
        self.assertEqual(ab.first_audio_file(self.root / "Book").name, "01.mp3")


class ImportFormatTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="abi_")
        self.addCleanup(lambda: shutil.rmtree(self.tmp, ignore_errors=True))
        old = os.environ.get("APPDATA")
        os.environ["APPDATA"] = self.tmp
        self.addCleanup(lambda: os.environ.__setitem__("APPDATA", old) if old is not None else os.environ.pop("APPDATA", None))
        from library_manager import LibraryDatabase
        self.db = LibraryDatabase()
        self.addCleanup(self.db.connection.close)
        self.root = Path(self.tmp) / "Audio"

    def test_import_mixed_formats_and_source_labels(self):
        for number in range(2):
            make_mp3(self.root / "Dune" / f"{number}.mp3", album="Dune", author="Frank Herbert")
        make_zip(self.root / "Tolkien_DB12345.zip", NCC)
        make_zip(self.root / "silent.zip", LA_NCC)
        make_m4b(self.root / "Weir" / "Hail Mary.m4a", "Project Hail Mary", "Andy Weir")
        items = ab.scan_audiobook_folder(self.root)
        result = ab.import_audiobooks(self.db, items)
        self.assertEqual(result["added"], 4)
        rows = {r[1]: r for r in self.db.all_books_for_duplicate_check()}
        self.assertEqual(rows["The Hobbit"][3], "NLS BARD")
        self.assertEqual(rows["Silent Spring"][3], "Learning Ally")
        self.assertEqual(rows["Dune"][6], "audio folder")
        self.assertEqual(rows["Project Hail Mary"][6], "m4a")
        self.assertEqual(ab.import_audiobooks(self.db, items)["already_linked"], 4)
        dune_id = rows["Dune"][0]
        self.db.delete_book(dune_id, delete_file=True)
        self.assertTrue((self.root / "Dune" / "0.mp3").exists())


if __name__ == "__main__":
    unittest.main()
