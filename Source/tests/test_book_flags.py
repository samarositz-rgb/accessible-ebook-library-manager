import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import book_flags
from book_flags import IMAGE_ONLY_PDF, describe_flags, extract_source_book_id, record_book_flags


class FakeDatabase:
    def __init__(self, flags=None, source_id=""):
        self.flags = list(flags or [])
        self.source_id = source_id
        self.flag_writes = 0
        self.source_writes = 0

    def get_content_flags(self, book_id):
        return list(self.flags)

    def set_content_flags(self, book_id, flags):
        self.flags = list(flags)
        self.flag_writes += 1

    def get_source_book_id(self, book_id):
        return self.source_id

    def set_source_book_id(self, book_id, value):
        self.source_id = value
        self.source_writes += 1


def make_epub(path, identifier, scheme=None):
    scheme_attr = f' opf:scheme="{scheme}"' if scheme else ""
    container = (
        '<?xml version="1.0"?><container version="1.0" '
        'xmlns="urn:oasis:names:tc:opendocument:xmlns:container">'
        '<rootfiles><rootfile full-path="content.opf" '
        'media-type="application/oebps-package+xml"/></rootfiles></container>'
    )
    opf = (
        '<?xml version="1.0"?><package xmlns="http://www.idpf.org/2007/opf" version="2.0">'
        '<metadata xmlns:dc="http://purl.org/dc/elements/1.1/" '
        'xmlns:opf="http://www.idpf.org/2007/opf">'
        f"<dc:identifier{scheme_attr}>{identifier}</dc:identifier>"
        "</metadata></package>"
    )
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("META-INF/container.xml", container)
        archive.writestr("content.opf", opf)


class SourceBookIdTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.dir = Path(self.tmp.name)

    def test_bookshare_identifier(self):
        path = self.dir / "a.epub"
        make_epub(path, "bookshare:1234567", "bookshare")
        self.assertEqual(extract_source_book_id(path), "bookshare:1234567")

    def test_learning_ally_identifier_is_uppercased(self):
        path = self.dir / "b.epub"
        make_epub(path, "learningally-ab12345")
        self.assertEqual(extract_source_book_id(path), "learningally:AB12345")

    def test_isbn_urn_identifier(self):
        path = self.dir / "e.epub"
        make_epub(path, "urn:isbn:9780306406157")
        self.assertEqual(extract_source_book_id(path), "isbn:9780306406157")

    def test_unrelated_identifier_is_ignored(self):
        path = self.dir / "c.epub"
        make_epub(path, "publisher-house-id-abc")
        self.assertEqual(extract_source_book_id(path), "")

    def test_bad_file_returns_empty(self):
        path = self.dir / "d.epub"
        path.write_bytes(b"not a zip")
        self.assertEqual(extract_source_book_id(path), "")


class RecordFlagsTests(unittest.TestCase):
    def test_describe_flags_ignores_unknown(self):
        self.assertEqual(len(describe_flags([IMAGE_ONLY_PDF, "mystery"])), 1)

    def test_pdf_gets_flag_when_image_only(self):
        db = FakeDatabase()
        with mock.patch.object(book_flags, "is_pdf_probably_image_only", return_value=True):
            record_book_flags(db, 1, "scan.pdf")
        self.assertEqual(db.flags, [IMAGE_ONLY_PDF])

    def test_pdf_flag_cleared_when_text_found(self):
        db = FakeDatabase(flags=[IMAGE_ONLY_PDF, "other"])
        with mock.patch.object(book_flags, "is_pdf_probably_image_only", return_value=False):
            record_book_flags(db, 1, "ocr.pdf")
        self.assertEqual(db.flags, ["other"])

    def test_unchanged_flags_are_not_rewritten(self):
        db = FakeDatabase(flags=[IMAGE_ONLY_PDF])
        with mock.patch.object(book_flags, "is_pdf_probably_image_only", return_value=True):
            record_book_flags(db, 1, "scan.pdf")
        self.assertEqual(db.flag_writes, 0)

    def test_epub_records_source_id(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "x.epub"
            make_epub(path, "bookshare:7654321", "bookshare")
            db = FakeDatabase()
            record_book_flags(db, 1, path)
        self.assertEqual(db.source_id, "bookshare:7654321")

    def test_database_errors_never_raise(self):
        db = FakeDatabase()
        db.get_content_flags = mock.Mock(side_effect=RuntimeError("boom"))
        record_book_flags(db, 1, "scan.pdf")


if __name__ == "__main__":
    unittest.main()
