"""Unit tests for the language, source-id, and DB helpers added in this pass."""

import sys
import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase, main


PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))
import document_text
import db


OPF_TEMPLATE = """<?xml version="1.0" encoding="UTF-8"?>
<package xmlns="http://www.idpf.org/2007/opf" version="3.0" unique-identifier="pub-id">
  <metadata xmlns:dc="http://purl.org/dc/elements/1.1/">
    <dc:title>Sample</dc:title>
    <dc:creator>Author</dc:creator>
    {identifiers}
  </metadata>
  <manifest/>
  <spine/>
</package>
"""

CONTAINER_XML = """<?xml version="1.0"?>
<container xmlns="urn:oasis:names:tc:opendocument:xmlns:container" version="1.0">
  <rootfiles><rootfile full-path="content.opf" media-type="application/oebps-package+xml"/></rootfiles>
</container>
"""


def make_epub(tmp: Path, identifiers_xml: str) -> Path:
    path = tmp / "book.epub"
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("META-INF/container.xml", CONTAINER_XML)
        zf.writestr("content.opf", OPF_TEMPLATE.format(identifiers=identifiers_xml))
    return path


class LanguageDetectionTests(TestCase):
    def test_english(self):
        text = "The quick brown fox jumps over the lazy dog and the cat is asleep in the sun."
        self.assertEqual(document_text.detect_language_from_text(text), "en")

    def test_german(self):
        text = "Der Hund und die Katze sind im Garten mit dem Ball für ein Spiel von zu Hause."
        self.assertEqual(document_text.detect_language_from_text(text), "de")

    def test_french(self):
        text = "Le chien et le chat sont dans le jardin avec un ballon pour un jeu dans la maison."
        self.assertEqual(document_text.detect_language_from_text(text), "fr")

    def test_defaults_when_empty(self):
        self.assertEqual(document_text.detect_language_from_text(""), "en")


class SourceBookIdTests(TestCase):
    def _import_library_manager(self):
        # Deferred: library_manager pulls in tk. Only import when needed.
        import library_manager
        return library_manager

    def test_extracts_bookshare_id(self):
        lm = self._import_library_manager()
        with TemporaryDirectory() as tmp:
            epub = make_epub(Path(tmp), '<dc:identifier id="pub-id" opf:scheme="bookshare" xmlns:opf="http://www.idpf.org/2007/opf">bookshare-12345</dc:identifier>')
            self.assertEqual(lm.extract_source_book_id(epub), "bookshare:12345")

    def test_extracts_isbn_from_urn(self):
        lm = self._import_library_manager()
        with TemporaryDirectory() as tmp:
            epub = make_epub(Path(tmp), '<dc:identifier id="pub-id">urn:isbn:9780306406157</dc:identifier>')
            self.assertEqual(lm.extract_source_book_id(epub), "isbn:9780306406157")

    def test_no_recognized_scheme_returns_empty(self):
        lm = self._import_library_manager()
        with TemporaryDirectory() as tmp:
            epub = make_epub(Path(tmp), '<dc:identifier id="pub-id">custom-house-id-abc</dc:identifier>')
            self.assertEqual(lm.extract_source_book_id(epub), "")

    def test_missing_epub_returns_empty(self):
        lm = self._import_library_manager()
        self.assertEqual(lm.extract_source_book_id(Path("/no/such/file.epub")), "")


class LibraryDatabaseFlagsTests(TestCase):
    """These allocate a temp APPDATA folder and clean up defensively.

    Windows keeps a file lock on the SQLite db briefly after close, so we use
    shutil.rmtree(ignore_errors=True) rather than TemporaryDirectory's
    strict cleanup.
    """

    def setUp(self):
        import os, shutil, tempfile
        self._tmp = tempfile.mkdtemp(prefix="ebooklib_test_db_")
        self._old_appdata = os.environ.get("APPDATA")
        os.environ["APPDATA"] = self._tmp
        self.addCleanup(lambda: shutil.rmtree(self._tmp, ignore_errors=True))
        if self._old_appdata is not None:
            self.addCleanup(lambda: os.environ.__setitem__("APPDATA", self._old_appdata))

    def test_content_flags_round_trip(self):
        database = db.LibraryDatabase()
        try:
            book_id = database.add_book(
                "T", "A", "S", "tag", "n", "/orig/x.epub",
                str(Path(self._tmp) / "stored_a.epub"),
            )
            self.assertEqual(database.get_content_flags(book_id), [])
            database.set_content_flags(book_id, "image_only_pdf,foo")
            self.assertEqual(database.get_content_flags(book_id), ["image_only_pdf", "foo"])
        finally:
            database.close()

    def test_source_book_id_lookup(self):
        database = db.LibraryDatabase()
        try:
            book_id = database.add_book(
                "T", "A", "S", "tag", "n", "/orig/y.epub",
                str(Path(self._tmp) / "stored_b.epub"),
            )
            self.assertIsNone(database.get_book_by_source_book_id("bookshare:9999"))
            database.set_source_book_id(book_id, "bookshare:9999")
            row = database.get_book_by_source_book_id("bookshare:9999")
            self.assertIsNotNone(row)
            self.assertEqual(row[0], book_id)
        finally:
            database.close()


if __name__ == "__main__":
    main()
