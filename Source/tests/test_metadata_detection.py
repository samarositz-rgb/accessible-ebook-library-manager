import importlib.util
import zipfile
from tempfile import TemporaryDirectory
from pathlib import Path
from unittest import TestCase, main
from unittest.mock import patch


PROJECT_ROOT = Path(__file__).resolve().parents[1]
MODULE_PATH = PROJECT_ROOT / "library_manager.py"
CALIBRE_TOOLS_PATH = PROJECT_ROOT / "calibre_tools.py"

spec = importlib.util.spec_from_file_location("library_manager", MODULE_PATH)
library_manager = importlib.util.module_from_spec(spec)
spec.loader.exec_module(library_manager)

calibre_spec = importlib.util.spec_from_file_location("calibre_tools", CALIBRE_TOOLS_PATH)
calibre_tools = importlib.util.module_from_spec(calibre_spec)
calibre_spec.loader.exec_module(calibre_tools)


def detect_for_pdf(filename, text):
    return detect_for_file(filename, text)


def detect_for_file(filename, text):
    path = Path(filename)
    existing = {
        "title": library_manager.clean_filename_title(path),
        "author": "",
        "source": "",
        "tags": "",
        "notes": "",
    }
    with patch.object(library_manager, "read_text_for_metadata_detection", return_value=text):
        return library_manager.detect_metadata_from_text(path, existing=existing)


class MetadataDetectionTests(TestCase):
    def test_format_never_generates_a_tag(self):
        for extension in ("epub", "pdf", "docx", "mobi"):
            with patch.object(library_manager, "read_text_for_metadata_detection", return_value=""), \
                    patch.object(library_manager, "read_epub_metadata", return_value={}):
                result = library_manager.detect_metadata_from_text(Path("Example." + extension))
                self.assertEqual(result["tags"], "")

    def test_epub_subjects_restricted_to_dropdown_categories(self):
        result = self.epub_metadata('<dc:subject>EPUB</dc:subject><dc:subject>fiction</dc:subject><dc:subject>Invented category</dc:subject>')
        self.assertEqual(result["tags"], "Fiction")

    def test_detection_preserves_existing_library_tags(self):
        with patch.object(library_manager, "read_text_for_metadata_detection", return_value="Subject: EPUB"):
            result = library_manager.detect_metadata_from_text(Path("Example.txt"), {"tags": "My old category"})
        self.assertEqual(result["tags"], "My old category")

    def test_byline_and_catalog_author_do_not_invent_source(self):
        for text in ("The Problem of Evil\nby Peter Vaninwagen\nChapter 1",
                     "Cataloging in Publication Data The problem of evil / Peter van Inwagen. p. cm. ISBN 123"):
            result = library_manager.detect_metadata_from_text_content(text)
            self.assertIn("Peter", result["author"])
            self.assertEqual(result["source"], "")
        self.assertEqual(library_manager.clean_author_value("Byron Smith"), "Byron Smith")

    def test_epub_metadata_reader_preserves_author_paragraph(self):
        with TemporaryDirectory() as folder:
            path = Path(folder) / "test.epub"
            with zipfile.ZipFile(path, "w") as archive:
                archive.writestr("chapter.xhtml", "<html><body><h1>The Problem of Evil</h1><p>by Peter Vaninwagen</p><p>" + "Body text. " * 100 + "</p></body></html>")
            result = library_manager.detect_metadata_from_text(path)
            self.assertEqual(result["author"], "Peter Vaninwagen")

    def epub_metadata(self, fields):
        with TemporaryDirectory() as folder:
            path = Path(folder) / "book.epub"
            with zipfile.ZipFile(path, "w") as archive:
                archive.writestr("META-INF/container.xml", '<container xmlns="urn:oasis:names:tc:opendocument:xmlns:container"><rootfiles><rootfile full-path="book.opf"/></rootfiles></container>')
                archive.writestr("book.opf", '<package xmlns="http://www.idpf.org/2007/opf" xmlns:dc="http://purl.org/dc/elements/1.1/" unique-identifier="ebook"><metadata>' + fields + '</metadata></package>')
            return library_manager.read_epub_metadata(path)

    def test_epub_source_print_isbn_is_not_library_source(self):
        metadata = self.epub_metadata('<dc:title>Is This Anything?</dc:title><dc:source>9781982112691</dc:source><dc:identifier id="ebook">9781982112745</dc:identifier>')
        self.assertEqual(metadata["source"], "")
        self.assertEqual(metadata["isbn"], "9781982112745")

    def test_epub_retains_provider_after_source_isbn(self):
        metadata = self.epub_metadata('<dc:source>urn:isbn:978-1-9821-1269-1</dc:source><dc:source>Bookshare</dc:source><dc:identifier>urn:uuid:1234567890</dc:identifier>')
        self.assertEqual(metadata["source"], "Bookshare")
        self.assertEqual(metadata["isbn"], "")

    def test_epub_identifier_prefers_package_isbn(self):
        metadata = self.epub_metadata('<dc:identifier>9781982112691</dc:identifier><dc:identifier id="ebook">ISBN-13: 978-1-9821-1274-5</dc:identifier>')
        self.assertEqual(metadata["isbn"], "9781982112745")

    def test_metadata_isbn_checks_whole_value_and_checksum(self):
        for value in ("Bookshare 9781982112691", "https://example.com/9781982112691", "9781982112692", "1234567890"):
            self.assertEqual(library_manager.isbn_from_metadata_value(value), "")
        self.assertEqual(library_manager.isbn_from_metadata_value("ISBN: 0-306-40615-2"), "0306406152")
        self.assertEqual(library_manager.isbn_from_metadata_value("0-8044-2957-X"), "080442957X")

    def test_pdf_rejects_abbreviated_embedded_title(self):
        detected = detect_for_pdf(
            "Contract Law Selected Source Materials Annotated.pdf",
            "\n".join([
                "Title: ea8 burton contr source matls 11",
                "CONTRACT LAW:",
                "Selected Source Materials",
                "Annotated",
                "2011 Edition",
                "ISBN 978-0-314-27426-7",
            ]),
        )

        self.assertEqual(detected["title"], "Contract Law Selected Source Materials Annotated")
        self.assertEqual(detected["year"], "2011")
        self.assertEqual(detected["isbn"], "9780314274267")

    def test_pdf_rejects_file_like_embedded_title(self):
        detected = detect_for_pdf(
            "California Criminal Law_CH07-CH13.pdf",
            "\n".join([
                "Title: A012636.pdf",
                "The court considered the issue in 1965.",
            ]),
        )

        self.assertEqual(detected["title"], "California Criminal Law CH07 - CH13")
        self.assertEqual(detected["year"], "")

    def test_legal_notice_fragments_are_not_authors(self):
        bad_authors = [
            "Author: persons licensed to",
            "Author: one of",
            "by the states. The Uniform Commer-",
            "by THE COURT:",
        ]

        for text in bad_authors:
            with self.subTest(text=text):
                detected = detect_for_pdf("Contract Law Selected Source Materials Annotated.pdf", text)
                self.assertEqual(detected["author"], "")

    def test_strong_visible_title_is_preserved(self):
        detected = detect_for_pdf(
            "Civil Procedure A Contemporary Approach.pdf",
            "\n".join([
                "Title: Civil Procedure A Contemporary Approach",
                "Author: Kaboya, Gisele N",
                "A Thomson Reuters business",
                "ISBN 9780314908643",
            ]),
        )

        self.assertEqual(detected["title"], "Civil Procedure A Contemporary Approach")
        self.assertEqual(detected["author"], "Kaboya, Gisele N")
        self.assertEqual(detected["publisher"], "Thomson Reuters")
        self.assertEqual(detected["isbn"], "9780314908643")

    def test_docx_rejects_reprinted_permission_boilerplate_as_title(self):
        detected = detect_for_file(
            "Family Law Cases Text Problems.docx",
            "\n".join([
                "FAMILY LAW: CASES, TEXT, PROBLEMS",
                "Fifth Edition",
                "Ira Mark Ellman",
                "LexisNexis",
                "Loose-leaf ISBN 978-1-4224-7664-2",
                "Excerpted material appearing in this book is reprinted by permission as listed below.",
                "Copyright 1995 by another publisher. Reprinted by permission.",
            ]),
        )

        self.assertEqual(detected["title"], "Family Law: Cases, Text, Problems")
        self.assertEqual(detected["edition"], "Fifth Edition")
        self.assertEqual(detected["isbn"], "9781422476642")
        self.assertEqual(detected["publisher"], "LexisNexis")
        self.assertEqual(detected["year"], "")

    def test_docx_uses_title_page_before_later_rule_headings(self):
        detected = detect_for_file(
            "Professional Responsibility.docx",
            "\n".join([
                "PROFESSIONAL RESPONSIBILITY STANDARDS, RULES & STATUTES",
                "2012-2013 Abridged Edition",
                "Selected and Edited",
                "By",
                "JOHN S. DZIENKOWSKI",
                "west",
                "A Thomson Reuters business",
                "ISBN 9780314281357",
                "2.3 Evaluation for Use by Third Persons",
            ]),
        )

        self.assertEqual(detected["title"], "Professional Responsibility Standards, Rules & Statutes")
        self.assertEqual(detected["author"], "John S. Dzienkowski")
        self.assertEqual(detected["edition"], "2012-2013 Abridged Edition")
        self.assertEqual(detected["year"], "2012")
        self.assertEqual(detected["isbn"], "9780314281357")
        self.assertEqual(detected["publisher"], "Thomson Reuters")

    def test_isbn13_is_preferred_over_isbn10(self):
        text = "\n".join([
            "ISBN 0-314-28135-2",
            "ISBN 978-0-314-28135-7",
        ])

        self.assertEqual(library_manager.extract_isbn_from_text(text), "9780314281357")

    def test_sync_folder_contents_copies_updates_and_removes_extra_files(self):
        with TemporaryDirectory() as temp:
            root = Path(temp)
            source = root / "source"
            destination = root / "destination"
            source.mkdir()
            destination.mkdir()

            (source / "book.txt").write_text("new text", encoding="utf-8")
            (source / "nested").mkdir()
            (source / "nested" / "chapter.txt").write_text("chapter", encoding="utf-8")
            (destination / "old.txt").write_text("remove me", encoding="utf-8")

            library_manager.sync_folder_contents(source, destination)

            self.assertEqual((destination / "book.txt").read_text(encoding="utf-8"), "new text")
            self.assertEqual((destination / "nested" / "chapter.txt").read_text(encoding="utf-8"), "chapter")
            self.assertFalse((destination / "old.txt").exists())

    def test_watched_scan_matches_existing_corrected_metadata_by_filename_title(self):
        class FakeDatabase:
            def all_books_for_duplicate_check(self):
                return [
                    (
                        7,
                        "Family Law: Cases, Text, Problems",
                        "Corrected Author",
                        "Personal",
                        "",
                        "",
                        "docx",
                        r"C:\Originals\family_law_cases_text_problems.docx",
                        r"C:\Library\Corrected Author - Family Law Cases Text Problems.docx",
                        "2026-05-01T00:00:00Z",
                        "Fifth Edition",
                        "",
                        "",
                        "LexisNexis",
                    )
                ]

        app = library_manager.LibraryApp.__new__(library_manager.LibraryApp)
        app.db = FakeDatabase()
        watched_file = Path(r"C:\VoiceDream\Family Law Cases Text Problems.docx")

        with patch.object(library_manager, "read_text_for_metadata_detection", return_value=""):
            row, reason = app.watched_file_matches_existing_metadata(watched_file)

        self.assertIsNotNone(row)
        self.assertEqual(row[0], 7)
        self.assertIn(reason, {"same normalized title", "same filename title"})

    def test_watched_scan_keeps_richer_corrected_metadata_without_isbn(self):
        class FakeDatabase:
            def all_books_for_duplicate_check(self):
                return [
                    (
                        11,
                        "Family Law: Cases, Text, Problems",
                        "Corrected Author",
                        "Bookshare",
                        "textbook, law",
                        "metadata corrected by hand",
                        "docx",
                        r"C:\Originals\Family Law Cases Text Problems.docx",
                        r"C:\Library\Corrected Author - Family Law Cases Text Problems.docx",
                        "2026-05-01T00:00:00Z",
                        "Fifth Edition",
                        "2010",
                        "",
                        "LexisNexis",
                    )
                ]

        app = library_manager.LibraryApp.__new__(library_manager.LibraryApp)
        app.db = FakeDatabase()
        watched_file = Path(r"C:\VoiceDream\Family Law Cases Text Problems Fifth Edition.docx")

        with patch.object(library_manager, "read_text_for_metadata_detection", return_value=""):
            row, reason = app.watched_file_matches_existing_metadata(watched_file)

        self.assertIsNotNone(row)
        self.assertEqual(row[0], 11)
        self.assertIn("richer corrected metadata", reason)

    def test_quickstart_files_are_ignored_for_import(self):
        app = library_manager.LibraryApp.__new__(library_manager.LibraryApp)
        with TemporaryDirectory() as temp:
            folder = Path(temp)
            quickstart = folder / "QuickStart.pdf"
            quickstart.write_text("not a book", encoding="utf-8")
            normal = folder / "Contracts.pdf"
            normal.write_text("book", encoding="utf-8")

            self.assertTrue(app.is_ignored_import_file(quickstart))
            self.assertFalse(app.is_ignored_import_file(normal))

    def test_calibre_opf_metadata_prefers_isbn13(self):
        opf = """<?xml version="1.0" encoding="utf-8"?>
<package xmlns:dc="http://purl.org/dc/elements/1.1/">
  <metadata>
    <dc:title>Kindle Sample Book</dc:title>
    <dc:creator>Jane Author</dc:creator>
    <dc:publisher>Sample Press</dc:publisher>
    <dc:date>2020-04-15</dc:date>
    <dc:subject>Law</dc:subject>
    <dc:subject>Study</dc:subject>
    <dc:identifier opf:scheme="ISBN" xmlns:opf="http://www.idpf.org/2007/opf">0-123-45678-9</dc:identifier>
    <dc:identifier opf:scheme="ISBN" xmlns:opf="http://www.idpf.org/2007/opf">978-0-123-45678-6</dc:identifier>
  </metadata>
</package>"""

        metadata = calibre_tools.parse_calibre_opf_metadata(opf)

        self.assertEqual(metadata["title"], "Kindle Sample Book")
        self.assertEqual(metadata["author"], "Jane Author")
        self.assertEqual(metadata["publisher"], "Sample Press")
        self.assertEqual(metadata["year"], "2020")
        self.assertEqual(metadata["isbn"], "9780123456786")
        self.assertEqual(metadata["tags"], "Law, Study")

    def test_kindle_metadata_uses_calibre_without_opening_calibre_interface(self):
        app = library_manager.LibraryApp.__new__(library_manager.LibraryApp)
        source_path = Path("Kindle Sample Book.azw3")

        with (
            patch.object(library_manager, "read_calibre_metadata", return_value={
                "title": "Kindle Sample Book",
                "author": "Jane Author",
                "publisher": "Sample Press",
                "year": "2020",
            }),
            patch.object(library_manager, "read_text_for_metadata_detection", return_value=""),
        ):
            metadata = app.guess_metadata_from_file(source_path)

        self.assertEqual(metadata["title"], "Kindle Sample Book")
        self.assertEqual(metadata["author"], "Jane Author")
        self.assertEqual(metadata["publisher"], "Sample Press")
        self.assertEqual(metadata["year"], "2020")
        self.assertEqual(metadata["source"], "Kindle")
        self.assertEqual(metadata["tags"], "")


if __name__ == "__main__":
    main()
