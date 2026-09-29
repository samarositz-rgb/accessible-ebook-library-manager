import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
import zipfile
import xml.etree.ElementTree as ET

from braille_export import (
    BrailleEngine, BrailleOptions, ExportCancelled, prepare_epub,
    publish_files, require_runtime, validate_output,
)
from tests.braille_samples import sample_epub, sample_epub2


class BrailleExportTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.folder = Path(self.temp.name)

    def tearDown(self):
        self.temp.cleanup()

    def test_options_validate_and_choose_ueb_tables(self):
        self.assertEqual(BrailleOptions().code, '(liblouis-table:"en-ueb-g2.ctb")')
        self.assertIn('en-ueb-g1.ctb', BrailleOptions(contracted=False).code)
        for kwargs in ({"format": "txt"}, {"cells_per_line": 0}, {"lines_per_page": 101}, {"cells_per_line": 40.5}, {"contracted": "false"}):
            with self.assertRaises(ValueError):
                BrailleOptions(**kwargs)

    def test_prepare_preserves_original_and_adds_xml_language(self):
        source = sample_epub(self.folder / "input.epub")
        original = source.read_bytes()
        dest = self.folder / "prepared.epub"
        self.assertEqual(prepare_epub(source, dest), "3.0")
        self.assertEqual(source.read_bytes(), original)
        with zipfile.ZipFile(source) as before, zipfile.ZipFile(dest) as after:
            self.assertEqual(before.namelist(), after.namelist())
            for name in ("chapter.xhtml", "nav.xhtml"):
                a, b = ET.fromstring(before.read(name)), ET.fromstring(after.read(name))
                self.assertEqual(''.join(a.itertext()), ''.join(b.itertext()))
                self.assertEqual(b.get('{http://www.w3.org/XML/1998/namespace}lang'), "en")

    def test_unsafe_epub_paths_are_rejected(self):
        source = sample_epub(self.folder / "input.epub")
        with zipfile.ZipFile(source, "a") as archive:
            archive.writestr("../outside.txt", "bad")
        with self.assertRaisesRegex(ValueError, "unsafe"):
            prepare_epub(source, self.folder / "prepared.epub")

    def test_publish_keeps_existing_file_and_exports_all_volumes(self):
        source = self.folder / "sample.brf"
        source.write_bytes(b",BOOK\r\n\f")
        dest = self.folder / "export"
        first = publish_files([source], dest, "Book")
        second = publish_files([source], dest, "Book")
        self.assertEqual(first[0].name, "Book.brf")
        self.assertEqual(second[0].name, "Book (2).brf")
        volumes = publish_files([source, source], dest, "Volumes")
        self.assertEqual([p.name for p in volumes], ["Volumes - Volume 1.brf", "Volumes - Volume 2.brf"])
        self.assertTrue(all(p.read_bytes() == source.read_bytes() for p in first + second + volumes))

    def test_cancel_rolls_back_partial_output(self):
        source = self.folder / "sample.brf"
        source.write_bytes(b"X" * (1024 * 1024 + 1))
        calls = 0
        def cancelled():
            nonlocal calls
            calls += 1
            return calls > 2
        dest = self.folder / "export"
        with self.assertRaises(ExportCancelled):
            publish_files([source], dest, "Book", cancelled)
        self.assertEqual(list(dest.iterdir()), [])
        self.assertEqual(source.stat().st_size, 1024 * 1024 + 1)

    def test_brf_validation_rejects_plain_lowercase_and_long_lines(self):
        source = self.folder / "sample.brf"
        for data in (b"This is not BRF", b"X" * 41, b""):
            source.write_bytes(data)
            with self.assertRaises(RuntimeError):
                validate_output(source, BrailleOptions())

    def test_unsupported_input_is_not_renamed_as_braille(self):
        source = self.folder / "input.pdf"
        source.write_bytes(b"sample")
        with self.assertRaisesRegex(ValueError, "EPUB"):
            BrailleEngine().convert(source, BrailleOptions())

    def test_missing_runtime_is_explained_without_downloading(self):
        with patch("braille_export.runtime_folder", return_value=self.folder):
            with self.assertRaisesRegex(RuntimeError, "complete application folder"):
                require_runtime()

    def test_preparation_checks_cancellation_before_copying(self):
        source = sample_epub(self.folder / "input.epub")
        original = source.read_bytes()
        def cancel():
            raise ExportCancelled()
        with self.assertRaises(ExportCancelled):
            prepare_epub(source, self.folder / "prepared.epub", cancel)
        self.assertEqual(source.read_bytes(), original)

    def test_brf_page_height_is_checked(self):
        path = self.folder / "tall.brf"
        path.write_bytes(b"ABC\r\n" * 26 + b"\f")
        with self.assertRaisesRegex(RuntimeError, "page height"):
            validate_output(path, BrailleOptions())

    def test_ebraille_requires_braille_in_body_not_just_navigation(self):
        path = self.folder / "empty.ebrl"
        with zipfile.ZipFile(path, "w") as z:
            z.writestr("mimetype", "application/epub+zip")
            z.writestr("META-INF/container.xml", '<container><rootfiles><rootfile full-path="book.opf"/></rootfiles></container>')
            z.writestr("book.opf", '<package xmlns:dc="http://purl.org/dc/elements/1.1/"><metadata><dc:format>eBraille 1.0</dc:format></metadata><manifest><item id="chapter" href="chapter.xhtml"/></manifest><spine><itemref idref="chapter"/></spine></package>')
            z.writestr("nav.xhtml", '<html><body>\u2801</body></html>')
            z.writestr("chapter.xhtml", '<html><body/></html>')
        with self.assertRaisesRegex(RuntimeError, "no translated braille"):
            validate_output(path, BrailleOptions("ebraille"))

    def test_previous_conversion_temporary_files_are_released(self):
        engine = BrailleEngine()
        engine.book_temp = tempfile.TemporaryDirectory(dir=self.folder)
        previous = Path(engine.book_temp.name)
        (previous / "old.brf").write_text("ABC")
        source = self.folder / "original.brf"
        source.write_text("DEF")
        self.assertEqual(engine.convert(source, BrailleOptions()), [source])
        self.assertFalse(previous.exists())
        self.assertEqual(source.read_text(), "DEF")

    def test_failure_on_later_volume_rolls_back_only_new_exports(self):
        source = self.folder / "first.brf"
        source.write_text("ABC")
        dest = self.folder / "exports"
        dest.mkdir()
        existing = dest / "Book - Volume 1.brf"
        existing.write_text("KEEP")
        with self.assertRaises(FileNotFoundError):
            publish_files([source, self.folder / "missing.brf"], dest, "Book")
        self.assertEqual(list(dest.iterdir()), [existing])
        self.assertEqual(existing.read_text(), "KEEP")

    def test_brf_volumes_sort_numerically(self):
        source = sample_epub(self.folder / "source.epub")
        engine = BrailleEngine()
        engine.work = self.folder
        output = self.folder / "result"
        output.mkdir()
        for number in (10, 2, 1):
            (output / f"volume-{number}.brf").write_text("ABC\n\f", encoding="ascii")
        try:
            with patch.object(engine, "run_script", return_value=output):
                files = engine.convert(source, BrailleOptions())
            self.assertEqual([p.name for p in files], ["volume-1.brf", "volume-2.brf", "volume-10.brf"])
        finally:
            engine.__exit__()


@unittest.skipUnless(os.environ.get("AELM_TEST_BRAILLE") == "1", "Real bundled-engine test is opt-in")
class BrailleIntegrationTests(unittest.TestCase):
    def test_cancelled_job_shuts_down_and_removes_temporary_files(self):
        with BrailleEngine() as engine:
            work = engine.work
            source = sample_epub(work / "book.epub")
            request = engine.request
            submitted = False
            def cancel_after_submit(path, *args, **kwargs):
                nonlocal submitted
                result = request(path, *args, **kwargs)
                if path == "jobs":
                    submitted = True
                return result
            engine.request = cancel_after_submit
            engine.cancelled = lambda: submitted
            with self.assertRaises(ExportCancelled):
                engine.convert(source, BrailleOptions())
        self.assertFalse(work.exists())
        self.assertIsNotNone(engine.process.poll())

    def test_both_formats_and_grades_preserve_structure_and_final_text(self):
        with BrailleEngine() as engine:
            source = sample_epub(engine.work / "book.epub")
            original = source.read_bytes()
            results = {}
            for fmt in ("brf", "ebraille"):
                for contracted in (True, False):
                    options = BrailleOptions(fmt, contracted, 40, 25)
                    files = engine.convert(source, options)
                    self.assertEqual(len(files), 1)
                    validate_output(files[0], options)
                    results[fmt, contracted] = files[0].read_bytes()
                    if fmt == "brf":
                        text = files[0].read_text("ascii")
                        self.assertIn("EXPORT$" if contracted else "EXPORTED", text)
                    else:
                        with zipfile.ZipFile(files[0]) as archive:
                            root = ET.fromstring(archive.read("ebraille/chapter.xhtml"))
                            self.assertEqual(len(root.findall('.//{*}h1')), 2)
                            self.assertIsNotNone(root.find('.//*[@id="p1"]'))
                            final = root.findall('.//{*}p')[-1].text
                            self.assertTrue(final and any('\u2800' <= c <= '\u28ff' for c in final))
                            package = ET.fromstring(archive.read("package.opf"))
                            code = package.find('.//{*}meta[@property="a11y:brailleSystem"]').text
                            self.assertIn("grade2" if contracted else "grade1", code)
            self.assertNotEqual(results["brf", True], results["brf", False])
            self.assertEqual(source.read_bytes(), original)
            older = sample_epub2(engine.work / "older.epub")
            for fmt in ("brf", "ebraille"):
                with self.subTest(epub_version=2, format=fmt):
                    files = engine.convert(older, BrailleOptions(fmt))
                    self.assertEqual(len(files), 1)
                    validate_output(files[0], BrailleOptions(fmt))

    @unittest.skipUnless(os.environ.get("AELM_TEST_BRAILLE_LARGE") == "1", "Large-book stress test is opt-in")
    def test_large_book_final_paragraph_is_not_truncated(self):
        with BrailleEngine() as engine:
            template = sample_epub(engine.work / "template.epub")
            source = engine.work / "large.epub"
            paragraph = "The quick brown fox jumps over the lazy dog. " * 10
            body = "<p>" + paragraph + "</p>"
            with zipfile.ZipFile(template) as original, zipfile.ZipFile(source, "w", zipfile.ZIP_DEFLATED) as output:
                for item in original.infolist():
                    content = original.read(item)
                    if item.filename == "chapter.xhtml":
                        text = content.decode("utf-8").replace("<section>", "<section>" + body * 650)
                        self.assertGreater(len(text), 500000)
                        content = text.encode("utf-8")
                    output.writestr(item, content)
            for fmt in ("brf", "ebraille"):
                with self.subTest(format=fmt):
                    files = engine.convert(source, BrailleOptions(fmt, False))
                    if fmt == "brf":
                        text = "".join(p.read_text("ascii") for p in files)
                        self.assertIn("EXPORTED", text[-500:])
                        self.assertGreater(text.count("QUICK"), 10000)
                    else:
                        with zipfile.ZipFile(files[0]) as result:
                            chapter = ET.fromstring(result.read("ebraille/chapter.xhtml"))
                            self.assertEqual(len(chapter.findall(".//{*}p")), 1303)
                            self.assertIn("\u2811\u282d\u280f\u2815\u2817\u281e\u2811\u2819", "".join(chapter.findall(".//{*}p")[-1].itertext()))


if __name__ == "__main__":
    unittest.main()
