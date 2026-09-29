from pathlib import Path
import tempfile
import unittest
import zipfile

from braille_export import ExportCancelled
from conversion_warnings import inspect_conversion_risks
from tests.braille_samples import sample_epub


class ConversionWarningTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.folder = Path(self.temp.name)
        self.original = sample_epub(self.folder / "original.epub")

    def tearDown(self):
        self.temp.cleanup()

    def book_with(self, content):
        target = self.folder / "test.epub"
        with zipfile.ZipFile(self.original) as source, zipfile.ZipFile(target, "w") as output:
            for item in source.infolist():
                data = source.read(item)
                if item.filename == "chapter.xhtml":
                    data = data.replace(b"</body>", content.encode() + b"</body>")
                output.writestr(item, data)
        return target

    def test_plain_text_book_has_no_specific_flags(self):
        self.assertEqual(inspect_conversion_risks(self.original), [])

    def test_images_tables_math_and_language_are_flagged_without_changing_book(self):
        source = self.book_with('<img src="diagram.png"/><img src="cover.png" alt="Cover"/><table><tr><td>Value</td></tr></table><math xmlns="http://www.w3.org/1998/Math/MathML"><mi>x</mi></math><p lang="fr">Bonjour</p>')
        before = source.read_bytes()
        warnings = "\n".join(inspect_conversion_risks(source))
        self.assertIn("2 image", warnings)
        self.assertIn("1 lack", warnings)
        self.assertIn("1 table", warnings)
        self.assertIn("1 mathematical", warnings)
        self.assertIn("fr", warnings)
        self.assertEqual(source.read_bytes(), before)

    def test_malformed_epub_does_not_claim_no_risks(self):
        source = self.folder / "broken.epub"
        source.write_bytes(b"not a zip")
        self.assertIn("could not be fully inspected", inspect_conversion_risks(source)[0])

    def test_cancellation_propagates(self):
        def cancel():
            raise ExportCancelled()
        with self.assertRaises(ExportCancelled):
            inspect_conversion_risks(self.original, cancel)


if __name__ == "__main__":
    unittest.main()
