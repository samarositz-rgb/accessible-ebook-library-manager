import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import help_file


class HelpFileTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        self.bundled = self.root / "bundled.txt"
        self.bundled.write_text("Default help", encoding="utf-8")
        self.data = self.root / "data"

    def test_bundled_help_exists_and_has_content(self):
        self.assertIn("Keyboard commands", help_file.bundled_help_path().read_text(encoding="utf-8"))

    def test_first_use_copies_default(self):
        self.assertEqual(help_file.load_help_text(self.data, self.bundled), "Default help")
        self.assertTrue(help_file.user_help_path(self.data).exists())

    def test_user_edits_are_kept(self):
        help_file.ensure_user_help(self.data, self.bundled)
        help_file.user_help_path(self.data).write_text("My notes", encoding="utf-8")
        self.assertEqual(help_file.load_help_text(self.data, self.bundled), "My notes")

    def test_empty_user_file_falls_back_to_default(self):
        help_file.ensure_user_help(self.data, self.bundled)
        help_file.user_help_path(self.data).write_text("  ", encoding="utf-8")
        self.assertEqual(help_file.load_help_text(self.data, self.bundled), "Default help")

    def test_missing_everything_uses_fallback(self):
        self.assertEqual(help_file.load_help_text(self.data, self.root / "none.txt"), help_file.FALLBACK_TEXT)


if __name__ == "__main__":
    unittest.main()
