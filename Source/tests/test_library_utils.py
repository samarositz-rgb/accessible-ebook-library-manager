"""Unit tests for library_utils.

These cover the small pure helpers that other modules rely on for
duplicate detection, ISBN normalization, filename sanitization, and
metadata scoring. They are safe to refactor against.
"""

import importlib.util
import os
import shutil
import time
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase, main


PROJECT_ROOT = Path(__file__).resolve().parents[1]
MODULE_PATH = PROJECT_ROOT / "library_utils.py"

spec = importlib.util.spec_from_file_location("library_utils", MODULE_PATH)
library_utils = importlib.util.module_from_spec(spec)
spec.loader.exec_module(library_utils)


class SafeFilenameTests(TestCase):
    def test_strips_reserved_characters(self):
        self.assertEqual(library_utils.safe_filename('a<b>c:d"e/f\\g|h?i*j'), "a_b_c_d_e_f_g_h_i_j")

    def test_returns_untitled_for_empty(self):
        self.assertEqual(library_utils.safe_filename(""), "Untitled")
        self.assertEqual(library_utils.safe_filename("   "), "Untitled")

    def test_passes_through_normal_text(self):
        self.assertEqual(library_utils.safe_filename("Pride and Prejudice"), "Pride and Prejudice")


class NormalizeDuplicateKeyTests(TestCase):
    def test_lowercases_and_strips_punctuation(self):
        self.assertEqual(library_utils.normalize_duplicate_key("Hello, World!"), "hello world")

    def test_drops_leading_articles(self):
        self.assertEqual(
            library_utils.normalize_duplicate_key("The Great Gatsby"),
            library_utils.normalize_duplicate_key("Great Gatsby"),
        )
        self.assertEqual(
            library_utils.normalize_duplicate_key("A Tale of Two Cities"),
            library_utils.normalize_duplicate_key("Tale of Two Cities"),
        )

    def test_handles_none_and_empty(self):
        self.assertEqual(library_utils.normalize_duplicate_key(None), "")
        self.assertEqual(library_utils.normalize_duplicate_key(""), "")


class NormalizeIsbnKeyTests(TestCase):
    def test_keeps_valid_isbn13(self):
        self.assertEqual(library_utils.normalize_isbn_key("978-0-306-40615-7"), "9780306406157")

    def test_keeps_valid_isbn10(self):
        self.assertEqual(library_utils.normalize_isbn_key("0-306-40615-2"), "0306406152")

    def test_keeps_isbn10_with_x_checksum(self):
        self.assertEqual(library_utils.normalize_isbn_key("123456789X"), "123456789x")

    def test_rejects_wrong_length(self):
        self.assertEqual(library_utils.normalize_isbn_key("12345"), "")
        self.assertEqual(library_utils.normalize_isbn_key("12345678901234"), "")

    def test_handles_none(self):
        self.assertEqual(library_utils.normalize_isbn_key(None), "")


class TitleKeysLookSameTests(TestCase):
    def test_identical_titles_match(self):
        self.assertTrue(library_utils.title_keys_look_same("Moby Dick", "Moby Dick"))

    def test_articles_ignored(self):
        self.assertTrue(library_utils.title_keys_look_same("The Hobbit", "Hobbit"))

    def test_substring_match_when_both_long(self):
        self.assertTrue(library_utils.title_keys_look_same(
            "Pride and Prejudice and Zombies",
            "Pride and Prejudice",
        ))

    def test_unrelated_titles_do_not_match(self):
        self.assertFalse(library_utils.title_keys_look_same("Moby Dick", "War and Peace"))

    def test_empty_inputs_return_false(self):
        self.assertFalse(library_utils.title_keys_look_same("", "anything"))
        self.assertFalse(library_utils.title_keys_look_same("anything", ""))


class MetadataScoreTests(TestCase):
    def test_no_fields_scores_zero(self):
        self.assertEqual(library_utils.metadata_score_from_values(), 0)

    def test_author_worth_three(self):
        self.assertEqual(library_utils.metadata_score_from_values(author="Austen"), 3)

    def test_isbn_only_counts_when_valid(self):
        self.assertEqual(library_utils.metadata_score_from_values(isbn="not-an-isbn"), 0)
        self.assertEqual(library_utils.metadata_score_from_values(isbn="9780306406157"), 2)

    def test_from_detection_dict(self):
        score = library_utils.metadata_score_from_detection({
            "author": "Austen", "year": "1813", "publisher": "Egerton", "isbn": "9780141439518",
        })
        # author 3 + year 2 + publisher 2 + isbn 2
        self.assertEqual(score, 9)


class FilesNeedCopyTests(TestCase):
    def test_missing_destination_needs_copy(self):
        with TemporaryDirectory() as tmp:
            source = Path(tmp) / "a.txt"
            source.write_text("hi")
            destination = Path(tmp) / "missing.txt"
            self.assertTrue(library_utils.files_need_copy(source, destination))

    def test_matching_size_and_mtime_skips_copy(self):
        with TemporaryDirectory() as tmp:
            source = Path(tmp) / "a.txt"
            destination = Path(tmp) / "b.txt"
            source.write_text("identical")
            shutil.copy2(source, destination)
            self.assertFalse(library_utils.files_need_copy(source, destination))

    def test_different_size_needs_copy(self):
        with TemporaryDirectory() as tmp:
            source = Path(tmp) / "a.txt"
            destination = Path(tmp) / "b.txt"
            source.write_text("hello")
            destination.write_text("hi")
            os.utime(destination, (source.stat().st_atime, source.stat().st_mtime))
            self.assertTrue(library_utils.files_need_copy(source, destination))


class SyncFolderContentsTests(TestCase):
    def test_copies_new_files_and_removes_extras(self):
        with TemporaryDirectory() as tmp:
            source = Path(tmp) / "src"
            destination = Path(tmp) / "dst"
            source.mkdir()
            destination.mkdir()
            (source / "keep.txt").write_text("keep")
            (source / "nested").mkdir()
            (source / "nested" / "deep.txt").write_text("deep")
            (destination / "stale.txt").write_text("stale")

            library_utils.sync_folder_contents(source, destination)

            self.assertEqual((destination / "keep.txt").read_text(), "keep")
            self.assertEqual((destination / "nested" / "deep.txt").read_text(), "deep")
            self.assertFalse((destination / "stale.txt").exists())


class FolderFileStatsTests(TestCase):
    def test_counts_files_and_bytes_recursively(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "a.txt").write_bytes(b"hello")
            (root / "sub").mkdir()
            (root / "sub" / "b.txt").write_bytes(b"world!!")
            count, size = library_utils.folder_file_stats(root)
            self.assertEqual(count, 2)
            self.assertEqual(size, 5 + 7)

    def test_missing_folder_returns_zeros(self):
        self.assertEqual(library_utils.folder_file_stats(Path("/no/such/folder/at/all")), (0, 0))


if __name__ == "__main__":
    main()
