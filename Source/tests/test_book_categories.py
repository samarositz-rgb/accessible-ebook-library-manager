import unittest

from book_categories import COMMON_TAGS, normalize_import_tags
from online_metadata import normalize_online_metadata


class BookCategoryTests(unittest.TestCase):
    def test_formats_and_sources_are_not_categories(self):
        self.assertEqual(normalize_import_tags("EPUB, PDF, DOCX, Kindle, Bookshare, unknown subject"), "")

    def test_supported_categories_and_aliases_are_canonical_and_unique(self):
        self.assertEqual(normalize_import_tags("fiction; Fiction, textbook, nonfiction, EPUB"),
                         "Fiction, Textbooks, Non-fiction")
        self.assertEqual(normalize_import_tags(COMMON_TAGS), ", ".join(COMMON_TAGS))

    def test_online_subjects_use_only_supported_categories(self):
        result = normalize_online_metadata({"tags": ["Fiction / General", "EPUB", "History"]})
        self.assertEqual(result["tags"], "Fiction, History")
