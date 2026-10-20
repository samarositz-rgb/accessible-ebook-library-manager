import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import find_elsewhere as fe

FEED = """<?xml version="1.0" encoding="utf-8"?>
<feed xmlns="http://www.w3.org/2005/Atom">
<title>Books: pride</title>
<entry><id>https://www.gutenberg.org/ebooks/1342.opds</id><title>Pride and Prejudice</title><content type="text">Jane Austen</content></entry>
<entry><id>https://www.gutenberg.org/ebooks/search.opds/?x=1</id><title>Not a book</title></entry>
<entry><id>https://www.gutenberg.org/ebooks/42671.opds</id><title>Pride and Prejudice</title><content type="text">Jane Austen</content></entry>
</feed>"""


class FindElsewhereTests(unittest.TestCase):
    def test_unset_means_all_services(self):
        self.assertEqual(fe.enabled_services(""), list(fe.SERVICES))

    def test_chosen_services_keep_fixed_order(self):
        self.assertEqual(fe.enabled_services("bard, gutenberg, bogus"), ["gutenberg", "bard"])

    def test_search_words_drop_subtitle_and_use_last_name(self):
        self.assertEqual(fe.search_words("Dune (Dune, #1)", "Frank Herbert"), "Dune Herbert")
        self.assertEqual(fe.search_words("Emma: A Novel", "Austen, Jane"), "Emma Austen")

    def test_site_with_search_address(self):
        url, label, paste = fe.service_target("bookshare", "Dune Herbert")
        self.assertEqual(url, "https://www.bookshare.org/search?keyword=Dune+Herbert")
        self.assertEqual(label, "Bookshare")
        self.assertFalse(paste)

    def test_login_only_site_opens_home_page_and_needs_paste(self):
        url, _label, paste = fe.service_target("bard", "Dune Herbert")
        self.assertEqual(url, "https://nlsbard.loc.gov")
        self.assertTrue(paste)

    def test_parse_gutenberg_feed_skips_non_books(self):
        results = fe.parse_gutenberg_feed(FEED)
        self.assertEqual([r["url"] for r in results],
                         ["https://www.gutenberg.org/ebooks/1342", "https://www.gutenberg.org/ebooks/42671"])
        self.assertEqual(results[0]["author"], "Jane Austen")


if __name__ == "__main__":
    unittest.main()
