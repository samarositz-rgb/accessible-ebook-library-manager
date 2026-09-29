import os
import unittest
from pathlib import Path
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QApplication, QWizard
from metadata_wizard import MetadataWizard, DiscoveryWorker, fill_missing, matching_book
from online_metadata import lookup_online_metadata


class MetadataWizardTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_merge_preserves_existing_and_source(self):
        merged, changes = fill_missing({"title": "Corrected", "source": "Bookshare", "author": ""},
                                       {"title": "Old", "source": "Google Books", "author": "Writer"})
        self.assertEqual(merged, {"title": "Corrected", "source": "Bookshare", "author": "Writer"})
        self.assertEqual(changes, ["Author: Writer"])

    def test_matching_rejects_different_book_or_edition(self):
        self.assertFalse(matching_book({"title": "Alpha"}, {"title": "Beta"}))
        self.assertFalse(matching_book({"title": "Alpha", "isbn": "123"}, {"title": "Alpha", "isbn": "456"}))
        self.assertTrue(matching_book({"title": "Alpha!"}, {"title": "ALPHA"}))

    def test_service_filter_does_not_query_other_provider(self):
        with patch("online_metadata.fetch_json", return_value={}) as fetch:
            lookup_online_metadata({"title": "Alpha", "isbn": "123"}, service_filter="google")
        self.assertEqual(fetch.call_count, 2)
        self.assertTrue(all("googleapis.com" in call.args[0] for call in fetch.call_args_list))

    def test_open_library_does_not_assign_work_publication_to_edition(self):
        worker = DiscoveryWorker(Path("unused"), {"title": "Alpha"}, "openlibrary", None)
        with patch("metadata_wizard.lookup_online_metadata", return_value=({"title": "Alpha", "author": "Writer", "year": "1900", "isbn": "123", "publisher": "Other edition"}, "Open Library")):
            worker.run()
        self.assertEqual(worker.found, {"title": "Alpha", "author": "Writer"})

    def wait_page(self, wizard):
        for _ in range(200):
            self.app.processEvents()
            if wizard.currentPage().done:
                return
            QTest.qWait(10)
        self.fail("Metadata page did not finish")

    def test_wizard_sequences_sources_and_handles_service_error(self):
        with patch("metadata_wizard.detect_metadata_from_text", return_value={"title": "Alpha", "author": "Writer"}) as local, \
             patch("metadata_wizard.lookup_online_metadata", side_effect=[({}, ""), RuntimeError("Service unavailable")]) as online:
            wizard = MetadataWizard(Path("unused"), {"title": "Alpha", "source": "Bookshare"})
            wizard.show()
            self.wait_page(wizard)
            self.assertEqual(wizard.metadata["author"], "Writer")
            self.assertEqual(local.call_args.kwargs["max_chars"], 250000)
            wizard.next()
            self.wait_page(wizard)
            wizard.next()
            self.wait_page(wizard)
            self.assertIn("Service unavailable", wizard.currentPage().report.toPlainText())
            self.assertEqual([call.kwargs["service_filter"] for call in online.call_args_list], ["openlibrary", "google"])
            self.assertEqual(wizard.metadata["source"], "Bookshare")
            self.assertTrue(wizard.button(QWizard.WizardButton.FinishButton).isEnabled())
            wizard.reject()


if __name__ == "__main__":
    unittest.main()
