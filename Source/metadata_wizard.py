"""Staged, non-destructive metadata discovery with native keyboard controls."""

import re

from PySide6.QtCore import QThread
from PySide6.QtWidgets import QPlainTextEdit, QVBoxLayout, QWizard, QWizardPage

from library_manager import detect_metadata_from_text
from online_metadata import lookup_online_metadata


FIELDS = ("title", "author", "series", "series_number", "edition", "year", "isbn", "publisher")


def matching_book(existing, found):
    def normalized(value):
        return re.sub(r"\W+", "", str(value).casefold())

    isbn = normalized(existing.get("isbn", ""))
    candidate = normalized(found.get("isbn", ""))
    if isbn and candidate:
        return isbn == candidate
    title = normalized(existing.get("title", ""))
    if not title or title != normalized(found.get("title", "")):
        return False
    author = normalized(existing.get("author", ""))
    return not author or author == normalized(found.get("author", ""))


def fill_missing(existing, found):
    result = dict(existing)
    changes = []
    for key in FIELDS + ("tags", "notes"):
        value = str(found.get(key) or "").strip()
        if not str(result.get(key) or "").strip() and value:
            result[key] = value
            changes.append(f"{key.replace('_', ' ').title()}: {value}")
    return result, changes


class DiscoveryWorker(QThread):
    def __init__(self, path, metadata, service, parent):
        super().__init__(parent)
        self.path = path
        self.metadata = dict(metadata)
        self.service = service
        self.found = {}
        self.error = ""

    def run(self):
        try:
            if self.service is None:
                self.found = detect_metadata_from_text(self.path, self.metadata, max_chars=250000)
            else:
                self.found, _ = lookup_online_metadata(self.metadata, service_filter=self.service, cancel_check=self.isInterruptionRequested)
                if self.found and not matching_book(self.metadata, self.found):
                    self.found = {}
                    self.error = "The result could not be confidently matched to this book. No fields were changed."
                # Open Library search results describe a work, not a specific edition.
                # Its first publication date and pooled ISBN/publisher must not be
                # presented as this copy's publication details.
                if self.service == "openlibrary":
                    for key in ("year", "isbn", "publisher", "edition"):
                        self.found.pop(key, None)
        except Exception as exc:
            self.error = str(exc)


class DiscoveryPage(QWizardPage):
    def __init__(self, title, service, parent):
        super().__init__(parent)
        self.setTitle(title)
        self.service = service
        self.done = False
        self.report = QPlainTextEdit(self)
        self.report.setReadOnly(True)
        self.report.setTabChangesFocus(True)
        self.report.setAccessibleName(title + " results")
        layout = QVBoxLayout(self)
        layout.addWidget(self.report)

    def isComplete(self):
        return self.done

    def initializePage(self):
        if self.done:
            return
        wizard = self.wizard()
        self.report.setPlainText("Searching book metadata..." if self.service is None else "Searching online metadata...")
        self.worker = DiscoveryWorker(wizard.path, wizard.metadata, self.service, self)
        wizard.worker = self.worker
        self.worker.finished.connect(self.finished)
        self.worker.start()

    def cleanupPage(self):
        # Back/Next revisits the saved results without repeating network requests.
        pass

    def finished(self):
        wizard = self.wizard()
        if wizard.cancel_requested:
            wizard.reject()
            return
        wizard.metadata, changes = fill_missing(wizard.metadata, self.worker.found)
        missing = [key.replace("_", " ") for key in FIELDS if not wizard.metadata.get(key)]
        report = self.worker.error or ("Added:\n" + "\n".join(changes) if changes else "No additional fields found.")
        report += "\n\nExisting values and the book's source were preserved."
        report += "\n\nStill blank: " + (", ".join(missing) if missing else "None")
        self.report.setPlainText(report)
        self.done = True
        self.completeChanged.emit()
        wizard.button(QWizard.WizardButton.BackButton).setEnabled(wizard.currentId() > 0)
        self.report.setFocus()


class MetadataWizard(QWizard):
    def __init__(self, path, metadata, parent=None):
        super().__init__(parent)
        self.path = path
        self.metadata = dict(metadata)
        self.worker = None
        self.cancel_requested = False
        self.review_callback = None
        self.setWindowTitle("Find Missing Metadata")
        self.currentIdChanged.connect(self.update_navigation)
        self.setOption(QWizard.WizardOption.NoBackButtonOnLastPage, False)
        self.setButtonText(QWizard.WizardButton.FinishButton, "&Review Metadata")
        self.resize(640, 420)
        for title, service in (("1. Read Book Metadata and Text", None),
                               ("2. Open Library", "openlibrary"),
                               ("3. Google Books", "google")):
            self.addPage(DiscoveryPage(title, service, self))

    def update_navigation(self):
        running = self.worker is not None and self.worker.isRunning()
        self.button(QWizard.WizardButton.BackButton).setEnabled(self.currentId() > 0 and not running)

    def reject(self):
        if self.worker and self.worker.isRunning():
            self.cancel_requested = True
            self.worker.requestInterruption()
            self.currentPage().report.setPlainText("Cancelling. Waiting for the current search to finish...")
            return
        super().reject()

    def accept(self):
        if self.review_callback and not self.review_callback(self):
            return
        super().accept()

    def closeEvent(self, event):
        if self.worker and self.worker.isRunning():
            event.ignore()
            self.reject()
        else:
            super().closeEvent(event)
