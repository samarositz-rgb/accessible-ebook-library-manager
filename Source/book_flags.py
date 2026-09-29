"""Per-book notes recorded while a book is indexed.

Two things are worth knowing about a book before its text is ever read aloud:

* Whether a PDF is only scanned pictures of pages. Such a file has no text for a
  screen reader to speak, so the person should hear about it up front and be
  able to list every such book.
* Which accessible-library title a file is. Bookshare and Learning Ally files
  are often re-downloaded under different file names; their package identifier
  is a far more reliable way to recognise the same book again than the title.

Everything here is read-only with respect to the book files and is best effort:
a failure to inspect a file is logged and never stops indexing or importing.
"""

from __future__ import annotations

import logging
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

try:
    from defusedxml.ElementTree import fromstring as _safe_fromstring
except ImportError:  # defusedxml is optional
    _safe_fromstring = ET.fromstring

from document_text import extract_text_for_indexing


logger = logging.getLogger(__name__)

IMAGE_ONLY_PDF = "image_only_pdf"
SCANNED_PDF_SAMPLE_PAGES = 10
SCANNED_PDF_MIN_CHARS_PER_PAGE = 100

_OPF_NS = "http://www.idpf.org/2007/opf"
_DC_NS = "http://purl.org/dc/elements/1.1/"
_CONTAINER_NS = "urn:oasis:names:tc:opendocument:xmlns:container"

FLAG_DESCRIPTIONS = {
    IMAGE_ONLY_PDF: "This PDF appears to be scanned pictures of pages with no readable text. It needs OCR before a screen reader can read it.",
}


def describe_flags(flags):
    """Return plain sentences for the flags a book carries."""
    return [FLAG_DESCRIPTIONS[flag] for flag in flags if flag in FLAG_DESCRIPTIONS]


def is_pdf_probably_image_only(pdf_path: Path) -> bool:
    """True when the first pages of a PDF contain almost no extractable text.

    Returns False whenever the answer is uncertain (pypdf missing, unreadable or
    password-protected file) so that good books are never flagged by mistake.
    """
    try:
        from pypdf import PdfReader
    except Exception:
        return False
    try:
        reader = PdfReader(str(pdf_path))
        if reader.is_encrypted:
            return False
        pages = reader.pages
        sample = min(SCANNED_PDF_SAMPLE_PAGES, len(pages))
        if sample == 0:
            return False
        characters = 0
        for index in range(sample):
            try:
                characters += len((pages[index].extract_text() or "").strip())
            except Exception:
                continue
    except Exception:
        logger.debug("Could not inspect PDF text for %s", pdf_path, exc_info=True)
        return False
    return characters / sample < SCANNED_PDF_MIN_CHARS_PER_PAGE


def extract_source_book_id(epub_path: Path) -> str:
    """Return ``bookshare:<id>`` or ``learningally:<id>`` from an EPUB, else ''.

    Only identifiers that name one of those libraries are accepted, so an
    unrelated publisher identifier can never be mistaken for a match.
    """
    try:
        with zipfile.ZipFile(epub_path, "r") as archive:
            container = _safe_fromstring(archive.read("META-INF/container.xml"))
            rootfile = container.find(f".//{{{_CONTAINER_NS}}}rootfile")
            if rootfile is None:
                return ""
            package = _safe_fromstring(archive.read(rootfile.attrib.get("full-path", "")))
    except Exception:
        return ""

    for identifier in package.iter(f"{{{_DC_NS}}}identifier"):
        text = (identifier.text or "").strip()
        if not text:
            continue
        scheme = identifier.attrib.get(f"{{{_OPF_NS}}}scheme", "")
        haystack = f"{text} {scheme}".casefold().replace(" ", "")
        if "bookshare" in haystack:
            runs = re.findall(r"\d{3,}", text)
            if runs:
                return f"bookshare:{max(runs, key=len)}"
        if "learningally" in haystack:
            tokens = re.findall(r"[A-Za-z]{0,3}\d{3,}", text)
            if tokens:
                return f"learningally:{tokens[-1].upper()}"
    return ""


def record_book_flags(database, book_id: int, path) -> None:
    """Inspect one stored book and save its flags. Never raises."""
    path = Path(path)
    suffix = path.suffix.lower()
    try:
        if suffix == ".pdf":
            flags = [flag for flag in database.get_content_flags(book_id) if flag != IMAGE_ONLY_PDF]
            if is_pdf_probably_image_only(path):
                flags.append(IMAGE_ONLY_PDF)
            if flags != database.get_content_flags(book_id):
                database.set_content_flags(book_id, flags)
        elif suffix == ".epub":
            source_book_id = extract_source_book_id(path)
            if source_book_id and source_book_id != database.get_source_book_id(book_id):
                database.set_source_book_id(book_id, source_book_id)
    except Exception:
        logger.debug("Could not record flags for book %s (%s)", book_id, path, exc_info=True)


def index_book_file(database, book_id: int, path) -> None:
    """Index a book's text for search and record what we learn about the file."""
    path = Path(path)
    database.index_book_content(book_id, extract_text_for_indexing(path))
    record_book_flags(database, book_id, path)
