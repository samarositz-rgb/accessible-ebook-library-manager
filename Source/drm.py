"""DRM detection helpers.

Inspects a book file and returns a short label describing the kind of DRM
protection found, or an empty string when the file is unprotected. This
module never decrypts and never bundles any circumvention tooling. Its
sole purpose is honest reporting so the accessibility library can
distinguish books that will actually parse from books that need to be
run through an external decryption step first.

Detected labels:

  ``adobe_adept``   -- Adobe ADEPT DRM (Adobe Digital Editions, many
                       library-lend EPUBs, some Kobo purchases).
  ``epub_encrypted`` -- Generic EPUB content encryption other than the
                        benign IDPF font-obfuscation scheme.
  ``kobo``          -- EPUB encryption.xml declares a Kobo-specific
                       algorithm URI.
  ``pdf_encrypted`` -- PDF whose trailer declares an /Encrypt entry.

Font obfuscation (IDPF and legacy Adobe font schemes) is *not* flagged
because it is intentionally reversible per the EPUB spec and does not
prevent reading.
"""

from __future__ import annotations

import logging
import re
import zipfile
from pathlib import Path


logger = logging.getLogger(__name__)


# Algorithm URIs used only for embedded-font obfuscation. Presence of
# only these algorithms in encryption.xml is not DRM.
_FONT_OBFUSCATION_ALGORITHMS = {
    "http://www.idpf.org/2008/embedding",
    "http://ns.adobe.com/pdf/enc#RC",
}


def detect_drm(path: Path) -> str:
    """Return a short DRM type string if the file is protected, else ''."""
    try:
        suffix = path.suffix.lower()
    except Exception:
        return ""
    if suffix == ".epub":
        return _detect_epub_drm(path)
    if suffix == ".pdf":
        return _detect_pdf_drm(path)
    return ""


def _detect_epub_drm(path: Path) -> str:
    try:
        with zipfile.ZipFile(path, "r") as archive:
            names = set(archive.namelist())

            # Strongest signal: Adobe ADEPT drops a rights.xml file next to
            # container.xml. Its presence is enough to conclude ADEPT DRM.
            if "META-INF/rights.xml" in names:
                return "adobe_adept"

            # Otherwise inspect META-INF/encryption.xml. It exists for both
            # real DRM and for benign IDPF font obfuscation; the difference
            # is the Algorithm URI referenced by each EncryptedData entry.
            if "META-INF/encryption.xml" in names:
                try:
                    enc_bytes = archive.read("META-INF/encryption.xml")
                except Exception:
                    logger.debug("Could not read encryption.xml from %s", path, exc_info=True)
                    return "epub_encrypted"
                enc = enc_bytes.decode("utf-8", errors="ignore")
                algorithms = re.findall(r'Algorithm=["\']([^"\']+)["\']', enc)
                real_encryption = [a for a in algorithms if a not in _FONT_OBFUSCATION_ALGORITHMS]
                if real_encryption:
                    if "kobo" in enc.lower():
                        return "kobo"
                    return "epub_encrypted"
    except zipfile.BadZipFile:
        return ""
    except Exception:
        logger.debug("EPUB DRM detection failed for %s", path, exc_info=True)
    return ""


def _detect_pdf_drm(path: Path) -> str:
    """Peek at head and tail of a PDF for an /Encrypt trailer entry.

    We deliberately do not open the PDF with a real parser here -- this
    check needs to work even without pypdf and must be very fast because
    it runs on every indexed PDF.
    """
    try:
        size = path.stat().st_size
        with open(path, "rb") as fh:
            head = fh.read(8192)
            if size > 4096:
                fh.seek(-4096, 2)
                tail = fh.read(4096)
            else:
                tail = b""
        blob = head + tail
        if re.search(rb"/Encrypt\s+\d+\s+\d+\s+R", blob):
            return "pdf_encrypted"
    except Exception:
        logger.debug("PDF DRM detection failed for %s", path, exc_info=True)
    return ""
