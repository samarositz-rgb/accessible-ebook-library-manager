"""Unit tests for DRM detection.

These build synthetic EPUB and PDF files with the specific markers each
detection path looks for, verifying the classifier without ever needing
a real DRM-protected book.
"""

import sys
import zipfile
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase, main


PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))
import drm


CONTAINER_XML = """<?xml version="1.0"?>
<container xmlns="urn:oasis:names:tc:opendocument:xmlns:container" version="1.0">
  <rootfiles><rootfile full-path="content.opf" media-type="application/oebps-package+xml"/></rootfiles>
</container>
"""

OPF_XML = """<?xml version="1.0" encoding="UTF-8"?>
<package xmlns="http://www.idpf.org/2007/opf" version="3.0" unique-identifier="pub-id">
  <metadata xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>T</dc:title></metadata>
  <manifest/><spine/>
</package>
"""

ADEPT_RIGHTS_XML = """<?xml version="1.0"?>
<adept:rights xmlns:adept="http://ns.adobe.com/adept">
  <adept:licenseToken>...</adept:licenseToken>
</adept:rights>
"""

FONT_OBFUSCATION_ENCRYPTION_XML = """<?xml version="1.0" encoding="UTF-8"?>
<encryption xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
  <EncryptedData xmlns="http://www.w3.org/2001/04/xmlenc#">
    <EncryptionMethod Algorithm="http://www.idpf.org/2008/embedding"/>
    <CipherData><CipherReference URI="OEBPS/fonts/BookFont.otf"/></CipherData>
  </EncryptedData>
</encryption>
"""

REAL_ENCRYPTION_XML = """<?xml version="1.0" encoding="UTF-8"?>
<encryption xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
  <EncryptedData xmlns="http://www.w3.org/2001/04/xmlenc#">
    <EncryptionMethod Algorithm="http://www.w3.org/2001/04/xmlenc#aes256-cbc"/>
    <CipherData><CipherReference URI="OEBPS/text/chapter1.xhtml"/></CipherData>
  </EncryptedData>
</encryption>
"""

KOBO_ENCRYPTION_XML = """<?xml version="1.0" encoding="UTF-8"?>
<encryption xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
  <EncryptedData xmlns="http://www.w3.org/2001/04/xmlenc#">
    <EncryptionMethod Algorithm="http://kobo.com/2013/xml/enc/kobo-content"/>
    <CipherData><CipherReference URI="OEBPS/text/chapter1.xhtml"/></CipherData>
  </EncryptedData>
</encryption>
"""


def make_epub(tmp: Path, name: str, extra_files: dict[str, str] | None = None) -> Path:
    path = tmp / name
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("META-INF/container.xml", CONTAINER_XML)
        zf.writestr("content.opf", OPF_XML)
        for member, body in (extra_files or {}).items():
            zf.writestr(member, body)
    return path


class EpubDrmDetectionTests(TestCase):
    def test_plain_epub_is_not_flagged(self):
        with TemporaryDirectory() as tmp:
            path = make_epub(Path(tmp), "plain.epub")
            self.assertEqual(drm.detect_drm(path), "")

    def test_adept_rights_xml_is_flagged(self):
        with TemporaryDirectory() as tmp:
            path = make_epub(Path(tmp), "adept.epub", {"META-INF/rights.xml": ADEPT_RIGHTS_XML})
            self.assertEqual(drm.detect_drm(path), "adobe_adept")

    def test_font_obfuscation_alone_is_not_drm(self):
        with TemporaryDirectory() as tmp:
            path = make_epub(Path(tmp), "font.epub", {"META-INF/encryption.xml": FONT_OBFUSCATION_ENCRYPTION_XML})
            self.assertEqual(drm.detect_drm(path), "")

    def test_generic_content_encryption_is_flagged(self):
        with TemporaryDirectory() as tmp:
            path = make_epub(Path(tmp), "enc.epub", {"META-INF/encryption.xml": REAL_ENCRYPTION_XML})
            self.assertEqual(drm.detect_drm(path), "epub_encrypted")

    def test_kobo_algorithm_uri_maps_to_kobo(self):
        with TemporaryDirectory() as tmp:
            path = make_epub(Path(tmp), "kobo.epub", {"META-INF/encryption.xml": KOBO_ENCRYPTION_XML})
            self.assertEqual(drm.detect_drm(path), "kobo")

    def test_bad_zip_returns_empty(self):
        with TemporaryDirectory() as tmp:
            path = Path(tmp) / "bad.epub"
            path.write_bytes(b"not a zip at all")
            self.assertEqual(drm.detect_drm(path), "")


class PdfDrmDetectionTests(TestCase):
    def test_encrypted_trailer_is_flagged(self):
        with TemporaryDirectory() as tmp:
            path = Path(tmp) / "enc.pdf"
            path.write_bytes(b"%PDF-1.4\n1 0 obj\n<<>>\nendobj\ntrailer\n<< /Encrypt 42 0 R /Size 5 >>\n%%EOF\n")
            self.assertEqual(drm.detect_drm(path), "pdf_encrypted")

    def test_unencrypted_pdf_is_not_flagged(self):
        with TemporaryDirectory() as tmp:
            path = Path(tmp) / "plain.pdf"
            path.write_bytes(b"%PDF-1.4\n1 0 obj\n<<>>\nendobj\ntrailer\n<< /Size 5 >>\n%%EOF\n")
            self.assertEqual(drm.detect_drm(path), "")

    def test_missing_pdf_returns_empty(self):
        self.assertEqual(drm.detect_drm(Path("/no/such/file.pdf")), "")


class NonBookExtensionTests(TestCase):
    def test_txt_returns_empty(self):
        with TemporaryDirectory() as tmp:
            path = Path(tmp) / "a.txt"
            path.write_text("just text")
            self.assertEqual(drm.detect_drm(path), "")


if __name__ == "__main__":
    main()
