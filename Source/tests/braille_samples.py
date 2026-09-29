"""Small structured book for real conversion tests, never a user's library."""
from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET


def sample_epub(path):
    path = Path(path)
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("mimetype", "application/epub+zip")
        archive.writestr("META-INF/container.xml", '''<?xml version="1.0"?>
<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container"><rootfiles><rootfile full-path="book.opf" media-type="application/oebps-package+xml"/></rootfiles></container>''')
        archive.writestr("book.opf", '''<?xml version="1.0"?>
<package xmlns="http://www.idpf.org/2007/opf" version="3.0" unique-identifier="id"><metadata xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:identifier id="id">urn:uuid:01a31dbd-9db2-4a94-9496-730461512110</dc:identifier><dc:title>Braille Export Test</dc:title><dc:creator>Test Author</dc:creator><dc:language>en</dc:language><meta property="dcterms:modified">2026-09-25T00:00:00Z</meta></metadata><manifest><item id="nav" href="nav.xhtml" media-type="application/xhtml+xml" properties="nav"/><item id="chapter" href="chapter.xhtml" media-type="application/xhtml+xml"/></manifest><spine><itemref idref="chapter"/></spine></package>''')
        archive.writestr("nav.xhtml", '''<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops" lang="en"><head><title>Contents</title></head><body><nav epub:type="toc"><h1>Contents</h1><ol><li><a href="chapter.xhtml#one">Chapter One</a></li><li><a href="chapter.xhtml#two">Chapter Two</a></li></ol></nav></body></html>''')
        archive.writestr("chapter.xhtml", '''<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops" lang="en"><head><title>Braille Export Test</title></head><body><section><h1 id="one">Chapter One</h1><span epub:type="pagebreak" role="doc-pagebreak" id="p1" aria-label="1">1</span><p>The quick brown fox jumps over the lazy dog.</p><p>This is a book with numbers 123 and punctuation: yes!</p></section><section><h1 id="two">Chapter Two</h1><p>Final paragraph. Every word should be exported.</p></section></body></html>''')
    return path


def sample_epub2(path):
    original = sample_epub(Path(path).with_name("epub3-source.epub"))
    ns = "http://www.idpf.org/2007/opf"
    with zipfile.ZipFile(original) as source, zipfile.ZipFile(path, "w") as output:
        for entry in source.infolist():
            data = source.read(entry)
            if entry.filename == "book.opf":
                root = ET.fromstring(data)
                root.set("version", "2.0")
                metadata = root.find(f"{{{ns}}}metadata")
                metadata.remove(metadata.find(f"{{{ns}}}meta"))
                item = root.find(f'.//{{{ns}}}item[@id="nav"]')
                item.attrib.clear()
                item.attrib.update(id="ncx", href="toc.ncx", **{"media-type": "application/x-dtbncx+xml"})
                root.find(f"{{{ns}}}spine").set("toc", "ncx")
                data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            if entry.filename != "nav.xhtml":
                output.writestr(entry, data)
        output.writestr("toc.ncx", '''<ncx xmlns="http://www.daisy.org/z3986/2005/ncx/" version="2005-1"><head><meta name="dtb:uid" content="urn:uuid:01a31dbd-9db2-4a94-9496-730461512110"/><meta name="dtb:depth" content="1"/><meta name="dtb:totalPageCount" content="0"/><meta name="dtb:maxPageNumber" content="0"/></head><docTitle><text>Braille Export Test</text></docTitle><navMap><navPoint id="one" playOrder="1"><navLabel><text>Chapter One</text></navLabel><content src="chapter.xhtml#one"/></navPoint><navPoint id="two" playOrder="2"><navLabel><text>Chapter Two</text></navLabel><content src="chapter.xhtml#two"/></navPoint></navMap></ncx>''')
    return Path(path)
