"""Conservative source-content checks, not a braille quality certification."""

from pathlib import Path
import posixpath
from urllib.parse import unquote
import xml.etree.ElementTree as ET
import zipfile


def inspect_conversion_risks(source, check_cancelled=lambda: None):
    if Path(source).suffix.lower() != ".epub":
        return ["Source layout was not inspected. Review images, tables, and mathematical notation in the converted book."]
    images = missing_alt = tables = math = 0
    languages = set()
    text_found = False
    try:
        with zipfile.ZipFile(source) as archive:
            container = ET.fromstring(archive.read("META-INF/container.xml"))
            rootfile = container.find(".//{*}rootfile")
            package_name = rootfile.get("full-path") if rootfile is not None else ""
            package = ET.fromstring(archive.read(package_name))
            for node in package.findall(".//{http://purl.org/dc/elements/1.1/}language"):
                if node.text:
                    languages.add(node.text.strip().lower())
            items = {node.get("id"): node for node in package.findall("./{*}manifest/{*}item")}
            visited = set()
            for reference in package.findall("./{*}spine/{*}itemref"):
                check_cancelled()
                item = items.get(reference.get("idref"))
                if item is None or item.get("media-type") != "application/xhtml+xml":
                    continue
                name = posixpath.normpath(posixpath.join(posixpath.dirname(package_name), unquote(item.get("href", ""))))
                if name in visited:
                    continue
                visited.add(name)
                if archive.getinfo(name).file_size > 32 * 1024 * 1024:
                    return ["A chapter is too large for the content-warning check. Its images, tables, and notation were not inspected."]
                root = ET.fromstring(archive.read(name))
                body = root.find("{*}body")
                if body is None:
                    continue
                text_found |= bool("".join(body.itertext()).strip())
                for node in root.iter():
                    tag = node.tag.rsplit("}", 1)[-1].lower()
                    lang = node.get("{http://www.w3.org/XML/1998/namespace}lang") or node.get("lang")
                    if lang:
                        languages.add(lang.strip().lower())
                    if tag in {"img", "svg", "object"}:
                        images += 1
                        if not (node.get("alt", "").strip() or node.get("aria-label", "").strip() or node.findtext("{*}title", "").strip()):
                            missing_alt += 1
                    if tag == "table":
                        tables += 1
                    if tag == "math" or node.get("role") == "math":
                        math += 1
    except (OSError, KeyError, ET.ParseError, zipfile.BadZipFile, ValueError):
        return ["The source could not be fully inspected. Review the converted book for missing or reordered content."]
    warnings = []
    if images:
        warnings.append(f"{images} image or graphic element(s) found; {missing_alt} lack a detected text alternative. Braille export does not perform OCR or create tactile graphics. Review image-only content and descriptions.")
    if tables:
        warnings.append(f"{tables} table(s) found. Verify row and column order, headings, and cell relationships in the output.")
    if math:
        warnings.append(f"{math} mathematical expression(s) found. Verify notation with a qualified braille reader; English literary UEB settings do not certify mathematical transcription.")
    non_english = sorted(lang for lang in languages if lang and lang.split("-")[0] not in {"en", "und"})
    if non_english:
        warnings.append("Non-English language labels found: " + ", ".join(non_english) + ". The selected braille settings are English UEB; check language-specific transcription.")
    if not text_found:
        warnings.append("No readable body text was detected. This may be an image-only book; braille conversion does not perform OCR.")
    return warnings
