import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from xml.etree import ElementTree as ET
try:
    from defusedxml.ElementTree import fromstring as _safe_fromstring
except ImportError:
    from xml.etree.ElementTree import fromstring as _safe_fromstring


CALIBRE_METADATA_EXTENSIONS = {
    ".azw", ".azw1", ".azw3", ".azw4", ".azw8", ".kfx", ".kfx-zip",
    ".mobi", ".pobi", ".prc", ".tpz",
}
WINDOWS_NO_CONSOLE_FLAGS = getattr(subprocess, "CREATE_NO_WINDOW", 0)


def calibre_tool_candidates(tool_name: str) -> list[Path]:
    executable = f"{tool_name}.exe" if sys.platform.startswith("win") else tool_name
    candidates = []
    found = shutil.which(executable) or shutil.which(tool_name)
    if found:
        candidates.append(Path(found))

    if sys.platform.startswith("win"):
        candidates.extend([
            Path(os.environ.get("PROGRAMFILES", "")) / "Calibre2" / executable,
            Path(os.environ.get("PROGRAMFILES(X86)", "")) / "Calibre2" / executable,
            Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "Calibre2" / executable,
        ])
    elif sys.platform == "darwin":
        candidates.append(Path("/Applications/calibre.app/Contents/MacOS") / tool_name)

    return candidates


def find_calibre_tool(tool_name: str) -> str | None:
    for candidate in calibre_tool_candidates(tool_name):
        if candidate and candidate.exists():
            return str(candidate)
    return None


def first_text(root: ET.Element, *names: str) -> str:
    wanted = {name.casefold() for name in names}
    for element in root.iter():
        tag = element.tag.rsplit("}", 1)[-1].casefold()
        if tag in wanted and element.text:
            value = re.sub(r"\s+", " ", element.text).strip()
            if value:
                return value
    return ""


def all_text(root: ET.Element, *names: str) -> list[str]:
    wanted = {name.casefold() for name in names}
    values = []
    for element in root.iter():
        tag = element.tag.rsplit("}", 1)[-1].casefold()
        if tag in wanted and element.text:
            value = re.sub(r"\s+", " ", element.text).strip()
            if value:
                values.append(value)
    return values


def isbn13_from_digits(value: str) -> str:
    digits = re.sub(r"[^0-9Xx]", "", value or "")
    if len(digits) == 13 and digits.startswith(("978", "979")):
        return digits
    return ""


def isbn10_from_digits(value: str) -> str:
    digits = re.sub(r"[^0-9Xx]", "", value or "")
    return digits.upper() if len(digits) == 10 else ""


def isbn_from_calibre_opf(root: ET.Element) -> str:
    isbn13 = ""
    isbn10 = ""

    for element in root.iter():
        tag = element.tag.rsplit("}", 1)[-1].casefold()
        text = re.sub(r"\s+", " ", element.text or "").strip()
        attributes = " ".join(str(value) for value in element.attrib.values())
        combined = f"{attributes} {text}"
        if "isbn" not in combined.casefold():
            continue

        candidates = re.findall(r"(?:97[89][\d\-\s]{10,20}|[\dXx][\dXx\-\s]{8,18}[\dXx])", combined)
        for candidate in candidates:
            normalized13 = isbn13_from_digits(candidate)
            normalized10 = isbn10_from_digits(candidate)
            if normalized13:
                isbn13 = normalized13
            elif normalized10 and not isbn10:
                isbn10 = normalized10

    return isbn13 or isbn10


def parse_calibre_opf_metadata(opf_text: str) -> dict:
    if not opf_text.strip():
        return {}

    root = _safe_fromstring(opf_text.encode("utf-8"))
    authors = all_text(root, "creator")
    tags = all_text(root, "subject")
    pubdate = first_text(root, "date")
    year_match = re.search(r"\b(1[5-9]\d{2}|20[0-4]\d)\b", pubdate)

    metadata = {
        "title": first_text(root, "title"),
        "author": " & ".join(authors),
        "publisher": first_text(root, "publisher"),
        "year": year_match.group(1) if year_match else "",
        "isbn": isbn_from_calibre_opf(root),
        "tags": ", ".join(dict.fromkeys(tags)),
        "notes": first_text(root, "description"),
    }
    return {key: value for key, value in metadata.items() if value}


def read_calibre_metadata(path: Path, ebook_meta_path: str | None = None, timeout: int = 60) -> dict:
    ebook_meta = ebook_meta_path or find_calibre_tool("ebook-meta")
    if not ebook_meta:
        return {}

    with tempfile.TemporaryDirectory(prefix="aelm_calibre_metadata_") as temp_folder:
        opf_path = Path(temp_folder) / "metadata.opf"
        completed = subprocess.run(
            [ebook_meta, str(path), "--to-opf", str(opf_path)],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
            creationflags=WINDOWS_NO_CONSOLE_FLAGS,
            timeout=timeout,
        )
        if completed.returncode != 0 or not opf_path.exists():
            return {}
        return parse_calibre_opf_metadata(opf_path.read_text(encoding="utf-8", errors="ignore"))
