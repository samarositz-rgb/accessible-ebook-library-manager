import html
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
try:
    from defusedxml.ElementTree import fromstring as _safe_fromstring
except ImportError:
    from xml.etree.ElementTree import fromstring as _safe_fromstring


WINDOWS_NO_CONSOLE_FLAGS = getattr(subprocess, "CREATE_NO_WINDOW", 0)


def strip_xml_html_tags(text: str) -> str:
    text = re.sub(r"<[^>]+>", " ", text)
    text = html.unescape(text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def strip_xml_html_tags_preserve_lines(text: str) -> str:
    text = re.sub(r"<\s*(?:br|/p|/h[1-6]|/li|/div|/section|/nav)\b[^>]*>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", " ", text)
    lines = [re.sub(r"\s+", " ", html.unescape(line)).strip() for line in text.splitlines()]
    return "\n".join(line for line in lines if line)


def read_text_from_epub(epub_path: Path, max_chars: int = 50000) -> str:
    chunks = []
    try:
        with zipfile.ZipFile(epub_path, "r") as archive:
            names = archive.namelist()
            text_names = [
                name for name in names
                if name.lower().endswith((".xhtml", ".html", ".htm", ".xml"))
                and "nav" not in name.lower()
                and "toc" not in name.lower()
                and "container.xml" not in name.lower()
            ]

            for name in text_names[:80]:
                try:
                    decoded = archive.read(name).decode("utf-8", errors="ignore")
                    cleaned = strip_xml_html_tags(decoded)
                    if cleaned:
                        chunks.append(cleaned)
                    if sum(len(chunk) for chunk in chunks) >= max_chars:
                        break
                except Exception:
                    continue
    except Exception:
        return ""

    return "\n".join(chunks)[:max_chars]


def read_text_from_epub_preserve_lines(epub_path: Path, max_chars: int = 50000) -> str:
    chunks = []
    try:
        with zipfile.ZipFile(epub_path, "r") as archive:
            names = archive.namelist()
            text_names = [
                name for name in names
                if name.lower().endswith((".xhtml", ".html", ".htm", ".xml"))
                and "nav" not in name.lower()
                and "toc" not in name.lower()
                and "container.xml" not in name.lower()
            ]

            for name in text_names[:80]:
                try:
                    decoded = archive.read(name).decode("utf-8", errors="ignore")
                    cleaned = strip_xml_html_tags_preserve_lines(decoded)
                    if cleaned:
                        chunks.append(cleaned)
                    if sum(len(chunk) for chunk in chunks) >= max_chars:
                        break
                except Exception:
                    continue
    except Exception:
        return ""

    return "\n".join(chunks)[:max_chars]


def read_text_from_docx(docx_path: Path, max_chars: int = 50000) -> str:
    try:
        with zipfile.ZipFile(docx_path, "r") as archive:
            raw = archive.read("word/document.xml")
        root = _safe_fromstring(raw)
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        paragraphs = []
        for paragraph in root.findall(".//w:p", ns):
            parts = []
            for element in paragraph.iter():
                tag = element.tag.rsplit("}", 1)[-1]
                if tag == "t" and element.text:
                    parts.append(element.text)
                elif tag == "tab":
                    parts.append(" ")
                elif tag in {"br", "cr"}:
                    parts.append("\n")
            text = "".join(parts).strip()
            if text:
                paragraphs.append(text)
            if sum(len(item) for item in paragraphs) >= max_chars:
                break
        if paragraphs:
            return "\n\n".join(paragraphs)[:max_chars]
        decoded = raw.decode("utf-8", errors="ignore")
        return strip_xml_html_tags(decoded)[:max_chars]
    except Exception:
        return ""


def read_text_from_legacy_doc(doc_path: Path, max_chars: int = 50000) -> str:
    if not sys.platform.startswith("win") or not shutil.which("powershell"):
        return ""

    script = r"""
$ErrorActionPreference = 'Stop'
$path = $args[0]
$maxChars = [int]$args[1]
$word = $null
$doc = $null
try {
    [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open($path, $false, $true, $false)
    $text = [string]$doc.Content.Text
    if ($text.Length -gt $maxChars) {
        $text = $text.Substring(0, $maxChars)
    }
    [Console]::Write($text)
}
finally {
    if ($doc -ne $null) {
        $doc.Close($false) | Out-Null
    }
    if ($word -ne $null) {
        $word.Quit() | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
"""
    try:
        with tempfile.TemporaryDirectory(prefix="aelm_doc_extract_") as temp_folder:
            script_path = Path(temp_folder) / "extract_doc_text.ps1"
            script_path.write_text(script, encoding="utf-8")
            completed = subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    str(script_path),
                    str(doc_path),
                    str(max_chars),
                ],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="ignore",
                creationflags=WINDOWS_NO_CONSOLE_FLAGS,
                timeout=90,
            )
        if completed.returncode != 0:
            return ""
        return re.sub(r"\r\n?|\v|\f", "\n", completed.stdout or "")[:max_chars]
    except Exception:
        return ""


def read_text_from_plain_file(path: Path, max_chars: int = 50000) -> str:
    try:
        raw = path.read_bytes()[:max_chars * 2]
        return raw.decode("utf-8", errors="ignore")[:max_chars]
    except Exception:
        return ""


def read_metadata_text_from_pdf(pdf_path: Path, max_chars: int = 12000) -> str:
    try:
        raw = pdf_path.read_bytes()[: max_chars * 20]
    except Exception:
        return ""

    text = raw.decode("latin-1", errors="ignore")
    lines = []
    for key in ["Title", "Author", "Subject", "Keywords", "Creator", "Producer"]:
        for match in re.finditer(rf"/{key}\s*\((.*?)\)", text, flags=re.DOTALL):
            value = match.group(1)
            value = value.replace(r"\(", "(").replace(r"\)", ")").replace(r"\\", "\\")
            value = re.sub(r"\s+", " ", value).strip()
            if value:
                lines.append(f"{key}: {value}")
    return "\n".join(lines)[:max_chars]


def read_text_from_pdf(pdf_path: Path, max_chars: int = 50000) -> str:
    try:
        from pypdf import PdfReader
    except Exception:
        return read_metadata_text_from_pdf(pdf_path, max_chars=max_chars)

    chunks = []
    try:
        reader = PdfReader(str(pdf_path))
        metadata = reader.metadata or {}
        for label, key in [
            ("Title", "/Title"),
            ("Author", "/Author"),
            ("Subject", "/Subject"),
            ("Keywords", "/Keywords"),
        ]:
            value = metadata.get(key)
            if value:
                chunks.append(f"{label}: {value}")

        for page in reader.pages:
            try:
                page_text = page.extract_text() or ""
            except Exception:
                page_text = ""
            if page_text.strip():
                chunks.append(page_text)
            if sum(len(chunk) for chunk in chunks) >= max_chars:
                break
    except Exception:
        return read_metadata_text_from_pdf(pdf_path, max_chars=max_chars)

    text = "\n\n".join(chunks).strip()
    if not text:
        return read_metadata_text_from_pdf(pdf_path, max_chars=max_chars)
    return text[:max_chars]


def read_text_for_metadata_detection(path: Path, max_chars: int = 50000) -> str:
    suffix = path.suffix.lower()
    if suffix == ".epub":
        return read_text_from_epub(path, max_chars=max_chars)
    if suffix == ".docx":
        return read_text_from_docx(path, max_chars=max_chars)
    if suffix == ".doc":
        return read_text_from_legacy_doc(path, max_chars=max_chars)
    if suffix == ".pdf":
        return read_text_from_pdf(path, max_chars=max_chars)
    if suffix in {".txt", ".rtf", ".html", ".htm"}:
        return read_text_from_plain_file(path, max_chars=max_chars)
    return ""


def extract_text_for_indexing(path: Path) -> str:
    suffix = path.suffix.lower()
    max_chars = 500_000
    if suffix == ".epub":
        return read_text_from_epub_preserve_lines(path, max_chars=max_chars)
    if suffix == ".docx":
        return read_text_from_docx(path, max_chars=max_chars)
    if suffix == ".doc":
        return read_text_from_legacy_doc(path, max_chars=max_chars)
    if suffix == ".pdf":
        return read_text_from_pdf(path, max_chars=max_chars)
    if suffix in {".txt", ".rtf", ".html", ".htm"}:
        return read_text_from_plain_file(path, max_chars=max_chars)
    return ""


def detect_language_from_text(text: str, default: str = "en") -> str:
    sample = re.sub(r"[^A-Za-zÃ€-Ã–Ã˜-Ã¶Ã¸-Ã¿' ]+", " ", text or "").casefold()
    words = re.findall(r"[a-zÃ€-Ã–Ã˜-Ã¶Ã¸-Ã¿']+", sample[:200000])
    if not words:
        return default

    word_set = set(words)
    language_markers = {
        "en": {"the", "and", "of", "to", "in", "that", "is", "for", "with", "as", "was", "are", "by", "from", "this"},
        "es": {"el", "la", "los", "las", "de", "que", "y", "en", "un", "una", "por", "con", "para", "es", "del"},
        "fr": {"le", "la", "les", "de", "des", "que", "et", "en", "un", "une", "pour", "avec", "est", "dans", "du"},
        "de": {"der", "die", "das", "und", "von", "zu", "den", "mit", "ist", "im", "dem", "ein", "eine", "nicht", "fÃ¼r"},
        "it": {"il", "lo", "la", "gli", "le", "di", "che", "e", "un", "una", "per", "con", "Ã¨", "del", "della"},
        "pt": {"o", "a", "os", "as", "de", "que", "e", "em", "um", "uma", "para", "com", "Ã©", "do", "da"},
    }
    scores = {}
    for language, markers in language_markers.items():
        scores[language] = sum(1 for marker in markers if marker in word_set)

    best_language, best_score = max(scores.items(), key=lambda item: item[1])
    if best_score < 3:
        return default
    return best_language
