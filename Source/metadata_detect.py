"""Heuristic metadata detection.

Pure-ish helpers that pull title/author/publisher/year/ISBN out of book text,
plus light wrappers around Open Library and Google Books for online lookup.

Two helpers (read_epub_metadata and is_import_boilerplate_line) still live
in library_manager.py, so we import them lazily inside the functions that
need them to avoid a circular import.
"""

import html
import json
import re
import urllib.parse
import urllib.request
from pathlib import Path

from calibre_tools import CALIBRE_METADATA_EXTENSIONS, read_calibre_metadata  # noqa: F401  (kept for symmetry)
from document_text import read_text_for_metadata_detection


def clean_metadata_line(line: str) -> str:
    line = html.unescape(line).strip()
    line = re.sub(r"\s+", " ", line)
    line = re.sub(r"^[#*\\\-\u2013\u2014: ]+", "", line)
    line = re.sub(r"^(?:bookshare|accessible book|daisy book)\s*[:\-]\s*", "", line, flags=re.IGNORECASE)
    return line.strip()


def is_bookshare_notice_line(line: str) -> bool:
    lowered = re.sub(r"\s+", " ", line or "").strip().lower()
    if not lowered:
        return False
    return (
        lowered == "notice"
        or "this accessible media has been made available to people with bona fide disabilities" in lowered
        or lowered.startswith("this notice tells you about restrictions on the use of this accessible media")
        or "bona fide disabilities that affect reading" in lowered
    )


def extract_isbn_from_text(text: str) -> str:
    """Find a likely ISBN, including Word superscript/check-digit artifacts."""
    if not text:
        return ""

    text = re.sub(r"[\u2010-\u2015]", "-", text)
    labeled_patterns = [
        r"\bISBN(?:-1[03])?[ \t]*:?[ \t]*([0-9Xx][0-9Xx \t\-\^\.]{8,35}[0-9Xx])",
        r"\bISBN(?:-1[03])?[ \t]*:?[ \t]*((?:97[89][-\t ]?)?[0-9][0-9Xx \t\-\^\.]{8,35})",
    ]
    general_patterns = [
        r"\b((?:97[89][-\s]?)?[0-9][-\s]?[0-9]{2,5}[-\s]?[0-9]{2,7}[-\s]?[0-9]{1,7}(?:[-\s\^]?[0-9Xx])?)\b",
    ]

    candidates = []
    for pattern in labeled_patterns:
        candidates.extend(re.findall(pattern, text, flags=re.IGNORECASE))
    for pattern in general_patterns:
        candidates.extend(re.findall(pattern, text, flags=re.IGNORECASE))

    fallback = ""
    isbn10 = ""
    for candidate in candidates:
        cleaned = re.sub(r"[^0-9Xx]", "", candidate)
        if len(cleaned) == 13:
            return cleaned
        if len(cleaned) == 10 and not isbn10:
            isbn10 = cleaned
        if not fallback and len(cleaned) in {9, 11, 12} and cleaned.startswith(("978", "979")):
            fallback = cleaned
    if isbn10:
        return isbn10
    return fallback


def extract_publication_year_from_text(text: str, require_label: bool = False, include_copyright: bool = True) -> str:
    if not text:
        return ""

    sample = text[:60000]
    edition_match = re.search(
        r"\b(19[5-9]\d|20[0-4]\d)(?:\s*[-/]\s*(?:19[5-9]\d|20[0-4]\d))?\s+(?:abridged\s+)?(?:edition|ed\.)\b",
        sample,
        flags=re.IGNORECASE,
    )
    if edition_match:
        return edition_match.group(1)

    labeled_patterns = [
        r"\b(?:published|publication date|date of publication|publication year)\D{0,80}(19[5-9]\d|20[0-4]\d)",
    ]
    if include_copyright:
        labeled_patterns.append(r"(?:\u00a9|\(c\)|copyright)\D{0,80}(19[5-9]\d|20[0-4]\d)")
    for pattern in labeled_patterns:
        match = re.search(pattern, sample, flags=re.IGNORECASE)
        if match:
            return match.group(1)

    if require_label:
        return ""

    match = re.search(r"\b(19[5-9]\d|20[0-4]\d)\b", sample, flags=re.IGNORECASE)
    return match.group(1) if match else ""


def extract_publisher_from_text(text: str) -> str:
    if not text:
        return ""

    sample = text[:60000]
    patterns = [
        r"(?:\u00a9|\(c\)|copyright)\s*(?:19[5-9]\d|20[0-4]\d)\s+([^\r\n]{2,80})",
        r"\bpublished by\s+([^\r\n]{2,100})",
        r"\bpublisher\s*[:\-]\s*([^\r\n]{2,100})",
    ]
    for pattern in patterns:
        match = re.search(pattern, sample, flags=re.IGNORECASE)
        if not match:
            continue
        publisher = clean_metadata_line(match.group(1))
        publisher = clean_publisher_value(publisher)
        if looks_like_publisher(publisher):
            return publisher

    if re.search(r"\bA\s+Thomson\s+Reuters\s+business\b", sample, flags=re.IGNORECASE):
        return "Thomson Reuters"
    return ""


def clean_filename_title(path: Path) -> str:
    title = path.stem
    title = re.sub(r"[_]+", " ", title)
    title = re.sub(r"\s*-\s*", " - ", title)
    title = re.sub(r"\s+", " ", title)
    title = re.sub(r"\b(?:bookshare|daisy|epub|pdf|docx|accessible)\b", "", title, flags=re.IGNORECASE)
    title = re.sub(r"\b(?:bs|bookshare)[-_ ]?\d+\b", "", title, flags=re.IGNORECASE)
    title = re.sub(r"\s+", " ", title).strip(" -_.,")
    return title or path.stem.replace("_", " ").replace("-", " ")


def is_weak_title(title: str, path: Path) -> bool:
    if not title:
        return True
    cleaned_title = re.sub(r"\s+", " ", title).strip().lower()
    cleaned_stem = clean_filename_title(path).lower()
    raw_stem = path.stem.replace("_", " ").replace("-", " ").lower()
    return cleaned_title in {cleaned_stem, raw_stem, path.stem.lower()} or len(cleaned_title) < 3


def clean_author_value(value: str) -> str:
    value = clean_metadata_line(value)
    value = re.sub(r"^(?:by|author|authors|written by|edited by)\s*[:\-]?\s*", "", value, flags=re.IGNORECASE)
    value = re.sub(r"\s+(?:narrated by|read by|bookshare|copyright|all rights reserved).*$", "", value, flags=re.IGNORECASE)
    value = re.sub(r"\s+", " ", value).strip(" .;,")
    if value.isupper() and len(value.split()) <= 8:
        value = value.title()
        value = re.sub(r"\b([A-Z])\.", lambda match: match.group(1).upper() + ".", value)
    return value


def clean_title_value(value: str) -> str:
    value = clean_metadata_line(value)
    value = re.sub(r"^(?:title|book title)\s*[:\-]?\s*", "", value, flags=re.IGNORECASE)
    value = re.sub(r"\s+", " ", value).strip(" .;,")
    return value


def title_words(value: str) -> list[str]:
    return [
        word.casefold()
        for word in re.findall(r"[A-Za-z][A-Za-z0-9']+", value or "")
        if len(word) > 2
    ]


def has_useful_filename_title(path: Path) -> bool:
    words = title_words(clean_filename_title(path))
    return len(words) >= 2


def looks_like_machine_pdf_title(value: str) -> bool:
    title = clean_title_value(value)
    if not title:
        return True

    lower = title.casefold()
    if re.search(r"\.(?:pdf|docx?|rtf|txt|epub)$", lower):
        return True
    if re.search(r"\b[a-z]\d{3,}\b", lower) or re.search(r"\b\d{4,}\b", lower):
        return True

    words = title_words(title)
    if not words:
        return True

    short_words = [word for word in words if len(word) <= 4]
    if len(words) >= 4 and len(short_words) / len(words) >= 0.75 and re.search(r"\d", title):
        return True

    machine_terms = {"matls", "contr", "ch", "fm", "em", "cb", "a012636"}
    if any(word in machine_terms for word in words):
        return True

    return False


def looks_like_boilerplate_title(value: str) -> bool:
    lower = clean_title_value(value).casefold()
    if not lower:
        return True
    if lower in {"published", "copyright", "contents", "table of contents", "chapter", "page"}:
        return True
    boilerplate = [
        "all rights reserved",
        "paragraph",
        "several pages",
        "photocopy",
        "recording",
        "licensed to",
        "permission",
        "publisher",
        "excerpted material",
        "reprinted by permission",
        "appearing in this book",
        "evaluation for use",
        "third persons",
        "respect for rights",
    ]
    if re.match(r"^\d+(?:\.\d+)+\s+", lower):
        return True
    return any(fragment in lower for fragment in boilerplate)


def title_candidate_matches_filename(candidate: str, path: Path) -> bool:
    filename_words = set(title_words(clean_filename_title(path)))
    candidate_words = set(title_words(candidate))
    if not filename_words or not candidate_words:
        return False
    overlap = filename_words & candidate_words
    needed = min(2, len(candidate_words), len(filename_words))
    return len(overlap) >= needed


def looks_like_useful_title_candidate(candidate: str, path: Path, prefer_filename: bool = False) -> bool:
    candidate = clean_title_value(candidate)
    if len(candidate) < 4 or looks_like_boilerplate_title(candidate):
        return False

    words = title_words(candidate)
    if len(words) > 18:
        return False

    if prefer_filename:
        if looks_like_machine_pdf_title(candidate):
            return False
        if has_useful_filename_title(path) and not title_candidate_matches_filename(candidate, path):
            return False

    return True


def should_replace_title(current_title: str, candidate: str, path: Path, prefer_filename: bool = False) -> bool:
    if not looks_like_useful_title_candidate(candidate, path, prefer_filename=prefer_filename):
        return False
    if not current_title:
        return True
    if prefer_filename and has_useful_filename_title(path) and not is_weak_title(current_title, path):
        return False
    return is_weak_title(current_title, path)


def title_page_candidate(lines: list[str], path: Path) -> str:
    filename_words = set(title_words(clean_filename_title(path)))
    for line in lines[:12]:
        candidate = clean_title_value(line)
        if not looks_like_useful_title_candidate(candidate, path, prefer_filename=False):
            continue
        candidate_words = set(title_words(candidate))
        if filename_words and filename_words.issubset(candidate_words):
            return candidate.title() if candidate.isupper() else candidate
    return ""


def clean_publisher_value(value: str) -> str:
    value = clean_metadata_line(value)
    value = re.sub(r"^(?:publisher|published by|imprint)\s*[:\-]?\s*", "", value, flags=re.IGNORECASE)
    value = re.sub(r"^(?:copyright|\(c\)|\u00a9)\s*(?:[a-z]\s*)?(?:19[5-9]\d|20[0-4]\d)(?:\s*,\s*(?:19[5-9]\d|20[0-4]\d))*\s*", "", value, flags=re.IGNORECASE)
    value = re.sub(r"^(?:[,;]\s*)?(?:19[5-9]\d|20[0-4]\d)(?:\s*,\s*(?:19[5-9]\d|20[0-4]\d))*\s*", "", value, flags=re.IGNORECASE)
    value = re.sub(r"[\s.]+(?:all rights.*|printed in .*)$", "", value, flags=re.IGNORECASE)
    value = re.sub(r"\s+", " ", value).strip(" .;,")
    return value


def looks_like_publisher(value: str) -> bool:
    if not value:
        return False
    lower = value.casefold()
    if len(value) > 120:
        return False
    if any(fragment in lower for fragment in [
        " across the street ", " friend", " gun", " saw ", " pages are indicated",
        "law school", "advisory board", "created this publication", "accurate and authoritative",
        "particular jurisdiction", "does not render", "professional advice",
        "reprinted by permission", " by permission", "excerpted material",
    ]):
        return False
    if value.rstrip().endswith("-"):
        return False
    publisher_pattern = (
        r"\b(?:press|publishing|publisher|publishers|books|house|pearson|mcgraw|"
        r"cengage|wiley|openstax|scholastic|harper|penguin|simon|houghton|"
        r"macmillan|oxford|cambridge|west|aspen|wolters|kluwer|lexisnexis)\b"
        r"|random house|thomson reuters|matthew bender|carolina academic"
    )
    return bool(re.search(publisher_pattern, lower, flags=re.IGNORECASE))


def labeled_value(line: str, labels: list[str]) -> str:
    label_pattern = "|".join(re.escape(label) for label in labels)
    match = re.match(rf"^(?:{label_pattern})\s*[:\-]\s*(.+)$", line, flags=re.IGNORECASE)
    if match:
        return clean_metadata_line(match.group(1))
    return ""


def line_after_label(lines: list[str], labels: list[str], max_index: int = 120) -> str:
    label_pattern = "|".join(re.escape(label) for label in labels)
    for index, line in enumerate(lines[:max_index]):
        value = labeled_value(line, labels)
        if value:
            return value
        if re.fullmatch(rf"(?:{label_pattern})\s*:?", line, flags=re.IGNORECASE):
            for next_line in lines[index + 1:index + 4]:
                if next_line and not re.fullmatch(rf"(?:{label_pattern})\s*:?", next_line, flags=re.IGNORECASE):
                    return clean_metadata_line(next_line)
    return ""


def looks_like_author(value: str) -> bool:
    if not value:
        return False
    lower = value.lower()
    blocked_fragments = [
        "chapter", "contents", "copyright", "isbn", "publisher", "bookshare",
        "licensed to", "persons licensed", "all rights reserved", "photocopy",
        "recording", "permission", "ellipsis", "several pages", "one of",
        "any means", "electronic or mechanical", "co-conspirators",
        "third persons", "evaluation for use", "respect for rights",
    ]
    if any(fragment in lower for fragment in blocked_fragments):
        return False
    words = [word for word in re.split(r"\s+", value) if word]
    if not 1 <= len(words) <= 8 or len(value) > 120:
        return False
    if lower.endswith(":") or re.search(r"(?<!\b[A-Z])\.(?!\s*[A-Z]\b)", value):
        return False
    if value[:1].islower():
        return False
    if re.fullmatch(r"[A-Z\s:.'-]+", value) and (len(words) <= 1 or value.rstrip().endswith(":") or value.casefold() in {"the court"}):
        return False
    if len(words) == 1 and not re.search(r"[A-Z][a-z]+", value):
        return False

    name_signal = bool(re.search(r"\b[A-Z]\.", value)) or len(re.findall(r"\b[A-Z][a-zA-Z'.-]+\b", value)) >= 2
    if not name_signal:
        return False
    return True


def fetch_json(url: str, timeout: int = 12) -> dict:
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": "AccessibleEbookLibraryManager/1.0",
            "Accept": "application/json",
        },
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8", errors="ignore"))


def normalize_online_metadata(metadata: dict) -> dict:
    result = {
        "title": metadata.get("title", ""),
        "author": metadata.get("author", ""),
        "edition": metadata.get("edition", ""),
        "year": metadata.get("year", ""),
        "isbn": metadata.get("isbn", ""),
        "publisher": metadata.get("publisher", ""),
        "source": metadata.get("source", ""),
        "tags": metadata.get("tags", ""),
        "notes": metadata.get("notes", ""),
        "cover_url": metadata.get("cover_url", ""),
    }
    for key, value in list(result.items()):
        if isinstance(value, list):
            result[key] = ", ".join(str(item) for item in value if item)
        elif value is None:
            result[key] = ""
        else:
            result[key] = str(value).strip()
    if result["year"]:
        match = re.search(r"\b(1[5-9]\d{2}|20\d{2})\b", result["year"])
        if match:
            result["year"] = match.group(1)
    return result


def metadata_from_open_library_doc(doc: dict) -> dict:
    isbn_values = doc.get("isbn") or []
    tags = doc.get("subject") or []
    cover_id = doc.get("cover_i")
    cover_url = f"https://covers.openlibrary.org/b/id/{cover_id}-L.jpg" if cover_id else ""
    return normalize_online_metadata({
        "title": doc.get("title", ""),
        "author": ", ".join(doc.get("author_name") or []),
        "year": str(doc.get("first_publish_year") or ""),
        "isbn": isbn_values[0] if isbn_values else "",
        "publisher": (doc.get("publisher") or [""])[0],
        "tags": ", ".join(tags[:8]),
        "cover_url": cover_url,
    })


def metadata_from_google_volume(volume: dict) -> dict:
    info = volume.get("volumeInfo", {})
    identifiers = info.get("industryIdentifiers") or []
    isbn = ""
    for identifier in identifiers:
        if identifier.get("type") == "ISBN_13":
            isbn = identifier.get("identifier", "")
            break
    if not isbn and identifiers:
        isbn = identifiers[0].get("identifier", "")
    categories = info.get("categories") or []
    image_links = info.get("imageLinks") or {}
    cover_url = image_links.get("large") or image_links.get("medium") or image_links.get("thumbnail") or ""
    return normalize_online_metadata({
        "title": info.get("title", ""),
        "author": ", ".join(info.get("authors") or []),
        "year": info.get("publishedDate", ""),
        "isbn": isbn,
        "publisher": info.get("publisher", ""),
        "tags": ", ".join(categories[:8]),
        "notes": info.get("description", ""),
        "cover_url": cover_url,
    })


def lookup_online_metadata(existing: dict) -> tuple[dict, str]:
    isbn = re.sub(r"[^0-9Xx]", "", existing.get("isbn", ""))
    title = existing.get("title", "").strip()
    author = existing.get("author", "").strip()

    queries = []
    if isbn:
        queries.append(("openlibrary", f"https://openlibrary.org/search.json?isbn={urllib.parse.quote(isbn)}&limit=1"))
        queries.append(("google", f"https://www.googleapis.com/books/v1/volumes?q=isbn:{urllib.parse.quote(isbn)}&maxResults=1"))
    if title:
        open_query = f"title:{title}"
        google_query = f"intitle:{title}"
        if author:
            open_query += f" author:{author}"
            google_query += f"+inauthor:{author}"
        queries.append(("openlibrary", "https://openlibrary.org/search.json?" + urllib.parse.urlencode({"q": open_query, "limit": "1"})))
        queries.append(("google", "https://www.googleapis.com/books/v1/volumes?" + urllib.parse.urlencode({"q": google_query, "maxResults": "1"})))

    last_error = ""
    for service, url in queries:
        try:
            data = fetch_json(url)
            if service == "openlibrary":
                docs = data.get("docs") or []
                if docs:
                    return metadata_from_open_library_doc(docs[0]), "Open Library"
            else:
                items = data.get("items") or []
                if items:
                    return metadata_from_google_volume(items[0]), "Google Books"
        except Exception as exc:
            last_error = str(exc)
            continue

    if last_error:
        raise RuntimeError(last_error)
    return {}, ""


def detect_metadata_from_text(path: Path, existing: dict | None = None) -> dict:
    """Best-effort local metadata detector.

    This is not a true AI model. It uses EPUB metadata, file names, and common
    title/author patterns found near the start of the book text. It is designed
    to help with sources such as Bookshare, where file names are often awkward.
    """
    from library_manager import read_epub_metadata, is_import_boilerplate_line  # lazy: avoid circular import
    existing = existing or {}
    result = {
        "title": existing.get("title", ""),
        "author": existing.get("author", ""),
        "source": existing.get("source", ""),
        "tags": existing.get("tags", ""),
        "notes": existing.get("notes", ""),
        "edition": existing.get("edition", ""),
        "year": existing.get("year", ""),
        "isbn": existing.get("isbn", ""),
        "publisher": existing.get("publisher", ""),
    }

    if path.suffix.lower() == ".epub":
        epub_metadata = read_epub_metadata(path)
        for key in result:
            if epub_metadata.get(key):
                result[key] = epub_metadata[key]

    prefer_filename_title = path.suffix.lower() in {".pdf", ".docx", ".doc"}
    text = read_text_for_metadata_detection(path)
    lines = [clean_metadata_line(line) for line in re.split(r"[\r\n]+", text)]
    lines = [
        line for line in lines
        if line and len(line) < 240 and not is_bookshare_notice_line(line) and not is_import_boilerplate_line(line)
    ]

    filename_title = clean_filename_title(path)
    if (
        is_weak_title(result["title"], path)
        or (prefer_filename_title and looks_like_machine_pdf_title(result["title"]) and has_useful_filename_title(path))
    ):
        result["title"] = filename_title

    if prefer_filename_title:
        front_title = title_page_candidate(lines, path)
        if front_title:
            result["title"] = front_title

    # Look harder for labeled Bookshare/front-matter fields.
    title_value = line_after_label(lines, ["title", "book title", "name"], max_index=180)
    if title_value:
        candidate = clean_title_value(title_value)
        if should_replace_title(result["title"], candidate, path, prefer_filename=prefer_filename_title):
            result["title"] = candidate

    author_value = line_after_label(
        lines,
        ["author", "authors", "creator", "creators", "by", "written by"],
        max_index=220,
    )
    if author_value and (not result["author"] or "unknown" in result["author"].lower()):
        cleaned_author = clean_author_value(author_value)
        if looks_like_author(cleaned_author):
            result["author"] = cleaned_author

    # Look for common author patterns and title-by-author front matter.
    author_patterns = [
        r"^by\s+(.+)$",
        r"^authors?[:\s]+(.+)$",
        r"^written by\s+(.+)$",
        r"^(.+?)\s*/\s*by\s+(.+)$",
        r"^(.+?)\s+by\s+([A-Z][A-Za-z .,'-]{2,120})$",
    ]
    for line in lines[:160]:
        for pattern in author_patterns:
            match = re.match(pattern, line, flags=re.IGNORECASE)
            if not match:
                continue
            if pattern.startswith("^(.+?)"):
                possible_title = clean_title_value(match.group(1))
                possible_author = clean_author_value(match.group(2))
                if should_replace_title(result["title"], possible_title, path, prefer_filename=prefer_filename_title):
                    result["title"] = possible_title
            else:
                possible_author = clean_author_value(match.group(1))
            if (not result["author"] or "unknown" in result["author"].lower()) and looks_like_author(possible_author):
                result["author"] = possible_author
                break
        if result["author"]:
            break

    # Look for common title labels.
    for line in lines[:160]:
        match = re.match(r"^(?:title|book title)[:\s]+(.+)$", line, flags=re.IGNORECASE)
        if match:
            candidate = clean_title_value(match.group(1))
            if should_replace_title(result["title"], candidate, path, prefer_filename=prefer_filename_title):
                result["title"] = candidate
            break

    # If there is still no good title, use the first substantial line that is not a boilerplate line.
    boilerplate_words = [
        "bookshare", "copyright", "all rights reserved", "dedication",
        "contents", "table of contents", "chapter", "isbn", "published"
    ]
    if not result["title"]:
        for line in lines[:100]:
            lower = line.lower()
            if len(line) < 4:
                continue
            if is_import_boilerplate_line(line):
                continue
            if any(word in lower for word in boilerplate_words):
                continue
            if lower.startswith("by "):
                continue
            result["title"] = line
            break

    if not result["title"]:
        result["title"] = filename_title

    # ISBN, year, publisher, and edition guesses.
    if text and not result.get("isbn"):
        result["isbn"] = extract_isbn_from_text(text)

    if text:
        text_year = extract_publication_year_from_text(
            text,
            require_label=prefer_filename_title or bool(result.get("year")),
            include_copyright=not prefer_filename_title,
        )
        if text_year:
            result["year"] = text_year

    if text and not result.get("publisher"):
        publisher_value = line_after_label(
            lines,
            ["publisher", "published by", "imprint"],
            max_index=220,
        )
        if publisher_value:
            cleaned_publisher = clean_publisher_value(publisher_value)
            if looks_like_publisher(cleaned_publisher):
                result["publisher"] = cleaned_publisher
        if not result.get("publisher"):
            result["publisher"] = extract_publisher_from_text(text)
        if not result.get("publisher"):
            for line in lines[:180]:
                cleaned_publisher = clean_publisher_value(line)
                if looks_like_publisher(cleaned_publisher):
                    result["publisher"] = cleaned_publisher
                    break

    if text and not result.get("edition"):
        edition_value = line_after_label(lines, ["edition"], max_index=220)
        if edition_value and "edition" in edition_value.lower():
            result["edition"] = clean_metadata_line(edition_value)
        for line in lines[:180]:
            if result.get("edition"):
                break
            match = re.search(r"\b(\d+(?:st|nd|rd|th)\s+edition|first edition|second edition|third edition|fourth edition|fifth edition|sixth edition|seventh edition|eighth edition|ninth edition|tenth edition|revised edition|international edition|teacher'?s edition|student edition)\b", line, flags=re.IGNORECASE)
            if match:
                result["edition"] = clean_metadata_line(match.group(1))
                break
            match = re.search(r"\b((?:19[5-9]\d|20[0-4]\d)(?:\s*[-/]\s*(?:19[5-9]\d|20[0-4]\d))?\s+(?:abridged\s+)?edition)\b", line, flags=re.IGNORECASE)
            if match:
                result["edition"] = clean_metadata_line(match.group(1))
                break

    # Source guess.
    combined = (str(path) + " " + text[:1000]).lower()
    if not result["source"]:
        if "bookshare" in combined:
            result["source"] = "Bookshare"
        elif "kindle" in combined or path.suffix.lower() in CALIBRE_METADATA_EXTENSIONS:
            result["source"] = "Kindle"
        else:
            result["source"] = "Personal"

    # Tags guess.
    if not result["tags"]:
        tags = []
        keyword_value = line_after_label(lines, ["keywords", "subjects", "subject"], max_index=120)
        if keyword_value:
            tags.extend([part.strip() for part in re.split(r"[,;]", keyword_value) if part.strip()])
        if "bookshare" in combined:
            tags.append("Bookshare")
        if path.suffix.lower() == ".epub":
            tags.append("EPUB")
        elif path.suffix.lower() in CALIBRE_METADATA_EXTENSIONS:
            tags.append("Kindle")
        elif path.suffix.lower() == ".docx":
            tags.append("DOCX")
        elif path.suffix.lower() == ".pdf":
            tags.append("PDF")
        deduped_tags = []
        for tag in tags:
            if tag and tag.lower() not in [existing_tag.lower() for existing_tag in deduped_tags]:
                deduped_tags.append(tag)
        result["tags"] = ", ".join(deduped_tags)

    # Do not auto-fill notes. Notes remain available for manual editing only.
    return result


def detect_metadata_from_text_content(text: str, existing: dict | None = None) -> dict:
    existing = existing or {}
    result = {
        "title": existing.get("title", ""),
        "author": existing.get("author", ""),
        "source": existing.get("source", ""),
        "tags": existing.get("tags", ""),
        "notes": existing.get("notes", ""),
        "edition": existing.get("edition", ""),
        "year": existing.get("year", ""),
        "isbn": existing.get("isbn", ""),
        "publisher": existing.get("publisher", ""),
    }

    from library_manager import is_import_boilerplate_line
    lines = [clean_metadata_line(line) for line in re.split(r"[\r\n]+", text or "")]
    lines = [
        line for line in lines
        if line and len(line) < 240 and not is_bookshare_notice_line(line) and not is_import_boilerplate_line(line)
    ]

    title_value = line_after_label(lines, ["title", "book title", "name"], max_index=180)
    if title_value:
        result["title"] = clean_title_value(title_value)

    author_value = line_after_label(
        lines,
        ["author", "authors", "creator", "creators", "by", "written by"],
        max_index=220,
    )
    if author_value:
        cleaned_author = clean_author_value(author_value)
        if looks_like_author(cleaned_author):
            result["author"] = cleaned_author

    if not result["author"]:
        for line in lines[:160]:
            match = re.match(r"^by\s+(.+)$", line, flags=re.IGNORECASE)
            if match:
                cleaned_author = clean_author_value(match.group(1))
                if looks_like_author(cleaned_author):
                    result["author"] = cleaned_author
                    break

    if not result["title"]:
        for line in lines[:100]:
            lower = line.lower()
            if len(line) < 4 or lower.startswith("by "):
                continue
            if is_import_boilerplate_line(line):
                continue
            if any(word in lower for word in ["bookshare", "copyright", "contents", "table of contents", "chapter", "isbn", "published"]):
                continue
            result["title"] = clean_title_value(line)
            break

    if text and not result.get("isbn"):
        result["isbn"] = extract_isbn_from_text(text)

    if text:
        text_year = extract_publication_year_from_text(text, require_label=bool(result.get("year")))
        if text_year:
            result["year"] = text_year

    if text and not result.get("publisher"):
        publisher_value = line_after_label(lines, ["publisher", "published by", "imprint"], max_index=220)
        if publisher_value:
            cleaned_publisher = clean_publisher_value(publisher_value)
            if looks_like_publisher(cleaned_publisher):
                result["publisher"] = cleaned_publisher
        if not result.get("publisher"):
            result["publisher"] = extract_publisher_from_text(text)

    if text and not result.get("edition"):
        edition_value = line_after_label(lines, ["edition"], max_index=220)
        if edition_value:
            result["edition"] = clean_metadata_line(edition_value)
        for line in lines[:180]:
            if result.get("edition"):
                break
            match = re.search(r"\b(\d+(?:st|nd|rd|th)\s+edition|first edition|second edition|third edition|fourth edition|fifth edition|sixth edition|seventh edition|eighth edition|ninth edition|tenth edition|revised edition|international edition|teacher'?s edition|student edition)\b", line, flags=re.IGNORECASE)
            if match:
                result["edition"] = clean_metadata_line(match.group(1))
                break

    if not result["source"]:
        result["source"] = "Bookshare" if "bookshare" in (text or "").lower() else "Personal"

    return result



