"""Notice the books in a Kindle for PC download folder and look for accessible copies.

Only the plain information at the start of each Kindle file (title, author and
Amazon ID) is read. The protected book text is never read, copied, imported or
opened, and nothing here removes or works around Amazon's protection. A Kindle
book is added to the library as a No File Available entry, then the free public
catalogues are checked for a copy that can be read with a screen reader or
braille display. Sources that need a member login, such as Bookshare, are one
keypress away from the entry through Book > Find Book on Other Services.
"""

from __future__ import annotations

import os
import re
import struct
import urllib.parse
from pathlib import Path

from goodreads_import import WISHLIST_TAG, find_match, is_wishlist_path, merge_tags, title_key, author_last_names

KINDLE_PREFIX = "wishlist:kindle:"
KINDLE_TAG = "On Kindle"
FREE_COPY_TAG = "Free Copy Found"
KINDLE_FOLDER_SETTING = "kindle_folder"
READABLE_EXTENSIONS = {".azw", ".azw3", ".azw4", ".mobi", ".prc"}
UNREADABLE_EXTENSIONS = {".kfx", ".azw8", ".kcr"}
HEADER_BYTES = 256 * 1024

EXTH_TYPES = {100: "author", 101: "publisher", 104: "isbn", 106: "published", 113: "asin", 504: "asin",
              503: "title", 501: "cdetype", 524: "language"}


def candidate_folders() -> list[Path]:
    """Places Kindle for PC keeps downloaded books, including synced Documents folders."""
    home = Path.home()
    roots = [home / "Documents", home / "OneDrive" / "Documents", home / "iCloudDrive" / "Documents" / "Documents",
             home / "iCloudDrive" / "Documents"]
    return [root / "My Kindle Content" for root in roots if (root / "My Kindle Content").is_dir()]


def is_online_only(path) -> bool:
    """True for a cloud placeholder (iCloud Drive or OneDrive) that is not stored on this computer.

    Opening one would make Windows download it, which can stall, so it is skipped and counted.
    """
    try:
        return bool(os.stat(path).st_file_attributes & (0x400000 | 0x40000 | 0x1000))
    except (AttributeError, OSError):
        return False


def _decode(raw: bytes, encoding: int) -> str:
    codec = "utf-8" if encoding == 65001 else "cp1252"
    return raw.decode(codec, errors="replace").strip("\x00 \r\n")


def read_kindle_header(path) -> dict | None:
    """Title, author, ASIN and similar from the unprotected header of a MOBI or AZW file.

    Returns None when the file is not in that format. Never reads the book text.
    """
    try:
        with open(path, "rb") as handle:
            data = handle.read(HEADER_BYTES)
        if len(data) < 90 or data[60:68] not in (b"BOOKMOBI", b"TEXtREAd"):
            return None
        record_count = struct.unpack(">H", data[76:78])[0]
        if record_count < 1:
            return None
        start = struct.unpack(">I", data[78:82])[0]
        record = data[start:]
        if record[16:20] != b"MOBI":
            return None
        header_length = struct.unpack(">I", record[20:24])[0]
        encryption = struct.unpack(">H", record[12:14])[0]
        encoding = struct.unpack(">I", record[28:32])[0]
        name_offset, name_length = struct.unpack(">II", record[84:92])
        details = {"title": _decode(record[name_offset:name_offset + name_length], encoding), "protected": encryption != 0}
        exth_flags = struct.unpack(">I", record[128:132])[0] if len(record) >= 132 else 0
        position = 16 + header_length
        if exth_flags & 0x40 and record[position:position + 4] == b"EXTH":
            count = struct.unpack(">I", record[position + 8:position + 12])[0]
            cursor = position + 12
            for _ in range(count):
                if cursor + 8 > len(record):
                    break
                kind, length = struct.unpack(">II", record[cursor:cursor + 8])
                if length < 8:
                    break
                value = record[cursor + 8:cursor + length]
                name = EXTH_TYPES.get(kind)
                if name and name not in details:
                    details[name] = _decode(value, encoding)
                cursor += length
        if details.get("title") == "" and "title" not in details:
            return None
        return details
    except Exception:
        return None


def scan_kindle_folder(folder) -> dict:
    """Find readable Kindle books under a folder. Newer formats are only counted."""
    root = Path(folder)
    books, unreadable, online_only, seen = [], 0, 0, set()
    for current, _dirs, files in os.walk(root):
        for name in sorted(files):
            suffix = Path(name).suffix.lower()
            path = Path(current) / name
            if suffix in UNREADABLE_EXTENSIONS:
                unreadable += 1
            elif suffix in READABLE_EXTENSIONS:
                if is_online_only(path):
                    online_only += 1
                    continue
                details = read_kindle_header(path)
                if not details:
                    continue
                asin = details.get("asin") or re.sub(r"_EBOK$", "", path.stem)
                if details.get("cdetype") == "PDOC" or asin in seen:
                    continue  # personal documents are not store purchases
                seen.add(asin)
                books.append({"asin": asin, "title": details.get("title", ""), "author": details.get("author", ""),
                              "publisher": details.get("publisher", ""), "isbn": re.sub(r"[^0-9Xx]", "", details.get("isbn", "")),
                              "protected": details.get("protected", True)})
    return {"books": books, "unreadable": unreadable, "online_only": online_only}


def kindle_note(book: dict, lookups: list[str]) -> str:
    parts = [f"Kindle book, Amazon ID {book['asin']}."]
    if book.get("protected"):
        parts.append("It is protected by Amazon and cannot be used outside the Kindle apps.")
    parts.extend(lookups)
    return " ".join(parts)


def find_free_copies(title: str, author: str, fetch_json, search_gutenberg) -> list[str]:
    """Check Project Gutenberg and Open Library. Returns plain sentences for the notes."""
    found = []
    wanted_title = title_key(title)
    wanted_authors = author_last_names(author)
    try:
        for hit in search_gutenberg(title, author):
            if title_key(hit["title"]) == wanted_title and (not wanted_authors or wanted_authors & author_last_names(hit["author"])):
                found.append(f"Free copy on Project Gutenberg: {hit['url']}.")
                break
    except Exception:
        pass
    try:
        query = {"title": title.split(":")[0], "limit": "5", "fields": "key,title,author_name,ebook_access"}
        if author:
            query["author"] = author
        data = fetch_json("https://openlibrary.org/search.json?" + urllib.parse.urlencode(query))
        labels = {"public": "free to read", "borrowable": "can be borrowed", "printdisabled": "available to readers with print disabilities"}
        for doc in data.get("docs") or []:
            access = doc.get("ebook_access", "")
            if access in labels and title_key(doc.get("title", "")) == wanted_title and (
                    not wanted_authors or wanted_authors & author_last_names(", ".join(doc.get("author_name") or []))):
                found.append(f"Open Library says it is {labels[access]}: https://openlibrary.org{doc['key']}.")
                break
    except Exception:
        pass
    return found


def add_kindle_books(db, books: list[dict], lookup=None, cancel_check=None, progress=None) -> dict:
    """Add new Kindle books as No File Available entries. lookup(title, author) returns note sentences."""
    rows = [list(r) for r in db.all_books_for_duplicate_check()]
    result = {"added": 0, "marked": 0, "already": 0, "free_found": 0}
    for number, book in enumerate(books, start=1):
        if cancel_check and cancel_check():
            break
        key = f"kindle:{book['asin']}"
        if db.find_book_id_by_source_book_id(key):
            result["already"] += 1
            continue
        if progress:
            progress(f"Checking Kindle book {number} of {len(books)}: {book['title']}")
        existing = find_match({"title": book["title"], "author": book["author"], "isbn": book["isbn"], "isbn10": ""}, rows)
        if existing is not None:
            db.set_source_book_id(existing[0], key)
            if not is_wishlist_path(existing[8]):
                result["already"] += 1
                continue
            db.update_book(existing[0], existing[1], existing[2], existing[3], merge_tags(existing[4], [KINDLE_TAG]), existing[5])
            result["marked"] += 1
            continue
        lines = lookup(book["title"], book["author"]) if lookup else []
        tags = [WISHLIST_TAG, KINDLE_TAG] + ([FREE_COPY_TAG] if lines else [])
        book_id = db.add_book(book["title"], book["author"], "Kindle", "", "", "", KINDLE_PREFIX + book["asin"])
        db.update_book(book_id, book["title"], book["author"], "Kindle", merge_tags("", tags), kindle_note(book, lines))
        db.set_source_book_id(book_id, key)
        if book["isbn"] or book["publisher"]:
            db.update_extra_fields(book_id, isbn=book["isbn"], publisher=book["publisher"])
        rows.append([book_id, book["title"], book["author"], "Kindle", ", ".join(tags), "", "", "", KINDLE_PREFIX + book["asin"]] + [None] * 14)
        result["added"] += 1
        result["free_found"] += 1 if lines else 0
    return result
