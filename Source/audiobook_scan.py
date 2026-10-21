"""Find audiobooks (M4B files, such as a Libation library) and link them into the library.

Audiobook files are large, so they are linked where they are and never copied or
changed. Removing a book from the library never deletes a linked file. If a book
is already in the library as a No File Available entry, the entry is linked to the file
rather than duplicated.
"""

from __future__ import annotations

import re
import struct
from pathlib import Path

from goodreads_import import NOTE_MARKER, find_match, is_wishlist_path, merge_tags

AUDIOBOOK_TAG = "Audiobook"
ASIN_PATTERN = re.compile(r"\[([A-Z0-9]{10})\]")
MAX_MOOV_BYTES = 64 * 1024 * 1024

TEXT_ATOMS = {
    b"\xa9nam": "title", b"\xa9ART": "author", b"aART": "album_artist", b"\xa9alb": "album",
    b"\xa9wrt": "narrator", b"\xa9nrt": "narrator", b"\xa9day": "year", b"\xa9gen": "genre",
    b"desc": "description", b"\xa9cmt": "comment", b"\xa9mvn": "series", b"\xa9mvi": "series_number",
}


def _atoms(data: bytes, start: int = 0, end: int | None = None):
    """Yield (type, body_start, body_end) for each atom in data[start:end]."""
    end = len(data) if end is None else end
    pos = start
    while pos + 8 <= end:
        size, kind = struct.unpack(">I4s", data[pos:pos + 8])
        header = 8
        if size == 1 and pos + 16 <= end:
            size = struct.unpack(">Q", data[pos + 8:pos + 16])[0]
            header = 16
        elif size == 0:
            size = end - pos
        if size < header or pos + size > end:
            return
        yield kind, pos + header, pos + size
        pos += size


def _data_text(data: bytes, start: int, end: int) -> str:
    for kind, body, stop in _atoms(data, start, end):
        if kind == b"data" and stop - body >= 8:
            return data[body + 8:stop].decode("utf-8", errors="replace").strip("\x00 \r\n")
    return ""


def _read_moov(path: Path) -> bytes:
    with open(path, "rb") as handle:
        while True:
            header = handle.read(8)
            if len(header) < 8:
                return b""
            size, kind = struct.unpack(">I4s", header)
            consumed = 8
            if size == 1:
                size = struct.unpack(">Q", handle.read(8))[0]
                consumed = 16
            if kind == b"moov":
                body = size - consumed if size else 0
                if 0 < body <= MAX_MOOV_BYTES:
                    return handle.read(body)
                return b""
            if size < consumed:
                return b""
            handle.seek(size - consumed, 1)


def read_m4b_tags(path) -> dict:
    """Read title, author, narrator and similar tags from an M4B. Best effort; never raises."""
    tags: dict = {}
    try:
        moov = _read_moov(Path(path))
        for kind, start, end in _atoms(moov):
            if kind != b"udta":
                continue
            for inner, istart, iend in _atoms(moov, start, end):
                if inner != b"meta":
                    continue
                for item, tstart, tend in _atoms(moov, istart + 4, iend):
                    if item != b"ilst":
                        continue
                    for name, bstart, bend in _atoms(moov, tstart, tend):
                        if name in TEXT_ATOMS:
                            value = _data_text(moov, bstart, bend)
                            if value:
                                tags.setdefault(TEXT_ATOMS[name], value)
                        elif name == b"----":
                            label, value = "", ""
                            for part, pstart, pend in _atoms(moov, bstart, bend):
                                if part == b"name":
                                    label = moov[pstart + 4:pend].decode("utf-8", errors="replace")
                            value = _data_text(moov, bstart, bend)
                            if label.upper() == "ASIN" and value:
                                tags["asin"] = value
    except Exception:
        return tags
    return tags


def describe_audiobook(path: Path, root: Path | None = None) -> dict:
    """Combine file tags with what the Libation folder and file names say."""
    path = Path(path)
    tags = read_m4b_tags(path)
    asin = tags.get("asin", "")
    match = ASIN_PATTERN.search(path.stem) or ASIN_PATTERN.search(path.parent.name)
    if not asin and match:
        asin = match.group(1)
    title = tags.get("title") or tags.get("album") or ASIN_PATTERN.sub("", path.stem).strip(" -_")
    author = tags.get("author") or tags.get("album_artist") or ""
    if not author and root is not None:
        try:
            parts = path.relative_to(root).parts
            if len(parts) > 2:
                author = parts[0]
        except ValueError:
            pass
    year_match = re.search(r"\b(1[5-9]\d{2}|20\d{2})\b", tags.get("year", ""))
    return {
        "path": str(path), "title": title.strip(), "author": author.strip(),
        "narrator": tags.get("narrator", ""), "year": year_match.group(1) if year_match else "",
        "genre": tags.get("genre", ""), "asin": asin,
    }


def scan_audiobook_folder(folder) -> list[dict]:
    root = Path(folder)
    return [describe_audiobook(path, root) for path in sorted(root.rglob("*.m4b")) if path.is_file()]


def audiobook_note(item: dict) -> str:
    parts = []
    if item["narrator"]:
        parts.append(f"Narrated by {item['narrator']}.")
    if item["asin"]:
        parts.append(f"Audible ASIN {item['asin']}.")
    return " ".join(parts)


def _append_note(existing: str, addition: str) -> str:
    if not addition or addition in (existing or ""):
        return existing or ""
    return (existing + "\n" if existing else "") + addition


def import_audiobooks(db, items: list[dict], books: list | None = None) -> dict:
    """Link audiobooks into the library. Returns counts of added, linked and skipped."""
    books = list(books if books is not None else db.all_books_for_duplicate_check())
    result = {"added": 0, "linked": 0, "already_linked": 0, "already_have": 0}
    known_paths = {row[8] for row in books}
    for item in items:
        if item["path"] in known_paths:
            result["already_linked"] += 1
            continue
        entry = {"title": item["title"], "author": item["author"], "isbn": "", "isbn10": ""}
        row = find_match(entry, books)
        if row is not None and is_wishlist_path(row[8]):
            db.link_book_file(row[0], item["path"])
            db.update_book(row[0], row[1], row[2], row[3], merge_tags(row[4], [AUDIOBOOK_TAG]),
                           _append_note(row[5], audiobook_note(item)))
            known_paths.add(item["path"])
            result["linked"] += 1
        elif row is not None:
            result["already_have"] += 1
        else:
            book_id = db.add_book(item["title"], item["author"], "Libation", "", "", item["path"], item["path"])
            db.update_book(book_id, item["title"], item["author"], "Libation", AUDIOBOOK_TAG, audiobook_note(item))
            if item["year"]:
                db.update_extra_fields(book_id, year=item["year"])
            known_paths.add(item["path"])
            result["added"] += 1
    return result
