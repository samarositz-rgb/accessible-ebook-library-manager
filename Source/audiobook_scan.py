"""Find audiobooks and link them into the library.

Recognised: M4B and M4A files (such as a Libation library), folders of MP3 or M4A
parts that make up one book, and ZIP files from NLS BARD and Learning Ally.
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


AUDIO_EXTENSIONS = {".m4b", ".m4a", ".mp3"}
SINGLE_FILE_EXTENSIONS = {".m4b"}
PART_FOLDER_PATTERN = re.compile(r"^(cd|disc|disk|part|volume|vol)\s*\.?\s*\d+", re.IGNORECASE)
DAISY_AUDIO_EXTENSIONS = {".mp3", ".m4a", ".m4b", ".wav", ".mp2", ".ogg", ".wma"}
NLS_NUMBER_PATTERN = re.compile(r"\b((?:DB|RC|BR|LS|DBC)\d{3,7})\b", re.IGNORECASE)
KIND_LABELS = {"file": "audio file", "folder": "folder of audio files", "zip": "ZIP file"}


# ---------------------------------------------------------------- MP3 and similar (ID3) tags

def _decode_id3_text(raw: bytes) -> str:
    if not raw:
        return ""
    encoding, body = raw[0], raw[1:]
    codec = {0: "latin-1", 1: "utf-16", 2: "utf-16-be", 3: "utf-8"}.get(encoding, "latin-1")
    return body.decode(codec, errors="replace").strip("\x00 \r\n").split("\x00")[0].strip()


def read_id3_tags(path) -> dict:
    """Read title, artist, album, composer and year from an ID3v2 tag. Never raises."""
    tags: dict = {}
    wanted3 = {b"TIT2": "title", b"TPE1": "author", b"TALB": "album", b"TCOM": "narrator", b"TYER": "year", b"TDRC": "year"}
    wanted2 = {b"TT2": "title", b"TP1": "author", b"TAL": "album", b"TCM": "narrator", b"TYE": "year"}
    try:
        with open(path, "rb") as handle:
            header = handle.read(10)
            if len(header) < 10 or header[:3] != b"ID3":
                return tags
            version = header[3]
            size = sum((header[6 + i] & 0x7F) << (7 * (3 - i)) for i in range(4))
            data = handle.read(min(size, 4 * 1024 * 1024))
        pos = 0
        if version == 2:
            while pos + 6 <= len(data):
                frame_id = data[pos:pos + 3]
                length = int.from_bytes(data[pos + 3:pos + 6], "big")
                if not frame_id.strip(b"\x00") or length <= 0:
                    break
                if frame_id in wanted2:
                    tags.setdefault(wanted2[frame_id], _decode_id3_text(data[pos + 6:pos + 6 + length]))
                pos += 6 + length
        else:
            while pos + 10 <= len(data):
                frame_id = data[pos:pos + 4]
                if version == 4:
                    length = sum((data[pos + 4 + i] & 0x7F) << (7 * (3 - i)) for i in range(4))
                else:
                    length = int.from_bytes(data[pos + 4:pos + 8], "big")
                if not frame_id.strip(b"\x00") or length <= 0:
                    break
                if frame_id in wanted3:
                    tags.setdefault(wanted3[frame_id], _decode_id3_text(data[pos + 10:pos + 10 + length]))
                pos += 10 + length
    except Exception:
        return tags
    return {key: value for key, value in tags.items() if value}


# ---------------------------------------------------------------- DAISY style ZIP files (NLS BARD, Learning Ally)

def _meta_values(text: str) -> dict:
    values: dict = {}
    for match in re.finditer(r"<meta\b[^>]*>", text, re.IGNORECASE):
        tag = match.group(0)
        name = re.search(r'name\s*=\s*["\']([^"\']+)["\']', tag, re.IGNORECASE)
        content = re.search(r'content\s*=\s*["\']([^"\']*)["\']', tag, re.IGNORECASE)
        if name and content:
            values.setdefault(name.group(1).casefold(), content.group(1).strip())
    for tag_name in ("title", "creator", "identifier", "date", "publisher", "narrator"):
        element = re.search(rf"<(?:dc:)?{tag_name}\b[^>]*>([^<]+)</(?:dc:)?{tag_name}>", text, re.IGNORECASE)
        if element:
            values.setdefault(f"dc:{tag_name}", element.group(1).strip())
    return values


def inspect_audio_zip(path) -> dict | None:
    """Return details when a ZIP is an audiobook or DAISY talking book, otherwise None.

    The ZIP is only read, never extracted or changed.
    """
    import zipfile
    try:
        with zipfile.ZipFile(path) as archive:
            names = archive.namelist()
            lowered = [name.casefold() for name in names]
            audio = [n for n in lowered if Path(n).suffix in DAISY_AUDIO_EXTENSIONS]
            daisy = [n for n in lowered if Path(n).name in {"ncc.html", "ncc.htm", "navigation.xml"} or n.endswith((".smil", ".opf"))]
            if not audio and not daisy:
                return None
            if any(n.endswith((".epub",)) or Path(n).name in {"mimetype"} for n in lowered) and not audio:
                return None
            meta: dict = {}
            for name in names:
                base = Path(name).name.casefold()
                if base in {"ncc.html", "ncc.htm"} or base.endswith(".opf") or base == "bookinfo.xml":
                    try:
                        meta.update({k: v for k, v in _meta_values(archive.read(name)[:300000].decode("utf-8", errors="replace")).items() if k not in meta})
                    except Exception:
                        continue
            return {"meta": meta, "audio_files": len(audio), "has_daisy": bool(daisy)}
    except Exception:
        return None


def _nls_or_learning_ally(path: Path, meta: dict) -> tuple[str, str]:
    """Return (source name, catalogue number) for a ZIP from NLS BARD or Learning Ally."""
    haystack = " ".join([path.stem, meta.get("dc:identifier", ""), meta.get("dc:publisher", ""), meta.get("ncc:sourcepublisher", "")]).casefold()
    number = NLS_NUMBER_PATTERN.search(path.stem) or NLS_NUMBER_PATTERN.search(meta.get("dc:identifier", ""))
    if "learning ally" in haystack or "learningally" in haystack or "rfb&d" in haystack or "rfbd" in haystack or "recording for the blind" in haystack:
        return "Learning Ally", number.group(1).upper() if number else meta.get("dc:identifier", "")
    if number or "national library service" in haystack or "nls" in haystack.split():
        return "NLS BARD", number.group(1).upper() if number else ""
    return "Audiobook", meta.get("dc:identifier", "")


# ---------------------------------------------------------------- describing items

def _year(text: str) -> str:
    found = re.search(r"\b(1[5-9]\d{2}|20\d{2})\b", text or "")
    return found.group(1) if found else ""


def _author_from_folders(path: Path, root: Path | None, depth_needed: int) -> str:
    if root is None:
        return ""
    try:
        parts = path.relative_to(root).parts
    except ValueError:
        return ""
    return parts[0] if len(parts) > depth_needed else ""


def describe_audiobook(path: Path, root: Path | None = None) -> dict:
    """A single M4B or M4A file: combine file tags with the Libation folder and file names."""
    path = Path(path)
    tags = read_m4b_tags(path)
    asin = tags.get("asin", "")
    match = ASIN_PATTERN.search(path.stem) or ASIN_PATTERN.search(path.parent.name)
    if not asin and match:
        asin = match.group(1)
    title = tags.get("title") or tags.get("album") or ASIN_PATTERN.sub("", path.stem).strip(" -_")
    author = tags.get("author") or tags.get("album_artist") or _author_from_folders(path, root, 2)
    return {
        "kind": "file", "path": str(path), "title": title.strip(), "author": author.strip(),
        "narrator": tags.get("narrator", ""), "year": _year(tags.get("year", "")),
        "genre": tags.get("genre", ""), "asin": asin, "source": "Libation" if asin else "Audiobook", "number": "",
    }


def describe_audio_folder(folder: Path, files: list[Path], root: Path | None = None) -> dict:
    """A folder of MP3 or M4A parts that make up one book."""
    folder = Path(folder)
    tags = {}
    for candidate in files:
        tags = read_id3_tags(candidate) if candidate.suffix.lower() == ".mp3" else read_m4b_tags(candidate)
        if tags:
            break
    name = folder.name
    if PART_FOLDER_PATTERN.match(name):
        name = folder.parent.name
    title = tags.get("album") or tags.get("title") or name
    author = tags.get("author") or tags.get("album_artist") or _author_from_folders(folder, root, 1)
    return {
        "kind": "folder", "path": str(folder), "title": title.strip(), "author": author.strip(),
        "narrator": tags.get("narrator", ""), "year": _year(tags.get("year", "")),
        "genre": tags.get("genre", ""), "asin": "", "source": "Audiobook", "number": "",
    }


def describe_audio_zip(path: Path, details: dict, root: Path | None = None) -> dict:
    """A ZIP downloaded from NLS BARD or Learning Ally (or any DAISY audio ZIP)."""
    path = Path(path)
    meta = details["meta"]
    source, number = _nls_or_learning_ally(path, meta)
    title = meta.get("dc:title") or meta.get("ncc:title") or NLS_NUMBER_PATTERN.sub("", path.stem).strip(" -_")
    author = meta.get("dc:creator") or meta.get("ncc:creator") or ""
    author = re.sub(r"\s*;\s*", " and ", author).strip()
    return {
        "kind": "zip", "path": str(path), "title": title.strip(), "author": author.strip(),
        "narrator": meta.get("dc:narrator", "") or meta.get("ncc:narrator", ""), "year": _year(meta.get("dc:date", "")),
        "genre": "", "asin": "", "source": source, "number": number,
    }


def scan_audiobook_folder(folder) -> list[dict]:
    """Find audiobooks under a folder or drive: M4B and M4A files, folders of MP3 or M4A parts, and audio ZIPs."""
    import os
    root = Path(folder)
    items: list[dict] = []
    claimed: list[Path] = []

    def under_claimed(directory: Path) -> bool:
        return any(directory == c or c in directory.parents for c in claimed)

    for current, dirs, files in os.walk(root):
        dirs.sort()
        directory = Path(current)
        if under_claimed(directory):
            dirs[:] = []
            continue
        audio = sorted(directory / f for f in files if Path(f).suffix.lower() in AUDIO_EXTENSIONS)
        singles = [p for p in audio if p.suffix.lower() in SINGLE_FILE_EXTENSIONS]
        parts = [p for p in audio if p.suffix.lower() not in SINGLE_FILE_EXTENSIONS]
        for single in singles:
            items.append(describe_audiobook(single, root))
        if len(parts) == 1 and not singles:
            items.append(describe_audiobook(parts[0], root))
        elif len(parts) >= 2:
            items.append(describe_audio_folder(directory, parts, root))
            claimed.append(directory)
            dirs[:] = []
        elif not audio:
            disc_dirs = [d for d in dirs if PART_FOLDER_PATTERN.match(d)]
            if disc_dirs and any(
                any(Path(f).suffix.lower() in {".mp3", ".m4a"} for f in os.listdir(directory / d)) for d in disc_dirs
            ):
                collected = []
                for d in disc_dirs:
                    collected += sorted((directory / d / f) for f in os.listdir(directory / d) if Path(f).suffix.lower() in {".mp3", ".m4a"})
                items.append(describe_audio_folder(directory, collected, root))
                claimed.append(directory)
                dirs[:] = []
                continue
        for name in sorted(f for f in files if f.lower().endswith(".zip")):
            details = inspect_audio_zip(directory / name)
            if details:
                items.append(describe_audio_zip(directory / name, details, root))
    return items


def audiobook_note(item: dict) -> str:
    parts = []
    if item["narrator"]:
        parts.append(f"Narrated by {item['narrator']}.")
    if item["asin"]:
        parts.append(f"Audible ASIN {item['asin']}.")
    if item.get("number"):
        parts.append(f"{item['source']} book number {item['number']}.")
    return " ".join(parts)


def _append_note(existing: str, addition: str) -> str:
    if not addition or addition in (existing or ""):
        return existing or ""
    return (existing + "\n" if existing else "") + addition


def import_audiobooks(db, items: list[dict], books: list | None = None, dry_run: bool = False) -> dict:
    """Link audiobooks into the library. Returns counts of added, linked and skipped.

    With dry_run the library is not changed, so the counts can be shown as a preview.
    A book that appears twice in the same scan, for example as an M4B and as a ZIP,
    is only added once.
    """
    books = [list(row) for row in (books if books is not None else db.all_books_for_duplicate_check())]
    result = {"added": 0, "linked": 0, "already_linked": 0, "already_have": 0}
    known_paths = {row[8] for row in books}
    for item in items:
        if item["path"] in known_paths:
            result["already_linked"] += 1
            continue
        entry = {"title": item["title"], "author": item["author"], "isbn": "", "isbn10": ""}
        row = find_match(entry, books)
        if row is not None and is_wishlist_path(row[8]):
            if not dry_run:
                db.link_book_file(row[0], item["path"], item_format(item))
                db.update_book(row[0], row[1], row[2], row[3], merge_tags(row[4], [AUDIOBOOK_TAG]),
                               _append_note(row[5], audiobook_note(item)))
            row[8] = item["path"]
            known_paths.add(item["path"])
            result["linked"] += 1
        elif row is not None:
            result["already_have"] += 1
        else:
            book_id = -1
            if not dry_run:
                book_id = db.add_book(item["title"], item["author"], item["source"], "", "", item["path"], item["path"])
                db.link_book_file(book_id, item["path"], item_format(item))
                db.update_book(book_id, item["title"], item["author"], item["source"], AUDIOBOOK_TAG, audiobook_note(item))
                if item["year"]:
                    db.update_extra_fields(book_id, year=item["year"])
            fresh = [None] * 23
            fresh[0], fresh[1], fresh[2], fresh[3], fresh[4], fresh[5] = book_id, item["title"], item["author"], item["source"], AUDIOBOOK_TAG, ""
            fresh[8], fresh[12] = item["path"], ""
            books.append(fresh)
            known_paths.add(item["path"])
            result["added"] += 1
    return result


def item_format(item: dict) -> str:
    if item["kind"] == "folder":
        return "audio folder"
    return Path(item["path"]).suffix.lower().lstrip(".")


def first_audio_file(folder) -> Path | None:
    """The first part of a folder audiobook, so a player opens the book from its start."""
    folder = Path(folder)
    candidates = sorted(p for p in folder.rglob("*") if p.is_file() and p.suffix.lower() in AUDIO_EXTENSIONS)
    return candidates[0] if candidates else None
