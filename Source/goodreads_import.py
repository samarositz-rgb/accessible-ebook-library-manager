"""Bring a Goodreads library export into the local library.

Goodreads no longer offers a public programming interface, so the supported
route is the CSV export from the Goodreads website (My Books, Import and
export, Export Library). This module reads that file and, for books already in
the library, records the rating, shelves, reading status and review. Nothing
is sent to Goodreads and no book files are touched.
"""

from __future__ import annotations

import csv
import re

NOTE_MARKER = "Goodreads:"
STATUS_LABELS = {"read": "read", "currently-reading": "currently reading", "to-read": "want to read"}


def clean_isbn(value: str) -> str:
    """Goodreads writes ISBNs as ="0441172717"; keep only the digits."""
    return re.sub(r"[^0-9Xx]", "", value or "").upper()


def title_key(value: str) -> str:
    text = re.sub(r"\(.*?\)", " ", (value or "").casefold())
    text = text.split(":")[0].strip()
    text = re.sub(r"^(the|a|an)\s+", "", text)
    text = re.sub(r",\s*(the|a|an)$", "", text)
    return re.sub(r"[^a-z0-9]+", " ", text).strip()


def author_key(value: str) -> str:
    """Last name of the first author, whether written 'First Last' or 'Last, First'."""
    first = re.split(r"\band\b|;|&", value or "", maxsplit=1)[0]
    if "," in first:
        return re.sub(r"[^a-z]", "", first.split(",")[0].casefold().split()[-1]) if first.split(",")[0].split() else ""
    parts = re.sub(r"[^a-z ]+", " ", first.casefold()).split()
    return parts[-1] if parts else ""


def parse_goodreads_csv(path) -> list[dict]:
    """Read a Goodreads export. Raises ValueError when it is not one."""
    with open(path, newline="", encoding="utf-8-sig", errors="replace") as handle:
        reader = csv.DictReader(handle)
        fields = set(reader.fieldnames or [])
        if not {"Title", "Author"} <= fields:
            raise ValueError("This does not look like a Goodreads library export. It needs Title and Author columns.")
        rows = []
        for raw in reader:
            title = (raw.get("Title") or "").strip()
            if not title:
                continue
            shelves = [s.strip() for s in (raw.get("Bookshelves") or "").split(",") if s.strip()]
            status = (raw.get("Exclusive Shelf") or "").strip()
            if status and status not in shelves:
                shelves.append(status)
            rating = (raw.get("My Rating") or "0").strip()
            rows.append({
                "title": title,
                "author": (raw.get("Author") or "").strip(),
                "isbn": clean_isbn(raw.get("ISBN13") or "") or clean_isbn(raw.get("ISBN") or ""),
                "isbn10": clean_isbn(raw.get("ISBN") or ""),
                "rating": int(rating) if rating.isdigit() else 0,
                "status": status,
                "shelves": shelves,
                "date_read": (raw.get("Date Read") or "").strip(),
                "review": (raw.get("My Review") or "").strip(),
            })
        return rows


def build_note(entry: dict) -> str:
    parts = []
    status = STATUS_LABELS.get(entry["status"], entry["status"])
    if status:
        parts.append(status)
    if entry["rating"]:
        parts.append(f"rated {entry['rating']} of 5")
    if entry["date_read"]:
        parts.append(f"finished {entry['date_read']}")
    line = f"{NOTE_MARKER} " + ", ".join(parts) + "." if parts else ""
    if entry["review"]:
        review = re.sub(r"<br\s*/?>", " ", entry["review"])
        line = (line + " " if line else f"{NOTE_MARKER} ") + "Review: " + re.sub(r"\s+", " ", review).strip()
    return line


def merge_note(existing: str, new_line: str) -> str:
    """Replace an earlier Goodreads line, keeping all of the user's own notes."""
    kept = [line for line in (existing or "").splitlines() if not line.startswith(NOTE_MARKER)]
    if new_line:
        kept.append(new_line)
    return "\n".join(line for line in kept if line.strip())


def merge_tags(existing: str, shelves: list[str]) -> str:
    tags = [t.strip() for t in (existing or "").split(",") if t.strip()]
    known = {t.casefold() for t in tags}
    for shelf in shelves:
        if shelf.casefold() not in known:
            tags.append(shelf)
            known.add(shelf.casefold())
    return ", ".join(tags)


def find_match(entry: dict, books: list):
    """Match a Goodreads row to a library row by ISBN, then by title and author."""
    isbns = {i for i in (entry["isbn"], entry["isbn10"]) if i}
    for row in books:
        isbn = clean_isbn(row[12])
        if isbn and isbn in isbns:
            return row
    wanted_title = title_key(entry["title"])
    wanted_author = author_key(entry["author"])
    for row in books:
        if title_key(row[1]) == wanted_title and (not wanted_author or author_key(row[2]) == wanted_author):
            return row
    return None


def apply_goodreads_entries(db, entries: list[dict], books: list | None = None) -> dict:
    """Update matching books. Returns counts of updated, unchanged and unmatched."""
    books = books if books is not None else db.all_books_for_duplicate_check()
    result = {"updated": 0, "unchanged": 0, "unmatched": 0, "unmatched_titles": [], "unmatched_entries": []}
    for entry in entries:
        row = find_match(entry, books)
        if row is None:
            result["unmatched"] += 1
            result["unmatched_titles"].append(entry["title"])
            result["unmatched_entries"].append(entry)
            continue
        tags = merge_tags(row[4], entry["shelves"])
        notes = merge_note(row[5], build_note(entry))
        if tags == (row[4] or "") and notes == (row[5] or ""):
            result["unchanged"] += 1
            continue
        db.update_book(row[0], row[1], row[2], row[3], tags, notes)
        result["updated"] += 1
    return result


WISHLIST_TAG = "No File Available"
LEGACY_WISHLIST_TAG = "Wishlist"
WISHLIST_PREFIX = "wishlist:goodreads:"


def is_wishlist_path(stored_path) -> bool:
    return str(stored_path or "").startswith(WISHLIST_PREFIX)


def wishlist_key(entry: dict) -> str:
    """A unique stand-in for the missing book file, stable across repeated imports."""
    ident = entry["isbn"] or f"{title_key(entry['title'])}-{author_key(entry['author'])}".replace(" ", "-")
    return WISHLIST_PREFIX + ident


def add_wishlist_entries(db, entries: list[dict]) -> int:
    """Add books that have no file yet, tagged No File Available. Returns how many were added."""
    added = 0
    for entry in entries:
        key = wishlist_key(entry)
        try:
            book_id = db.add_book(entry["title"], entry["author"], "Goodreads", "", "", "", key)
        except Exception:
            continue  # already present from an earlier import
        # add_book only keeps standard category tags, so set the shelves directly.
        tags = merge_tags("", [WISHLIST_TAG] + entry["shelves"])
        db.update_book(book_id, entry["title"], entry["author"], "Goodreads", tags, build_note(entry))
        if entry["isbn"]:
            db.update_extra_fields(book_id, isbn=entry["isbn"])
        added += 1
    return added


def write_unmatched_file(path, entries: list[dict]) -> None:
    """Write a plain text list of books to look for elsewhere."""
    lines = [f"Goodreads books not in your library: {len(entries)}", ""]
    for entry in sorted(entries, key=lambda e: title_key(e["title"])):
        line = f"{entry['title']} by {entry['author']}"
        extras = []
        if entry["isbn"]:
            extras.append(f"ISBN {entry['isbn']}")
        status = STATUS_LABELS.get(entry["status"], entry["status"])
        if status:
            extras.append(status)
        if entry["rating"]:
            extras.append(f"rated {entry['rating']} of 5")
        if extras:
            line += " (" + ", ".join(extras) + ")"
        lines.append(line)
    with open(path, "w", encoding="utf-8", newline="\r\n") as handle:
        handle.write("\n".join(lines) + "\n")


def rename_legacy_wishlist_tags(db) -> int:
    """Rename the old Wishlist tag to No File Available. Returns how many books changed."""
    changed = 0
    for row in db.all_books_for_duplicate_check():
        tags = [t.strip() for t in (row[4] or "").split(",") if t.strip()]
        if LEGACY_WISHLIST_TAG.casefold() in {t.casefold() for t in tags} and is_wishlist_path(row[8]):
            renamed = [WISHLIST_TAG if t.casefold() == LEGACY_WISHLIST_TAG.casefold() else t for t in tags]
            db.update_book(row[0], row[1], row[2], row[3], merge_tags("", renamed), row[5])
            changed += 1
    return changed
