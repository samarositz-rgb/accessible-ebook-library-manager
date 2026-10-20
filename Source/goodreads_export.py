"""Write the library as a CSV file that Goodreads can import.

On Goodreads choose My Books, Import and export, then Import. Goodreads does not
allow other programs to add books directly, so a file you upload yourself is
the only supported route. Ratings, reading status, dates and reviews that came
from an earlier Goodreads import are written back out.
"""

from __future__ import annotations

import csv
import re

from goodreads_import import NOTE_MARKER, STATUS_LABELS, WISHLIST_TAG, clean_isbn

COLUMNS = [
    "Title", "Author", "ISBN", "ISBN13", "My Rating", "Publisher", "Year Published",
    "Date Read", "Bookshelves", "Exclusive Shelf", "My Review", "Private Notes",
]
STATUS_BY_LABEL = {label: key for key, label in STATUS_LABELS.items()}
SKIP_SHELVES = {WISHLIST_TAG.casefold(), "audiobook"}


def parse_goodreads_note(notes: str) -> dict:
    """Read back the status, rating, date and review from the note the importer wrote."""
    info = {"status": "", "rating": "", "date_read": "", "review": "", "other": []}
    for line in (notes or "").splitlines():
        if not line.startswith(NOTE_MARKER):
            if line.strip():
                info["other"].append(line.strip())
            continue
        body = line[len(NOTE_MARKER):].strip()
        review = ""
        if "Review:" in body:
            body, review = body.split("Review:", 1)
        info["review"] = review.strip()
        for label, key in STATUS_BY_LABEL.items():
            if re.search(rf"(^|[ ,]){re.escape(label)}(?=[,.]|$)", body):
                info["status"] = key
                break
        rated = re.search(r"rated (\d) of 5", body)
        if rated:
            info["rating"] = rated.group(1)
        finished = re.search(r"finished ([0-9/\-]+)", body)
        if finished:
            info["date_read"] = finished.group(1).rstrip(".")
    return info


def shelf_name(tag: str) -> str:
    return re.sub(r"[^a-z0-9\-]+", "-", tag.strip().casefold()).strip("-")


def export_row(row, include_without_status: bool = True) -> dict | None:
    """Build one CSV row from a library row, or None when it should be left out."""
    note = parse_goodreads_note(row[5])
    status = note["status"]
    if not status and str(row[8] or "").startswith("wishlist:"):
        status = "to-read"
    if not status and not include_without_status:
        return None
    isbn = clean_isbn(row[12])
    shelves = []
    for tag in (row[4] or "").split(","):
        name = shelf_name(tag)
        if name and name not in SKIP_SHELVES and name not in STATUS_LABELS and name not in shelves:
            shelves.append(name)
    return {
        "Title": row[1], "Author": row[2],
        "ISBN": isbn if len(isbn) == 10 else "", "ISBN13": isbn if len(isbn) == 13 else "",
        "My Rating": note["rating"] or "0", "Publisher": row[13] or "",
        "Year Published": row[11] or "", "Date Read": note["date_read"],
        "Bookshelves": ", ".join(shelves), "Exclusive Shelf": status,
        "My Review": note["review"], "Private Notes": " ".join(note["other"]),
    }


def export_goodreads_csv(path, rows, include_without_status: bool = True) -> int:
    """Write the CSV. Returns how many books were written."""
    written = 0
    with open(path, "w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=COLUMNS)
        writer.writeheader()
        for row in rows:
            record = export_row(row, include_without_status)
            if record:
                writer.writerow(record)
                written += 1
    return written
