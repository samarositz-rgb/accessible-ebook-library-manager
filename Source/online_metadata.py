import json
import re
import urllib.parse
import urllib.request
import urllib.error
from book_categories import normalize_import_tags


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
        "series": metadata.get("series", ""),
        "series_number": metadata.get("series_number", ""),
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
    result["tags"] = normalize_import_tags(result["tags"])
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


def metadata_from_open_library_data(record: dict) -> dict:
    """Convert an Open Library ``api/books`` data record (looked up by ISBN)."""
    identifiers = record.get("identifiers") or {}
    isbns = (identifiers.get("isbn_13") or []) + (identifiers.get("isbn_10") or [])
    cover = record.get("cover") or {}
    return normalize_online_metadata({
        "title": record.get("title", ""),
        "author": ", ".join(a.get("name", "") for a in record.get("authors") or []),
        "year": record.get("publish_date", ""),
        "isbn": isbns[0] if isbns else "",
        "publisher": ", ".join(p.get("name", "") for p in (record.get("publishers") or [])[:1]),
        "tags": ", ".join(s.get("name", "") for s in (record.get("subjects") or [])[:8]),
        "cover_url": cover.get("large") or cover.get("medium") or "",
    })


FILL_FIELDS = ("author", "year", "isbn", "publisher", "tags", "notes", "cover_url")
SERVICE_LABELS = {"openlibrary": "Open Library", "google": "Google Books", "openlibrary_data": "Open Library ISBN record"}


def fill_missing_fields(primary: dict, extra: dict) -> dict:
    """Return primary with any blank fields filled from extra."""
    merged = dict(primary)
    for key in FILL_FIELDS:
        if not merged.get(key) and extra.get(key):
            merged[key] = extra[key]
    return merged


def build_queries(existing: dict) -> list:
    isbn = re.sub(r"[^0-9Xx]", "", existing.get("isbn", ""))
    title = existing.get("title", "").strip()
    author = existing.get("author", "").strip()
    queries = []
    if isbn:
        queries.append(("openlibrary", f"https://openlibrary.org/search.json?isbn={urllib.parse.quote(isbn)}&limit=1"))
        queries.append(("google", f"https://www.googleapis.com/books/v1/volumes?q=isbn:{urllib.parse.quote(isbn)}&maxResults=1"))
        queries.append(("openlibrary_data", f"https://openlibrary.org/api/books?bibkeys=ISBN:{urllib.parse.quote(isbn)}&jscmd=data&format=json"))
    if title:
        open_query = f"title:{title}"
        google_query = f"intitle:{title}"
        if author:
            open_query += f" author:{author}"
            google_query += f"+inauthor:{author}"
        queries.append(("openlibrary", "https://openlibrary.org/search.json?" + urllib.parse.urlencode({"q": open_query, "limit": "1"})))
        queries.append(("google", "https://www.googleapis.com/books/v1/volumes?" + urllib.parse.urlencode({"q": google_query, "maxResults": "1"})))
    return queries


def parse_service_result(service: str, data: dict) -> dict:
    if service == "openlibrary":
        docs = data.get("docs") or []
        return metadata_from_open_library_doc(docs[0]) if docs else {}
    if service == "openlibrary_data":
        records = [value for value in data.values() if isinstance(value, dict)]
        return metadata_from_open_library_data(records[0]) if records else {}
    items = data.get("items") or []
    return metadata_from_google_volume(items[0]) if items else {}


def lookup_online_metadata(existing: dict, service_filter=None, cancel_check=None, fill_gaps=True) -> tuple[dict, str]:
    """Look a book up online.

    The first source that finds the book supplies the result. Unless a single
    service is requested, blank fields are then filled from the other sources
    and the returned label names every source that contributed.
    """
    last_error = ""
    limited_services = set()
    result: dict = {}
    contributed: list = []
    tried_services = set()
    for service, url in build_queries(existing):
        if cancel_check and cancel_check():
            return {}, ""
        if service_filter and service != service_filter:
            continue
        if service in limited_services or service in contributed:
            continue
        if result and not fill_gaps:
            break
        if result and all(result.get(key) for key in FILL_FIELDS):
            break
        try:
            found = parse_service_result(service, fetch_json(url))
        except urllib.error.HTTPError as exc:
            if exc.code == 429:
                limited_services.add(service)
                label = SERVICE_LABELS.get(service, service)
                last_error = f"{label} is temporarily limiting requests. Your metadata is unchanged. Continue to review the available results, or try again later."
            else:
                last_error = str(exc)
            exc.close()
            continue
        except Exception as exc:
            last_error = str(exc)
            continue
        tried_services.add(service)
        if not found:
            continue
        if not result:
            result = found
        else:
            result = fill_missing_fields(result, found)
        contributed.append(service)

    if result:
        return result, " and ".join(SERVICE_LABELS.get(service, service) for service in contributed)
    if last_error:
        raise RuntimeError(last_error)
    return {}, ""
