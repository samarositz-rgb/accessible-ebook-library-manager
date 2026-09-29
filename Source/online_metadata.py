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


def lookup_online_metadata(existing: dict, service_filter=None, cancel_check=None) -> tuple[dict, str]:
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
    limited_services = set()
    for service, url in queries:
        if cancel_check and cancel_check():
            return {}, ""
        if service_filter and service != service_filter:
            continue
        if service in limited_services:
            continue
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
        except urllib.error.HTTPError as exc:
            if exc.code == 429:
                limited_services.add(service)
                label = "Google Books" if service == "google" else "Open Library"
                last_error = f"{label} is temporarily limiting requests. Your metadata is unchanged. Continue to review the available results, or try again later."
            else:
                last_error = str(exc)
            exc.close()
        except Exception as exc:
            last_error = str(exc)
            continue

    if last_error:
        raise RuntimeError(last_error)
    return {}, ""
