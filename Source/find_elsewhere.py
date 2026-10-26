"""Find a book that is not in the library on free and print-disabled services.

Project Gutenberg can be searched directly because it needs no account. The
other services (Bookshare, NLS BARD, Learning Ally and similar) need a member
login, so the app opens their site for the person to search while signed in.
Nothing is downloaded and no account details are ever handled here.
"""

from __future__ import annotations

import re
import urllib.parse
import urllib.request
from xml.etree import ElementTree as ET

try:
    from defusedxml.ElementTree import fromstring as _safe_fromstring
except ImportError:  # defusedxml is optional
    _safe_fromstring = ET.fromstring

SETTING_KEY = "find_elsewhere_services"  # free and accessible services (kept for older libraries)
LIBRARY_KEY_SETTING = "libby_library_key"
GROUP_SETTINGS = {
    "free": SETTING_KEY,
    "commercial": "find_commercial_services",
    "library": "find_library_services",
}
GROUP_TITLES = {
    "free": "Free and accessible services",
    "commercial": "Commercial services",
    "library": "Your local library",
}

# key: (label, search address with {q} for the encoded search words, or "" when the
# site has no public search address, site address, needs a member account, group)
SERVICES = {
    "gutenberg": ("Project Gutenberg", "https://www.gutenberg.org/ebooks/search/?query={q}", "https://www.gutenberg.org", False, "free"),
    "bookshare": ("Bookshare", "https://www.bookshare.org/search?keyword={q}", "https://www.bookshare.org", True, "free"),
    "bard": ("NLS BARD", "", "https://nlsbard.loc.gov", True, "free"),
    "learningally": ("Learning Ally", "", "https://learningally.org", True, "free"),
    "archive": ("Internet Archive", "https://archive.org/search?query={q}", "https://archive.org", False, "free"),
    "standardebooks": ("Standard Ebooks", "https://standardebooks.org/ebooks?query={q}", "https://standardebooks.org", False, "free"),
    "openlibrary": ("Open Library", "https://openlibrary.org/search?q={q}", "https://openlibrary.org", False, "free"),
    "hathitrust": ("HathiTrust", "https://babel.hathitrust.org/cgi/ls?q1={q}", "https://www.hathitrust.org", True, "free"),
    "kindle": ("Kindle Store on Amazon", "https://www.amazon.com/s?k={q}&i=digital-text", "https://www.amazon.com", False, "commercial"),
    "applebooks": ("Apple Books", "https://books.apple.com/us/search?term={q}", "https://books.apple.com", False, "commercial"),
    "kobo": ("Kobo", "https://www.kobo.com/us/en/search?query={q}", "https://www.kobo.com", False, "commercial"),
    "googleplay": ("Google Play Books", "https://play.google.com/store/search?q={q}&c=books", "https://play.google.com/store/books", False, "commercial"),
    "barnesnoble": ("Barnes and Noble", "https://www.barnesandnoble.com/s/{q}", "https://www.barnesandnoble.com", False, "commercial"),
    "bookshop": ("Bookshop.org", "https://bookshop.org/search?keywords={q}", "https://bookshop.org", False, "commercial"),
    "libby": ("Libby and OverDrive", "", "https://libbyapp.com", True, "library"),
    "hoopla": ("Hoopla", "https://www.hoopladigital.com/search?q={q}", "https://www.hoopladigital.com", True, "library"),
}


def services_in_group(group: str) -> list[str]:
    return [key for key, value in SERVICES.items() if value[4] == group]


def enabled_services(setting_value: str, group: str = "free") -> list[str]:
    """Services chosen in settings for a group, in a fixed order. An unset value means all."""
    members = services_in_group(group)
    if not (setting_value or "").strip():
        return members
    chosen = {part.strip() for part in setting_value.split(",")}
    return [key for key in members if key in chosen]


def search_words(title: str, author: str = "") -> str:
    """Title without a subtitle or series note, plus the author's last name."""
    clean = re.sub(r"\(.*?\)", " ", title or "").split(":")[0].strip()
    words = clean
    first_author = re.split(r"\band\b|;|&", author or "", maxsplit=1)[0].strip()
    if first_author:
        if "," in first_author:
            last = first_author.split(",")[0].strip()
        else:
            last = first_author.split()[-1] if first_author.split() else ""
        words = f"{clean} {last}".strip()
    return re.sub(r"\s+", " ", words)


def service_target(key: str, words: str, library_key: str = "") -> tuple[str, str, bool]:
    """Return (address to open, label, whether the search words must be pasted by hand)."""
    label, template, site, _account, _group = SERVICES[key]
    if key == "libby" and library_key.strip():
        name = re.sub(r"[^a-z0-9\-]", "", library_key.strip().casefold())
        return f"https://{name}.overdrive.com/search?query={urllib.parse.quote_plus(words)}", label, False
    if template:
        return template.format(q=urllib.parse.quote_plus(words)), label, False
    return site, label, True


def parse_gutenberg_feed(xml_text: str, limit: int = 5) -> list[dict]:
    """Read results from a Project Gutenberg OPDS search feed."""
    ns = {"a": "http://www.w3.org/2005/Atom"}
    root = _safe_fromstring(xml_text.encode("utf-8") if isinstance(xml_text, str) else xml_text)
    results = []
    for entry in root.findall("a:entry", ns):
        ident = (entry.findtext("a:id", default="", namespaces=ns) or "")
        match = re.search(r"/ebooks/(\d+)", ident)
        if not match:
            continue
        results.append({
            "title": (entry.findtext("a:title", default="", namespaces=ns) or "").strip(),
            "author": (entry.findtext("a:content", default="", namespaces=ns) or "").strip(),
            "url": f"https://www.gutenberg.org/ebooks/{match.group(1)}",
        })
        if len(results) >= limit:
            break
    return results


def search_gutenberg(title: str, author: str = "", timeout: int = 10) -> list[dict]:
    """Search Project Gutenberg. Raises OSError-family errors when offline."""
    url = "https://www.gutenberg.org/ebooks/search.opds/?query=" + urllib.parse.quote_plus(search_words(title, author))
    request = urllib.request.Request(url, headers={"User-Agent": "AccessibleEbookLibraryManager/1.0"})
    with urllib.request.urlopen(request, timeout=timeout) as response:
        return parse_gutenberg_feed(response.read())
