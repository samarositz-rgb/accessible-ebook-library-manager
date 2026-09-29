"""The supported categories for newly imported or discovered metadata."""

import re


COMMON_TAGS = ("Fiction", "Non-fiction", "Textbooks", "Reference", "Biography and Memoir",
               "History", "Science and Technology", "Law", "Business", "Other")


def normalize_import_tags(value):
    known = {tag.casefold(): tag for tag in COMMON_TAGS}
    aliases = {"nonfiction": "non-fiction", "non fiction": "non-fiction",
               "textbook": "textbooks", "biography": "biography and memoir",
               "memoir": "biography and memoir", "science & technology": "science and technology"}
    parts = value if isinstance(value, (list, tuple)) else [value or ""]
    result = []
    for part in parts:
        for tag in re.split(r"[,;/]", str(part)):
            key = re.sub(r"\s+", " ", tag.strip()).casefold()
            canonical = known.get(aliases.get(key, key))
            if canonical and canonical not in result:
                result.append(canonical)
    return ", ".join(result)
