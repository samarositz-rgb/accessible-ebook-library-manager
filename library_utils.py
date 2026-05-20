import re
import shutil
from pathlib import Path


def folder_file_stats(folder: Path) -> tuple[int, int]:
    if not folder or not folder.exists():
        return 0, 0
    file_count = 0
    byte_count = 0
    for path in folder.rglob("*"):
        if not path.is_file():
            continue
        try:
            byte_count += path.stat().st_size
            file_count += 1
        except OSError:
            continue
    return file_count, byte_count


def files_need_copy(source: Path, destination: Path) -> bool:
    if not destination.exists():
        return True
    try:
        source_stat = source.stat()
        destination_stat = destination.stat()
    except OSError:
        return True
    return (
        source_stat.st_size != destination_stat.st_size
        or int(source_stat.st_mtime) != int(destination_stat.st_mtime)
    )


def sync_folder_contents(source_folder: Path, destination_folder: Path):
    source_folder.mkdir(parents=True, exist_ok=True)
    destination_folder.mkdir(parents=True, exist_ok=True)
    source_files = set()

    for source_path in source_folder.rglob("*"):
        relative = source_path.relative_to(source_folder)
        destination_path = destination_folder / relative
        if source_path.is_dir():
            destination_path.mkdir(parents=True, exist_ok=True)
            continue
        if not source_path.is_file():
            continue
        source_files.add(relative)
        destination_path.parent.mkdir(parents=True, exist_ok=True)
        if files_need_copy(source_path, destination_path):
            shutil.copy2(source_path, destination_path)

    for destination_path in sorted(destination_folder.rglob("*"), key=lambda item: len(item.parts), reverse=True):
        relative = destination_path.relative_to(destination_folder)
        if destination_path.is_file() and relative not in source_files:
            destination_path.unlink()
        elif destination_path.is_dir():
            try:
                destination_path.rmdir()
            except OSError:
                pass


def replace_folder_from_backup(backup_folder: Path, target_folder: Path):
    if not backup_folder.exists() or not backup_folder.is_dir():
        return
    target_folder.mkdir(parents=True, exist_ok=True)
    sync_folder_contents(backup_folder, target_folder)


def safe_filename(text: str) -> str:
    bad = '<>:"/\\|?*'
    cleaned = "".join("_" if ch in bad else ch for ch in text).strip()
    return cleaned or "Untitled"


def normalize_duplicate_key(text: str) -> str:
    text = (text or "").casefold()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    text = re.sub(r"\b(?:the|a|an)\b", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def normalize_isbn_key(value: str) -> str:
    value = re.sub(r"[^0-9Xx]", "", value or "")
    return value.casefold() if len(value) in {10, 13} else ""


def title_match_tokens(value: str) -> set[str]:
    normalized = normalize_duplicate_key(value)
    ignored = {
        "bookshare", "daisy", "accessible", "ebook", "epub", "pdf", "doc", "docx",
        "edition", "ed", "textbook", "book", "copy", "owner", "vaio",
    }
    return {token for token in normalized.split() if len(token) > 1 and token not in ignored}


def title_keys_look_same(left: str, right: str, allow_richer_metadata_match: bool = False) -> bool:
    left_key = normalize_duplicate_key(left)
    right_key = normalize_duplicate_key(right)
    if not left_key or not right_key:
        return False
    if left_key == right_key:
        return True
    if len(left_key) >= 8 and len(right_key) >= 8 and (left_key in right_key or right_key in left_key):
        return True

    left_tokens = title_match_tokens(left)
    right_tokens = title_match_tokens(right)
    if len(left_tokens) < 3 or len(right_tokens) < 3:
        return False
    overlap = len(left_tokens & right_tokens)
    smaller_ratio = overlap / min(len(left_tokens), len(right_tokens))
    larger_ratio = overlap / max(len(left_tokens), len(right_tokens))
    if smaller_ratio >= 0.9 and larger_ratio >= 0.7:
        return True
    return allow_richer_metadata_match and smaller_ratio >= 0.8 and larger_ratio >= 0.6


def metadata_score_from_values(author="", source="", tags="", notes="", edition="", year="", isbn="", publisher="") -> int:
    score = 0
    score += 3 if (author or "").strip() else 0
    score += 2 if edition else 0
    score += 2 if year else 0
    score += 2 if normalize_isbn_key(isbn or "") else 0
    score += 2 if publisher else 0
    score += 1 if source else 0
    score += 1 if tags else 0
    score += 1 if notes else 0
    return score


def metadata_score_from_detection(metadata: dict) -> int:
    return metadata_score_from_values(
        metadata.get("author", ""),
        metadata.get("source", ""),
        metadata.get("tags", ""),
        metadata.get("notes", ""),
        metadata.get("edition", ""),
        metadata.get("year", ""),
        metadata.get("isbn", ""),
        metadata.get("publisher", ""),
    )


def metadata_score_from_row(row) -> int:
    return metadata_score_from_values(row[2], row[3], row[4], row[5], row[10], row[11], row[12], row[13])
