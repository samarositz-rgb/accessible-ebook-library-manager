"""User-defined destinations for sending books to preferred readers.

Alongside the built-in Send To targets (Voice Dream, Kindle, NLS, …) the
user can define custom destinations: a name, a folder, an optional target
format, and an optional post-copy notice. Sending a book to a destination
optionally converts it to the destination's target format via Calibre's
``ebook-convert`` before copying, so a book stored as PDF can be delivered
as EPUB to a reader that prefers EPUB, all from one keystroke.

Destinations are stored as a JSON list in the ``settings`` table under
the key ``custom_destinations`` — no schema migration needed.
"""

from __future__ import annotations

import json
import shutil
import subprocess
import tempfile
from dataclasses import asdict, dataclass, field
from pathlib import Path

from calibre_tools import WINDOWS_NO_CONSOLE_FLAGS, find_calibre_tool


SETTING_KEY = "custom_destinations"


@dataclass
class Destination:
    name: str
    path: str
    target_format: str = ""      # e.g. "epub", "pdf", "mobi"; "" means no conversion
    post_copy_notice: str = ""   # text to also drop in the folder as a .txt file

    def normalized_format(self) -> str:
        fmt = self.target_format.strip().lower().lstrip(".")
        return fmt


def load_destinations(db) -> list[Destination]:
    raw = db.get_setting(SETTING_KEY, "")
    if not raw:
        return []
    try:
        parsed = json.loads(raw)
    except (json.JSONDecodeError, ValueError):
        return []
    result: list[Destination] = []
    if not isinstance(parsed, list):
        return result
    for entry in parsed:
        if not isinstance(entry, dict):
            continue
        name = str(entry.get("name", "")).strip()
        path = str(entry.get("path", "")).strip()
        if not name or not path:
            continue
        result.append(Destination(
            name=name,
            path=path,
            target_format=str(entry.get("target_format", "")).strip(),
            post_copy_notice=str(entry.get("post_copy_notice", "")).strip(),
        ))
    return result


def save_destinations(db, destinations: list[Destination]) -> None:
    payload = json.dumps([asdict(d) for d in destinations], ensure_ascii=False)
    db.set_setting(SETTING_KEY, payload)


def _run_ebook_convert(ebook_convert: str, source: Path, target: Path) -> None:
    """Invoke Calibre's ebook-convert. Raises RuntimeError on failure."""
    proc = subprocess.run(
        [ebook_convert, str(source), str(target)],
        capture_output=True,
        text=True,
        creationflags=WINDOWS_NO_CONSOLE_FLAGS,
        timeout=600,
    )
    if proc.returncode != 0:
        detail = (proc.stderr or proc.stdout or "").strip().splitlines()[-1:]
        message = detail[0] if detail else f"exit code {proc.returncode}"
        raise RuntimeError(f"ebook-convert failed: {message}")


def prepare_for_destination(
    source: Path,
    target_format: str,
    ebook_convert_path: str | None,
) -> tuple[Path, Path | None]:
    """Return the file to copy plus a temp dir to clean up (or None).

    If ``target_format`` is empty or the source already matches, no
    conversion happens and the source itself is returned. If conversion
    is required but Calibre is unavailable, RuntimeError is raised so
    the caller can surface the error before overwriting a destination.
    """
    fmt = target_format.strip().lower().lstrip(".")
    if not fmt:
        return source, None
    source_ext = source.suffix.lower().lstrip(".")
    if source_ext == fmt:
        return source, None
    if not ebook_convert_path:
        raise RuntimeError(
            f"Calibre's ebook-convert is required to convert {source.suffix} to .{fmt}, "
            "but it was not found. Install Calibre or clear the destination's target format."
        )
    temp_dir = Path(tempfile.mkdtemp(prefix="ebooklib_convert_"))
    target = temp_dir / f"{source.stem}.{fmt}"
    try:
        _run_ebook_convert(ebook_convert_path, source, target)
    except Exception:
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise
    if not target.exists():
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise RuntimeError("ebook-convert reported success but produced no output file.")
    return target, temp_dir


def send_book_to_destination(
    source: Path,
    destination: Destination,
    *,
    ebook_convert_path: str | None = None,
    overwrite: bool = False,
) -> Path:
    """Copy (and optionally convert) ``source`` into ``destination``.

    Returns the final destination path. Raises FileExistsError if the
    target already exists and ``overwrite`` is False, FileNotFoundError
    if the destination folder is missing, or RuntimeError if conversion
    is needed but Calibre is unavailable / fails.
    """
    dest_folder = Path(destination.path)
    if not dest_folder.exists() or not dest_folder.is_dir():
        raise FileNotFoundError(f"Destination folder does not exist: {dest_folder}")

    prepared, temp_dir = prepare_for_destination(
        source, destination.target_format, ebook_convert_path
    )
    try:
        target = dest_folder / prepared.name
        if target.exists() and not overwrite:
            raise FileExistsError(str(target))
        shutil.copy2(prepared, target)
        _write_post_copy_notice(dest_folder, destination.post_copy_notice)
        return target
    finally:
        if temp_dir is not None:
            shutil.rmtree(temp_dir, ignore_errors=True)


def _write_post_copy_notice(folder: Path, notice: str) -> None:
    text = (notice or "").strip()
    if not text:
        return
    notice_path = folder / "IMPORTANT NOTICE.txt"
    try:
        # Only write it once per folder — do not stomp on an existing user file.
        if notice_path.exists():
            return
        notice_path.write_text(text, encoding="utf-8")
    except OSError:
        # Non-fatal; the copy already succeeded.
        pass


def get_ebook_convert_path() -> str | None:
    return find_calibre_tool("ebook-convert")
