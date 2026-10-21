"""Convert audiobooks to a single chaptered M4B file.

This needs FFmpeg, a free command line tool that is not part of this app. The
original files are never changed or deleted: the M4B is written to a separate
output folder. Folders of MP3 or M4A parts become one M4B with a chapter for each
part. A single M4A is copied into an M4B without re-encoding. ZIP files from
NLS BARD or Learning Ally are read without being extracted permanently.
"""

from __future__ import annotations

import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from audiobook_scan import DAISY_AUDIO_EXTENSIONS, read_id3_tags, read_m4b_tags
from library_utils import safe_filename

FFMPEG_SETTING = "ffmpeg_path"
OUTPUT_FOLDER_SETTING = "audiobook_output_folder"
DEFAULT_BITRATE = "64k"
PART_EXTENSIONS = {".mp3", ".m4a", ".wav", ".ogg", ".wma", ".mp2"}
CREATE_NO_WINDOW = 0x08000000 if sys.platform.startswith("win") else 0

INSTALL_HELP = (
    "FFmpeg was not found. It is a free tool that does the audio conversion.\n\n"
    "To install it on Windows, open PowerShell and type: winget install Gyan.FFmpeg\n"
    "Then close and reopen this app. You can also choose ffmpeg.exe yourself if it is already on your computer."
)


def find_ffmpeg(configured: str = "") -> str | None:
    """Locate ffmpeg: the saved path first, then the PATH and the usual install folders."""
    if configured and Path(configured).is_file():
        return configured
    found = shutil.which("ffmpeg")
    if found:
        return found
    local = os.environ.get("LOCALAPPDATA", "")
    candidates = [
        Path(local) / "Microsoft" / "WinGet" / "Links" / "ffmpeg.exe",
        Path(os.environ.get("ProgramData", "C:/ProgramData")) / "chocolatey" / "bin" / "ffmpeg.exe",
        Path("C:/ffmpeg/bin/ffmpeg.exe"),
        Path(os.environ.get("ProgramFiles", "C:/Program Files")) / "ffmpeg" / "bin" / "ffmpeg.exe",
    ]
    for candidate in candidates:
        if candidate.is_file():
            return str(candidate)
    if local:
        packages = Path(local) / "Microsoft" / "WinGet" / "Packages"
        if packages.is_dir():
            for candidate in packages.glob("*FFmpeg*/**/ffmpeg.exe"):
                return str(candidate)
    return None


def can_convert(stored_path) -> bool:
    """True for audiobooks this can turn into M4B: folders of parts, MP3 or M4A files, audio ZIPs."""
    path = Path(str(stored_path or ""))
    if not path.exists():
        return False
    if path.is_dir():
        return True
    return path.suffix.lower() in {".mp3", ".m4a", ".zip"}


def collect_parts(source: Path, scratch: Path) -> list[Path]:
    """Return the audio parts of a book in listening order. ZIP contents go in the scratch folder."""
    if source.is_dir():
        files = [p for p in source.rglob("*") if p.is_file() and p.suffix.lower() in PART_EXTENSIONS]
    elif source.suffix.lower() == ".zip":
        files = []
        with zipfile.ZipFile(source) as archive:
            for info in archive.infolist():
                if Path(info.filename).suffix.lower() in DAISY_AUDIO_EXTENSIONS and not info.is_dir():
                    target = scratch / Path(info.filename).name
                    counter = 2
                    while target.exists():
                        target = scratch / f"{Path(info.filename).stem}_{counter}{Path(info.filename).suffix}"
                        counter += 1
                    with archive.open(info) as src, open(target, "wb") as dst:
                        shutil.copyfileobj(src, dst)
                    files.append(target)
    else:
        files = [source]
    return sorted(files, key=lambda p: [int(t) if t.isdigit() else t.casefold() for t in re.split(r"(\d+)", str(p))])


def part_duration(ffmpeg: str, path: Path) -> float:
    """Length of an audio file in seconds, read from ffmpeg's own report."""
    result = subprocess.run([ffmpeg, "-hide_banner", "-i", str(path)], capture_output=True, text=True,
                            check=False, creationflags=CREATE_NO_WINDOW)
    match = re.search(r"Duration:\s*(\d+):(\d+):(\d+(?:\.\d+)?)", result.stderr)
    if not match:
        return 0.0
    hours, minutes, seconds = match.groups()
    return int(hours) * 3600 + int(minutes) * 60 + float(seconds)


def chapter_title(path: Path, number: int) -> str:
    tags = read_id3_tags(path) if path.suffix.lower() == ".mp3" else read_m4b_tags(path)
    title = tags.get("title", "") if tags else ""
    return title or re.sub(r"[_]+", " ", path.stem).strip() or f"Chapter {number}"


def _escape_meta(value: str) -> str:
    return re.sub(r"([=;#\\\n])", r"\\\1", value or "")


def metadata_text(title: str, author: str, narrator: str, year: str, chapters: list[tuple[int, int, str]]) -> str:
    lines = [";FFMETADATA1", f"title={_escape_meta(title)}", f"album={_escape_meta(title)}"]
    if author:
        lines += [f"artist={_escape_meta(author)}", f"album_artist={_escape_meta(author)}"]
    if narrator:
        lines.append(f"composer={_escape_meta(narrator)}")
    if year:
        lines.append(f"date={_escape_meta(year)}")
    lines.append("genre=Audiobook")
    for start, end, name in chapters:
        lines += ["[CHAPTER]", "TIMEBASE=1/1000", f"START={start}", f"END={end}", f"title={_escape_meta(name)}"]
    return "\n".join(lines) + "\n"


def concat_list_text(parts: list[Path]) -> str:
    return "".join("file '" + str(part).replace("\\", "/").replace("'", "'\\''") + "'\n" for part in parts)


def unique_output(folder: Path, author: str, title: str) -> Path:
    base = safe_filename(f"{author} - {title}" if author else title).strip(" -.") or "Audiobook"
    candidate = folder / f"{base}.m4b"
    counter = 2
    while candidate.exists():
        candidate = folder / f"{base} ({counter}).m4b"
        counter += 1
    return candidate


def run_ffmpeg(command: list[str], cancel_check=None) -> tuple[int, str]:
    """Run ffmpeg, stopping it promptly if the person cancels."""
    process = subprocess.Popen(command, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE, text=True,
                               encoding="utf-8", errors="replace", creationflags=CREATE_NO_WINDOW)
    errors = []
    import threading

    def drain():
        for line in process.stderr:
            errors.append(line)
            del errors[:-40]
    reader = threading.Thread(target=drain, daemon=True)
    reader.start()
    while process.poll() is None:
        if cancel_check and cancel_check():
            process.kill()
            process.wait()
            reader.join(timeout=2)
            process.stderr.close()
            return -1, "Cancelled"
        try:
            process.wait(timeout=0.5)
        except subprocess.TimeoutExpired:
            pass
    reader.join(timeout=5)
    process.stderr.close()
    return process.returncode, "".join(errors)[-800:]


def convert_to_m4b(ffmpeg: str, source, output_folder, title: str, author: str = "", narrator: str = "",
                   year: str = "", bitrate: str = DEFAULT_BITRATE, cancel_check=None, progress=None) -> Path:
    """Convert one audiobook to a chaptered M4B and return the new file's path.

    Raises RuntimeError with a plain message when the conversion cannot be done.
    """
    source = Path(source)
    output_folder = Path(output_folder)
    output_folder.mkdir(parents=True, exist_ok=True)
    destination = unique_output(output_folder, author, title)
    with tempfile.TemporaryDirectory(prefix="aelm-m4b-") as scratch_name:
        scratch = Path(scratch_name)
        parts = collect_parts(source, scratch)
        if not parts:
            raise RuntimeError("No audio files were found in this book.")
        if len(parts) == 1 and parts[0].suffix.lower() == ".m4a":
            command = [ffmpeg, "-y", "-hide_banner", "-loglevel", "error", "-i", str(parts[0]), "-map", "0:a",
                       "-map_metadata", "0", "-map_chapters", "0", "-c:a", "copy", "-f", "ipod", str(destination)]
        else:
            if progress:
                progress(f"Reading {len(parts)} parts")
            chapters, cursor = [], 0
            for number, part in enumerate(parts, start=1):
                length = int(part_duration(ffmpeg, part) * 1000)
                if length <= 0:
                    raise RuntimeError(f"FFmpeg could not read {part.name}.")
                chapters.append((cursor, cursor + length, chapter_title(part, number)))
                cursor += length
            (scratch / "parts.txt").write_text(concat_list_text(parts), encoding="utf-8")
            (scratch / "meta.txt").write_text(metadata_text(title, author, narrator, year, chapters), encoding="utf-8")
            command = [ffmpeg, "-y", "-hide_banner", "-loglevel", "error", "-f", "concat", "-safe", "0",
                       "-i", str(scratch / "parts.txt"), "-i", str(scratch / "meta.txt"), "-map", "0:a",
                       "-map_metadata", "1", "-map_chapters", "1", "-c:a", "aac", "-b:a", bitrate,
                       "-f", "ipod", str(destination)]
        if progress:
            progress("Writing M4B")
        code, errors = run_ffmpeg(command, cancel_check)
    if code != 0 or not destination.is_file():
        destination.unlink(missing_ok=True)
        if errors == "Cancelled":
            raise RuntimeError("Cancelled")
        raise RuntimeError((errors.strip() or "FFmpeg did not create the M4B file.")[:500])
    return destination
