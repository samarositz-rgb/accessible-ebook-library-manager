import os
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import audiobook_convert as conv
import audiobook_scan as ab

FFMPEG = conv.find_ffmpeg(os.environ.get("FFMPEG_TEST_PATH", ""))


class PureHelperTests(unittest.TestCase):
    def test_metadata_text_has_chapters_and_escapes(self):
        text = conv.metadata_text("A=B", "Author", "Narrator", "2001", [(0, 1500, "One"), (1500, 3000, "Two; part")])
        self.assertIn("title=A\\=B", text)
        self.assertIn("START=1500", text)
        self.assertIn("title=Two\\; part", text)
        self.assertIn("composer=Narrator", text)

    def test_concat_list_escapes_quotes(self):
        text = conv.concat_list_text([Path("C:/a/It's.mp3")])
        self.assertEqual(text, "file 'C:/a/It'\\''s.mp3'\n")

    def test_can_convert(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp) / "Book"
            folder.mkdir()
            (folder / "a.mp3").write_bytes(b"x")
            (Path(tmp) / "b.m4b").write_bytes(b"x")
            self.assertTrue(conv.can_convert(folder))
            self.assertTrue(conv.can_convert(folder / "a.mp3"))
            self.assertFalse(conv.can_convert(Path(tmp) / "b.m4b"))
            self.assertFalse(conv.can_convert(Path(tmp) / "missing.mp3"))

    def test_unique_output_never_overwrites(self):
        with tempfile.TemporaryDirectory() as tmp:
            first = conv.unique_output(Path(tmp), "A", "T")
            first.write_bytes(b"x")
            self.assertEqual(conv.unique_output(Path(tmp), "A", "T").name, "A - T (2).m4b")

    def test_missing_ffmpeg_is_reported_as_none(self):
        self.assertIsNone(conv.find_ffmpeg("C:/nope/ffmpeg.exe") if not shutil.which("ffmpeg") and FFMPEG is None else None)


@unittest.skipUnless(FFMPEG, "FFmpeg is not installed; set FFMPEG_TEST_PATH to run the conversion tests")
class ConversionTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="conv_")
        self.addCleanup(lambda: shutil.rmtree(self.tmp, ignore_errors=True))
        self.root = Path(self.tmp)

    def tone(self, path: Path, seconds=2, title=""):
        path.parent.mkdir(parents=True, exist_ok=True)
        command = [FFMPEG, "-y", "-loglevel", "error", "-f", "lavfi", "-i", f"sine=frequency=440:duration={seconds}"]
        if title:
            command += ["-metadata", f"title={title}"]
        subprocess.run(command + [str(path)], check=True)

    def test_folder_of_mp3_becomes_one_m4b_with_chapters(self):
        for number, name in ((1, "Opening"), (2, "Middle"), (10, "Ending")):
            self.tone(self.root / "Book" / f"{number:02d}.mp3", title=name)
        out = conv.convert_to_m4b(FFMPEG, self.root / "Book", self.root / "out", "My Book", "An Author", "A Voice", "1999")
        self.assertEqual(out.suffix, ".m4b")
        self.assertTrue(out.stat().st_size > 0)
        tags = ab.read_m4b_tags(out)
        self.assertEqual((tags["title"], tags["author"], tags["narrator"]), ("My Book", "An Author", "A Voice"))
        report = subprocess.run([FFMPEG, "-hide_banner", "-i", str(out)], capture_output=True, text=True).stderr
        self.assertEqual(report.count("Chapter #"), 3)
        for name in ("Opening", "Middle", "Ending"):
            self.assertIn(name, report)
        self.assertTrue((self.root / "Book" / "01.mp3").exists())

    def test_single_m4a_is_copied_to_m4b(self):
        source = self.root / "Solo.m4a"
        self.tone(source)
        out = conv.convert_to_m4b(FFMPEG, source, self.root / "out", "Solo", "A")
        self.assertTrue(out.is_file())
        self.assertTrue(source.exists())

    def test_zip_of_mp3s_converts_and_leaves_zip_alone(self):
        self.tone(self.root / "src" / "track1.mp3")
        self.tone(self.root / "src" / "track2.mp3")
        archive = self.root / "Book_DB123.zip"
        with zipfile.ZipFile(archive, "w") as z:
            z.write(self.root / "src" / "track1.mp3", "track1.mp3")
            z.write(self.root / "src" / "track2.mp3", "track2.mp3")
            z.writestr("ncc.html", "<html></html>")
        before = archive.read_bytes()
        out = conv.convert_to_m4b(FFMPEG, archive, self.root / "out", "Zipped", "Z")
        self.assertTrue(out.is_file())
        self.assertEqual(archive.read_bytes(), before)

    def test_cancel_stops_and_leaves_no_file(self):
        self.tone(self.root / "Book" / "1.mp3", seconds=3)
        self.tone(self.root / "Book" / "2.mp3", seconds=3)
        with self.assertRaisesRegex(RuntimeError, "Cancelled"):
            conv.convert_to_m4b(FFMPEG, self.root / "Book", self.root / "out", "C", cancel_check=lambda: True)
        self.assertEqual(list((self.root / "out").glob("*.m4b")), [])

    def test_empty_folder_gives_plain_error(self):
        (self.root / "Empty").mkdir()
        with self.assertRaisesRegex(RuntimeError, "No audio files"):
            conv.convert_to_m4b(FFMPEG, self.root / "Empty", self.root / "out", "E")


if __name__ == "__main__":
    unittest.main()
