"""Build-time provisioning only. End users never download a conversion engine."""

import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import urllib.request
import zipfile


ROOT = Path(__file__).resolve().parent
CACHE = ROOT / "build" / "braille-downloads"
RUNTIME = ROOT / "vendor" / "braille"
VERSION = "pipeline-ui-1.12.0-temurin-17.0.20.1"
ASSETS = {
    "pipeline-setup.exe": (
        "https://github.com/daisy/pipeline-ui/releases/download/1.12.0/daisy-pipeline-setup-1.12.0.exe",
        "91eca67facffcaaa8fcb120227932ae2d1310c6819194e5a5b83cdb4fb55a3b0",
    ),
    "7zr.exe": (
        "https://github.com/ip7z/7zip/releases/download/26.03/7zr.exe",
        "ad4c82fadcbdf93c03b4fc440f300509c7d60c5c2f4d183e35d9d70d6957037d",
    ),
    "7z-setup.exe": (
        "https://github.com/ip7z/7zip/releases/download/26.03/7z2603-x64.exe",
        "0859c524b8a63551848f0c246abddcb1d0b7b656b0fbfe879f8d85e61a9e6edd",
    ),
    "java.zip": (
        "https://github.com/adoptium/temurin17-binaries/releases/download/jdk-17.0.20.1%2B1/OpenJDK17U-jre_x64_windows_hotspot_17.0.20.1_1.zip",
        "bc21a93923103cdaac93ee337b0ae4365e739fde36df823dd456bc67c8a9d352",
    ),
}


def digest(path):
    with path.open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


def download(name):
    url, expected = ASSETS[name]
    target = CACHE / name
    if not target.is_file() or digest(target) != expected:
        partial = target.with_suffix(".part")
        print(f"Downloading built-in braille dependency: {name}", flush=True)
        with urllib.request.urlopen(url, timeout=90) as response, partial.open("wb") as output:
            shutil.copyfileobj(response, output)
        if digest(partial) != expected:
            raise RuntimeError(f"Checksum verification failed: {name}")
        partial.replace(target)
    return target


def main():
    marker = RUNTIME / "bundle.json"
    if marker.is_file() and json.loads(marker.read_text())["version"] == VERSION:
        manifest = json.loads(marker.read_text())["files"]
        if all((RUNTIME / name).is_file() and digest(RUNTIME / name) == value for name, value in manifest.items()):
            print("Bundled braille engine verified.")
            return
        raise RuntimeError("Braille runtime was changed or damaged; restore it before building.")
    CACHE.mkdir(parents=True, exist_ok=True)
    for name in ASSETS:
        download(name)
    subprocess.run([str(CACHE / "7zr.exe"), "x", str(CACHE / "7z-setup.exe"), "-o" + str(CACHE / "7zip"), "-y"], check=True, capture_output=True)
    seven = str(CACHE / "7zip/7z.exe")
    subprocess.run([seven, "x", str(CACHE / "pipeline-setup.exe"), "-o" + str(CACHE / "installer"), "$PLUGINSDIR\\app-64.7z", "-y"], check=True, capture_output=True)
    stage = ROOT / "build" / "b"
    subprocess.run([seven, "x", str(CACHE / "installer/$PLUGINSDIR/app-64.7z"), "-o" + str(stage), "resources\\app.asar.unpacked\\resources\\daisy-pipeline\\*", "-y"], check=True, capture_output=True)
    shutil.copytree(stage / "resources/app.asar.unpacked/resources/daisy-pipeline", RUNTIME, dirs_exist_ok=True)
    java_stage = CACHE / "java"
    with zipfile.ZipFile(CACHE / "java.zip") as archive:
        archive.extractall(java_stage)
    java = next(java_stage.glob("*/bin/java.exe")).parent.parent
    # Replace only our staged JRE, never a system Java installation.
    target = (RUNTIME / "jre").resolve()
    if not target.is_relative_to(RUNTIME.resolve()):
        raise RuntimeError("Invalid runtime staging path")
    shutil.rmtree(target)
    shutil.copytree(java, target)
    files = {str(path.relative_to(RUNTIME)).replace("\\", "/"): digest(path) for path in RUNTIME.rglob("*") if path.is_file() and path != marker}
    marker.write_text(json.dumps({"version": VERSION, "downloads": ASSETS, "files": files}, indent=2), encoding="utf-8")
    print("Built-in braille engine prepared and verified.")


if __name__ == "__main__":
    main()
