"""Local braille conversion using the bundled, unmodified DAISY Pipeline engine."""

from dataclasses import dataclass
from datetime import datetime, timezone
import base64
import hashlib
import hmac
import io
import os
from pathlib import Path
import posixpath
import re
import secrets
import socket
import subprocess
import sys
import tempfile
import time
import urllib.request
import urllib.parse
import urllib.error
import xml.etree.ElementTree as ET
from xml.dom import minidom
import zipfile


class ExportCancelled(Exception):
    pass


@dataclass(frozen=True)
class BrailleOptions:
    format: str = "brf"
    contracted: bool = True
    cells_per_line: int = 40
    lines_per_page: int = 25

    def __post_init__(self):
        if self.format not in {"brf", "ebraille"}:
            raise ValueError("Choose BRF or eBraille.")
        if type(self.contracted) is not bool or type(self.cells_per_line) is not int or type(self.lines_per_page) is not int:
            raise ValueError("Braille settings must use a grade choice and whole-number page dimensions.")
        if not 10 <= self.cells_per_line <= 120 or not 5 <= self.lines_per_page <= 100:
            raise ValueError("Braille page dimensions are outside the supported range.")

    @property
    def code(self):
        return '(liblouis-table:"en-ueb-g%d.ctb")' % (2 if self.contracted else 1)


def runtime_folder():
    return Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent)) / "vendor" / "braille"


def require_runtime():
    folder = runtime_folder()
    for relative in ("jre/bin/java.exe", "etc/pipeline.properties"):
        if not (folder / relative).is_file():
            raise RuntimeError("The built-in braille converter is missing. Please use the complete application folder.")
    return folder


def stop_process(process):
    if process is not None and process.poll() is None:
        process.terminate()
        try:
            process.wait(timeout=10)
        except subprocess.TimeoutExpired:
            process.kill()
            process.wait()


class BrailleEngine:
    """One private, authenticated loopback service for an export batch."""

    def __init__(self, cancelled=lambda: False):
        self.cancelled = cancelled
        self.process = None
        self.temp = None
        self.book_temp = None
        self.log = None
        self.warnings = []

    def check_cancelled(self):
        if self.cancelled():
            raise ExportCancelled("Export cancelled.")

    def __enter__(self):
        self.runtime = require_runtime()
        self.temp = tempfile.TemporaryDirectory(prefix="aelm-brl-")
        self.work = Path(self.temp.name)
        self.env = os.environ.copy()
        self.env["APPDATA"] = str(self.work)
        self.env.pop("JAVA_TOOL_OPTIONS", None)
        self.env.pop("_JAVA_OPTIONS", None)
        with socket.socket() as reservation:
            reservation.bind(("127.0.0.1", 0))
            port = reservation.getsockname()[1]
        key, secret = secrets.token_hex(16), secrets.token_hex(32)
        self.key, self.secret = key, secret
        self.base_url = f"http://127.0.0.1:{port}/ws"
        logging = self.work / "logging.xml"
        logging.write_text('<configuration><appender name="stderr" class="ch.qos.logback.core.ConsoleAppender"><target>System.err</target><encoder><pattern>%level %msg%n</pattern></encoder></appender><root level="WARN"><appender-ref ref="stderr"/></root></configuration>', encoding="utf-8")
        props = {
            "org.daisy.pipeline.ws.host": "127.0.0.1",
            "org.daisy.pipeline.ws.port": str(port),
            "org.daisy.pipeline.ws.localfs": "true",
            "org.daisy.pipeline.ws.authentication": "true",
            "org.daisy.pipeline.ws.authentication.key": key,
            "org.daisy.pipeline.ws.authentication.secret": secret,
            "org.daisy.pipeline.ws.cors": "false",
            "org.daisy.pipeline.ws.tmpdir": str(self.work),
            "org.daisy.pipeline.data": str(self.work / "data"),
            "org.daisy.pipeline.logdir": str(self.work),
            "java.io.tmpdir": str(self.work),
            "logback.configurationFile": logging.as_uri(),
        }
        command = [str(self.runtime / "jre/bin/java.exe"), "-Xmx768m"]
        for module in ("java.base/java.security", "java.base/java.net", "java.base/java.lang", "java.base/java.util", "java.naming/javax.naming.spi", "java.rmi/sun.rmi.transport.tcp"):
            command.extend(["--add-opens", module + "=ALL-UNNAMED"])
        for module in ("java.base/sun.net.www.protocol.http", "java.base/sun.net.www.protocol.https", "java.base/sun.net.www.protocol.jar", "jdk.xml.dom/org.w3c.dom.html", "jdk.naming.rmi/com.sun.jndi.url.rmi"):
            command.append("--add-exports=" + module + "=ALL-UNNAMED")
        command.extend(f"-D{name}={value}" for name, value in props.items())
        command.extend(["-cp", os.pathsep.join(str(self.runtime / part / "*") for part in ("system/common", "system/webservice")), "org.daisy.pipeline.webservice.restlet.impl.PipelineWebService"])
        self.log = (self.work / "engine.log").open("wb")
        try:
            self.process = subprocess.Popen(command, cwd=self.runtime, env=self.env, stdout=self.log, stderr=subprocess.STDOUT, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
            deadline = time.monotonic() + 90
            opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
            while time.monotonic() < deadline:
                self.check_cancelled()
                if self.process.poll() is not None:
                    raise RuntimeError("Braille engine could not start: " + (self.work / "engine.log").read_text(errors="replace")[-2000:])
                try:
                    with opener.open(f"http://127.0.0.1:{port}/ws/alive", timeout=1) as response:
                        if response.status == 200:
                            return self
                except OSError:
                    pass
                time.sleep(0.2)
            raise RuntimeError("The braille engine took too long to start.")
        except BaseException:
            self.__exit__(None, None, None)
            raise

    def __exit__(self, *_args):
        stop_process(self.process)
        if self.log:
            self.log.close()
        if self.book_temp:
            self.book_temp.cleanup()
        if self.temp:
            self.temp.cleanup()

    def request(self, path, data=None, method=None, timeout=60):
        self.check_cancelled()
        timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        uri = f"{self.base_url}/{path}?authid={self.key}&time={timestamp}&nonce={secrets.randbelow(10**30):030d}"
        signature = base64.b64encode(hmac.new(self.secret.encode(), uri.encode(), hashlib.sha1).digest()).decode()
        uri += "&sign=" + urllib.parse.quote_plus(signature)
        request = urllib.request.Request(uri, data=data, method=method, headers={"Content-Type": "application/xml"})
        opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
        try:
            with opener.open(request, timeout=timeout) as response:
                return response.read()
        except urllib.error.HTTPError as exc:
            details = exc.read().decode("utf-8", errors="replace")
            try:
                root = ET.fromstring(details)
                details = " ".join(text.strip() for text in root.itertext() if text.strip())
            except ET.ParseError:
                pass
            raise RuntimeError(f"Braille converter error {exc.code}: {details[:1000]}") from None

    def run_script(self, script, source, options=None, timeout=1800):
        ns = "http://www.daisy.org/ns/pipeline/data"
        root = ET.Element(f"{{{ns}}}jobRequest")
        ET.SubElement(root, f"{{{ns}}}script", href=f"{self.base_url}/scripts/{script}")
        inp = ET.SubElement(root, f"{{{ns}}}input", name="source")
        ET.SubElement(inp, f"{{{ns}}}item", value=Path(source).resolve().as_uri())
        for name, value in (options or {}).items():
            ET.SubElement(root, f"{{{ns}}}option", name=name).text = str(value)
        job = ET.fromstring(self.request("jobs", ET.tostring(root, encoding="utf-8", xml_declaration=True)))
        job_id = job.get("id")
        if not job_id or not all(c.isalnum() or c == "-" for c in job_id):
            raise RuntimeError("The converter did not return a valid job identifier.")
        deadline = time.monotonic() + timeout
        try:
            while True:
                self.check_cancelled()
                if time.monotonic() > deadline:
                    raise RuntimeError("Braille conversion timed out.")
                job = ET.fromstring(self.request(f"jobs/{job_id}"))
                status = job.get("status")
                if status == "SUCCESS":
                    for message in job.iter(f"{{{ns}}}message"):
                        content = message.get("content", "").strip()
                        if message.get("level") == "WARN" and content and content not in self.warnings:
                            self.warnings.append(content[:1000])
                    break
                if status in {"ERROR", "FAIL"}:
                    messages = [m.get("content", "") for m in job.iter(f"{{{ns}}}message") if m.get("level") in {"ERROR", "WARN"}]
                    raise RuntimeError("; ".join(messages)[-2000:] or f"{script} conversion failed.")
                time.sleep(0.3)
            output = Path(self.book_temp.name) / job_id if self.book_temp else self.work / job_id
            output.mkdir()
            with zipfile.ZipFile(io.BytesIO(self.request(f"jobs/{job_id}/result"))) as archive:
                for entry in archive.infolist():
                    target = (output / entry.filename).resolve()
                    if not target.is_relative_to(output.resolve()):
                        raise RuntimeError("Unsafe filename in conversion output.")
                archive.extractall(output)
            return output
        finally:
            try:
                # Cancellation must not prevent removal of this batch's temporary job.
                old_cancelled = self.cancelled
                self.cancelled = lambda: False
                self.request(f"jobs/{job_id}", method="DELETE", timeout=5)
            except (OSError, RuntimeError):
                pass
            finally:
                self.cancelled = old_cancelled

    def convert(self, source, options):
        """Return files valid until the next conversion or engine shutdown."""
        self.check_cancelled()
        self.warnings = []
        if self.book_temp:
            self.book_temp.cleanup()
            self.book_temp = None
        source = Path(source)
        if not source.is_file():
            raise ValueError("The stored book file is missing.")
        suffix = ".brf" if options.format == "brf" else ".ebrl"
        if source.suffix.lower() == suffix:
            return [source]
        if source.suffix.lower() != ".epub":
            raise ValueError("Braille conversion currently accepts EPUB books. Convert this book to EPUB first.")
        # Java caches some source ZIP handles until shutdown. Remove unlocked
        # outputs now; the enclosing engine directory removes the rest on exit.
        self.book_temp = tempfile.TemporaryDirectory(prefix="book-", dir=self.work, ignore_cleanup_errors=True)
        prepared = Path(self.book_temp.name) / "source.epub"
        version = prepare_epub(source, prepared, self.check_cancelled)
        if version.startswith("2") and options.format == "ebraille":
            upgraded = self.run_script("epub2-to-epub3", prepared)
            candidates = list(upgraded.rglob("*.epub"))
            if len(candidates) != 1:
                raise RuntimeError("The converter could not prepare this older EPUB for eBraille.")
            prepare_epub(candidates[0], prepared, self.check_cancelled)
        settings = {"braille-code": options.code}
        if options.format == "brf":
            settings.update({
                "output-file-format": "embossed AND (-daisy-locale: en-US)",
                "page-width": str(options.cells_per_line),
                "page-height": str(options.lines_per_page),
            })
            script = "epub3-to-pef"
        else:
            settings["epub-package"] = "true"
            script = "epub3-to-ebraille"
        output = self.run_script(script, prepared, settings)
        files = sorted(output.rglob("*" + suffix), key=lambda p: [
            (0, int(part)) if part.isdigit() else (1, part.casefold())
            for part in re.split(r"(\d+)", str(p.relative_to(output)))
        ])
        if not files or (options.format == "ebraille" and len(files) != 1):
            raise RuntimeError("The converter did not produce the expected braille file.")
        for path in files:
            self.check_cancelled()
            validate_output(path, options)
        return files


def prepare_epub(source, destination, check_cancelled=lambda: None):
    """Normalize language attributes on a private copy, preserving book content."""
    with zipfile.ZipFile(source) as original:
        if sum(item.file_size for item in original.infolist()) > 1024 ** 3:
            raise ValueError("This EPUB is too large for braille conversion (over 1 GB unpacked).")
        for entry in original.infolist():
            parts = entry.filename.replace("\\", "/").split("/")
            if ".." in parts or entry.filename.startswith(("/", "\\")) or ":" in entry.filename:
                raise ValueError("This EPUB contains an unsafe file path.")
        container = ET.fromstring(original.read("META-INF/container.xml"))
        rootfile = container.find(".//{*}rootfile")
        if rootfile is None:
            raise ValueError("This EPUB has no package document.")
        package = ET.fromstring(original.read(rootfile.get("full-path")))
        language = package.findtext(".//{http://purl.org/dc/elements/1.1/}language") or "en"
        # Pipeline's XML translation requires xml:lang even for EPUBs using HTML lang.
        with zipfile.ZipFile(destination, "w") as output:
            for entry in original.infolist():
                check_cancelled()
                if not entry.filename.lower().endswith((".xhtml", ".html", ".htm")):
                    with original.open(entry) as incoming, output.open(entry, "w") as outgoing:
                        while chunk := incoming.read(1024 * 1024):
                            check_cancelled()
                            outgoing.write(chunk)
                    continue
                content = original.read(entry)
                if entry.filename.lower().endswith((".xhtml", ".html", ".htm")):
                    if b"<!ENTITY" in content.upper():
                        raise ValueError("EPUB documents with custom XML entities are not supported.")
                    with minidom.parseString(content) as document:
                        for node in document.getElementsByTagName("*"):
                            if node.hasAttribute("lang") and not node.hasAttribute("xml:lang"):
                                node.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:lang", node.getAttribute("lang"))
                        root = document.documentElement
                        if not root.getAttribute("xml:lang").strip():
                            root.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:lang", language)
                        content = document.toxml(encoding="utf-8")
                output.writestr(entry, content)
        return package.get("version", "3.0")


def validate_output(path, options):
    if not path.stat().st_size:
        raise RuntimeError("The converter produced an empty file.")
    if options.format == "brf":
        text = path.read_text(encoding="ascii")
        if not text.strip() or any(ord(c) > 95 or (ord(c) < 32 and c not in "\r\n\f") for c in text):
            raise RuntimeError("The converter produced invalid BRF characters.")
        if any(len(line) > options.cells_per_line for line in text.splitlines()):
            raise RuntimeError("The BRF output exceeds the selected line width.")
        if any(len(page.strip("\r\n").splitlines()) > options.lines_per_page for page in text.split("\f")):
            raise RuntimeError("The BRF output exceeds the selected page height.")
    else:
        with zipfile.ZipFile(path) as archive:
            if archive.read("mimetype") != b"application/epub+zip" or archive.testzip():
                raise RuntimeError("The converter produced an invalid eBraille package.")
            container = ET.fromstring(archive.read("META-INF/container.xml"))
            rootfile = container.find(".//{*}rootfile")
            if rootfile is None:
                raise RuntimeError("The eBraille package has no book document.")
            package_path = rootfile.get("full-path", "")
            package = ET.fromstring(archive.read(package_path))
            if not (package.findtext(".//{http://purl.org/dc/elements/1.1/}format") or "").startswith("eBraille"):
                raise RuntimeError("The result is not identified as an eBraille publication.")
            manifest = {item.get("id"): item.get("href", "") for item in package.findall("./{*}manifest/{*}item")}
            spine = package.findall("./{*}spine/{*}itemref")
            if not spine:
                raise RuntimeError("The eBraille file has no reading order.")
            has_braille = False
            for item in spine:
                href = manifest.get(item.get("idref"))
                if not href:
                    raise RuntimeError("The eBraille reading order refers to missing content.")
                name = posixpath.normpath(posixpath.join(posixpath.dirname(package_path), urllib.parse.unquote(href)))
                if name not in archive.namelist():
                    raise RuntimeError("The eBraille file is missing a chapter.")
                chapter = ET.fromstring(archive.read(name))
                body = chapter.find("{*}body")
                if body is not None:
                    has_braille |= any("\u2801" <= char <= "\u28ff" for char in "".join(body.itertext()))
            if not has_braille:
                raise RuntimeError("The eBraille file contains no translated braille.")


def publish_files(files, destination, base, cancelled=lambda: False):
    """Copy all volumes without overwriting anything; roll back only our own copies."""
    destination = Path(destination)
    if not base or base in {".", ".."} or any(c in base for c in "/\\:"):
        raise ValueError("Invalid export filename.")
    destination.mkdir(parents=True, exist_ok=True)
    created = []
    try:
        for volume, source in enumerate(files, 1):
            label = base + (f" - Volume {volume}" if len(files) > 1 else "")
            number = 1
            while True:
                if cancelled():
                    raise ExportCancelled("Export cancelled.")
                target = destination / (label + (f" ({number})" if number > 1 else "") + source.suffix)
                try:
                    stream = target.open("xb")
                    break
                except FileExistsError:
                    number += 1
            created.append(target)
            with stream, source.open("rb") as incoming:
                while chunk := incoming.read(1024 * 1024):
                    if cancelled():
                        raise ExportCancelled("Export cancelled.")
                    stream.write(chunk)
        return created
    except BaseException:
        for path in created:
            path.unlink(missing_ok=True)
        raise
