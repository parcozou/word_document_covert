"""Update Word fields in a DOCX through LibreOffice before public delivery."""

from __future__ import annotations

import argparse
import os
from pathlib import Path
import shutil
import socket
import subprocess
import tempfile
import time
from zipfile import ZIP_DEFLATED, ZipFile
import xml.etree.ElementTree as ET

import uno


WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
UPDATE_FIELDS = f"{{{WORD_NS}}}updateFields"
VALUE_ATTRIBUTE = f"{{{WORD_NS}}}val"


def _property(name: str, value: object):
    item = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
    item.Name = name
    item.Value = value
    return item


def _find_soffice(explicit_path: str | None) -> str:
    candidates = [
        explicit_path,
        os.getenv("SOFFICE_PATH"),
        shutil.which("soffice"),
        shutil.which("soffice.exe"),
        r"C:\Program Files\LibreOffice\program\soffice.exe",
    ]
    for candidate in candidates:
        if candidate and Path(candidate).exists():
            return str(Path(candidate))
    raise RuntimeError("LibreOffice soffice executable is not available.")


def _open_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as server:
        server.bind(("127.0.0.1", 0))
        return int(server.getsockname()[1])


def _connect_desktop(port: int, timeout_seconds: int):
    local_context = uno.getComponentContext()
    resolver = local_context.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_context
    )
    url = (
        f"uno:socket,host=127.0.0.1,port={port};"
        "urp;StarOffice.ComponentContext"
    )
    deadline = time.monotonic() + timeout_seconds
    while time.monotonic() < deadline:
        try:
            context = resolver.resolve(url)
            return context.ServiceManager.createInstanceWithContext(
                "com.sun.star.frame.Desktop", context
            )
        except Exception:
            time.sleep(0.25)
    raise RuntimeError("LibreOffice did not start within the finalization timeout.")


def _disable_open_refresh(docx_path: Path) -> None:
    temp_path = docx_path.with_suffix(".finalized.tmp.docx")
    with ZipFile(docx_path, "r") as source, ZipFile(
        temp_path, "w", compression=ZIP_DEFLATED
    ) as destination:
        for info in source.infolist():
            content = source.read(info.filename)
            if info.filename == "word/settings.xml":
                root = ET.fromstring(content)
                update_fields = root.find(UPDATE_FIELDS)
                if update_fields is None:
                    update_fields = ET.SubElement(root, UPDATE_FIELDS)
                update_fields.set(VALUE_ATTRIBUTE, "false")
                content = ET.tostring(
                    root, encoding="utf-8", xml_declaration=True
                )
            destination.writestr(info, content)
    temp_path.replace(docx_path)


def finalize_docx(
    docx_path: Path, soffice_path: str | None = None, timeout_seconds: int = 45
) -> None:
    docx_path = docx_path.resolve()
    if not docx_path.exists():
        raise FileNotFoundError(docx_path)
    soffice = _find_soffice(soffice_path)
    port = _open_port()
    with tempfile.TemporaryDirectory(prefix="docx_field_refresh_") as profile_dir:
        profile_uri = Path(profile_dir).resolve().as_uri()
        process = subprocess.Popen(
            [
                soffice,
                "--headless",
                "--invisible",
                "--nologo",
                "--nodefault",
                "--nolockcheck",
                "--nofirststartwizard",
                "--norestore",
                "--quickstart=no",
                f"-env:UserInstallation={profile_uri}",
                (
                    f"--accept=socket,host=127.0.0.1,port={port};"
                    "urp;StarOffice.ComponentContext"
                ),
            ],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        try:
            desktop = _connect_desktop(port, timeout_seconds)
            document = desktop.loadComponentFromURL(
                docx_path.as_uri(),
                "_blank",
                0,
                (
                    _property("Hidden", True),
                    # Avoid a duplicate full update on load; refresh fields explicitly below.
                    _property("UpdateDocMode", 0),
                ),
            )
            document.TextFields.refresh()
            indexes = document.getDocumentIndexes()
            for index in range(indexes.getCount()):
                indexes.getByIndex(index).update()
            document.storeAsURL(
                docx_path.as_uri(),
                (
                    _property("FilterName", "Office Open XML Text"),
                    _property("Overwrite", True),
                ),
            )
        finally:
            try:
                if "document" in locals() and document is not None:
                    try:
                        document.close(True)
                    except Exception:
                        document.dispose()
            except Exception:
                pass
            process.terminate()
            try:
                process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                process.kill()
                process.wait(timeout=5)
    _disable_open_refresh(docx_path)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("docx_path", type=Path)
    parser.add_argument("--soffice")
    parser.add_argument("--timeout", type=int, default=45)
    args = parser.parse_args()
    finalize_docx(args.docx_path, args.soffice, args.timeout)
    print(str(args.docx_path.resolve()))


if __name__ == "__main__":
    main()
