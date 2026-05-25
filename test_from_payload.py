"""Generate and audit a DOCX from the sample Coze payload."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from docx import Document

from docx_converter import convert_markdown_to_docx, safe_file_stem


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--input",
        type=Path,
        default=Path(__file__).resolve().parent.parent / "Untitled-1.md",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path(__file__).resolve().parent / "generated_files",
    )
    parser.add_argument("--skip-images", action="store_true")
    args = parser.parse_args()

    payload = json.loads(args.input.read_text(encoding="utf-8"))
    output = args.output_dir / f"{safe_file_stem(payload['title'])}_sample.docx"
    result = convert_markdown_to_docx(
        payload["formatted_markdown"],
        payload["title"],
        output,
        include_images=not args.skip_images,
    )

    document = Document(str(output))
    invalid_table_cells = [
        cell.text.strip()[:80]
        for table in document.tables
        for row in table.rows
        for cell in row.cells
        if cell.text.strip().startswith(("#", "---", "!["))
    ]
    suspicious_headings = [
        paragraph.text.strip()[:100]
        for paragraph in document.paragraphs
        if paragraph.style.name.startswith("Heading")
        and len(paragraph.text.split()) > 18
    ]
    audit = {
        "output_path": str(output.resolve()),
        "native_word_tables": len(document.tables),
        "reported_tables": result.table_count,
        "embedded_images": result.image_count,
        "skipped_images": result.skipped_image_count,
        "headings": result.heading_count,
        "invalid_table_cells": invalid_table_cells,
        "suspicious_headings": suspicious_headings,
    }
    if not document.tables or len(document.tables) != result.table_count:
        raise RuntimeError(f"Table audit failed: {audit}")
    if invalid_table_cells:
        raise RuntimeError(f"Content was incorrectly embedded in a table: {audit}")
    if suspicious_headings:
        raise RuntimeError(f"Long narrative content was incorrectly styled as a heading: {audit}")
    print(json.dumps(audit, indent=2))


if __name__ == "__main__":
    main()
