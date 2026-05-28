"""Isolated DOCX conversion worker.

The web process calls this script as a short-lived subprocess so conversion
memory is released before the LibreOffice field finalizer starts.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("job_file", type=Path)
    parser.add_argument("output_file", type=Path)
    parser.add_argument("stats_file", type=Path)
    args = parser.parse_args()

    print(f"[docx-worker] load job={args.job_file.name}", flush=True)
    job = json.loads(args.job_file.read_text(encoding="utf-8-sig"))
    print(f"[docx-worker] import converter output={args.output_file.name}", flush=True)
    from docx_converter import convert_markdown_to_docx

    print(f"[docx-worker] convert start output={args.output_file.name}", flush=True)
    result = convert_markdown_to_docx(
        job["formatted_markdown"],
        job["title"],
        args.output_file,
        include_images=bool(job.get("include_images", True)),
    )
    print(f"[docx-worker] convert done output={args.output_file.name}", flush=True)
    args.stats_file.write_text(
        json.dumps(
            {
                "table_count": result.table_count,
                "image_count": result.image_count,
                "skipped_image_count": result.skipped_image_count,
                "heading_count": result.heading_count,
            }
        ),
        encoding="utf-8",
    )


if __name__ == "__main__":
    main()
