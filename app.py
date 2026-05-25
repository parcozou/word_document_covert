"""HTTP plugin endpoint for Coze: Markdown report in, DOCX download URL out."""

from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from typing import Optional
from urllib.parse import quote
import os
import uuid

from fastapi import FastAPI, Header, HTTPException
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

from docx_converter import convert_markdown_to_docx, safe_file_stem


GENERATED_DIR = Path(os.getenv("GENERATED_DIR", "generated_files")).resolve()
GENERATED_DIR.mkdir(parents=True, exist_ok=True)
PUBLIC_BASE_URL = os.getenv("PUBLIC_BASE_URL", "http://localhost:8000").rstrip("/")
STORAGE_MODE = os.getenv("STORAGE_MODE", "local").lower()
DOCX_API_KEY = os.getenv("DOCX_API_KEY", "")

app = FastAPI(
    title="Markdown Report to DOCX Plugin",
    description="Creates a Word report with native tables and returns a download link.",
    version="1.0.0",
)
app.mount("/files", StaticFiles(directory=str(GENERATED_DIR)), name="files")


class DocxRequest(BaseModel):
    formatted_markdown: str = Field(
        ...,
        min_length=1,
        description="Complete Markdown report, including Markdown table syntax and image links.",
    )
    title: str = Field(..., min_length=1, max_length=180)
    to_format: str = Field("docx", description="Must be docx.")
    include_images: bool = Field(
        True, description="Embed permitted HTTPS chart images when they are accessible."
    )


class DocxResponse(BaseModel):
    download_url: str
    file_name: str
    table_count: int
    image_count: int
    skipped_image_count: int
    generated_at_utc: str
    message: str


def _authenticate(x_api_key: Optional[str]) -> None:
    if DOCX_API_KEY and x_api_key != DOCX_API_KEY:
        raise HTTPException(status_code=401, detail="Invalid API key.")


def _s3_download_url(local_file: Path, file_name: str) -> str:
    try:
        import boto3
    except ImportError as exc:
        raise RuntimeError("boto3 is required when STORAGE_MODE=s3.") from exc

    bucket = os.environ["S3_BUCKET"]
    key_prefix = os.getenv("S3_KEY_PREFIX", "generated-docx").strip("/")
    key = f"{key_prefix}/{file_name}"
    client = boto3.client(
        "s3",
        endpoint_url=os.getenv("S3_ENDPOINT_URL") or None,
        region_name=os.getenv("S3_REGION", "auto"),
        aws_access_key_id=os.environ["S3_ACCESS_KEY_ID"],
        aws_secret_access_key=os.environ["S3_SECRET_ACCESS_KEY"],
    )
    client.upload_file(
        str(local_file),
        bucket,
        key,
        ExtraArgs={
            "ContentType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "ContentDisposition": f'attachment; filename="{file_name}"',
        },
    )
    public_base_url = os.getenv("S3_PUBLIC_BASE_URL", "").rstrip("/")
    if public_base_url:
        return f"{public_base_url}/{quote(key)}"
    expires_in = int(os.getenv("S3_URL_EXPIRY_SECONDS", "86400"))
    return client.generate_presigned_url(
        "get_object",
        Params={"Bucket": bucket, "Key": key},
        ExpiresIn=expires_in,
    )


def _publish_file(local_file: Path, file_name: str) -> str:
    if STORAGE_MODE == "s3":
        return _s3_download_url(local_file, file_name)
    if STORAGE_MODE != "local":
        raise RuntimeError("STORAGE_MODE must be either local or s3.")
    return f"{PUBLIC_BASE_URL}/files/{quote(file_name)}"


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok", "storage_mode": STORAGE_MODE}


@app.post("/generate-docx", response_model=DocxResponse)
def generate_docx(payload: DocxRequest, x_api_key: Optional[str] = Header(None)) -> DocxResponse:
    _authenticate(x_api_key)
    if payload.to_format.lower() != "docx":
        raise HTTPException(status_code=400, detail="Only docx output is supported.")

    unique_suffix = uuid.uuid4().hex[:10]
    file_name = f"{safe_file_stem(payload.title)}_{unique_suffix}.docx"
    local_file = GENERATED_DIR / file_name
    try:
        result = convert_markdown_to_docx(
            payload.formatted_markdown,
            payload.title,
            local_file,
            include_images=payload.include_images,
        )
        download_url = _publish_file(local_file, file_name)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"DOCX generation failed: {exc}") from exc

    return DocxResponse(
        download_url=download_url,
        file_name=file_name,
        table_count=result.table_count,
        image_count=result.image_count,
        skipped_image_count=result.skipped_image_count,
        generated_at_utc=datetime.now(timezone.utc).isoformat(),
        message="DOCX report generated with editable Word tables.",
    )
