"""HTTP plugin endpoint for Coze: Markdown report in, DOCX download URL out."""

from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from typing import Optional
from urllib.parse import quote
import hashlib
import hmac
import os
import subprocess
import sys
import time
import uuid

from fastapi import FastAPI, Header, HTTPException
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field
from dotenv import load_dotenv

from docx_converter import convert_markdown_to_docx, safe_file_stem


load_dotenv()

GENERATED_DIR = Path(os.getenv("GENERATED_DIR", "generated_files")).resolve()
GENERATED_DIR.mkdir(parents=True, exist_ok=True)
PUBLIC_BASE_URL = os.getenv("PUBLIC_BASE_URL", "http://localhost:8000").rstrip("/")
STORAGE_MODE = os.getenv("STORAGE_MODE", "local").lower()
DOCX_API_KEY = os.getenv("DOCX_API_KEY", "")
FINALIZE_FIELDS = os.getenv("FINALIZE_FIELDS", "true").lower() not in {"0", "false", "no"}
FINALIZER_PYTHON = os.getenv("DOCX_FINALIZER_PYTHON", sys.executable)
FINALIZER_SCRIPT = Path(__file__).with_name("finalize_docx.py")
FIELD_FINALIZATION_TIMEOUT = int(os.getenv("FIELD_FINALIZATION_TIMEOUT", "90"))
USE_PROXY_DOWNLOAD_URLS = (
    os.getenv("USE_PROXY_DOWNLOAD_URLS", "true").lower() not in {"0", "false", "no"}
)
DOWNLOAD_LINK_EXPIRY_SECONDS = int(os.getenv("DOWNLOAD_LINK_EXPIRY_SECONDS", "7776000"))
S3_PRESIGNED_URL_MAX_SECONDS = 604800

app = FastAPI(
    title="Markdown Report to DOCX Plugin",
    description="Creates a Word report with native tables and returns a download link.",
    version="1.0.0",
)
if STORAGE_MODE == "local":
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


class RefreshDownloadRequest(BaseModel):
    file_name: str = Field(..., min_length=1, max_length=220)


class RefreshDownloadResponse(BaseModel):
    download_url: str
    file_name: str
    generated_at_utc: str
    message: str


def _authenticate(x_api_key: Optional[str]) -> None:
    if DOCX_API_KEY and x_api_key != DOCX_API_KEY:
        raise HTTPException(status_code=401, detail="Invalid API key.")


def _s3_client():
    try:
        import boto3
    except ImportError as exc:
        raise RuntimeError("boto3 is required when STORAGE_MODE=s3.") from exc

    return boto3.client(
        "s3",
        endpoint_url=os.getenv("S3_ENDPOINT_URL") or None,
        region_name=os.getenv("S3_REGION", "auto"),
        aws_access_key_id=os.environ["S3_ACCESS_KEY_ID"],
        aws_secret_access_key=os.environ["S3_SECRET_ACCESS_KEY"],
    )


def _s3_object_key(file_name: str) -> str:
    key_prefix = os.getenv("S3_KEY_PREFIX", "generated-docx").strip("/")
    return f"{key_prefix}/{file_name}" if key_prefix else file_name


def _download_secret() -> str:
    secret = os.getenv("DOWNLOAD_LINK_SECRET") or DOCX_API_KEY
    if not secret:
        raise RuntimeError("DOWNLOAD_LINK_SECRET or DOCX_API_KEY is required for proxy downloads.")
    return secret


def _download_signature(file_name: str, expires_at: int) -> str:
    message = f"{file_name}:{expires_at}".encode("utf-8")
    return hmac.new(_download_secret().encode("utf-8"), message, hashlib.sha256).hexdigest()


def _proxy_download_url(file_name: str) -> str:
    expires_at = int(time.time()) + max(DOWNLOAD_LINK_EXPIRY_SECONDS, 1)
    signature = _download_signature(file_name, expires_at)
    return (
        f"{PUBLIC_BASE_URL}/download/{quote(file_name)}"
        f"?expires={expires_at}&signature={signature}"
    )


def _safe_download_name(file_name: str) -> str:
    if Path(file_name).name != file_name or "/" in file_name or "\\" in file_name:
        raise HTTPException(status_code=400, detail="Invalid file name.")
    return file_name


def _s3_download_url(local_file: Path, file_name: str) -> str:
    bucket = os.environ["S3_BUCKET"]
    key = _s3_object_key(file_name)
    client = _s3_client()
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
    if USE_PROXY_DOWNLOAD_URLS:
        return _proxy_download_url(file_name)
    expires_in = min(int(os.getenv("S3_URL_EXPIRY_SECONDS", "86400")), S3_PRESIGNED_URL_MAX_SECONDS)
    return client.generate_presigned_url(
        "get_object",
        Params={"Bucket": bucket, "Key": key},
        ExpiresIn=expires_in,
    )


def _publish_file(local_file: Path, file_name: str) -> str:
    if STORAGE_MODE == "s3":
        download_url = _s3_download_url(local_file, file_name)
        local_file.unlink(missing_ok=True)
        return download_url
    if STORAGE_MODE != "local":
        raise RuntimeError("STORAGE_MODE must be either local or s3.")
    return f"{PUBLIC_BASE_URL}/files/{quote(file_name)}"


def _refreshed_download_url(file_name: str) -> str:
    file_name = _safe_download_name(file_name)
    if STORAGE_MODE == "s3":
        bucket = os.environ["S3_BUCKET"]
        key = _s3_object_key(file_name)
        client = _s3_client()
        try:
            client.head_object(Bucket=bucket, Key=key)
        except Exception as exc:
            raise HTTPException(status_code=404, detail="Generated DOCX file was not found.") from exc
        public_base_url = os.getenv("S3_PUBLIC_BASE_URL", "").rstrip("/")
        if public_base_url:
            return f"{public_base_url}/{quote(key)}"
        if USE_PROXY_DOWNLOAD_URLS:
            return _proxy_download_url(file_name)
        expires_in = min(int(os.getenv("S3_URL_EXPIRY_SECONDS", "86400")), S3_PRESIGNED_URL_MAX_SECONDS)
        return client.generate_presigned_url(
            "get_object",
            Params={"Bucket": bucket, "Key": key},
            ExpiresIn=expires_in,
        )

    if STORAGE_MODE != "local":
        raise HTTPException(status_code=500, detail="STORAGE_MODE must be either local or s3.")
    local_file = (GENERATED_DIR / file_name).resolve()
    if not str(local_file).startswith(str(GENERATED_DIR) + os.sep) or not local_file.exists():
        raise HTTPException(status_code=404, detail="Generated DOCX file was not found.")
    return f"{PUBLIC_BASE_URL}/files/{quote(file_name)}"


def _finalize_fields(local_file: Path) -> None:
    if not FINALIZE_FIELDS:
        return
    command = [
        FINALIZER_PYTHON,
        str(FINALIZER_SCRIPT),
        str(local_file),
        "--timeout",
        str(min(FIELD_FINALIZATION_TIMEOUT, 60)),
    ]
    soffice_path = os.getenv("SOFFICE_PATH")
    if soffice_path:
        command.extend(["--soffice", soffice_path])
    try:
        subprocess.run(
            command,
            check=True,
            capture_output=True,
            text=True,
            timeout=FIELD_FINALIZATION_TIMEOUT,
        )
    except subprocess.CalledProcessError as exc:
        detail = (exc.stderr or exc.stdout or "unknown LibreOffice error").strip()
        raise RuntimeError(f"Contents-page finalization failed: {detail}") from exc
    except subprocess.TimeoutExpired as exc:
        raise RuntimeError("Contents-page finalization timed out.") from exc


@app.get("/")
def root() -> dict[str, str]:
    return {
        "status": "ok",
        "service": "Markdown Report to DOCX Plugin",
        "health_check": "/health",
        "generate_docx": "/generate-docx",
        "refresh_download_url": "/refreshDownloadUrl",
    }


@app.get("/health")
def health() -> dict[str, str]:
    return {
        "status": "ok",
        "storage_mode": STORAGE_MODE,
        "field_finalization": "enabled" if FINALIZE_FIELDS else "disabled",
        "proxy_download_urls": "enabled" if USE_PROXY_DOWNLOAD_URLS else "disabled",
        "download_link_expiry_seconds": str(DOWNLOAD_LINK_EXPIRY_SECONDS),
    }


@app.get("/download/{file_name}")
def download_generated_docx(file_name: str, expires: int, signature: str):
    file_name = _safe_download_name(file_name)
    if int(time.time()) > expires:
        raise HTTPException(status_code=403, detail="Download link has expired.")
    expected_signature = _download_signature(file_name, expires)
    if not hmac.compare_digest(expected_signature, signature):
        raise HTTPException(status_code=403, detail="Invalid download signature.")

    media_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    content_disposition = f'attachment; filename="{file_name}"'
    if STORAGE_MODE == "s3":
        try:
            obj = _s3_client().get_object(Bucket=os.environ["S3_BUCKET"], Key=_s3_object_key(file_name))
        except Exception as exc:
            raise HTTPException(status_code=404, detail="Generated DOCX file was not found.") from exc
        return StreamingResponse(
            obj["Body"],
            media_type=media_type,
            headers={"Content-Disposition": content_disposition},
        )

    local_file = (GENERATED_DIR / file_name).resolve()
    if not str(local_file).startswith(str(GENERATED_DIR) + os.sep) or not local_file.exists():
        raise HTTPException(status_code=404, detail="Generated DOCX file was not found.")
    return FileResponse(local_file, media_type=media_type, filename=file_name)


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
        _finalize_fields(local_file)
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
        message="DOCX report generated with editable Word tables and finalized contents.",
    )


@app.post("/refreshDownloadUrl", response_model=RefreshDownloadResponse)
@app.post("/refresh-download-url", response_model=RefreshDownloadResponse, include_in_schema=False)
def refresh_download_url(
    payload: RefreshDownloadRequest, x_api_key: Optional[str] = Header(None)
) -> RefreshDownloadResponse:
    _authenticate(x_api_key)
    download_url = _refreshed_download_url(payload.file_name)
    return RefreshDownloadResponse(
        download_url=download_url,
        file_name=payload.file_name,
        generated_at_utc=datetime.now(timezone.utc).isoformat(),
        message="Download URL refreshed for an existing generated DOCX file.",
    )
