"""
Lola Service — FastAPI application for document processing via LibreOffice.
"""

import time
from typing import Optional

from fastapi import FastAPI, Query, HTTPException, Depends
from fastapi.responses import JSONResponse
from pydantic import BaseModel, Field

from app.config import config
from app.uno_client import UnoClient

app = FastAPI(
    title="Lola Document Service",
    description="Document processing via LibreOffice UNO API",
    version="0.1.0",
)

_start_time = time.time()
_uno_client: Optional[UnoClient] = None


def get_uno_client() -> UnoClient:
    global _uno_client
    if _uno_client is None:
        _uno_client = UnoClient(host=config.LO_HOST, port=config.LO_PORT)
    return _uno_client


# --- Request/Response Models ---


class ConvertRequest(BaseModel):
    input_path: str = Field(..., description="Path to source document (relative to /documents)")
    output_format: str = Field(default="pdf", description="Target format: pdf, docx, odt, html, rtf")
    output_path: Optional[str] = Field(default=None, description="Output file path (auto-generated if omitted)")


class ConvertResponse(BaseModel):
    output_path: str
    format: str
    size_bytes: int
    duration_ms: int


class FieldsResponse(BaseModel):
    template_path: str
    fields: list[str]
    field_count: int


class MergeRequest(BaseModel):
    template_path: str = Field(..., description="Path to template (relative to /documents)")
    data: list[dict[str, str]] = Field(..., description="Array of data records")
    output_dir: str = Field(..., description="Output directory (relative to /documents)")
    output_format: str = Field(default="pdf", description="Output format: pdf, docx, odt")
    filename_field: Optional[str] = Field(default=None, description="Data field for output filenames")
    timeout_seconds: int = Field(default=300, ge=1, le=600, description="Merge timeout in seconds (1–600)")


class MergeResponse(BaseModel):
    template_path: str
    output_dir: str
    output_format: str
    record_count: int
    output_files: list[str]
    duration_ms: int
    warnings: list[str]


class ErrorResponse(BaseModel):
    error: str
    code: str


class HealthResponse(BaseModel):
    status: str
    libreoffice: str
    version: str
    uptime_seconds: int


# --- Endpoints ---


@app.get("/health")
async def health(uno_client: UnoClient = Depends(get_uno_client)):
    """Health check — verifies LibreOffice UNO connection."""
    uptime = int(time.time() - _start_time)
    if uno_client.is_connected():
        return HealthResponse(
            status="ok",
            libreoffice="connected",
            version="0.1.0",
            uptime_seconds=uptime,
        )
    return JSONResponse(
        status_code=503,
        content={
            "status": "degraded",
            "libreoffice": "disconnected",
            "error": "Cannot connect to LibreOffice UNO socket",
            "code": "LIBREOFFICE_ERROR",
        },
    )


@app.post("/convert", response_model=ConvertResponse, responses={
    404: {"model": ErrorResponse},
    422: {"model": ErrorResponse},
})
async def convert(
    request: ConvertRequest,
    uno_client: UnoClient = Depends(get_uno_client),
):
    """Convert a document to another format."""
    from app.routes.convert import handle_convert
    return await handle_convert(request, uno_client)


@app.get("/fields", response_model=FieldsResponse, responses={
    404: {"model": ErrorResponse},
    422: {"model": ErrorResponse},
})
async def fields(template_path: str = Query(..., description="Path to template")):
    """Extract merge field names from a document template."""
    from app.routes.fields import handle_fields
    return await handle_fields(template_path)


@app.post("/mail_merge", response_model=MergeResponse, responses={
    400: {"model": ErrorResponse},
    404: {"model": ErrorResponse},
    422: {"model": ErrorResponse},
    503: {"model": ErrorResponse},
    504: {"model": ErrorResponse},
})
async def mail_merge(
    request: MergeRequest,
    uno_client: UnoClient = Depends(get_uno_client),
):
    """Execute a mail merge with data records."""
    from app.routes.mail_merge import handle_mail_merge
    return await handle_mail_merge(request, uno_client)
