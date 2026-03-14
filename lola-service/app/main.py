"""
Lola Service — FastAPI application for document processing via LibreOffice.

This is the skeleton entry point. Each endpoint will be implemented
in subsequent development phases.
"""

import time

from fastapi import FastAPI, Query, HTTPException
from pydantic import BaseModel, Field
from typing import Optional

app = FastAPI(
    title="Lola Document Service",
    description="Document processing via LibreOffice UNO API",
    version="0.1.0",
)

_start_time = time.time()


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


@app.get("/health", response_model=HealthResponse)
async def health():
    """Health check — verifies LibreOffice UNO connection."""
    # TODO: Phase 1 — implement UNO connection check
    uptime = int(time.time() - _start_time)
    return HealthResponse(
        status="ok",
        libreoffice="not_implemented",
        version="0.1.0",
        uptime_seconds=uptime,
    )


@app.post("/convert", response_model=ConvertResponse, responses={
    404: {"model": ErrorResponse},
    422: {"model": ErrorResponse},
})
async def convert(request: ConvertRequest):
    """Convert a document to another format."""
    # TODO: Phase 2 — implement conversion via UNO
    raise HTTPException(status_code=501, detail="Not yet implemented")


@app.get("/fields", response_model=FieldsResponse, responses={
    404: {"model": ErrorResponse},
    422: {"model": ErrorResponse},
})
async def fields(template_path: str = Query(..., description="Path to template")):
    """Extract merge field names from a document template."""
    # TODO: Phase 3 — implement field extraction via UNO
    raise HTTPException(status_code=501, detail="Not yet implemented")


@app.post("/mail_merge", response_model=MergeResponse, responses={
    400: {"model": ErrorResponse},
    404: {"model": ErrorResponse},
    422: {"model": ErrorResponse},
    503: {"model": ErrorResponse},
    504: {"model": ErrorResponse},
})
async def mail_merge(request: MergeRequest):
    """Execute a mail merge with data records."""
    # TODO: Phase 4 — implement mail merge via UNO
    raise HTTPException(status_code=501, detail="Not yet implemented")
