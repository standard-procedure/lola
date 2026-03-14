"""POST /mail_merge — Execute a mail merge via LibreOffice UNO."""

import os
import time

from fastapi.responses import JSONResponse

from app.config import config
from app.exceptions import LibreOfficeError, MergeError, LolaError
from app.uno_client import UnoClient

MAX_RECORDS = 1000
SUPPORTED_FORMATS = ["pdf", "docx", "odt"]


def resolve_path(relative_path: str) -> str:
    """Resolve a path relative to the documents volume."""
    return os.path.join(config.DOCUMENTS_PATH, relative_path)


async def handle_mail_merge(request, uno_client: UnoClient):
    from app.main import MergeResponse

    # Validate data
    if not request.data:
        return JSONResponse(
            status_code=400,
            content={
                "error": "Data array must contain at least one record",
                "code": "INVALID_REQUEST",
            },
        )

    if len(request.data) > MAX_RECORDS:
        return JSONResponse(
            status_code=400,
            content={
                "error": f"Data array exceeds maximum of {MAX_RECORDS} records",
                "code": "INVALID_REQUEST",
            },
        )

    output_format = request.output_format.lower()
    if output_format not in SUPPORTED_FORMATS:
        return JSONResponse(
            status_code=400,
            content={
                "error": f"Unsupported output format: '{output_format}'. Supported: {', '.join(SUPPORTED_FORMATS)}",
                "code": "INVALID_FORMAT",
            },
        )

    template_abs = resolve_path(request.template_path)
    if not os.path.exists(template_abs):
        return JSONResponse(
            status_code=404,
            content={
                "error": f"File not found: {request.template_path}",
                "code": "TEMPLATE_NOT_FOUND",
            },
        )

    output_abs = resolve_path(request.output_dir)
    os.makedirs(output_abs, exist_ok=True)

    # Warn about fields in template not present in data
    warnings = []
    try:
        from app.routes.fields import extract_merge_fields
        template_fields = set(extract_merge_fields(template_abs))
        data_keys = set(request.data[0].keys()) if request.data else set()
        for field in sorted(template_fields - data_keys):
            warnings.append(f"Template field '{field}' not found in data — will render as blank")
    except Exception:
        pass  # Field extraction is best-effort

    start = time.time()
    try:
        output_files = uno_client.mail_merge(
            template_abs,
            request.data,
            output_abs,
            output_format=output_format,
        )
    except LibreOfficeError as e:
        return JSONResponse(
            status_code=503,
            content={"error": e.message, "code": e.code},
        )
    except MergeError as e:
        return JSONResponse(
            status_code=422,
            content={"error": e.message, "code": e.code},
        )
    except LolaError as e:
        return JSONResponse(
            status_code=422,
            content={"error": e.message, "code": e.code},
        )
    except Exception as e:
        return JSONResponse(
            status_code=422,
            content={"error": f"Mail merge failed: {e}", "code": "MERGE_ERROR"},
        )

    duration_ms = int((time.time() - start) * 1000)

    # Convert absolute paths to relative paths for the response
    def to_relative(abs_path):
        try:
            return os.path.relpath(abs_path, config.DOCUMENTS_PATH)
        except ValueError:
            return abs_path

    return MergeResponse(
        template_path=request.template_path,
        output_dir=request.output_dir,
        output_format=output_format,
        record_count=len(request.data),
        output_files=[to_relative(p) for p in output_files],
        duration_ms=duration_ms,
        warnings=warnings,
    )
