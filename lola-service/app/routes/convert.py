"""POST /convert — Convert a document to another format via LibreOffice UNO."""

import os
import time

from fastapi.responses import JSONResponse

from app.config import config
from app.exceptions import ConversionError, InvalidFormatError, TemplateNotFoundError, LolaError
from app.uno_client import UnoClient

SUPPORTED_FORMATS = ["pdf", "docx", "odt", "html", "rtf"]

FILTER_MAP = {
    "pdf": "writer_pdf_Export",
    "docx": "MS Word 2007 XML",
    "odt": "writer8",
    "html": "HTML (StarWriter)",
    "rtf": "Rich Text Format",
}


def resolve_path(relative_path: str) -> str:
    """Resolve a path relative to the documents volume."""
    return os.path.join(config.DOCUMENTS_PATH, relative_path)


async def handle_convert(request, uno_client: UnoClient):
    from app.main import ConvertResponse

    output_format = request.output_format.lower()
    if output_format not in SUPPORTED_FORMATS:
        return JSONResponse(
            status_code=400,
            content={
                "error": f"Unsupported output format: '{output_format}'. Supported formats: {', '.join(SUPPORTED_FORMATS)}",
                "code": "INVALID_FORMAT",
            },
        )

    input_abs = resolve_path(request.input_path)
    if not os.path.exists(input_abs):
        return JSONResponse(
            status_code=404,
            content={
                "error": f"File not found: {request.input_path}",
                "code": "TEMPLATE_NOT_FOUND",
            },
        )

    # Determine output path
    if request.output_path:
        output_abs = resolve_path(request.output_path)
        output_rel = request.output_path
    else:
        base = os.path.splitext(request.input_path)[0]
        output_rel = f"{base}.{output_format}"
        output_abs = resolve_path(output_rel)

    os.makedirs(os.path.dirname(output_abs), exist_ok=True)

    start = time.time()
    try:
        uno_client.convert_to_pdf(input_abs, output_abs, save_filter=FILTER_MAP[output_format])
    except LolaError as e:
        return JSONResponse(
            status_code=422,
            content={"error": e.message, "code": e.code},
        )
    except Exception as e:
        return JSONResponse(
            status_code=422,
            content={"error": f"Conversion failed: {e}", "code": "CONVERSION_ERROR"},
        )

    duration_ms = int((time.time() - start) * 1000)
    size_bytes = os.path.getsize(output_abs) if os.path.exists(output_abs) else 0

    return ConvertResponse(
        output_path=output_rel,
        format=output_format,
        size_bytes=size_bytes,
        duration_ms=duration_ms,
    )
