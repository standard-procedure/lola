# Lola Service — API Specification

**Base URL:** `http://localhost:8080`

**Content Type:** `application/json` (all requests and responses)

**File Path Convention:** All file paths are relative to the `/documents` volume mount. For example, `templates/invoice.docx` refers to `/documents/templates/invoice.docx` inside the container.

---

## Error Response Format

All error responses follow this structure:

```json
{
  "error": "Human-readable error message",
  "code": "ERROR_CODE"
}
```

### Error Codes

| Code | HTTP Status | Description |
|------|-------------|-------------|
| `TEMPLATE_NOT_FOUND` | 404 | Template file does not exist |
| `TEMPLATE_ERROR` | 422 | Template file is invalid or corrupt |
| `CONVERSION_ERROR` | 422 | Document conversion failed |
| `MERGE_ERROR` | 422 | Mail merge execution failed |
| `FIELD_NOT_FOUND` | 422 | Template contains no merge fields |
| `MISSING_FIELDS` | 422 | Data is missing required merge field columns |
| `INVALID_FORMAT` | 400 | Unsupported output format requested |
| `INVALID_REQUEST` | 400 | Request body is malformed or missing required fields |
| `LIBREOFFICE_ERROR` | 503 | LibreOffice is unavailable or crashed |
| `TIMEOUT` | 504 | Operation timed out |
| `INTERNAL_ERROR` | 500 | Unexpected server error |

---

## Endpoints

### `GET /health`

Health check endpoint. Returns the status of the service and its LibreOffice connection.

**Request:** No parameters.

**Response (200):**

```json
{
  "status": "ok",
  "libreoffice": "connected",
  "version": "0.1.0",
  "uptime_seconds": 3621
}
```

**Response (503) — LibreOffice unavailable:**

```json
{
  "status": "degraded",
  "libreoffice": "disconnected",
  "error": "Cannot connect to LibreOffice UNO socket",
  "code": "LIBREOFFICE_ERROR"
}
```

---

### `POST /convert`

Convert a document from one format to another. Primarily DOCX→PDF but supports other LibreOffice-compatible formats.

**Request Body:**

```json
{
  "input_path": "templates/invoice.docx",
  "output_format": "pdf",
  "output_path": "output/invoice.pdf"
}
```

| Field | Type | Required | Default | Description |
|-------|------|----------|---------|-------------|
| `input_path` | string | ✅ | — | Path to source document (relative to /documents) |
| `output_format` | string | ❌ | `"pdf"` | Target format: `"pdf"`, `"docx"`, `"odt"`, `"html"`, `"rtf"` |
| `output_path` | string | ❌ | auto | Output file path. If omitted, derived from input path with new extension |

**Response (200):**

```json
{
  "output_path": "output/invoice.pdf",
  "format": "pdf",
  "size_bytes": 142857,
  "duration_ms": 2340
}
```

**Response (404) — Input file not found:**

```json
{
  "error": "File not found: templates/invoice.docx",
  "code": "TEMPLATE_NOT_FOUND"
}
```

**Response (422) — Conversion failed:**

```json
{
  "error": "Failed to convert document: LibreOffice could not open the file",
  "code": "CONVERSION_ERROR"
}
```

**Response (400) — Invalid format:**

```json
{
  "error": "Unsupported output format: 'xlsx'. Supported formats: pdf, docx, odt, html, rtf",
  "code": "INVALID_FORMAT"
}
```

---

### `GET /fields`

Extract merge field names from a document template.

**Query Parameters:**

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `template_path` | string | ✅ | Path to template (relative to /documents) |

**Example:** `GET /fields?template_path=templates/letter.docx`

**Response (200):**

```json
{
  "template_path": "templates/letter.docx",
  "fields": [
    "Address",
    "City",
    "CustomerName",
    "Date",
    "PostCode"
  ],
  "field_count": 5
}
```

Fields are returned in **alphabetical order**, deduplicated (a field used multiple times appears once).

**Response (200) — No fields found:**

```json
{
  "template_path": "templates/plain.docx",
  "fields": [],
  "field_count": 0
}
```

**Response (404) — Template not found:**

```json
{
  "error": "File not found: templates/letter.docx",
  "code": "TEMPLATE_NOT_FOUND"
}
```

**Response (422) — Invalid template:**

```json
{
  "error": "Cannot open template: file is not a valid document",
  "code": "TEMPLATE_ERROR"
}
```

---

### `POST /mail_merge`

Execute a mail merge: populate a template with data records and produce output files.

**Request Body:**

```json
{
  "template_path": "templates/letter.docx",
  "data": [
    {
      "CustomerName": "Alice Smith",
      "Address": "123 High Street",
      "City": "London",
      "PostCode": "SW1A 1AA",
      "Date": "14 March 2026"
    },
    {
      "CustomerName": "Bob Jones",
      "Address": "456 Oak Road",
      "City": "Leeds",
      "PostCode": "LS1 1BA",
      "Date": "14 March 2026"
    }
  ],
  "output_dir": "output/letters/batch_001",
  "output_format": "pdf",
  "filename_field": null
}
```

| Field | Type | Required | Default | Description |
|-------|------|----------|---------|-------------|
| `template_path` | string | ✅ | — | Path to template (relative to /documents) |
| `data` | array of objects | ✅ | — | Array of data records. Each object's keys are field names, values are strings. |
| `output_dir` | string | ✅ | — | Directory for output files (relative to /documents). Created if it doesn't exist. |
| `output_format` | string | ❌ | `"pdf"` | Output format: `"pdf"`, `"docx"`, `"odt"` |
| `filename_field` | string | ❌ | `null` | Data field to use for output filenames. If null, files are numbered `0001.pdf`, `0002.pdf`, etc. |

**Response (200):**

```json
{
  "template_path": "templates/letter.docx",
  "output_dir": "output/letters/batch_001",
  "output_format": "pdf",
  "record_count": 2,
  "output_files": [
    "output/letters/batch_001/0001.pdf",
    "output/letters/batch_001/0002.pdf"
  ],
  "duration_ms": 5142,
  "warnings": []
}
```

**Response with warnings (200):**

If some fields in the template don't have corresponding data columns, the merge still proceeds but includes warnings:

```json
{
  "template_path": "templates/letter.docx",
  "output_dir": "output/letters/batch_001",
  "output_format": "pdf",
  "record_count": 2,
  "output_files": [
    "output/letters/batch_001/0001.pdf",
    "output/letters/batch_001/0002.pdf"
  ],
  "duration_ms": 5142,
  "warnings": [
    "Template field 'PostCode' not found in data — will render as blank"
  ]
}
```

**Response (400) — Empty data:**

```json
{
  "error": "Data array must contain at least one record",
  "code": "INVALID_REQUEST"
}
```

**Response (404) — Template not found:**

```json
{
  "error": "File not found: templates/letter.docx",
  "code": "TEMPLATE_NOT_FOUND"
}
```

**Response (422) — Merge failed:**

```json
{
  "error": "Mail merge failed: LibreOffice could not process the template",
  "code": "MERGE_ERROR"
}
```

**Response (503) — LibreOffice unavailable:**

```json
{
  "error": "LibreOffice is not available. The service may be starting up or recovering from a crash.",
  "code": "LIBREOFFICE_ERROR"
}
```

**Response (504) — Timeout:**

```json
{
  "error": "Mail merge timed out after 120 seconds",
  "code": "TIMEOUT"
}
```

---

## Data Conventions

### File Paths

- All paths are relative to the `/documents` volume mount
- Use forward slashes only (`templates/letter.docx`, not `templates\letter.docx`)
- No leading slash (`templates/letter.docx`, not `/templates/letter.docx`)
- Path traversal (`../`) is rejected with `INVALID_REQUEST`

### Data Records

- All values must be strings (numbers should be pre-formatted: `"£1,234.56"` not `1234.56`)
- All records in a merge must have the same keys
- Missing keys are treated as empty strings
- Keys should match the MERGEFIELD names in the template exactly (case-sensitive)

### Output Files

- Output directories are created automatically if they don't exist
- When `filename_field` is null, files are numbered with zero-padded 4-digit prefixes: `0001.pdf`, `0002.pdf`, etc.
- When `filename_field` is set, the value from that field is sanitised (non-alphanumeric chars replaced with `_`) and used as the filename
- Existing files in the output directory are NOT overwritten — a numeric suffix is added if there's a conflict

### Supported Formats

| Format | Extension | LibreOffice Filter | Notes |
|--------|-----------|-------------------|-------|
| PDF | `.pdf` | `writer_pdf_Export` | Best for final output |
| DOCX | `.docx` | `MS Word 2007 XML` | Editable Word format |
| ODT | `.odt` | `writer8` | LibreOffice native format |
| HTML | `.html` | `HTML (StarWriter)` | Basic HTML output |
| RTF | `.rtf` | `Rich Text Format` | Legacy format |

---

## Rate Limiting & Concurrency

- The service processes **one merge/conversion at a time** (LibreOffice is single-threaded)
- Concurrent requests are queued and processed sequentially
- The queue depth is limited to **10 requests** — additional requests receive `429 Too Many Requests`
- Each operation has a configurable timeout (default 120 seconds)

---

## Examples

### cURL Examples

**Health check:**
```bash
curl http://localhost:8080/health
```

**Convert DOCX to PDF:**
```bash
curl -X POST http://localhost:8080/convert \
  -H "Content-Type: application/json" \
  -d '{
    "input_path": "templates/invoice.docx",
    "output_format": "pdf"
  }'
```

**Extract merge fields:**
```bash
curl "http://localhost:8080/fields?template_path=templates/letter.docx"
```

**Execute mail merge:**
```bash
curl -X POST http://localhost:8080/mail_merge \
  -H "Content-Type: application/json" \
  -d '{
    "template_path": "templates/letter.docx",
    "data": [
      {"CustomerName": "Alice Smith", "City": "London"},
      {"CustomerName": "Bob Jones", "City": "Leeds"}
    ],
    "output_dir": "output/letters/batch_001",
    "output_format": "pdf"
  }'
```
