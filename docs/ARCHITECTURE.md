# Lola — Architecture Overview

## What Is Lola?

Lola is a document processing service that drives LibreOffice to execute mail merges against Word `.docx` templates containing native MERGEFIELD codes. It exposes this capability as a simple HTTP API, with a Ruby gem client for Rails integration.

## Why Does This Exist?

Real-world clients have existing Word templates with mail merge fields (the `«FieldName»` markers). These templates use Word's native MERGEFIELD codes stored in the docx XML as `w:fldChar`/`w:instrText` structures.

Simple template libraries (python-docx, docxtpl) can't process these because they don't understand Word field codes. LibreOffice is one of the few tools that:
1. Correctly interprets Word MERGEFIELD codes when opening .docx files
2. Has a programmable mail merge engine (the UNO API)
3. Can run headless on Linux in Docker

The trade-off: UNO is complex and fragile. Lola wraps this complexity in a managed service so application code just makes HTTP calls.

## System Diagram

```
┌─────────────────────────────────────────────────────────────────┐
│  Host / Kubernetes / Docker Compose                             │
│                                                                 │
│  ┌─────────────────┐         ┌────────────────────────────────┐│
│  │  Ruby App        │   HTTP  │  lola-service container        ││
│  │  (Rails, etc.)   │────────>│                                ││
│  │                  │  :8080  │  ┌──────────┐  ┌────────────┐ ││
│  │  Lola::Client    │         │  │  FastAPI  │──│ LibreOffice│ ││
│  │                  │         │  │  (Python) │  │ (headless) │ ││
│  │                  │         │  │           │  │ UNO :2002  │ ││
│  └────────┬─────────┘         │  └──────────┘  └────────────┘ ││
│           │                   │                                ││
│           │                   │  /documents (volume)           ││
│           └───── shared ──────│    ├── templates/              ││
│                  volume       │    ├── output/                 ││
│                               │    └── temp/                   ││
│                               └────────────────────────────────┘│
└─────────────────────────────────────────────────────────────────┘
```

## Component Responsibilities

### Ruby Gem (`lola/`)

- HTTP client wrapping the service API
- File path management (placing templates, retrieving output)
- Error mapping (HTTP errors → Ruby exceptions)
- Configuration (service URL, timeouts, document paths)
- **No document processing logic** — all work delegated to the service

### Python Service (`lola-service/`)

- FastAPI web server handling HTTP requests
- UNO connection management (connect, reconnect, health check)
- Document conversion (DOCX→PDF via LibreOffice)
- Merge field extraction (parse template fields via UNO)
- Mail merge execution (register data source, execute, cleanup)
- LibreOffice process supervision (detect crashes, restart)
- Request serialization (LibreOffice is single-threaded)

### LibreOffice (inside `lola-service`)

- Runs headless with UNO socket listener on port 2002
- Opens and interprets Word .docx files
- Executes mail merge engine with registered data sources
- Exports documents in various formats (PDF, DOCX, ODT)
- **Managed by the Python service** — not directly accessed by anything else

## Data Flow: Mail Merge

```
1. Ruby app calls client.mail_merge(template:, data:, output_dir:)
       │
2. Client POSTs JSON to /mail_merge
       │
3. FastAPI validates request
       │
4. Service writes data to temp CSV file
       │
5. Service registers CSV directory as UNO data source
       │
6. Service creates MailMerge UNO object with:
   - DocumentURL (template)
   - DataSourceName (registered name)
   - Command ("data" — CSV filename sans extension)
   - OutputType (FILE)
   - OutputURL (output directory)
   - SaveFilter ("writer_pdf_Export" for PDF)
       │
7. Service calls oMailMerge.execute(())
       │
8. LibreOffice opens template, substitutes fields, writes output files
       │
9. Service unregisters data source, cleans up temp files
       │
10. Service returns JSON with output file paths
       │
11. Ruby client parses response, returns result object
       │
12. Ruby app reads output files from shared volume
```

## Key Design Decisions

### One Container (Not Two)

LibreOffice + Python in one container because:
- `python3-uno` must match LibreOffice version exactly
- UNO socket is localhost-only (127.0.0.1)
- Simplifies health checks and lifecycle
- No inter-container networking complexity

### Shared Volume (Not HTTP File Upload)

Templates and output files are shared via a Docker volume, not uploaded/downloaded via HTTP because:
- Large files (multi-MB documents) are expensive to transfer via HTTP
- Volume mount is simpler and faster
- Same approach used by other document processing services
- The Ruby app and service must run on the same host (or use NFS/S3)

### Serial Processing (Not Parallel)

One merge at a time because:
- LibreOffice is fundamentally single-threaded for document operations
- Concurrent UNO operations cause crashes or corrupted output
- The service uses a mutex/queue to serialize requests
- Horizontal scaling = multiple service instances (separate LO processes)

### CSV Data Source (Not Direct ResultSet)

CSV file + data source registration because:
- It's the most well-documented and tested UNO mail merge path
- Creating an in-memory UNO ResultSet is complex and poorly documented
- CSV files are easy to debug (just cat the file)
- Cleanup is straightforward (delete file, revoke data source)

## Scaling

### Vertical
- More CPU/RAM for the container → faster LibreOffice processing
- No benefit from multiple cores (LO is single-threaded per operation)

### Horizontal
- Run multiple `lola-service` instances behind a load balancer
- Each instance has its own LibreOffice process
- Shared volume must be accessible by all instances (NFS, EFS, etc.)
- Use request routing to spread load

### Caching
- Template parsing results could be cached (field names)
- Not worth caching merge results (data always changes)
- LibreOffice document cache is per-process (no cross-instance benefit)
