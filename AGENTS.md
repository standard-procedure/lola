# AGENTS.md — Lola Project

## What Is This?

Lola is a Ruby gem + Python microservice that drives LibreOffice for document processing. The primary use case is executing mail merges against Word `.docx` templates with native MERGEFIELD codes.

## Project Structure

```
lola/
  docs/              ← Research, API spec, architecture, implementation plan
  lola/              ← Ruby gem (Lola::Client)
  lola-service/      ← Python FastAPI microservice + LibreOffice
  docker-compose.yml ← Local development
  documents/         ← Shared volume (templates + output, gitignored)
```

## Key Documentation

**Read these before making changes:**

1. `docs/UNO-MAILMERGE.md` — How LibreOffice UNO mail merge works, known gotchas
2. `docs/IMPLEMENTATION-PLAN.md` — Architecture, development phases, testing strategy
3. `docs/API-SPEC.md` — HTTP API endpoint specifications
4. `docs/ARCHITECTURE.md` — System design and data flow

## Architecture

```
Ruby App → Lola::Client (HTTP) → FastAPI (Python) → LibreOffice (UNO) → /documents
```

- LibreOffice runs headless inside the Docker container
- Python connects via UNO socket (port 2002)
- FastAPI exposes HTTP API (port 8080)
- Files shared via `/documents` volume mount
- **LibreOffice is single-threaded** — all requests serialized

## Development Phases

1. ✅ Planning & Research (this phase)
2. ⬜ Docker image + health endpoint
3. ⬜ DOCX→PDF conversion
4. ⬜ Merge field extraction
5. ⬜ Mail merge execution
6. ⬜ Ruby gem

## Critical Things To Know

### LibreOffice Quirks (from UNO-MAILMERGE.md)

- **UNO `execute()` takes a tuple, not a list** — `oMailMerge.execute(())` not `execute([])`
- **All paths must be file:// URLs** — use `uno.systemPathToFileUrl()`
- **Data sources must be registered then cleaned up** — always revoke after merge
- **Single-threaded** — serialize all UNO operations with a lock
- **Process crashes** — detect and restart LibreOffice automatically
- **Field codes in output** — use `SaveFilter = "writer_pdf_Export"` to flatten fields

### Python UNO

- `uno` module is NOT a pip package — comes from `python3-uno` system package
- Must use LibreOffice's bundled Python or system Python with UNO path
- Connection via `com.sun.star.bridge.UnoUrlResolver`

### Testing

- Python: pytest with real LibreOffice in Docker (no mocking LO)
- Ruby: RSpec with WebMock for unit tests, integration tests against running service
- Test fixtures: real `.docx` files with merge fields

## Conventions

- Python: FastAPI + Pydantic, type hints, async where possible
- Ruby: standard gem layout, net/http for HTTP client
- All file paths relative to `/documents` volume
- Error responses: `{ "error": "message", "code": "ERROR_CODE" }`
