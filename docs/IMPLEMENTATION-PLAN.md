# Lola — Implementation Plan

## Overview

Lola is a Ruby gem that wraps a Python FastAPI microservice to drive LibreOffice for document processing. The key use case: executing mail merges against Word `.docx` templates that use native MERGEFIELD codes.

## Architecture

```
┌─────────────────────────┐
│  Ruby App / Rails        │
│                         │
│  Lola::Client (HTTP)    │──── HTTP (JSON) ────┐
│                         │                     │
└─────────────────────────┘                     │
                                                ▼
┌───────────────────────────────────────────────────────┐
│  lola-service container                               │
│                                                       │
│  ┌─────────────────────┐    ┌──────────────────────┐ │
│  │  Python FastAPI      │    │  LibreOffice         │ │
│  │  (port 8080)         │◄──►│  (headless, UNO      │ │
│  │                     │    │   socket port 2002)  │ │
│  │  - /convert         │    │                      │ │
│  │  - /mail_merge      │    │  Started by          │ │
│  │  - /fields          │    │  entrypoint script   │ │
│  │  - /health          │    │                      │ │
│  └─────────────────────┘    └──────────────────────┘ │
│                                                       │
│  /documents (shared volume)                           │
│    ├── templates/                                     │
│    └── output/                                        │
└───────────────────────────────────────────────────────┘
```

### Why One Container?

LibreOffice and the Python service run in the **same container** because:
1. `python3-uno` is tightly coupled to LibreOffice's Python — version must match exactly
2. UNO socket connection is localhost-only — no network overhead
3. Simpler deployment and lifecycle management
4. The entrypoint script starts LibreOffice, waits for it, then starts FastAPI

### File Flow

All document I/O goes through a shared `/documents` volume:
```
/documents/
  templates/     ← Upload .docx templates here
  output/        ← Merged/converted files appear here
  temp/          ← Temporary working files (auto-cleaned)
```

The Ruby app and the service share this volume. The Ruby gem handles file placement and retrieval.

---

## Docker Compose

```yaml
version: "3.8"

services:
  lola-service:
    build:
      context: ./lola-service
      dockerfile: Dockerfile
    ports:
      - "8080:8080"
    volumes:
      - ./documents:/documents
    environment:
      - LO_PORT=2002
      - API_PORT=8080
      - LOG_LEVEL=info
      - MAX_MERGE_TIMEOUT=120
    healthcheck:
      test: ["CMD", "curl", "-f", "http://localhost:8080/health"]
      interval: 30s
      timeout: 10s
      retries: 3
      start_period: 30s
    restart: unless-stopped
```

---

## Docker Image

```dockerfile
FROM ubuntu:22.04

ENV DEBIAN_FRONTEND=noninteractive

# Install LibreOffice + Python UNO bridge + Microsoft fonts
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer \
    libreoffice-calc \
    python3 \
    python3-pip \
    python3-uno \
    curl \
    ttf-mscorefonts-installer \
    fonts-liberation \
    fonts-dejavu-core \
    fontconfig \
    && fc-cache -f -v \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Create app directory and documents volume
WORKDIR /app
RUN mkdir -p /documents/templates /documents/output /documents/temp

# Install Python dependencies
COPY requirements.txt .
RUN pip3 install --no-cache-dir -r requirements.txt

# Copy application
COPY . .

# Entrypoint starts LibreOffice then FastAPI
COPY entrypoint.sh /entrypoint.sh
RUN chmod +x /entrypoint.sh

EXPOSE 8080

ENTRYPOINT ["/entrypoint.sh"]
```

### Entrypoint Script

```bash
#!/bin/bash
set -e

LO_PORT=${LO_PORT:-2002}
API_PORT=${API_PORT:-8080}

# Start LibreOffice in headless mode with UNO socket
soffice --headless --norestore --nologo \
  --env:UserInstallation=file:///tmp/lo_user_profile \
  --accept="socket,host=localhost,port=${LO_PORT};urp;StarOffice.ServiceManager" &

LO_PID=$!

# Wait for LibreOffice to be ready
echo "Waiting for LibreOffice to start on port ${LO_PORT}..."
for i in $(seq 1 30); do
    if python3 -c "
import uno
ctx = uno.getComponentContext()
resolver = ctx.ServiceManager.createInstanceWithContext(
    'com.sun.star.bridge.UnoUrlResolver', ctx)
resolver.resolve(
    'uno:socket,host=localhost,port=${LO_PORT};urp;StarOffice.ComponentContext')
print('connected')
" 2>/dev/null; then
        echo "LibreOffice is ready."
        break
    fi
    echo "  attempt $i/30..."
    sleep 2
done

# Start FastAPI
exec uvicorn app.main:app --host 0.0.0.0 --port ${API_PORT}
```

---

## Python FastAPI Service

### Project Structure

```
lola-service/
  Dockerfile
  entrypoint.sh
  requirements.txt
  app/
    __init__.py
    main.py           # FastAPI app, routes
    uno_client.py      # UNO connection management
    mail_merge.py      # Mail merge logic
    converter.py       # DOCX→PDF conversion
    field_extractor.py # Extract merge field names
    models.py          # Pydantic request/response models
    config.py          # Configuration
    exceptions.py      # Custom exceptions
  tests/
    __init__.py
    conftest.py
    test_health.py
    test_convert.py
    test_fields.py
    test_mail_merge.py
    fixtures/
      simple_template.docx
      merge_fields_template.docx
  docs/
```

### Key Dependencies

```
# requirements.txt
fastapi==0.115.*
uvicorn[standard]==0.34.*
pydantic==2.*
python-multipart==0.0.*
```

Note: `uno` is NOT a pip package — it comes from `python3-uno` system package.

### Core Modules

#### `uno_client.py` — UNO Connection Manager

```python
import uno
import threading
import time

class UnoClient:
    """Manages the connection to LibreOffice via UNO."""
    
    def __init__(self, host="localhost", port=2002):
        self.host = host
        self.port = port
        self._ctx = None
        self._lock = threading.Lock()  # LO is single-threaded
    
    def connect(self):
        """Establish or re-establish UNO connection."""
        localContext = uno.getComponentContext()
        resolver = localContext.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", localContext
        )
        self._ctx = resolver.resolve(
            f"uno:socket,host={self.host},port={self.port};"
            "urp;StarOffice.ComponentContext"
        )
    
    @property
    def ctx(self):
        """Get context, reconnecting if needed."""
        if self._ctx is None:
            self.connect()
        return self._ctx
    
    @property  
    def smgr(self):
        """Get ServiceManager."""
        return self.ctx.ServiceManager
    
    def execute_with_lock(self, fn):
        """Execute a function with the UNO lock held."""
        with self._lock:
            try:
                return fn(self)
            except Exception:
                # Connection may be dead — reset and retry once
                self._ctx = None
                self.connect()
                return fn(self)
```

#### `mail_merge.py` — Core Merge Logic

Implements the merge workflow from UNO-MAILMERGE.md:
1. Write CSV to temp directory
2. Register as data source
3. Execute merge
4. Clean up data source
5. Return output file paths

#### `converter.py` — Simple Conversion

Uses `soffice --convert-to` for simple format conversion (no merge fields). Falls back to UNO document open + export for more control.

#### `field_extractor.py` — Extract Field Names

Opens the template via UNO, iterates text fields using `XEnumerationAccess`, identifies database fields, and returns their names.

```python
def extract_fields(uno_client, template_path):
    """Extract MERGEFIELD names from a document."""
    desktop = uno_client.smgr.createInstanceWithContext(
        "com.sun.star.frame.Desktop", uno_client.ctx
    )
    
    # Open document hidden
    props = (make_property("Hidden", True),)
    doc = desktop.loadComponentFromURL(
        uno.systemPathToFileUrl(template_path),
        "_blank", 0, props
    )
    
    try:
        fields = set()
        text_fields = doc.getTextFields()
        enum = text_fields.createEnumeration()
        
        while enum.hasMoreElements():
            field = enum.nextElement()
            if field.supportsService("com.sun.star.text.TextField.Database"):
                # field.Content has the field name
                # field.FieldSubType or field.getPropertyValue("FieldSubType")
                master = field.getTextFieldMaster()
                name = master.getPropertyValue("DataColumnName")
                fields.add(name)
        
        return sorted(fields)
    finally:
        doc.close(True)
```

---

## Ruby Gem

### Project Structure

```
lola/
  lola.gemspec
  Gemfile
  Rakefile
  README.md
  lib/
    lola.rb
    lola/
      version.rb
      client.rb
      configuration.rb
      errors.rb
      response.rb
  spec/
    spec_helper.rb
    lola/
      client_spec.rb
    fixtures/
      vcr_cassettes/
  docs/
```

### Public API

```ruby
# Configuration
Lola.configure do |config|
  config.url = "http://localhost:8080"   # or ENV["LOLA_URL"]
  config.timeout = 120                    # seconds
  config.documents_path = "/documents"   # shared volume path
end

# Client
client = Lola::Client.new  # uses global config
# or
client = Lola::Client.new(url: "http://lola:8080", timeout: 60)

# Convert DOCX to PDF
result = client.convert("templates/invoice.docx", format: :pdf)
# => Lola::Response(output_path: "output/invoice.pdf", duration: 2.3)

# Extract merge field names
fields = client.fields("templates/letter.docx")
# => ["CustomerName", "Address", "City", "PostCode", "Date"]

# Mail merge
result = client.mail_merge(
  template: "templates/letter.docx",
  data: [
    { "CustomerName" => "Alice Smith", "City" => "London" },
    { "CustomerName" => "Bob Jones", "City" => "Leeds" },
  ],
  output_dir: "output/letters/batch_001",
  output_format: :pdf
)
# => Lola::MergeResult(
#      output_files: ["output/letters/batch_001/0001.pdf", ...],
#      record_count: 2,
#      duration: 5.1
#    )
```

### Error Hierarchy

```ruby
module Lola
  class Error < StandardError; end
  
  # Connection/HTTP errors
  class ConnectionError < Error; end
  class TimeoutError < Error; end
  
  # Service errors
  class ServiceError < Error
    attr_reader :code
    def initialize(message, code: nil)
      @code = code
      super(message)
    end
  end
  
  class ConversionError < ServiceError; end
  class MergeError < ServiceError; end
  class TemplateError < ServiceError; end
  class FieldNotFoundError < MergeError; end
end
```

### HTTP Client

Use `net/http` (stdlib) or `faraday` for HTTP. Keep dependencies minimal.

```ruby
# lib/lola/client.rb
require "net/http"
require "json"
require "uri"

module Lola
  class Client
    def initialize(url: nil, timeout: nil)
      @url = url || Lola.configuration.url
      @timeout = timeout || Lola.configuration.timeout
    end
    
    def health
      get("/health")
    end
    
    def convert(input_path, format: :pdf, output_path: nil)
      post("/convert", {
        input_path: input_path,
        output_format: format.to_s,
        output_path: output_path
      }.compact)
    end
    
    def fields(template_path)
      response = get("/fields", template_path: template_path)
      response["fields"]
    end
    
    def mail_merge(template:, data:, output_dir:, output_format: :pdf)
      post("/mail_merge", {
        template_path: template,
        data: data,
        output_dir: output_dir,
        output_format: output_format.to_s
      })
    end
    
    private
    
    def get(path, params = {})
      uri = build_uri(path, params)
      request = Net::HTTP::Get.new(uri)
      execute(request, uri)
    end
    
    def post(path, body)
      uri = build_uri(path)
      request = Net::HTTP::Post.new(uri)
      request.content_type = "application/json"
      request.body = JSON.generate(body)
      execute(request, uri)
    end
    
    def execute(request, uri)
      http = Net::HTTP.new(uri.host, uri.port)
      http.read_timeout = @timeout
      response = http.request(request)
      handle_response(response)
    rescue Errno::ECONNREFUSED, Errno::ECONNRESET => e
      raise Lola::ConnectionError, "Cannot connect to Lola service: #{e.message}"
    rescue Net::ReadTimeout => e
      raise Lola::TimeoutError, "Request timed out: #{e.message}"
    end
    
    def handle_response(response)
      body = JSON.parse(response.body)
      
      case response.code.to_i
      when 200..299
        body
      when 400..499
        raise error_for_code(body["code"], body["error"])
      when 500..599
        raise Lola::ServiceError.new(body["error"], code: body["code"])
      end
    end
    
    def error_for_code(code, message)
      case code
      when "CONVERSION_ERROR" then Lola::ConversionError.new(message, code: code)
      when "MERGE_ERROR"      then Lola::MergeError.new(message, code: code)
      when "TEMPLATE_ERROR"   then Lola::TemplateError.new(message, code: code)
      when "FIELD_NOT_FOUND"  then Lola::FieldNotFoundError.new(message, code: code)
      else Lola::ServiceError.new(message, code: code)
      end
    end
    
    def build_uri(path, params = {})
      uri = URI.join(@url, path)
      uri.query = URI.encode_www_form(params) unless params.empty?
      uri
    end
  end
end
```

---

## Error Handling Strategy

### LibreOffice Process Crash Recovery

```python
# In the FastAPI service
class LoProcessManager:
    def __init__(self):
        self.process = None
        
    def ensure_running(self):
        """Check if LO is alive, restart if not."""
        if self.process is None or self.process.poll() is not None:
            self.start()
            self.wait_for_ready()
    
    def start(self):
        """Start LibreOffice process."""
        self.process = subprocess.Popen([
            "soffice", "--headless", "--norestore", "--nologo",
            "--env:UserInstallation=file:///tmp/lo_user_profile",
            "--accept=socket,host=localhost,port=2002;urp;StarOffice.ServiceManager"
        ])
    
    def wait_for_ready(self, max_attempts=15):
        """Wait for UNO socket to accept connections."""
        for i in range(max_attempts):
            try:
                uno_client.connect()
                return
            except Exception:
                time.sleep(2)
        raise RuntimeError("LibreOffice failed to start")
```

### Timeout Handling

- FastAPI endpoint timeout: configurable (default 120s)
- UNO operation timeout: handled by the Python service (not built into UNO)
- Implemented via threading: run merge in a thread, join with timeout, kill LO if stuck

### Corrupt Template Handling

- Validate file is a valid ZIP (docx) before sending to LibreOffice
- Catch UNO exceptions during document load
- Return `TEMPLATE_ERROR` with descriptive message

### Missing Merge Field Detection

- Extract fields from template (`/fields` endpoint)
- Compare with provided data columns before merge
- Warn (not error) on missing fields — they'll render as blank or field codes

---

## Testing Strategy

### Python Service (pytest)

```
tests/
  conftest.py           # Fixtures, LO connection setup
  test_health.py        # Health endpoint
  test_convert.py       # DOCX→PDF conversion
  test_fields.py        # Field extraction
  test_mail_merge.py    # Full merge workflow
  fixtures/
    simple.docx         # Plain document (no fields)
    merge_template.docx # Template with MERGEFIELD codes
    conditional.docx    # Template with IF fields
    broken.docx         # Intentionally malformed
```

**Test environment:** Tests run inside the Docker container with a real LibreOffice instance. No mocking of LibreOffice — we test the real thing.

```python
# conftest.py
import pytest
from app.main import app
from fastapi.testclient import TestClient

@pytest.fixture
def client():
    return TestClient(app)

@pytest.fixture
def simple_template(tmp_path):
    """Create a simple .docx with merge fields for testing."""
    # Use python-docx to create test fixtures programmatically
    pass
```

### Ruby Gem (RSpec)

```
spec/
  spec_helper.rb
  lola/
    client_spec.rb       # HTTP client tests
    configuration_spec.rb
  integration/
    convert_spec.rb      # Integration tests (need running service)
  fixtures/
    vcr_cassettes/       # Recorded HTTP responses
```

**Unit tests:** Use WebMock to stub HTTP responses. Fast, no service needed.

**Integration tests:** Run against a real lola-service container. Tagged `integration`, skipped in CI by default.

---

## Development Phases

### Phase 1: Foundation (Docker + Health)

**Goal:** Working Docker container with LibreOffice + FastAPI skeleton.

**Deliverables:**
- `Dockerfile` that builds and runs
- `entrypoint.sh` that starts LO and FastAPI
- `GET /health` endpoint returning `{ "status": "ok", "libreoffice": "connected" }`
- `docker-compose.yml` for local dev
- Basic pytest setup with health test

**Acceptance criteria:**
- `docker compose up` starts the service
- `curl http://localhost:8080/health` returns 200
- LibreOffice is running and UNO connection works

### Phase 2: DOCX→PDF Conversion

**Goal:** Convert documents between formats.

**Deliverables:**
- `POST /convert` endpoint
- UNO-based conversion (open document, export with filter)
- Support PDF, DOCX, ODT output formats
- Error handling for invalid files
- Tests with real document fixtures

**Acceptance criteria:**
- Upload a `.docx`, get a `.pdf` back
- Handles errors gracefully (bad file, unsupported format)
- Output matches LibreOffice GUI conversion quality

### Phase 3: Merge Field Extraction

**Goal:** Extract field names from Word templates.

**Deliverables:**
- `GET /fields` endpoint
- UNO-based field enumeration
- Support both simple and complex field representations
- Tests with various template styles

**Acceptance criteria:**
- Given a template with fields `CustomerName`, `Address`, `City` → returns all three
- Works with `.docx` files created in different Word versions
- Returns empty list for documents with no merge fields

### Phase 4: Mail Merge

**Goal:** Execute mail merges with data.

**Deliverables:**
- `POST /mail_merge` endpoint
- CSV data source registration workflow
- Merge execution with cleanup
- Support PDF and DOCX output
- Concurrent request serialization (mutex)
- Tests with real merge operations

**Acceptance criteria:**
- Given a template + 5 data records → produces 5 PDFs
- Field values correctly substituted
- Data source cleaned up after merge
- Concurrent requests don't crash

### Phase 5: Ruby Gem

**Goal:** Polished Ruby client gem.

**Deliverables:**
- `lola.gemspec` with proper metadata
- `Lola::Client` with full API coverage
- `Lola::Configuration` for global setup
- Error hierarchy
- RSpec tests with WebMock
- README with usage examples

**Acceptance criteria:**
- `gem build lola.gemspec` produces valid gem
- All public methods documented
- Error handling covers all service error codes
- Integration tests pass against running service

---

## Multi-Agent TDD Workflow

Each phase follows the Planner → Developer → Tester → Reviewer cycle:

### Planner (this document)
- Defines what to build and acceptance criteria
- Provides code structure and API contracts
- Documents known gotchas and edge cases

### Developer
- Implements the code following the plan
- Writes unit tests alongside implementation
- Commits working code with passing tests

### Tester
- Writes additional integration/edge-case tests
- Verifies acceptance criteria
- Tests error paths and boundary conditions
- Creates test fixtures (sample .docx files)

### Reviewer
- Reviews code for correctness and style
- Checks error handling completeness
- Validates against UNO research (does the implementation handle known gotchas?)
- Ensures cleanup of temp files and data sources

---

## Configuration

### Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `LO_PORT` | `2002` | LibreOffice UNO socket port |
| `API_PORT` | `8080` | FastAPI listen port |
| `LOG_LEVEL` | `info` | Logging level |
| `MAX_MERGE_TIMEOUT` | `120` | Max seconds per merge operation |
| `DOCUMENTS_PATH` | `/documents` | Shared volume path |
| `LO_RESTART_AFTER` | `100` | Restart LO after N operations |

### Ruby Gem Configuration

```ruby
# config/initializers/lola.rb (Rails)
Lola.configure do |config|
  config.url = ENV.fetch("LOLA_URL", "http://localhost:8080")
  config.timeout = ENV.fetch("LOLA_TIMEOUT", 120).to_i
  config.documents_path = ENV.fetch("LOLA_DOCUMENTS_PATH", "/documents")
end
```
