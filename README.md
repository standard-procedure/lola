# Lola

A Ruby gem + Python microservice that drives LibreOffice for document processing. Execute mail merges against Word `.docx` templates with native MERGEFIELD codes, convert documents between formats, and extract merge field names.

## Why?

Real-world Word templates use native mail merge fields (`«FieldName»`). Simple template libraries can't process these — they're stored as complex field codes in the docx XML. LibreOffice understands them natively and has a programmable mail merge engine (UNO API).

Lola wraps this capability in a Docker-based microservice with a clean HTTP API and a Ruby gem client.

## Architecture

```
Ruby App / Rails
  → Lola::Client (HTTP)
    → Python FastAPI service (port 8080)
      → LibreOffice (headless, UNO API)
        → reads/writes /documents volume
```

See [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) for the full picture.

## Font Support

Lola's Docker image includes a comprehensive set of fonts for real-world Word template compatibility:

| Font package | Covers |
|---|---|
| `ttf-mscorefonts-installer` | Arial, Verdana, Georgia, Courier New, Times New Roman, and other Microsoft core fonts |
| `fonts-crosextra-carlito` | **Calibri** (metric-compatible replacement — Word's default font since 2007) |
| `fonts-crosextra-caladea` | **Cambria** (metric-compatible replacement — common Word heading font) |
| `fonts-liberation` | Arial, Times New Roman, Courier New (metric-compatible replacements) |
| `fonts-dejavu-core` | General-purpose coverage |

This means most real-world Word templates will produce correctly laid-out PDFs without font substitution artifacts.

### Custom fonts

If your templates use branded or custom fonts, mount them at `/fonts`:

```yaml
# docker-compose.yml
volumes:
  - ./documents:/documents
  - ./fonts:/fonts:ro   # drop .ttf/.otf files here
```

Custom fonts are automatically picked up and registered when the container starts.

## Quick Start

### 1. Start the service

```bash
docker compose up -d
```

### 2. Use the Ruby gem

```ruby
# Gemfile
gem "lola", path: "./lola"

# Configuration
Lola.configure do |config|
  config.url = "http://localhost:8080"
end

# Convert DOCX to PDF
client = Lola::Client.new
client.convert("templates/invoice.docx", format: :pdf)

# Extract merge fields from a template
client.fields("templates/letter.docx")
# => ["CustomerName", "Address", "City", "PostCode"]

# Execute a mail merge
client.mail_merge(
  template: "templates/letter.docx",
  data: [
    { "CustomerName" => "Alice Smith", "City" => "London" },
    { "CustomerName" => "Bob Jones", "City" => "Leeds" },
  ],
  output_dir: "output/letters/batch_001",
  output_format: :pdf
)
```

### 3. Or use the HTTP API directly

```bash
# Health check
curl http://localhost:8080/health

# Convert
curl -X POST http://localhost:8080/convert \
  -H "Content-Type: application/json" \
  -d '{"input_path": "templates/invoice.docx", "output_format": "pdf"}'

# Extract fields
curl "http://localhost:8080/fields?template_path=templates/letter.docx"

# Mail merge
curl -X POST http://localhost:8080/mail_merge \
  -H "Content-Type: application/json" \
  -d '{
    "template_path": "templates/letter.docx",
    "data": [{"CustomerName": "Alice", "City": "London"}],
    "output_dir": "output/batch_001",
    "output_format": "pdf"
  }'
```

## Documentation

- [API Specification](docs/API-SPEC.md) — endpoint details, request/response schemas
- [Architecture](docs/ARCHITECTURE.md) — system design and data flow
- [Implementation Plan](docs/IMPLEMENTATION-PLAN.md) — development phases and testing strategy
- [UNO Mail Merge Research](docs/UNO-MAILMERGE.md) — LibreOffice UNO API deep dive

## Project Structure

```
lola/
  lola/              ← Ruby gem
  lola-service/      ← Python FastAPI + LibreOffice
  docs/              ← Documentation
  docker-compose.yml ← Local development setup
```

## Development

### Prerequisites

- Docker & Docker Compose
- Ruby 3.x (for the gem)
- The service runs LibreOffice inside Docker — no local LO installation needed

### Running Tests

```bash
# Python service tests (inside Docker)
docker compose run --rm lola-service pytest

# Ruby gem tests
cd lola && bundle exec rspec
```

## License

MIT
