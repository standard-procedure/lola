#!/bin/bash
set -e

LO_PORT=${LO_PORT:-2002}
API_PORT=${API_PORT:-8080}

echo "Starting LibreOffice in headless mode on UNO port ${LO_PORT}..."

# Start LibreOffice with a disposable user profile
soffice --headless --norestore --nologo \
  --env:UserInstallation=file:///tmp/lo_user_profile \
  --accept="socket,host=localhost,port=${LO_PORT};urp;StarOffice.ServiceManager" &

LO_PID=$!

# Wait for LibreOffice to be ready (up to 60 seconds)
echo "Waiting for LibreOffice to accept connections..."
READY=false
for i in $(seq 1 30); do
    if python3 -c "
import uno
ctx = uno.getComponentContext()
resolver = ctx.ServiceManager.createInstanceWithContext(
    'com.sun.star.bridge.UnoUrlResolver', ctx)
resolver.resolve(
    'uno:socket,host=localhost,port=${LO_PORT};urp;StarOffice.ComponentContext')
print('OK')
" 2>/dev/null; then
        echo "LibreOffice is ready (attempt ${i})."
        READY=true
        break
    fi
    echo "  waiting... (attempt ${i}/30)"
    sleep 2
done

if [ "$READY" = false ]; then
    echo "ERROR: LibreOffice failed to start within 60 seconds."
    exit 1
fi

echo "Starting Lola API service on port ${API_PORT}..."

# Start FastAPI — exec replaces the shell so signals propagate
exec uvicorn app.main:app --host 0.0.0.0 --port ${API_PORT} --log-level ${LOG_LEVEL:-info}
