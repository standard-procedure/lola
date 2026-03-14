#!/bin/bash
set -e

export LO_PORT=${LO_PORT:-2002}
export API_PORT=${API_PORT:-8080}
export LOG_LEVEL=${LOG_LEVEL:-info}

echo "Starting LibreOffice via supervisord on UNO port ${LO_PORT}..."

# Start supervisord in background (manages LibreOffice process)
supervisord -c /etc/supervisor/conf.d/lola.conf &
SUPERVISORD_PID=$!

# Wait for LibreOffice to be ready (up to 60 seconds)
echo "Waiting for LibreOffice UNO socket to be ready..."
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

# Tell supervisord to start uvicorn now that LibreOffice is ready
supervisorctl -c /etc/supervisor/conf.d/lola.conf start uvicorn

# Keep the container alive by waiting on supervisord
wait $SUPERVISORD_PID
