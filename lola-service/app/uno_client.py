"""
UNO connection manager for LibreOffice.

Handles connecting to LibreOffice via the UNO bridge,
reconnecting on failure, and serializing access (LO is single-threaded).
"""

# NOTE: The `uno` module is NOT a pip package.
# It comes from the `python3-uno` system package installed alongside LibreOffice.
# This module will only import successfully inside the Docker container
# or on a system with LibreOffice + python3-uno installed.

import threading
import logging

logger = logging.getLogger(__name__)


class UnoClient:
    """Manages the UNO connection to a headless LibreOffice instance."""

    def __init__(self, host: str = "localhost", port: int = 2002):
        self.host = host
        self.port = port
        self._ctx = None
        self._lock = threading.Lock()
        self._operation_count = 0

    def connect(self) -> None:
        """Establish a UNO connection to LibreOffice."""
        import uno

        local_context = uno.getComponentContext()
        resolver = local_context.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_context
        )
        connection_string = (
            f"uno:socket,host={self.host},port={self.port};"
            "urp;StarOffice.ComponentContext"
        )
        self._ctx = resolver.resolve(connection_string)
        logger.info(f"Connected to LibreOffice at {self.host}:{self.port}")

    @property
    def ctx(self):
        """Get the UNO component context, reconnecting if needed."""
        if self._ctx is None:
            self.connect()
        return self._ctx

    @property
    def smgr(self):
        """Get the UNO ServiceManager."""
        return self.ctx.ServiceManager

    def is_connected(self) -> bool:
        """Check if the UNO connection is alive."""
        try:
            # Simple operation to test the connection
            _ = self.smgr
            return True
        except Exception:
            self._ctx = None
            return False

    def execute_with_lock(self, fn):
        """
        Execute a function with the UNO lock held.
        Retries once on connection failure.
        """
        with self._lock:
            try:
                result = fn(self)
                self._operation_count += 1
                return result
            except Exception as e:
                logger.warning(f"UNO operation failed: {e}. Reconnecting...")
                self._ctx = None
                try:
                    self.connect()
                    result = fn(self)
                    self._operation_count += 1
                    return result
                except Exception as retry_error:
                    logger.error(f"UNO retry also failed: {retry_error}")
                    raise
