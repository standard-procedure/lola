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

    def mail_merge(
        self,
        template_path: str,
        data: list,
        output_dir: str,
        output_format: str = "pdf",
    ) -> list:
        """
        Execute a mail merge via UNO.

        Args:
            template_path: Absolute path to .docx template.
            data: List of dicts — each dict is one record.
            output_dir: Absolute path to output directory.
            output_format: "pdf", "docx", or "odt".

        Returns:
            List of absolute paths to output files.
        """
        from app.uno_mail_merge import mail_merge as _mail_merge

        def _do(client):
            return _mail_merge(
                client.ctx,
                template_path,
                data,
                output_dir,
                output_format=output_format,
            )

        return self.execute_with_lock(_do)

    def convert_to_pdf(self, input_path: str, output_path: str, save_filter: str = "writer_pdf_Export") -> None:
        """
        Convert a document to the specified format using LibreOffice UNO.

        Args:
            input_path: Absolute filesystem path to the source document.
            output_path: Absolute filesystem path for the output file.
            save_filter: LibreOffice export filter name (default: writer_pdf_Export).
        """
        import uno
        from com.sun.star.beans import PropertyValue

        def _do(client):
            desktop = client.smgr.createInstanceWithContext(
                "com.sun.star.frame.Desktop", client.ctx
            )

            # Open document hidden
            hidden_prop = PropertyValue()
            hidden_prop.Name = "Hidden"
            hidden_prop.Value = True

            doc = desktop.loadComponentFromURL(
                uno.systemPathToFileUrl(input_path),
                "_blank",
                0,
                (hidden_prop,),
            )

            try:
                # Export with specified filter
                filter_prop = PropertyValue()
                filter_prop.Name = "FilterName"
                filter_prop.Value = save_filter

                doc.storeToURL(
                    uno.systemPathToFileUrl(output_path),
                    (filter_prop,),
                )
            finally:
                doc.close(True)

        self.execute_with_lock(_do)
