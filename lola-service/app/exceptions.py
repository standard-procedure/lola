"""Custom exceptions for the Lola service."""


class LolaError(Exception):
    """Base exception for Lola service errors."""

    def __init__(self, message: str, code: str = "INTERNAL_ERROR"):
        self.message = message
        self.code = code
        super().__init__(message)


class TemplateNotFoundError(LolaError):
    def __init__(self, path: str):
        super().__init__(f"File not found: {path}", code="TEMPLATE_NOT_FOUND")


class TemplateError(LolaError):
    def __init__(self, message: str):
        super().__init__(message, code="TEMPLATE_ERROR")


class ConversionError(LolaError):
    def __init__(self, message: str):
        super().__init__(message, code="CONVERSION_ERROR")


class MergeError(LolaError):
    def __init__(self, message: str):
        super().__init__(message, code="MERGE_ERROR")


class InvalidFormatError(LolaError):
    def __init__(self, format: str, supported: list[str]):
        supported_str = ", ".join(supported)
        super().__init__(
            f"Unsupported output format: '{format}'. Supported formats: {supported_str}",
            code="INVALID_FORMAT",
        )


class LibreOfficeError(LolaError):
    def __init__(self, message: str = "LibreOffice is not available"):
        super().__init__(message, code="LIBREOFFICE_ERROR")


class TimeoutError(LolaError):
    def __init__(self, seconds: int):
        super().__init__(
            f"Operation timed out after {seconds} seconds", code="TIMEOUT"
        )
