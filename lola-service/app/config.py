"""Configuration loaded from environment variables."""

import os


class Config:
    LO_HOST: str = os.getenv("LO_HOST", "localhost")
    LO_PORT: int = int(os.getenv("LO_PORT", "2002"))
    API_PORT: int = int(os.getenv("API_PORT", "8080"))
    LOG_LEVEL: str = os.getenv("LOG_LEVEL", "info")
    MAX_MERGE_TIMEOUT: int = int(os.getenv("MAX_MERGE_TIMEOUT", "120"))
    DOCUMENTS_PATH: str = os.getenv("DOCUMENTS_PATH", "/documents")
    LO_RESTART_AFTER: int = int(os.getenv("LO_RESTART_AFTER", "100"))


config = Config()
