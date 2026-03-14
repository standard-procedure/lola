"""Shared test fixtures for Lola service tests."""

import pytest
from fastapi.testclient import TestClient
from app.main import app


@pytest.fixture(autouse=True)
def clear_dependency_overrides():
    """Ensure dependency overrides are cleared after each test."""
    yield
    app.dependency_overrides.clear()


@pytest.fixture
def client():
    """FastAPI test client with no dependency overrides (no UNO)."""
    return TestClient(app, raise_server_exceptions=False)
