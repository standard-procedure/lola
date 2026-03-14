"""Tests for the /health endpoint."""

import pytest
from unittest.mock import MagicMock
from fastapi.testclient import TestClient
from app.main import app, get_uno_client


def make_client(lo_connected: bool) -> TestClient:
    mock = MagicMock()
    mock.is_connected.return_value = lo_connected
    app.dependency_overrides[get_uno_client] = lambda: mock
    client = TestClient(app, raise_server_exceptions=False)
    return client


def test_health_returns_ok_when_libreoffice_connected():
    """Health endpoint returns 200 with status=ok when LO is reachable."""
    client = make_client(lo_connected=True)
    response = client.get("/health")
    assert response.status_code == 200
    data = response.json()
    assert data["status"] == "ok"
    assert data["libreoffice"] == "connected"
    assert data["version"] == "0.1.0"
    assert "uptime_seconds" in data


def test_health_returns_degraded_when_libreoffice_disconnected():
    """Health endpoint returns 503 with status=degraded when LO is unreachable."""
    client = make_client(lo_connected=False)
    response = client.get("/health")
    assert response.status_code == 503
    data = response.json()
    assert data["status"] == "degraded"
    assert data["libreoffice"] == "disconnected"
    assert data["code"] == "LIBREOFFICE_ERROR"
