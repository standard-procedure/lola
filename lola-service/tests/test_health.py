"""Tests for the /health endpoint."""


def test_health_returns_ok(client):
    """Health endpoint should return 200 with status info."""
    response = client.get("/health")
    assert response.status_code == 200
    data = response.json()
    assert data["status"] == "ok"
    assert "version" in data
    assert "uptime_seconds" in data
    assert "libreoffice" in data
