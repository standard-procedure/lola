"""Tests for the POST /convert endpoint."""

import os
import pytest
from unittest.mock import MagicMock, patch
from fastapi.testclient import TestClient
from app.main import app, get_uno_client


def make_client(convert_result=None, convert_side_effect=None):
    mock = MagicMock()
    mock.is_connected.return_value = True
    if convert_side_effect:
        mock.convert_to_pdf.side_effect = convert_side_effect
    else:
        mock.convert_to_pdf.return_value = convert_result
    app.dependency_overrides[get_uno_client] = lambda: mock
    return TestClient(app, raise_server_exceptions=False)


def test_convert_docx_to_pdf(tmp_path):
    """POST /convert converts a document and returns output path."""
    # Create a fake input file so path-exists check passes
    input_file = tmp_path / "invoice.docx"
    input_file.write_bytes(b"fake docx")
    output_file = tmp_path / "invoice.pdf"
    output_file.write_bytes(b"fake pdf content")

    client = make_client(convert_result=str(output_file))

    with patch("app.routes.convert.resolve_path", side_effect=lambda p: str(tmp_path / p)):
        response = client.post("/convert", json={
            "input_path": "invoice.docx",
            "output_format": "pdf",
        })

    assert response.status_code == 200
    data = response.json()
    assert data["format"] == "pdf"
    assert "output_path" in data
    assert "size_bytes" in data
    assert "duration_ms" in data


def test_convert_returns_404_for_missing_file(tmp_path):
    """POST /convert returns 404 when the input file does not exist."""
    client = make_client()

    with patch("app.routes.convert.resolve_path", side_effect=lambda p: str(tmp_path / p)):
        response = client.post("/convert", json={"input_path": "missing.docx"})

    assert response.status_code == 404
    data = response.json()
    assert data["code"] == "TEMPLATE_NOT_FOUND"


def test_convert_returns_400_for_invalid_format(tmp_path):
    """POST /convert returns 400 for unsupported output format."""
    input_file = tmp_path / "doc.docx"
    input_file.write_bytes(b"fake")
    client = make_client()

    with patch("app.routes.convert.resolve_path", side_effect=lambda p: str(tmp_path / p)):
        response = client.post("/convert", json={
            "input_path": "doc.docx",
            "output_format": "xlsx",
        })

    assert response.status_code == 400
    data = response.json()
    assert data["code"] == "INVALID_FORMAT"


def test_convert_returns_422_on_conversion_failure(tmp_path):
    """POST /convert returns 422 when LibreOffice conversion fails."""
    input_file = tmp_path / "bad.docx"
    input_file.write_bytes(b"corrupt")
    from app.exceptions import ConversionError
    client = make_client(convert_side_effect=ConversionError("LibreOffice could not open the file"))

    with patch("app.routes.convert.resolve_path", side_effect=lambda p: str(tmp_path / p)):
        response = client.post("/convert", json={"input_path": "bad.docx"})

    assert response.status_code == 422
    data = response.json()
    assert data["code"] == "CONVERSION_ERROR"
