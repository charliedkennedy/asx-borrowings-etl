"""Regression tests for the intentionally offline document workflow."""
from __future__ import annotations

import os
import socket
import subprocess
import sys
from pathlib import Path

import fitz

import src.asx_maturity_screen as screen


ROOT = Path(__file__).resolve().parents[1]


def _pdf(path: Path, text: str) -> None:
    document = fitz.open()
    page = document.new_page()
    page.insert_text((72, 72), text)
    document.save(path)
    document.close()


def test_active_pipeline_has_no_web_retrieval_surface() -> None:
    source = Path(screen.__file__).read_text(encoding="utf-8").lower()
    assert "asx.com.au" not in source
    assert "asxclient" not in source
    assert "announcement_items" not in source
    assert "allow-web-fallback" not in source
    assert "download_pdf" not in source


def test_missing_local_document_never_uses_network(monkeypatch, tmp_path: Path) -> None:
    targets = tmp_path / "targets.csv"
    targets.write_text("ticker,target_name,aliases\nZZZ,Missing Entity,\n", encoding="utf-8")
    root = tmp_path / "PDF files with spaces"
    root.mkdir()
    output = tmp_path / "outputs" / "screen.xlsx"

    def forbidden(*args, **kwargs):
        raise AssertionError("local matching must not open a network connection")

    monkeypatch.setattr(socket, "create_connection", forbidden)
    assert screen.main_for_args([
        "--match-only", "--pdf-root", str(root), "--targets", str(targets), "--output", str(output),
    ]) == 0
    register = (output.parent / "document_register.csv").read_text(encoding="utf-8")
    assert "NO_LOCAL_DOCUMENT" in register


def test_match_only_requires_no_api_key_and_module_help_works(tmp_path: Path) -> None:
    root = tmp_path / "PDF files with spaces"
    root.mkdir()
    _pdf(root / "Integral Diagnostics Limited_FY2026_ACN123456789.Pdf", "Integral Diagnostics\nAnnual Report")
    targets = tmp_path / "targets.csv"
    targets.write_text("ticker,target_name,aliases\nIDX,Integral Diagnostics,Integral Diagnostics Limited\n", encoding="utf-8")
    output = tmp_path / "outputs" / "screen.xlsx"
    old = os.environ.pop("OPENAI_API_KEY", None)
    try:
        assert screen.main_for_args(["--match-only", "--pdf-root", str(root), "--targets", str(targets), "--output", str(output)]) == 0
    finally:
        if old is not None:
            os.environ["OPENAI_API_KEY"] = old
    assert (output.parent / "pdf_index.csv").exists()
    help_result = subprocess.run([sys.executable, "-m", "src.asx_maturity_screen", "--help"], cwd=ROOT, capture_output=True, text=True)
    assert help_result.returncode == 0
    assert "--pdf-root" in help_result.stdout
