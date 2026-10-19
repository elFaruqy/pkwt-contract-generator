"""
pdf_service: mengubah file Word menjadi PDF.

Urutan mesin konversi yang dicoba:
1. Microsoft Word (Windows, lewat docx2pdf) -> hasil paling identik
   dengan tampilan di Word. Ini mesin yang dipakai di kantor.
2. LibreOffice (jika terpasang di komputer/server) -> cadangan.
3. Tidak ada -> aplikasi hanya membuat file Word (mode demo online).
"""

from __future__ import annotations

import shutil
import subprocess
from pathlib import Path

from config.settings import IS_WINDOWS
from utils.logger import get_logger

logger = get_logger(__name__)

ENGINE_WORD = "Microsoft Word"
ENGINE_LIBREOFFICE = "LibreOffice"
LIBREOFFICE_TIMEOUT_SECONDS = 120


def _word_available() -> bool:
    """True jika di Windows dan paket docx2pdf + pywin32 terpasang."""
    if not IS_WINDOWS:
        return False
    try:
        import docx2pdf  # noqa: F401
        import pythoncom  # noqa: F401
    except ImportError:
        return False
    return True


def _libreoffice_path() -> str | None:
    """Lokasi program LibreOffice jika terpasang."""
    return shutil.which("soffice") or shutil.which("libreoffice")


def get_pdf_engine() -> str | None:
    """Nama mesin PDF yang tersedia, atau None jika tidak ada."""
    if _word_available():
        return ENGINE_WORD
    if _libreoffice_path():
        return ENGINE_LIBREOFFICE
    return None


def convert_to_pdf(docx_path: Path, pdf_path: Path) -> Path:
    """
    Mengonversi satu file .docx menjadi .pdf.

    Raises:
        RuntimeError: jika tidak ada mesin PDF atau konversi gagal.
    """
    engine = get_pdf_engine()
    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    if engine == ENGINE_WORD:
        _convert_with_word(docx_path, pdf_path)
    elif engine == ENGINE_LIBREOFFICE:
        _convert_with_libreoffice(docx_path, pdf_path)
    else:
        raise RuntimeError("Tidak ada mesin PDF (Microsoft Word / LibreOffice) di komputer ini.")

    if not pdf_path.exists():
        raise RuntimeError(f"PDF tidak terbentuk untuk {docx_path.name}")
    return pdf_path


def _convert_with_word(docx_path: Path, pdf_path: Path) -> None:
    """Konversi lewat Microsoft Word (COM automation, khusus Windows)."""
    import pythoncom
    from docx2pdf import convert

    # Streamlit menjalankan kode di thread terpisah, jadi COM harus
    # diinisialisasi di thread ini dan dilepas lagi setelah selesai.
    pythoncom.CoInitialize()
    try:
        convert(str(docx_path), str(pdf_path))
    finally:
        pythoncom.CoUninitialize()


def _convert_with_libreoffice(docx_path: Path, pdf_path: Path) -> None:
    """Konversi lewat LibreOffice mode tanpa layar (headless)."""
    command = [
        _libreoffice_path() or "soffice",
        "--headless", "--convert-to", "pdf",
        "--outdir", str(pdf_path.parent), str(docx_path),
    ]
    completed = subprocess.run(
        command, capture_output=True, text=True, timeout=LIBREOFFICE_TIMEOUT_SECONDS
    )
    produced = pdf_path.parent / f"{docx_path.stem}.pdf"
    if completed.returncode != 0 or not produced.exists():
        raise RuntimeError(f"LibreOffice gagal: {completed.stderr.strip()[:200]}")
    if produced != pdf_path:
        produced.replace(pdf_path)
