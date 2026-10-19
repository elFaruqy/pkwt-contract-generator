"""
generate_service: mengatur proses generate kontrak secara massal.

Alur per pekerja:  data -> isi template (.docx) -> konversi PDF (opsional)
Jika satu pekerja gagal, proses TETAP lanjut ke pekerja berikutnya,
dan kegagalan dicatat di log + laporan.

Setiap batch disimpan di folder sendiri, mis. output/20261019_141530/,
lengkap dengan laporan_generate.csv sebagai jejak audit.
"""

from __future__ import annotations

import io
import zipfile
from collections.abc import Callable
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import pandas as pd

from services.document_service import build_context, render_contract
from services.excel_service import COL_BARIS_EXCEL
from services.pdf_service import convert_to_pdf
from utils.formatters import build_output_filename
from utils.logger import get_logger

logger = get_logger(__name__)

ProgressCallback = Callable[[int, int, str], None]


@dataclass
class GenerateResult:
    """Ringkasan hasil satu kali proses generate."""

    batch_folder: Path
    total: int = 0
    docx_ok: int = 0
    pdf_ok: int = 0
    pdf_requested: bool = False
    report_rows: list[dict[str, str]] = field(default_factory=list)

    @property
    def failures(self) -> list[dict[str, str]]:
        """Baris laporan yang statusnya bukan 'Berhasil'."""
        return [r for r in self.report_rows if r["Status"] != "Berhasil"]


def _unique_name(base_name: str, used: set[str]) -> str:
    """Mencegah file tertimpa jika ada NOPEK + nama yang sama."""
    name, counter = base_name, 2
    while name in used:
        name = f"{base_name}_{counter}"
        counter += 1
    used.add(name)
    return name


def generate_contracts(
    df: pd.DataFrame,
    template_path: Path,
    output_root: Path,
    make_pdf: bool,
    on_progress: ProgressCallback | None = None,
) -> GenerateResult:
    """Membuat kontrak untuk semua baris di `df`."""
    batch_folder = output_root / datetime.now().strftime("%Y%m%d_%H%M%S")
    word_folder, pdf_folder = batch_folder / "word", batch_folder / "pdf"
    result = GenerateResult(batch_folder=batch_folder, total=len(df), pdf_requested=make_pdf)
    used_names: set[str] = set()
    logger.info("Mulai generate %d kontrak (PDF=%s) ke %s", len(df), make_pdf, batch_folder)

    for position, (_, row) in enumerate(df.iterrows(), start=1):
        file_name = _unique_name(
            build_output_filename(row.get("NOPEK", ""), row.get("NAMA PEKERJA", "")), used_names
        )
        if on_progress:
            on_progress(position, len(df), file_name)

        report = {
            "Baris Excel": str(row[COL_BARIS_EXCEL]),
            "NOPEK": str(row.get("NOPEK", "")),
            "Nama Pekerja": str(row.get("NAMA PEKERJA", "")),
            "No Kontrak": str(row.get("NO KONTRAK", "")),
            "File": file_name,
            "Status": "Berhasil",
            "Keterangan": "",
        }
        try:
            docx_path = render_contract(template_path, build_context(row), word_folder / f"{file_name}.docx")
            result.docx_ok += 1
        except Exception as exc:  # satu kontrak gagal tidak boleh menghentikan batch
            logger.exception("Gagal membuat Word untuk %s", file_name)
            report.update(Status="Gagal Word", Keterangan=str(exc)[:200])
            result.report_rows.append(report)
            continue

        if make_pdf:
            try:
                convert_to_pdf(docx_path, pdf_folder / f"{file_name}.pdf")
                result.pdf_ok += 1
            except Exception as exc:
                logger.error("Gagal konversi PDF %s: %s", file_name, exc)
                report.update(Status="Gagal PDF", Keterangan=str(exc)[:200])
        result.report_rows.append(report)

    _write_report(result)
    logger.info(
        "Selesai: Word %d/%d, PDF %d/%d", result.docx_ok, result.total,
        result.pdf_ok, result.total if make_pdf else 0,
    )
    return result


def _write_report(result: GenerateResult) -> None:
    """Menyimpan laporan_generate.csv (bisa dibuka di Excel) sebagai jejak audit."""
    result.batch_folder.mkdir(parents=True, exist_ok=True)
    report_path = result.batch_folder / "laporan_generate.csv"
    pd.DataFrame(result.report_rows).to_csv(report_path, index=False, sep=";", encoding="utf-8-sig")


def build_zip(batch_folder: Path) -> bytes:
    """Mengemas seluruh isi folder batch menjadi satu file ZIP (di memori) untuk diunduh."""
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for file_path in sorted(batch_folder.rglob("*")):
            if file_path.is_file():
                archive.write(file_path, file_path.relative_to(batch_folder))
    return buffer.getvalue()
