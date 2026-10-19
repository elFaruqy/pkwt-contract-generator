"""
document_service: mengisi template Word (.docx) dengan data pekerja
menggunakan docxtpl, plus "Template Inspector" untuk mengecek
apakah placeholder di template cocok dengan kolom Excel.
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd
from docxtpl import DocxTemplate

from config.settings import COLUMN_TO_PLACEHOLDER, CURRENCY_COLUMNS
from utils.formatters import clean_text, format_rupiah, normalize_date_text
from utils.logger import get_logger

logger = get_logger(__name__)


def build_context(row: pd.Series) -> dict[str, str]:
    """
    Menyusun data satu pekerja menjadi dictionary {placeholder: nilai}
    yang siap dimasukkan ke template.
    """
    context: dict[str, str] = {}
    for column, placeholder in COLUMN_TO_PLACEHOLDER.items():
        value = row.get(column, "")
        if column in CURRENCY_COLUMNS:
            context[placeholder] = format_rupiah(value)
        elif column in ("TMT AWAL", "TMT AKHIR"):
            context[placeholder] = normalize_date_text(value)
        elif column == "WAKTU PEMBUATAN KONTRAK":
            context[placeholder] = normalize_date_text(value, with_weekday=True)
        else:
            context[placeholder] = clean_text(value)
    return context


def render_contract(template_path: Path, context: dict[str, str], output_path: Path) -> Path:
    """
    Mengisi template dan menyimpan hasilnya sebagai .docx.
    Template dibuka ulang untuk setiap kontrak agar data antar pekerja
    tidak saling tercampur.
    """
    document = DocxTemplate(str(template_path))
    document.render(context)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(str(output_path))
    return output_path


def inspect_template(template_path: Path) -> dict[str, list[str]]:
    """
    Template Inspector: membandingkan placeholder di template dengan
    placeholder yang disediakan aplikasi.

    Returns:
        {
          "found":   placeholder yang ada di template,
          "unknown": placeholder di template yang TIDAK punya kolom Excel
                     (akan tercetak kosong -> kemungkinan salah ketik),
          "unused":  kolom Excel yang tidak dipakai di template.
        }
    """
    document = DocxTemplate(str(template_path))
    found = set(document.get_undeclared_template_variables())
    known = set(COLUMN_TO_PLACEHOLDER.values())
    result = {
        "found": sorted(found),
        "unknown": sorted(found - known),
        "unused": sorted(known - found),
    }
    if result["unknown"]:
        logger.warning("Template berisi placeholder tak dikenal: %s", result["unknown"])
    return result
