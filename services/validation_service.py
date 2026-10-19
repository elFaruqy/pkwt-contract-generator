"""
validation_service: memeriksa data sebelum kontrak dibuat.

Dua tingkat temuan:
- ERROR      : generate DIBLOKIR (mis. No Kontrak kosong, TMT Akhir < TMT Awal).
- PERINGATAN : generate tetap boleh, tapi HR sebaiknya mengecek
               (mis. NIK tidak 16 digit, alamat kosong).

Prinsipnya: lebih baik aplikasi menolak 1 kontrak yang salah
daripada 1 kontrak salah terlanjur ditandatangani.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass

import pandas as pd

from config.settings import (
    COLUMN_TO_PLACEHOLDER,
    CURRENCY_COLUMNS,
    DATE_END_COLUMN,
    DATE_START_COLUMN,
    MANDATORY_FILLED_COLUMNS,
    NIK_LENGTH,
)
from services.excel_service import COL_BARIS_EXCEL, COL_NO_DATA
from utils.formatters import is_empty, parse_indonesian_date, parse_number

LEVEL_ERROR = "ERROR"
LEVEL_WARNING = "PERINGATAN"


@dataclass
class ValidationIssue:
    """Satu temuan validasi pada satu sel."""

    no_data: int
    baris_excel: int
    nopek: str
    nama_pekerja: str
    kolom: str
    masalah: str
    level: str


def validate_batch(df: pd.DataFrame) -> list[ValidationIssue]:
    """Menjalankan semua pemeriksaan pada data yang akan di-generate."""
    issues: list[ValidationIssue] = []
    for _, row in df.iterrows():
        issues.extend(_validate_row(row))
    issues.extend(_check_duplicates(df, "NO KONTRAK", LEVEL_ERROR))
    issues.extend(_check_duplicates(df, "NOPEK", LEVEL_WARNING))
    return sorted(issues, key=lambda item: (item.level != LEVEL_ERROR, item.no_data))


def _issue(row: pd.Series, kolom: str, masalah: str, level: str) -> ValidationIssue:
    """Membuat objek temuan dari satu baris data."""
    return ValidationIssue(
        no_data=int(row[COL_NO_DATA]),
        baris_excel=int(row[COL_BARIS_EXCEL]),
        nopek=str(row.get("NOPEK", "")),
        nama_pekerja=str(row.get("NAMA PEKERJA", "")),
        kolom=kolom,
        masalah=masalah,
        level=level,
    )


def _validate_row(row: pd.Series) -> list[ValidationIssue]:
    """Pemeriksaan per baris: wajib isi, angka, tanggal, dan NIK."""
    issues: list[ValidationIssue] = []

    # 1. Kolom wajib isi
    for col in MANDATORY_FILLED_COLUMNS:
        if is_empty(row.get(col)):
            issues.append(_issue(row, col, "Wajib diisi, tetapi kosong", LEVEL_ERROR))

    # 2. Kolom lain yang dipakai template tapi kosong -> peringatan
    optional_cols = [
        c for c in COLUMN_TO_PLACEHOLDER
        if c not in MANDATORY_FILLED_COLUMNS and c not in CURRENCY_COLUMNS
    ]
    for col in optional_cols:
        if is_empty(row.get(col)):
            issues.append(_issue(row, col, "Kosong, akan tercetak kosong di kontrak", LEVEL_WARNING))

    # 3. Nominal rupiah harus berupa angka
    for col in CURRENCY_COLUMNS:
        value = row.get(col)
        if not is_empty(value) and parse_number(value) is None:
            issues.append(_issue(row, col, f"Bukan angka: '{value}'", LEVEL_ERROR))
    upah = parse_number(row.get("UPAH TETAP"))
    if upah is not None and upah <= 0:
        issues.append(_issue(row, "UPAH TETAP", "Upah tetap harus lebih dari 0", LEVEL_ERROR))

    # 4. Tanggal TMT harus terbaca dan urutannya benar
    issues.extend(_check_dates(row))

    # 5. NIK KTP harus 16 digit angka
    nik = str(row.get("NIK PEKERJA", "")).strip()
    if nik and (not nik.isdigit() or len(nik) != NIK_LENGTH):
        issues.append(
            _issue(row, "NIK PEKERJA", f"NIK harus {NIK_LENGTH} digit angka (terbaca: {len(nik)} karakter)", LEVEL_WARNING)
        )
    return issues


def _check_dates(row: pd.Series) -> list[ValidationIssue]:
    """TMT Awal & TMT Akhir harus format tanggal yang dikenali, dan Akhir >= Awal."""
    issues: list[ValidationIssue] = []
    parsed: dict[str, object] = {}
    for col in (DATE_START_COLUMN, DATE_END_COLUMN):
        value = row.get(col)
        if is_empty(value):
            continue  # sudah ditangkap pemeriksaan wajib isi
        parsed[col] = parse_indonesian_date(value)
        if parsed[col] is None:
            issues.append(
                _issue(row, col, f"Format tanggal tidak dikenali: '{value}' (contoh benar: 01 Januari 2026)", LEVEL_ERROR)
            )
    start, end = parsed.get(DATE_START_COLUMN), parsed.get(DATE_END_COLUMN)
    if start and end and end < start:
        issues.append(_issue(row, DATE_END_COLUMN, "TMT Akhir lebih awal dari TMT Awal", LEVEL_ERROR))
    return issues


def _check_duplicates(df: pd.DataFrame, column: str, level: str) -> list[ValidationIssue]:
    """Menandai nilai yang muncul lebih dari sekali dalam data yang dipilih."""
    if column not in df.columns:
        return []
    filled = df[~df[column].apply(is_empty)]
    duplicated = filled[filled[column].duplicated(keep=False)]
    return [
        _issue(row, column, f"Nilai '{row[column]}' dipakai lebih dari satu kali", level)
        for _, row in duplicated.iterrows()
    ]


def count_by_level(issues: list[ValidationIssue]) -> tuple[int, int]:
    """Mengembalikan (jumlah ERROR, jumlah PERINGATAN)."""
    errors = sum(1 for item in issues if item.level == LEVEL_ERROR)
    return errors, len(issues) - errors


def issues_to_dataframe(issues: list[ValidationIssue]) -> pd.DataFrame:
    """Mengubah daftar temuan menjadi tabel untuk ditampilkan/diunduh."""
    columns = ["level", "no_data", "baris_excel", "nopek", "nama_pekerja", "kolom", "masalah"]
    df = pd.DataFrame([asdict(item) for item in issues], columns=columns)
    return df.rename(columns={
        "level": "Level", "no_data": "No Data", "baris_excel": "Baris Excel",
        "nopek": "NOPEK", "nama_pekerja": "Nama Pekerja", "kolom": "Kolom", "masalah": "Masalah",
    })
