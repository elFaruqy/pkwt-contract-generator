"""
excel_service: membaca data pekerja dari Excel dan menyiapkannya
agar siap divalidasi dan dipakai untuk generate kontrak.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import BinaryIO

import pandas as pd

from config.settings import REQUIRED_COLUMNS
from utils.logger import get_logger

logger = get_logger(__name__)

# Nama kolom bantu yang ditambahkan aplikasi (bukan dari Excel).
COL_NO_DATA = "NO DATA"        # urutan data mulai 1 (dipakai untuk pilih rentang)
COL_BARIS_EXCEL = "BARIS EXCEL"  # nomor baris di file Excel (baris 1 = judul kolom)


def _normalize_column_name(name: object) -> str:
    """'  Nama   pekerja ' -> 'NAMA PEKERJA' supaya beda spasi/kapital tidak bikin error."""
    return re.sub(r"\s+", " ", str(name)).strip().upper()


def load_excel(source: str | Path | BinaryIO) -> pd.DataFrame:
    """
    Membaca sheet pertama Excel. Semua sel dibaca sebagai teks supaya
    NOPEK/NIK tidak berubah jadi angka ilmiah (1.23E+15) atau kehilangan nol di depan.

    Raises:
        ValueError: jika file tidak bisa dibaca sebagai Excel.
    """
    try:
        df = pd.read_excel(source, dtype=str, engine="openpyxl")
    except Exception as exc:
        logger.error("Gagal membaca Excel: %s", exc)
        raise ValueError(f"File Excel tidak bisa dibaca: {exc}") from exc

    df.columns = [_normalize_column_name(col) for col in df.columns]
    df = df.fillna("")
    df = df.apply(lambda col: col.str.strip())

    # Buang baris yang benar-benar kosong di bagian bawah sheet.
    non_empty = df.ne("").any(axis=1)
    if non_empty.any():
        last_filled = non_empty[non_empty].index[-1]
        df = df.loc[:last_filled]

    df = df.reset_index(drop=True)
    df.insert(0, COL_NO_DATA, range(1, len(df) + 1))
    df.insert(1, COL_BARIS_EXCEL, range(2, len(df) + 2))

    logger.info("Excel terbaca: %d baris, %d kolom", len(df), len(df.columns) - 2)
    return df


def find_missing_columns(df: pd.DataFrame) -> list[str]:
    """Mengembalikan daftar kolom wajib yang tidak ditemukan di Excel."""
    return [col for col in REQUIRED_COLUMNS if col not in df.columns]


def search_rows(df: pd.DataFrame, keyword: str) -> pd.DataFrame:
    """
    Mencari kata kunci di semua kolom (tidak peka huruf besar/kecil).
    Pencarian dilakukan sebagai teks biasa, jadi karakter seperti "/" atau "("
    pada nomor kontrak tidak menyebabkan error.
    """
    keyword = keyword.strip()
    if not keyword:
        return df.iloc[0:0]
    data_cols = [c for c in df.columns if c not in (COL_NO_DATA, COL_BARIS_EXCEL)]
    mask = df[data_cols].apply(
        lambda col: col.astype(str).str.contains(keyword, case=False, regex=False)
    ).any(axis=1)
    return df[mask]


def slice_rows(df: pd.DataFrame, start_no: int, end_no: int) -> pd.DataFrame:
    """Mengambil data nomor `start_no` s.d. `end_no` (1-based, inklusif)."""
    start_no, end_no = sorted((int(start_no), int(end_no)))
    return df[(df[COL_NO_DATA] >= start_no) & (df[COL_NO_DATA] <= end_no)].copy()
