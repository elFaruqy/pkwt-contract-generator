"""
Fungsi bantu untuk merapikan data: angka rupiah, tanggal Indonesia,
dan nama file. Tidak ada kode Streamlit di sini, jadi mudah diuji.
"""

from __future__ import annotations

import re
from datetime import date, datetime

_ILLEGAL_FILENAME_CHARS = r'[\\/:*?"<>|]'
_EMPTY_VALUES = {"", "NAN", "NONE", "NAT", "-"}

BULAN: list[str] = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember",
]
HARI: list[str] = ["Senin", "Selasa", "Rabu", "Kamis", "Jum'at", "Sabtu", "Minggu"]

# Ejaan bulan yang sering muncul di data HR (lengkap, singkatan, ejaan lama).
_MONTH_LOOKUP: dict[str, int] = {
    "januari": 1, "jan": 1,
    "februari": 2, "pebruari": 2, "feb": 2, "peb": 2,
    "maret": 3, "mar": 3,
    "april": 4, "apr": 4,
    "mei": 5,
    "juni": 6, "jun": 6,
    "juli": 7, "jul": 7,
    "agustus": 8, "agu": 8, "ags": 8, "agt": 8,
    "september": 9, "sep": 9, "sept": 9,
    "oktober": 10, "okt": 10,
    "november": 11, "nopember": 11, "nov": 11, "nop": 11,
    "desember": 12, "des": 12,
}


def is_empty(value: object) -> bool:
    """True jika sel Excel dianggap kosong ("", nan, None, "-")."""
    return value is None or str(value).strip().upper() in _EMPTY_VALUES


# ------------------------------------------------------------------
# Angka & Rupiah
# ------------------------------------------------------------------
def parse_number(value: object) -> float | None:
    """
    Mengubah isi sel menjadi angka. Mengembalikan None jika kosong
    atau bukan angka.

    Contoh yang didukung:
        5000000        -> 5000000.0
        "5000000.0"    -> 5000000.0   (angka Excel yang terbaca sebagai teks)
        "Rp 5.000.000" -> 5000000.0   (format Indonesia)
        "5,000,000"    -> 5000000.0   (format Inggris)
    """
    if is_empty(value):
        return None
    text = str(value).strip().upper().replace("RP.", "").replace("RP", "").replace(" ", "")

    if re.fullmatch(r"\d{1,3}(,\d{3})+", text):          # 5,000,000
        text = text.replace(",", "")
    elif re.fullmatch(r"\d{1,3}(\.\d{3})+(,\d+)?", text):  # 5.000.000 atau 5.000.000,00
        text = text.split(",")[0].replace(".", "")
    elif re.fullmatch(r"\d+,\d+", text):                  # 5000000,50
        text = text.replace(",", ".")

    try:
        return float(text)
    except ValueError:
        return None


def format_rupiah(value: object) -> str:
    """
    Format angka menjadi "Rp. 5.000.000".
    Nilai kosong atau 0 dikembalikan sebagai "-" (sesuai kebiasaan kontrak).
    Nilai yang bukan angka dikembalikan apa adanya (validasi sudah menandainya).
    """
    number = parse_number(value)
    if number is None:
        return "-" if is_empty(value) else str(value)
    if number == 0:
        return "-"
    return "Rp. " + f"{int(round(number)):,}".replace(",", ".")


# ------------------------------------------------------------------
# Tanggal
# ------------------------------------------------------------------
def parse_indonesian_date(value: object) -> date | None:
    """
    Membaca tanggal dari berbagai format yang biasa dipakai HR:
        "01 Januari 2026", "1 Jan 2026", "Jum'at, 02 Januari 2026",
        "2026-01-01", "2026-01-01 00:00:00", "01/01/2026", "01-01-2026".
    Mengembalikan None jika formatnya tidak dikenali.
    """
    if is_empty(value):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value

    text = str(value).strip()
    if "," in text:  # buang nama hari di depan, mis. "Senin, "
        text = text.split(",", 1)[1].strip()

    iso = re.fullmatch(r"(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T].*)?", text)
    if iso:
        year, month, day = (int(part) for part in iso.groups())
        return _safe_date(year, month, day)

    numeric = re.fullmatch(r"(\d{1,2})[/-](\d{1,2})[/-](\d{4})", text)
    if numeric:
        day, month, year = (int(part) for part in numeric.groups())
        return _safe_date(year, month, day)

    words = re.fullmatch(r"(\d{1,2})[\s-]+([A-Za-z]+)\.?[\s-]+(\d{4})", text)
    if words:
        month = _MONTH_LOOKUP.get(words.group(2).lower())
        if month:
            return _safe_date(int(words.group(3)), month, int(words.group(1)))
    return None


def _safe_date(year: int, month: int, day: int) -> date | None:
    """Membuat objek date; None jika tanggal mustahil (mis. 31 Februari)."""
    try:
        return date(year, month, day)
    except ValueError:
        return None


def format_indonesian_date(value: date, with_weekday: bool = False) -> str:
    """date(2026, 1, 2) -> "02 Januari 2026" atau "Jum'at, 02 Januari 2026"."""
    text = f"{value.day:02d} {BULAN[value.month - 1]} {value.year}"
    return f"{HARI[value.weekday()]}, {text}" if with_weekday else text


def normalize_date_text(value: object, with_weekday: bool = False) -> str:
    """
    Teks tanggal yang diketik HR dibiarkan apa adanya.
    Hanya tanggal bawaan Excel (mis. "2026-01-01 00:00:00") yang diubah
    menjadi format Indonesia, supaya tidak tercetak aneh di kontrak.
    """
    if is_empty(value):
        return ""
    text = str(value).strip()
    if re.fullmatch(r"\d{4}-\d{1,2}-\d{1,2}(?:[ T].*)?", text):
        parsed = parse_indonesian_date(text)
        if parsed:
            return format_indonesian_date(parsed, with_weekday)
    return text


# ------------------------------------------------------------------
# Teks & Nama File
# ------------------------------------------------------------------
def clean_text(value: object) -> str:
    """Merapikan isi sel: kosong -> "", angka bulat "123.0" -> "123"."""
    if is_empty(value):
        return ""
    text = str(value).strip()
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return text


def build_output_filename(nopek: str, nama_pekerja: str) -> str:
    """
    Nama file sesuai rule bisnis: NOPEK_NAMA_PEKERJA
    Contoh: ("123456", "Budi Santoso") -> "123456_BUDI_SANTOSO"
    Karakter ilegal \\/:*?"<>| dihapus, spasi diganti garis bawah.
    """
    raw = f"{clean_text(nopek)}_{clean_text(nama_pekerja)}"
    raw = re.sub(_ILLEGAL_FILENAME_CHARS, "", raw)
    raw = re.sub(r"\s+", "_", raw.strip()).upper()
    raw = re.sub(r"_+", "_", raw).strip("_.")
    return raw[:120] or "TANPA_NAMA"
