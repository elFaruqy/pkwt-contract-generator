"""
Konfigurasi pusat aplikasi PKWT Contract Generator.

Semua "angka ajaib" (path folder, nama kolom Excel, aturan wajib isi)
dikumpulkan di sini supaya kalau ada perubahan format Excel, cukup
ubah file ini saja tanpa menyentuh logika aplikasi.
"""

from pathlib import Path
import platform

# ------------------------------------------------------------------
# Identitas aplikasi
# ------------------------------------------------------------------
APP_TITLE: str = "PKWT Contract Generator"
APP_VERSION: str = "1.5"
APP_ICON: str = "📄"

# Isi dengan nama & link LinkedIn Anda (dipakai di halaman "Tentang").
AUTHOR_NAME: str = "M Fadhillah Al Faruq"
AUTHOR_LINKEDIN: str = ""  # contoh: "https://www.linkedin.com/in/username-anda"

# ------------------------------------------------------------------
# Lokasi folder (otomatis relatif terhadap folder proyek,
# jadi tidak ada lagi path "D:\..." yang di-hardcode)
# ------------------------------------------------------------------
BASE_DIR: Path = Path(__file__).resolve().parent.parent
TEMPLATE_DIR: Path = BASE_DIR / "templates"
DATA_DIR: Path = BASE_DIR / "data"
OUTPUT_DIR: Path = BASE_DIR / "output"
LOG_DIR: Path = BASE_DIR / "logs"

SAMPLE_TEMPLATE: Path = TEMPLATE_DIR / "contoh_template_pkwt.docx"
SAMPLE_EXCEL: Path = DATA_DIR / "contoh_data_pekerja.xlsx"

# ------------------------------------------------------------------
# Deteksi lingkungan
# Konversi PDF lewat Microsoft Word hanya bisa di Windows.
# Di server Linux (mis. Streamlit Community Cloud) aplikasi otomatis
# berjalan dalam "mode demo": hasil berupa file Word.
# ------------------------------------------------------------------
IS_WINDOWS: bool = platform.system() == "Windows"
IS_DEMO_MODE: bool = not IS_WINDOWS

# ------------------------------------------------------------------
# Pemetaan kolom Excel -> placeholder di template Word
# Kiri  : judul kolom di Excel (huruf besar)
# Kanan : nama placeholder {{ ... }} di template .docx
# ------------------------------------------------------------------
COLUMN_TO_PLACEHOLDER: dict[str, str] = {
    "NO KONTRAK": "NO_KONTRAK",
    "WAKTU PEMBUATAN KONTRAK": "WAKTU_PEMBUATAN_KONTRAK",
    "MANAGER": "MANAGER",
    "NAMA PEKERJA": "NAMA_PEKERJA",
    "TTL PEKERJA": "TTL_PEKERJA",
    "ALAMAT PEKERJA": "ALAMAT_PEKERJA",
    "NIK PEKERJA": "NIK_PEKERJA",
    "NAMA PROYEK": "NAMA_PROYEK",
    "NO IO PROYEK": "NO_IO_PROYEK",
    "NOPEK": "NOPEK",
    "JABATAN PEKERJA": "JABATAN_PEKERJA",
    "TMT AWAL": "TMT_AWAL",
    "TMT AKHIR": "TMT_AKHIR",
    "LOKASI KERJA": "LOKASI_KERJA",
    "NAMA BAWAH": "NAMA_BAWAH",
    "UPAH TETAP": "UPAH_TETAP",
    "TUNJANGAN DAERAH": "TUNJANGAN_DAERAH",
    "TUNJANGAN LAINNYA": "TUNJANGAN_LAINNYA",
}

# Kolom yang wajib ADA di Excel (judul kolomnya harus tersedia).
REQUIRED_COLUMNS: list[str] = ["NO"] + list(COLUMN_TO_PLACEHOLDER.keys())

# Kolom yang wajib TERISI. Jika kosong -> generate diblokir.
MANDATORY_FILLED_COLUMNS: list[str] = [
    "NO KONTRAK",
    "WAKTU PEMBUATAN KONTRAK",
    "NAMA PEKERJA",
    "NOPEK",
    "TMT AWAL",
    "TMT AKHIR",
    "UPAH TETAP",
]

# Kolom nominal rupiah. Kosong/0 akan dicetak sebagai tanda "-".
CURRENCY_COLUMNS: list[str] = ["UPAH TETAP", "TUNJANGAN DAERAH", "TUNJANGAN LAINNYA"]

# Kolom tanggal yang dicek formatnya dan dibandingkan (awal <= akhir).
DATE_START_COLUMN: str = "TMT AWAL"
DATE_END_COLUMN: str = "TMT AKHIR"

# Kolom yang ditampilkan pada tabel preview.
PREVIEW_COLUMNS: list[str] = [
    "NO",
    "NAMA PEKERJA",
    "NOPEK",
    "NO KONTRAK",
    "JABATAN PEKERJA",
    "TMT AWAL",
    "TMT AKHIR",
    "WAKTU PEMBUATAN KONTRAK",
]

# Panjang NIK KTP yang benar.
NIK_LENGTH: int = 16

# Batas aman jumlah dokumen per sekali generate di mode demo online
# (supaya server gratis tidak kehabisan memori).
DEMO_MAX_ROWS: int = 30
