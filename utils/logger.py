"""
Pengaturan logging aplikasi.

Semua kejadian penting (data dibaca, kontrak dibuat, error) dicatat ke
file logs/app.log. File log otomatis "diputar" (rotate) saat mencapai
1 MB supaya tidak membengkak, dan disimpan maksimal 5 file cadangan.
"""

import logging
from logging.handlers import RotatingFileHandler

from config.settings import LOG_DIR

_LOG_FORMAT = "%(asctime)s | %(levelname)-8s | %(name)s | %(message)s"


def get_logger(name: str) -> logging.Logger:
    """
    Mengembalikan logger yang menulis ke console dan ke logs/app.log.

    Aman dipanggil berulang kali: Streamlit menjalankan ulang script
    setiap ada interaksi, jadi handler hanya dipasang sekali agar baris
    log tidak tercatat dobel.
    """
    logger = logging.getLogger(name)
    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)
    formatter = logging.Formatter(_LOG_FORMAT)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    try:
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        file_handler = RotatingFileHandler(
            LOG_DIR / "app.log", maxBytes=1_000_000, backupCount=5, encoding="utf-8"
        )
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    except OSError:
        # Jika folder logs tidak bisa ditulis (jarang terjadi),
        # aplikasi tetap jalan dengan log di console saja.
        logger.warning("Folder log tidak bisa ditulis, log hanya tampil di console.")

    logger.propagate = False
    return logger
