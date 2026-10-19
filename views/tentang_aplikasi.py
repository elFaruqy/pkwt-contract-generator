"""
Halaman "Tentang Aplikasi": menjelaskan masalah, solusi, dan dampak
untuk pengunjung demo (rekruter, atasan, rekan kerja).
"""

import streamlit as st

from config.settings import APP_ICON, APP_TITLE, APP_VERSION, AUTHOR_LINKEDIN, AUTHOR_NAME

st.title("📖 Tentang Aplikasi")
st.markdown(f"**{APP_ICON} {APP_TITLE}** · versi {APP_VERSION}")

st.header("Masalah")
st.markdown(
    """
Tim HR perlu membuat **ratusan kontrak PKWT** (Perjanjian Kerja Waktu Tertentu) untuk pekerja proyek.
Sebelumnya kontrak dibuat satu per satu dengan menyalin data dari Excel ke Word:

- **Lambat** — setiap kontrak diketik dan disimpan manual.
- **Rawan salah** — salah nomor kontrak, tanggal terbalik, atau upah tertukar baru ketahuan setelah dicetak.
- **Sulit diaudit** — tidak ada catatan kontrak mana yang sudah dibuat dan mana yang gagal.
"""
)

st.header("Solusi")
st.markdown(
    """
Aplikasi web internal yang berjalan di komputer HR (localhost), dengan alur:

**Excel → Validasi → Preview → Isi Template Word → Konversi PDF → Unduh / Simpan**

| Fitur | Manfaat untuk HR |
|---|---|
| Validasi sebelum cetak | Kontrak dengan data kosong, upah bukan angka, TMT terbalik, atau nomor kontrak dobel **diblokir** |
| Pencarian & pilih rentang | Cetak sebagian data saja tanpa mengedit file Excel |
| Template Inspector | Mendeteksi placeholder salah ketik di template Word sebelum proses massal |
| Generate massal Word + PDF | Puluhan kontrak dibuat dalam sekali klik, nama file seragam `NOPEK_NAMA.pdf` |
| Lanjut walau ada yang gagal | Satu dokumen bermasalah tidak menghentikan proses lainnya |
| Laporan & log | Setiap proses menghasilkan `laporan_generate.csv` dan log untuk jejak audit |
"""
)

st.header("Teknologi")
st.markdown(
    """
Python · Streamlit · Pandas · docxtpl (template Word) · docx2pdf + Microsoft Word (PDF di Windows) · logging.

Kode dipisah menjadi **UI** (folder `views/`) dan **services** (baca Excel, validasi, dokumen, PDF, generate)
sehingga mudah dirawat oleh satu orang.
"""
)

st.header("Tentang Demo Ini")
st.info(
    "Seluruh data, perusahaan, proyek, dan template di demo ini **fiktif**. "
    "Versi produksi dipakai secara internal di Windows; versi online ini membuat file Word "
    "karena server demo tidak memiliki Microsoft Word."
)

st.header("Pengembang")
author = f"[{AUTHOR_NAME}]({AUTHOR_LINKEDIN})" if AUTHOR_LINKEDIN else AUTHOR_NAME
st.markdown(f"Dibangun oleh **{author}** — HC Officer (Recruitment & Database) yang mengotomasi pekerjaan administrasi HR.")
