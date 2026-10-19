"""
Halaman "Generate Kontrak" (UI Streamlit).

File ini HANYA berisi tampilan. Semua logika ada di folder services/
dan utils/, sehingga tampilan bisa diubah tanpa mengganggu proses inti.
"""

from __future__ import annotations

import hashlib
import io
import os
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from config.settings import (
    APP_ICON, APP_TITLE, APP_VERSION, DEMO_MAX_ROWS, IS_DEMO_MODE,
    OUTPUT_DIR, PREVIEW_COLUMNS, SAMPLE_EXCEL, SAMPLE_TEMPLATE,
)
from services.document_service import inspect_template
from services.excel_service import (
    COL_BARIS_EXCEL, COL_NO_DATA, find_missing_columns, load_excel, search_rows, slice_rows,
)
from services.generate_service import GenerateResult, build_zip, generate_contracts
from services.pdf_service import get_pdf_engine
from services.validation_service import (
    LEVEL_ERROR, ValidationIssue, count_by_level, issues_to_dataframe, validate_batch,
)
from utils.formatters import normalize_date_text
from utils.logger import get_logger

logger = get_logger("generate_kontrak")
DATE_COLUMNS_TO_DISPLAY = ("TMT AWAL", "TMT AKHIR", "WAKTU PEMBUATAN KONTRAK")


# ------------------------------------------------------------------
# Bantuan kecil
# ------------------------------------------------------------------
@st.cache_data(show_spinner="Membaca Excel...")
def _load_excel_cached(file_bytes: bytes) -> pd.DataFrame:
    """Excel hanya dibaca ulang jika isi filenya berubah (lebih cepat)."""
    return load_excel(io.BytesIO(file_bytes))


def _for_display(df: pd.DataFrame) -> pd.DataFrame:
    """Tanggal bawaan Excel ditampilkan dalam format Indonesia, sama seperti di kontrak."""
    shown = df.copy()
    for col in DATE_COLUMNS_TO_DISPLAY:
        if col in shown.columns:
            shown[col] = shown[col].map(lambda v: normalize_date_text(v, with_weekday=col.startswith("WAKTU")))
    return shown


def _save_uploaded_template(file_bytes: bytes) -> Path:
    """docxtpl butuh file di disk, jadi template upload disimpan ke folder sementara."""
    digest = hashlib.md5(file_bytes).hexdigest()[:12]
    path = Path(tempfile.gettempdir()) / f"pkwt_template_{digest}.docx"
    if not path.exists():
        path.write_bytes(file_bytes)
    return path


# ------------------------------------------------------------------
# Bagian-bagian halaman
# ------------------------------------------------------------------
def render_sidebar() -> tuple[bytes | None, Path | None]:
    """Pilihan sumber data & template. Mengembalikan (isi Excel, path template)."""
    st.sidebar.header("⚙️ Sumber File")
    data_choice = st.sidebar.radio("Data pekerja", ["Data contoh (fiktif)", "Upload Excel sendiri"])
    excel_bytes: bytes | None = SAMPLE_EXCEL.read_bytes()
    if data_choice == "Upload Excel sendiri":
        uploaded = st.sidebar.file_uploader("File Excel (.xlsx)", type=["xlsx"])
        excel_bytes = uploaded.getvalue() if uploaded else None

    template_choice = st.sidebar.radio("Template kontrak", ["Template contoh", "Upload template sendiri"])
    template_path: Path | None = SAMPLE_TEMPLATE
    if template_choice == "Upload template sendiri":
        uploaded_tpl = st.sidebar.file_uploader("Template Word (.docx)", type=["docx"])
        template_path = _save_uploaded_template(uploaded_tpl.getvalue()) if uploaded_tpl else None

    st.sidebar.divider()
    engine = get_pdf_engine()
    st.sidebar.markdown(f"**Mesin PDF:** {engine or 'Word saja di demo online (PDF aktif di versi Windows)'}")
    if IS_DEMO_MODE:
        st.sidebar.warning("🌐 Mode demo online. Jangan upload data pekerja asli.")
    st.sidebar.caption(f"Versi {APP_VERSION}")
    return excel_bytes, template_path


def render_template_check(template_path: Path) -> None:
    """Template Inspector: memastikan placeholder template cocok dengan kolom Excel."""
    try:
        report = inspect_template(template_path)
    except Exception as exc:
        st.error(f"❌ Template tidak bisa dibaca: {exc}")
        st.stop()
    if report["unknown"]:
        st.warning(
            "⚠️ Template berisi placeholder yang tidak dikenal (kemungkinan salah ketik), "
            f"akan tercetak kosong: `{', '.join(report['unknown'])}`"
        )
    with st.expander(f"🧩 Template Inspector — {len(report['found'])} placeholder ditemukan"):
        st.write(", ".join(f"`{name}`" for name in report["found"]) or "-")
        if report["unused"]:
            st.caption("Kolom Excel yang tidak dipakai template: " + ", ".join(report["unused"]))


def render_search(df: pd.DataFrame) -> None:
    """Pencarian cepat untuk menemukan nomor data yang akan dicetak."""
    st.subheader("🔍 Cari Data")
    keyword = st.text_input("Kata kunci (NOPEK / nama / proyek / no kontrak)", placeholder="Contoh: 90000005")
    if not keyword:
        return
    found = search_rows(df, keyword)
    if found.empty:
        st.warning("Data tidak ditemukan.")
        return
    cols = [COL_NO_DATA, COL_BARIS_EXCEL] + [c for c in PREVIEW_COLUMNS if c in df.columns]
    st.success(f"Ditemukan {len(found)} data. Gunakan angka di kolom **{COL_NO_DATA}** untuk rentang cetak.")
    st.dataframe(_for_display(found[cols]), hide_index=True, width="stretch")


def render_range_selector(df: pd.DataFrame) -> pd.DataFrame:
    """Memilih rentang data (No Data awal s.d. akhir) yang akan dicetak."""
    st.subheader("🎯 Pilih Data yang Akan Dicetak")
    total = len(df)
    col_info, col_start, col_end = st.columns([2, 1, 1])
    col_info.info(f"📂 Total data di Excel: **{total}**")
    start = col_start.number_input("Mulai No Data", min_value=1, max_value=total, value=1)
    end = col_end.number_input("Sampai No Data", min_value=1, max_value=total, value=min(10, total))
    return slice_rows(df, start, end)


def render_validation(batch: pd.DataFrame, issues: list[ValidationIssue]) -> None:
    """Ringkasan validasi + tabel temuan + preview antrean cetak."""
    n_error, n_warning = count_by_level(issues)
    st.subheader("✅ Validasi Data")
    m1, m2, m3 = st.columns(3)
    m1.metric("Data dipilih", len(batch))
    m2.metric("Error (blokir)", n_error)
    m3.metric("Peringatan", n_warning)

    if issues:
        issues_df = issues_to_dataframe(issues)
        st.dataframe(issues_df, hide_index=True, width="stretch")
        st.download_button(
            "⬇️ Unduh daftar temuan (CSV)", issues_df.to_csv(index=False, sep=";").encode("utf-8-sig"),
            file_name="temuan_validasi.csv", mime="text/csv",
        )
    else:
        st.success("Semua data lengkap dan konsisten. Siap diproses!")

    worst: dict[int, str] = {}
    for item in issues:
        if worst.get(item.no_data) != LEVEL_ERROR:
            worst[item.no_data] = item.level
    preview = batch[[COL_NO_DATA] + [c for c in PREVIEW_COLUMNS if c in batch.columns]].copy()
    preview.insert(1, "STATUS", preview[COL_NO_DATA].map(
        lambda no: "❌ Error" if worst.get(no) == LEVEL_ERROR else ("⚠️ Cek" if no in worst else "✅ OK")
    ))
    with st.expander("📋 Preview antrean cetak", expanded=True):
        st.dataframe(_for_display(preview), hide_index=True, width="stretch")


def render_generate(batch: pd.DataFrame, template_path: Path, n_error: int) -> None:
    """Tombol generate. Diblokir jika masih ada error atau melebihi batas demo."""
    st.subheader("🚀 Generate Kontrak")
    engine = get_pdf_engine()
    make_pdf = st.checkbox(
        "Buat PDF juga", value=engine is not None, disabled=engine is None,
        help="Membutuhkan Microsoft Word (Windows) atau LibreOffice.",
    )
    over_limit = IS_DEMO_MODE and len(batch) > DEMO_MAX_ROWS
    if n_error:
        st.error(f"Generate diblokir: perbaiki dulu {n_error} error pada tabel validasi di atas.")
    if over_limit:
        st.warning(f"Mode demo dibatasi maksimal {DEMO_MAX_ROWS} kontrak per proses.")

    clicked = st.button(
        f"🚀 Generate {len(batch)} Kontrak", type="primary",
        disabled=bool(n_error) or over_limit, width="stretch",
    )
    if not clicked:
        return

    progress = st.progress(0.0, text="Menyiapkan...")

    def on_progress(position: int, total: int, file_name: str) -> None:
        progress.progress(position / total, text=f"Memproses [{position}/{total}] {file_name}")

    try:
        result = generate_contracts(batch, template_path, OUTPUT_DIR, make_pdf, on_progress)
        st.session_state["last_result"] = result
        st.session_state["last_zip"] = build_zip(result.batch_folder)
        progress.progress(1.0, text="Selesai!")
    except Exception as exc:
        logger.exception("Proses generate berhenti")
        st.error(f"Proses berhenti karena error tak terduga: {exc}")


def render_results() -> None:
    """Hasil generate terakhir (tetap tampil walau halaman di-refresh oleh Streamlit)."""
    result: GenerateResult | None = st.session_state.get("last_result")
    if result is None:
        return
    st.divider()
    st.subheader("📦 Hasil Generate Terakhir")
    c1, c2, c3 = st.columns(3)
    c1.metric("Total diproses", result.total)
    c2.metric("Word berhasil", f"{result.docx_ok}/{result.total}")
    c3.metric("PDF berhasil", f"{result.pdf_ok}/{result.total}" if result.pdf_requested else "-")

    if result.failures:
        st.error(f"{len(result.failures)} dokumen bermasalah (proses lain tetap berjalan):")
        st.dataframe(pd.DataFrame(result.failures), hide_index=True, width="stretch")

    st.download_button(
        "⬇️ Unduh Semua Hasil (ZIP)", st.session_state["last_zip"],
        file_name=f"kontrak_pkwt_{result.batch_folder.name}.zip", mime="application/zip",
        type="primary", width="stretch",
    )
    with st.expander("🧾 Laporan generate (juga tersimpan sebagai laporan_generate.csv)"):
        st.dataframe(pd.DataFrame(result.report_rows), hide_index=True, width="stretch")
    if not IS_DEMO_MODE:
        st.code(str(result.batch_folder), language=None)
        if st.button("📂 Buka Folder Hasil"):
            os.startfile(result.batch_folder)  # type: ignore[attr-defined]  # khusus Windows


# ------------------------------------------------------------------
# Alur utama halaman
# ------------------------------------------------------------------
def main() -> None:
    """Menyusun halaman dari atas ke bawah sesuai alur kerja HR."""
    st.title(f"{APP_ICON} {APP_TITLE}")
    st.caption("Generate kontrak PKWT massal dari Excel ke Word/PDF, dengan validasi data sebelum cetak.")
    if IS_DEMO_MODE:
        st.info(
            "**Versi demo publik** — memakai data & template fiktif. Versi produksi berjalan "
            "di Windows (localhost) dengan konversi PDF lewat Microsoft Word.\n\n"
            "💡 **Coba ini:** ubah **Sampai No Data** menjadi **25** untuk melihat validasi "
            "memblokir 5 data yang sengaja dibuat salah."
        )

    excel_bytes, template_path = render_sidebar()
    if excel_bytes is None or template_path is None:
        st.info("⬅️ Silakan upload file Excel dan template Word di sidebar.")
        st.stop()

    try:
        df = _load_excel_cached(excel_bytes)
    except ValueError as exc:
        st.error(f"❌ {exc}")
        st.stop()

    missing = find_missing_columns(df)
    if missing:
        st.error("❌ **Format Excel tidak sesuai.** Kolom berikut tidak ditemukan: " + ", ".join(missing))
        st.info("Pastikan judul kolom sama dengan file contoh di folder data/.")
        st.stop()
    if df.empty:
        st.warning("File Excel tidak berisi data.")
        st.stop()

    render_template_check(template_path)
    st.divider()
    render_search(df)
    st.divider()
    batch = render_range_selector(df)
    issues = validate_batch(batch)
    render_validation(batch, issues)
    st.divider()
    render_generate(batch, template_path, count_by_level(issues)[0])
    render_results()


main()
