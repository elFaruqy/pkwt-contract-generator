"""
PKWT Contract Generator — pintu masuk aplikasi.

File ini hanya mengatur menu halaman. Tampilan tiap halaman ada di
folder views/, logika proses ada di folder services/.

Jalankan dengan:  streamlit run app.py
"""

import streamlit as st

from config.settings import APP_ICON, APP_TITLE

st.set_page_config(page_title=APP_TITLE, page_icon=APP_ICON, layout="wide")

navigation = st.navigation([
    st.Page("views/generate_kontrak.py", title="Generate Kontrak", icon="📄", default=True),
    st.Page("views/tentang_aplikasi.py", title="Tentang Aplikasi", icon="📖"),
])
navigation.run()
