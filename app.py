import streamlit as st
import pandas as pd

# 1. Pengaturan Konfigurasi Halaman Streamlit
st.set_page_config(
    page_title="Viewer Jadwal Sekolah SMP N 1 Bambanglipuro",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("📅 Aplikasi Viewer Jadwal Sekolah")
st.caption("Membaca file: Jadwal_SMP_N_1_Bambanglipuro_TA_2025-2026_narwi_final_LATIHAN.xlsx")

# Nama file target (pastikan file ini berada di folder yang sama dengan app.py)
FILE_NAME = "Jadwal_SMP_N_1_Bambanglipuro_TA_2025-2026_narwi_final_LATIHAN.xlsx"

# 2. Sidebar untuk Navigasi Sheet
st.sidebar.header("Navigasi Data")
sheet_pilihan = st.sidebar.selectbox(
    "Pilih Sheet yang ingin ditampilkan:",
    ["HARIAN", "REK-KELAS", "DATA_WALIKELAS", "Tbh", "REK-PERGURU", "ABSENT"]
)

# 3. Proses Membaca Data dengan Error Handling agar tidak Tampil Blank
try:
    # Menggunakan engine 'openpyxl' untuk membaca format .xlsx dengan aman
    # Kita gunakan try-except agar jika satu sheet error, aplikasi tidak mati total (blank)
    df = pd.read_excel(FILE_NAME, sheet_name=sheet_pilihan, engine="openpyxl")
    
    # 4. Pembersihan Data Otomatis (Data Cleaning)
    # Menghapus baris atau kolom yang seluruhnya kosong (NaN) agar tidak memberatkan render Streamlit
    df_clean = df.dropna(how='all')
    
    # Menghapus kolom yang seluruhnya kosong
    df_clean = df_clean.dropna(axis=1, how='all')
    
    # Mengisi nilai NaN yang tersisa dengan string kosong agar tampilan lebih rapi
    df_clean = df_clean.fillna("")

    # 5. Menampilkan Data ke Streamlit
    st.subheader(f"📄 Data pada Lembar: {sheet_pilihan}")
    st.write(f"Menampilkan {df_clean.shape[0]} baris data yang terisi.")
    
    # Menggunakan st.dataframe agar tabel bisa di-scroll, di-search, dan di-download oleh user
    st.dataframe(df_clean, use_container_width=True)

except FileNotFoundError:
    st.error(f"❌ **File Tidak Ditemukan!** Pastikan file `{FILE_NAME}` sudah diletakkan di folder yang sama dengan kode `app.py` ini.")
    
except Exception as e:
    st.error(f"❌ **Terjadi Kesalahan saat Membaca Sheet '{sheet_pilihan}':**")
    st.info("Hal ini biasanya terjadi karena rumus Excel yang terlalu kompleks atau adanya proteksi cell.")
    # Menampilkan detail error di bawah box agar memudahkan debugging
    st.code(str(e), language="python")

# 6. Tips Tambahan di Sidebar
st.sidebar.markdown("---")
st.sidebar.info(
    "💡 **Tips Jika Mengalami Blank:**\n"
    "1. Pastikan library `openpyxl` sudah terinstal (`pip install openpyxl`).\n"
    "2. Jika data masih kosong, pastikan identitas sekolah pada file Excel asli sudah diisi."
)
