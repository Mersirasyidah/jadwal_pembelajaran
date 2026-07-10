import streamlit as st
import pandas as pd

# 1. Pengaturan Halaman Utama
st.set_page_config(
    page_title="Penyusun Jadwal SMP N 1 Bambanglipuro",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("📅 Sistem Informasi & Alokasi JP Jadwal Pelajaran")
st.write("Aplikasi visualisasi susunan jadwal pelajaran dan pengaturan target Jam Pelajaran (JP).")

# 2. Fungsi Memuat Data Master (Sheet SUSUN)
@st.cache_data
def load_data():
    try:
        # Mengambil data dari file CSV yang diunggah
        data = pd.read_csv("Jadwal Pelajaran Semua Kelas.xlsx - SUSUN.csv", header=None)
        return data
    except Exception as e:
        st.error(f"Gagal memuat file CSV: {e}")
        return None

df = load_data()

# 3. Sidebar: Input Alokasi JP per Mata Pelajaran
st.sidebar.header("⚙️ Pengaturan Alokasi JP")
st.sidebar.write("Masukkan target Jam Pelajaran (JP) per minggu untuk setiap mapel:")

# Input manual jumlah JP menggunakan dictionary
target_jp = {
    "Bahasa Indonesia": st.sidebar.number_input("Bahasa Indonesia (JP)", min_value=0, max_value=10, value=4),
    "Matematika": st.sidebar.number_input("Matematika (JP)", min_value=0, max_value=10, value=4),
    "IPA": st.sidebar.number_input("IPA (JP)", min_value=0, max_value=10, value=4),
    "Bahasa Inggris": st.sidebar.number_input("Bahasa Inggris (JP)", min_value=0, max_value=10, value=4),
    "PAI (Agama)": st.sidebar.number_input("PAI / Agama (JP)", min_value=0, max_value=10, value=3),
    "PPKn": st.sidebar.number_input("PPKn (JP)", min_value=0, max_value=10, value=2),
    "IPS": st.sidebar.number_input("IPS (JP)", min_value=0, max_value=10, value=4),
    "Seni Budaya / Prakarya": st.sidebar.number_input("Seni Budaya / Prakarya (JP)", min_value=0, max_value=10, value=3),
    "PJOK": st.sidebar.number_input("PJOK (JP)", min_value=0, max_value=10, value=3),
    "Bahasa Jawa": st.sidebar.number_input("Bahasa Jawa (JP)", min_value=0, max_value=10, value=2),
    "Informatika / TIK": st.sidebar.number_input("Informatika / TIK (JP)", min_value=0, max_value=10, value=2),
}

# Menampilkan ringkasan target JP yang dimasukkan pengguna dalam bentuk tabel kecil di sidebar
with st.sidebar.expander("📊 Lihat Ringkasan Target JP"):
    df_target = pd.DataFrame(list(target_jp.items()), columns=["Mata Pelajaran", "Target JP per Minggu"])
    st.dataframe(df_target, use_container_width=True, hide_index=True)


# 4. Menampilkan Tampilan Jadwal Format SUSUN
if df is not None:
    st.subheader("📋 Grid Susunan Jadwal Pelajaran (Format Excel)")
    st.info("💡 Geser tabel ke kanan (scroll horizontal) untuk melihat kelas lain dan hari berikutnya.")
    
    # Mengisi nilai kosong agar rapi secara visual
    df_filled = df.fillna("")
    
    # Tampilkan grid utama jadwal pelajaran
    st.dataframe(df_filled, use_container_width=True, height=450)
    
    # 5. Fitur Filter Cepat Berdasarkan Hari
    st.markdown("---")
    st.subheader("🔍 Filter Jadwal Berdasarkan Hari")
    pilihan_hari = st.selectbox("Pilih Hari:", ["Semua Hari", "SENIN", "SELASA"])
    
    if pilihan_hari == "SENIN":
        st.write("**Menampilkan Jadwal Khusus Hari SENIN (Semua Kelas 7, 8, 9)**")
        df_hari = df_filled.iloc[:, :46]
        st.dataframe(df_hari, use_container_width=True)
    elif pilihan_hari == "SELASA":
        st.write("**Menampilkan Jadwal Khusus Hari SELASA (Semua Kelas 7, 8, 9)**")
        kolom_selasa = [0] + list(range(46, 92))
        kolom_selasa = [k for k in kolom_selasa if k < len(df_filled.columns)]
        df_hari = df_filled.iloc[:, kolom_selasa]
        st.dataframe(df_hari, use_container_width=True)

else:
    st.warning("Pastikan file `Jadwal Pelajaran Semua Kelas.xlsx - SUSUN.csv` sudah diletakkan di folder aplikasi Anda.")
