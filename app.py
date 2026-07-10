import streamlit as st
import pandas as pd

# Pengaturan halaman dasar
st.set_page_config(
    page_title="Aplikasi Jadwal Pelajaran SMP N 1 Bambanglipuro",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("📅 Jadwal Pelajaran Berdasarkan Format Susun")
st.write("Menampilkan susunan jadwal pelajaran per jam dan per kelas sesuai dengan data master.")

# Fungsi untuk memuat data
@st.cache_data
def load_data():
    # Menggunakan file CSV hasil konversi dari sheet SUSUN Anda
    # Ganti nama file ini sesuai dengan file yang Anda simpan di direktori script Anda
    try:
        data = pd.read_csv("Jadwal Pelajaran Semua Kelas.xlsx - SUSUN.csv", header=None)
        return data
    except Exception as e:
        st.error(f"Gagal memuat file: {e}")
        return None

df = load_data()

if df is not None:
    # Memproses header seperti tampilan excel Anda
    # Baris 0: Semester/Tahun Ajaran (Info judul tambahan)
    # Baris 1: Nama Hari (SENIN, SELASA, dst)
    # Baris 2: Nama Kelas (7A, 7B, 8A, dst)
    
    st.info("💡 Gunakan fitur scroll horizontal pada tabel di bawah untuk melihat jadwal kelas 8, kelas 9, serta hari berikutnya.")
    
    # Membersihkan tampilan DataFrame agar baris atas menjadi header visual
    # Mengganti baris NaN dengan string kosong untuk keindahan visual
    df_filled = df.fillna("")
    
    # Kustomisasi CSS agar tabel streamlit menyerupai grid Excel yang rapi
    st.markdown("""
        <style>
        .reportview-container .main .block-container{
            max-width: 95%;
        }
        div[data-testid="stDataFrame"] {
            font-family: 'Courier New', Courier, monospace;
        }
        </style>
        """, unsafe_allow_html=True)
    
    # Menampilkan keseluruhan grid jadwal sesuai struktur file asli Anda
    # Kita menggunakan st.dataframe dengan ukuran kontainer penuh
    st.dataframe(df_filled, use_container_width=True, height=600)
    
    # Fitur Tambahan: Filter Berdasarkan Hari agar lebih mudah dibaca
    st.sidebar.header("🛠️ Panel Kontrol Tampilan")
    pilihan_hari = st.sidebar.selectbox("Pilih Hari untuk Difilter:", ["Semua Hari", "SENIN", "SELASA"])
    
    if pilihan_hari != "Semua Hari":
        st.subheader(f"📌 Tampilan Khusus Hari {pilihan_hari}")
        # Logika sederhana memotong kolom berdasarkan struktur koordinat file Anda
        if pilihan_hari == "SENIN":
            # Hari Senin biasanya ada di kolom awal setelah JAM KE (kolom index 0 sampai 45)
            df_hari = df_filled.iloc[:, :46]
            st.dataframe(df_hari, use_container_width=True)
        elif pilihan_hari == "SELASA":
            # Hari Selasa dimulai dari pembatas kolom berikutnya (kolom index 0 untuk JAM KE, dan kolom 46 ke atas)
            kolom_selasa = [0] + list(range(46, 92))
            # Memastikan index tidak out of bounds
            kolom_selasa = [k for k in kolom_selasa if k < len(df_filled.columns)]
            df_hari = df_filled.iloc[:, kolom_selasa]
            st.dataframe(df_hari, use_container_width=True)
else:
    st.warning("Silakan pastikan file `Jadwal Pelajaran Semua Kelas.xlsx - SUSUN.csv` berada dalam satu folder dengan aplikasi ini.")
