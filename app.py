import streamlit as st
import pandas as pd

# 1. Pengaturan Halaman Utama
st.set_page_config(
    page_title="Sistem Jadwal SMP N 1 Bambanglipuro",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("📅 Sistem Informasi Kurikulum & Alokasi JP")
st.write("Aplikasi manajemen data master guru, pembagian kelas ampunan, dan pengaturan target Jam Pelajaran (JP).")

# 2. DATA MASTER GURU & JP (Dibuat langsung di dalam kode agar terhindar dari Error No Such File)
@st.cache_data
def get_data_master_default():
    # Data ini disesuaikan dengan contoh sampel data SMP N 1 Bambanglipuro
    data_guru = [
        {"No": 1, "Nama Guru": "KIRNO WIDARSO, M.Pd., M.M", "Mata Pelajaran": "Kepala Sekolah / PAI", "Kelas Ampuan": "7A, 7B", "JP per Kelas": 2, "Total JP": 4},
        {"No": 2, "Nama Guru": "Hartini, M.Pd", "Mata Pelajaran": "Bahasa Indonesia", "Kelas Ampuan": "7A, 7B, 7C, 7D", "JP per Kelas": 4, "Total JP": 16},
        {"No": 3, "Nama Guru": "Bartina, S.Pd", "Mata Pelajaran": "Matematika", "Kelas Ampuan": "8D, 9F", "JP per Kelas": 4, "Total JP": 8},
        {"No": 4, "Nama Guru": "Ani Pujiastuti, M.Hum", "Mata Pelajaran": "Bahasa Inggris", "Kelas Ampuan": "7A, 7B, 7C", "JP per Kelas": 4, "Total JP": 12},
        {"No": 5, "Nama Guru": "Agus Fuadi, M.Pd", "Mata Pelajaran": "IPA", "Kelas Ampuan": "9E, 9F", "JP per Kelas": 5, "Total JP": 10},
        {"No": 6, "Nama Guru": "Dra. Isti Widayanti", "Mata Pelajaran": "IPS", "Kelas Ampuan": "7D, 7F", "JP per Kelas": 4, "Total JP": 8},
        {"No": 7, "Nama Guru": "Martina Supraptini, S.Pd", "Mata Pelajaran": "Seni Budaya", "Kelas Ampuan": "8A, 8B, 8C", "JP per Kelas": 3, "Total JP": 9},
        {"No": 8, "Nama Guru": "Dwi Indriyani, S.Pd", "Mata Pelajaran": "Bahasa Jawa", "Kelas Ampuan": "9B, 9C", "JP per Kelas": 2, "Total JP": 4},
        {"No": 9, "Nama Guru": "Sugiyanto, M.Pd", "Mata Pelajaran": "PJOK", "Kelas Ampuan": "8C, 8G", "JP per Kelas": 3, "Total JP": 6},
        {"No": 10, "Nama Guru": "Siti Herwulan, S.Pd", "Mata Pelajaran": "Matematika", "Kelas Ampuan": "7G, 8A", "JP per Kelas": 4, "Total JP": 8},
        {"No": 11, "Nama Guru": "Sunarwi, S.Pd", "Mata Pelajaran": "PPKn", "Kelas Ampuan": "7E, 7F", "JP per Kelas": 2, "Total JP": 4},
        {"No": 12, "Nama Guru": "Herlita Dewi Setyawati, S.Pd", "Mata Pelajaran": "Informatika / TIK", "Kelas Ampuan": "7F, 7G", "JP per Kelas": 2, "Total JP": 4}
    ]
    return pd.DataFrame(data_guru)

# Menyimpan data master ke dalam session state agar perubahan data bersifat interaktif
if 'df_master' not in st.session_state:
    st.session_state.df_master = get_data_master_default()

# 3. FITUR UTAMA: DATA MASTER GURU (Bisa Diedit Langsung)
st.subheader("📋 Data Master Guru & Alokasi Mengampu")
st.info("💡 Anda dapat mengubah langsung isi tabel di bawah ini (klik dua kali pada cell). Total JP otomatis terhitung jika Anda mengubah Kelas Ampuan atau JP per Kelas.")

# Menampilkan editor tabel data master
edited_df = st.data_editor(
    st.session_state.df_master, 
    use_container_width=True, 
    num_rows="dynamic",
    hide_index=True
)

# Tombol untuk kalkulasi ulang total JP berdasarkan input pengguna
if st.button("🔄 Hitung Ulang & Simpan Data Master"):
    # Fungsi menghitung jumlah kelas (misal "7A, 7B" = 2 kelas)
    def hitung_total_jp(row):
        try:
            list_kelas = str(row["Kelas Ampuan"]).split(",")
            jumlah_kelas = len([k.strip() for k in list_kelas if k.strip() != ""])
            return jumlah_kelas * int(row["JP per Kelas"])
        except:
            return 0

    edited_df["Total JP"] = edited_df.apply(hitung_total_jp, axis=1)
    st.session_state.df_master = edited_df
    st.success("Data Master berhasil diperbarui dan disinkronkan!")
    st.rerun()

# 4. RINGKASAN AGREGASI JP (Metrik Dashboard)
st.markdown("---")
st.subheader("📊 Analisis Distribusi Jam Pelajaran (JP)")

col1, col2, col3 = st.columns(3)
with col1:
    total_guru = len(st.session_state.df_master)
    st.metric(label="Jumlah Guru Aktif", value=f"{total_guru} Orang")
with col2:
    total_kumulatif_jp = st.session_state.df_master["Total JP"].sum()
    st.metric(label="Total Beban JP Sekolah / Minggu", value=f"{total_kumulatif_jp} JP")
with col3:
    rata_rata_jp = st.session_state.df_master["JP per Kelas"].mean()
    st.metric(label="Rata-rata JP per Mapel", value=f"{rata_rata_jp:.1f} JP")

# 5. PILIHAN BACKUP: JIKA TETAP INGIN MEMUAT FILE EXCEL/CSV JADWAL
st.markdown("---")
with st.expander("📂 Opsi Lanjutan: Muat File Jadwal Sinkronisasi Excel (Format SUSUN)"):
    st.write("Jika Anda memiliki file CSV jadwal mingguan, Anda bisa mengunggahnya secara langsung di bawah ini untuk menghindari error direktori folder:")
    uploaded_file = st.file_uploader("Pilih file CSV Jadwal (Format SUSUN)", type=["csv"])
    
    if uploaded_file is not None:
        try:
            df_excel = pd.read_csv(uploaded_file, header=None).fillna("")
            st.write("✅ File Berhasil Dimuat:")
            st.dataframe(df_excel, use_container_width=True)
        except Exception as e:
            st.error(f"Gagal membaca file: {e}")
