import streamlit as st
import pandas as pd

# 1. Pengaturan Layout Halaman
st.set_page_config(
    page_title="Data Master & JP - SMP N 1 Bambanglipuro",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Judul Aplikasi
st.title("📅 Sistem Informasi Kurikulum & Alokasi JP")
st.write("Aplikasi Manajemen Data Master Guru, Pembagian Kelas Ampuan, dan Target Jam Pelajaran (JP) tanpa ketergantungan file eksternal.")

# 2. DATABASE INTERNAL (Data Master Riil SMP N 1 Bambanglipuro TA 2025/2026)
@st.cache_data
def load_database_guru():
    # Mengompilasi data guru, mapel, dan kelas ampuannya dari berkas rancangan kurikulum sekolah
    data = [
        {"No": 1, "Nama Guru": "KIRNO WIDARSO, M.Pd., M.M", "Mata Pelajaran": "Kepala Sekolah / PAI", "Kelas Ampuan": "7A, 7B", "JP per Kelas": 2},
        {"No": 2, "Nama Guru": "Hartini, M.Pd", "Mata Pelajaran": "Bahasa Indonesia", "Kelas Ampuan": "7A, 7B, 7C, 7D", "JP per Kelas": 4},
        {"No": 3, "Nama Guru": "Bartina, S.Pd", "Mata Pelajaran": "Matematika", "Kelas Ampuan": "8D, 9F", "JP per Kelas": 4},
        {"No": 4, "Nama Guru": "Ani Pujiastuti, M.Hum", "Mata Pelajaran": "Bahasa Inggris", "Kelas Ampuan": "7A, 7B, 7C", "JP per Kelas": 4},
        {"No": 5, "Nama Guru": "Agus Fuadi, M.Pd", "Mata Pelajaran": "IPA", "Kelas Ampuan": "9E, 9F", "JP per Kelas": 5},
        {"No": 6, "Nama Guru": "Dra. Isti Widayanti", "Mata Pelajaran": "IPS", "Kelas Ampuan": "7D, 7F", "JP per Kelas": 4},
        {"No": 7, "Nama Guru": "Martina Supraptini, S.Pd", "Mata Pelajaran": "Seni Budaya", "Kelas Ampuan": "8A, 8B, 8C", "JP per Kelas": 3},
        {"No": 8, "Nama Guru": "Dwi Indriyani, S.Pd", "Mata Pelajaran": "Bahasa Jawa", "Kelas Ampuan": "9B, 9C", "JP per Kelas": 2},
        {"No": 9, "Nama Guru": "Sugiyanto, M.Pd", "Mata Pelajaran": "PJOK", "Kelas Ampuan": "8C, 8G", "JP per Kelas": 3},
        {"No": 10, "Nama Guru": "Siti Herwulan, S.Pd", "Mata Pelajaran": "Matematika", "Kelas Ampuan": "7G, 8A", "JP per Kelas": 4},
        {"No": 11, "Nama Guru": "Sunarwi, S.Pd", "Mata Pelajaran": "PPKn", "Kelas Ampuan": "7E, 7F", "JP per Kelas": 2},
        {"No": 12, "Nama Guru": "Herlita Dewi Setyawati, S.Pd", "Mata Pelajaran": "Informatika / TIK", "Kelas Ampuan": "7F, 7G", "JP per Kelas": 2},
        {"No": 13, "Nama Guru": "Khudlori, S.Ag", "Mata Pelajaran": "PAI / Agama", "Kelas Ampuan": "7A, 7C, 7D", "JP per Kelas": 3},
        {"No": 14, "Nama Guru": "Fitri Kurnia Pangestuti, S.Pd", "Mata Pelajaran": "Bahasa Indonesia", "Kelas Ampuan": "7C, 7D", "JP per Kelas": 4},
        {"No": 15, "Nama Guru": "Candra Kholifatun, S.Pd", "Mata Pelajaran": "IPA", "Kelas Ampuan": "7A, 7B", "JP per Kelas": 5},
        {"No": 16, "Nama Guru": "Primastuti Setyaningsih, S.Pd", "Mata Pelajaran": "IPS", "Kelas Ampuan": "7A, 7B, 7C, 7E, 7F", "JP per Kelas": 4},
        {"No": 17, "Nama Guru": "Rifki Ardhi Fahruzi, S.Pd", "Mata Pelajaran": "PJOK", "Kelas Ampuan": "7A, 7B, 7C", "JP per Kelas": 3}
    ]
    df = pd.DataFrame(data)
    
    # Menghitung Total JP awal secara otomatis berdasarkan jumlah kelas yang tertera
    def hitung_awal(row):
        kelas_list = [k.strip() for k in str(row["Kelas Ampuan"]).split(",") if k.strip() != ""]
        return len(kelas_list) * int(row["JP per Kelas"])
        
    df["Total JP"] = df.apply(hitung_awal, axis=1)
    return df

# Memasukkan data ke dalam session state Streamlit agar aplikasi mengingat perubahan data dari user
if 'database_kurikulum' not in st.session_state:
    st.session_state.database_kurikulum = load_database_guru()

# 3. BAGIAN METRIK / SUMMARY DASHBOARD
st.subheader("📊 Analisis Distribusi Jam Pelajaran (JP) Sekolah")
col1, col2, col3, col4 = st.columns(4)

df_aktif = st.session_state.database_kurikulum

with col1:
    st.metric(label="Jumlah Guru Terdaftar", value=f"{len(df_aktif)} Guru")
with col2:
    st.metric(label="Total Beban Kerja", value=f"{df_aktif['Total JP'].sum()} JP / Minggu")
with col3:
    st.metric(label="Rata-rata Mengajar", value=f"{df_aktif['Total JP'].mean():.1f} JP")
with col4:
    st.metric(label="Maksimal JP Individu", value=f"{df_aktif['Total JP'].max()} JP")

# 4. TABEL INTERAKTIF DATA MASTER
st.markdown("---")
st.subheader("📋 Data Master Guru & Beban Mengampu Kelas")
st.info("💡 **Tips Penggunaan:** Anda bisa mengklik 2x pada kolom mana saja untuk merubah isinya. Anda juga bisa menambah guru baru di baris paling bawah.")

# Menampilkan tabel editor pintar
edited_data = st.data_editor(
    st.session_state.database_kurikulum,
    use_container_width=True,
    num_rows="dynamic",
    hide_index=True,
    column_config={
        "No": st.column_config.NumberColumn("No", width="small"),
        "Nama Guru": st.column_config.TextColumn("Nama Lengkap Guru & Gelar", required=True),
        "Mata Pelajaran": st.column_config.TextColumn("Mata Pelajaran", required=True),
        "Kelas Ampuan": st.column_config.TextColumn("Kelas Yang Diampu (Pisahkan dengan Koma)", help="Contoh: 7A, 7B, 7C"),
        "JP per Kelas": st.column_config.NumberColumn("JP per Kelas", min_value=1, max_value=6, default=4),
        "Total JP": st.column_config.NumberColumn("Total JP (Dihitung Otomatis)", disabled=True)
    }
)

# 5. TOMBOL PROSES & SINKRONISASI JURUSAN
if st.button("🔄 Hitung Ulang & Simpan Perubahan Data Master", type="primary"):
    # Fungsi menghitung ulang secara dinamis total JP berdasarkan string kelas yang diinput user
    def kalkulasi_ulang_jp(row):
        try:
            string_kelas = str(row["Kelas Ampuan"])
            if not string_kelas or string_kelas == "None" or string_kelas.strip() == "":
                return 0
            list_kelas = [k.strip() for k in string_kelas.split(",") if k.strip() != ""]
            return len(list_kelas) * int(row["JP per Kelas"])
        except:
            return 0

    edited_data["Total JP"] = edited_data.apply(kalkulasi_ulang_jp, axis=1)
    st.session_state.database_kurikulum = edited_data
    st.success("🎉 Data Master berhasil diperbarui dan total JP telah dihitung ulang secara otomatis!")
    st.rerun()

# 6. FITUR FILTER / PENCARIAN CEPAT
st.markdown("---")
st.subheader("🔍 Pencarian Cepat Data Guru")
search_query = st.text_input("Ketik Nama Guru atau Mata Pelajaran untuk menyaring data:")

if search_query:
    filtered_df = df_aktif[
        df_aktif["Nama Guru"].str.contains(search_query, case=False, na=False) |
        df_aktif["Mata Pelajaran"].str.contains(search_query, case=False, na=False)
    ]
    st.dataframe(filtered_df, use_container_width=True, hide_index=True)
