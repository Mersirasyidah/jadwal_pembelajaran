import streamlit as st
import pandas as pd

# 1. Konfigurasi Halaman & State
st.set_page_config(page_title="Sistem Database Jadwal SMP", layout="wide")

# Database awal (Mapping Guru, Mapel, dan Kelas yang diampu)
# Kamu bisa menambah/mengubah daftar ini langsung dari aplikasi
if 'database_guru' not in st.session_state:
    st.session_state.database_guru = [
        {"kode": "4", "nama": "Ani Pujiastuti, M.Hum", "mapel": "Bahasa Inggris", "kelas_diampu": ["7A", "7B", "7C"]},
        {"kode": "5", "nama": "Agus Fuadi, M.Pd", "mapel": "IPA", "kelas_diampu": ["9E", "9F"]},
        {"kode": "42", "nama": "Sunarwi, S.Pd", "mapel": "PKN", "kelas_diampu": ["7E", "7F"]},
        {"kode": "2", "nama": "Hartini, M.Pd", "mapel": "Matematika", "kelas_diampu": ["8A", "8B"]},
        {"kode": "3", "nama": "Bartina, S.Pd", "mapel": "Bahasa Indonesia", "kelas_diampu": ["8D", "8E"]}
    ]

if 'plotting_jadwal' not in st.session_state:
    st.session_state.plotting_jadwal = []

# Data Master Sekolah
tingkatan = ['7', '8', '9']
abjad = ['A', 'B', 'C', 'D', 'E']
list_kelas = [f"{t}{a}" for t in tingkatan for a in abjad]
list_hari = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat']

st.title("🗃️ Sistem Database Pengajar & Generator Jadwal")
st.caption("Aplikasi otomatis mencocokkan guru dengan kelas yang diampunya untuk menghindari salah input.")

# 2. TABS UNTUK MEMISAHKAN MENU
tab1, tab2, tab3 = st.tabs(["👥 1. Database Guru & Ampuan", "📅 2. Plotting Jadwal Harian", "📊 3. Hasil Jadwal"])

# ==================== TAB 1: DATABASE GURU ====================
with tab1:
    st.subheader("Manajemen Guru dan Kelas yang Diampu")
    
    # Form Tambah Guru Baru ke Database
    with st.expander("➕ Tambah Guru & Kelas Ampuan Baru"):
        c1, c2, c3 = st.columns(3)
        with c1:
            new_nama = st.text_input("Nama Guru Baru:")
        with c2:
            new_mapel = st.text_input("Mata Pelajaran:")
        with c3:
            new_kelas = st.multiselect("Kelas yang Diampu:", list_kelas)
            
        if st.button("Simpan ke Database"):
            if new_nama and new_mapel and new_kelas:
                st.session_state.database_guru.append({
                    "kode": str(len(st.session_state.database_guru) + 1),
                    "nama": new_nama,
                    "mapel": new_mapel,
                    "kelas_diampu": new_kelas
                })
                st.success(f"Berhasil menambahkan {new_nama} ke database!")
                st.rerun()
            else:
                st.error("Semua data form wajib diisi/dipilih!")

    # Tampilkan Data Tabel Guru Saat Ini
    df_guru = pd.DataFrame(st.session_state.database_guru)
    df_guru['kelas_diampu'] = df_guru['kelas_diampu'].apply(lambda x: ", ".join(x))
    st.dataframe(df_guru, use_container_width=True)

# ==================== TAB 2: PLOTTING JADWAL ====================
with tab2:
    st.subheader("⏰ Input Jadwal Berdasarkan Database Guru")
    
    col_input, col_info = st.columns([1, 1])
    
    with col_input:
        # Pilihan Guru mengambil dari Database di Tab 1
        list_nama_guru = [g['nama'] for g in st.session_state.database_guru]
        pilihan_guru = st.selectbox("Pilih Guru pengajar:", list_nama_guru)
        
        # Ambil data mapel dan kelas berdasarkan guru yang dipilih secara otomatis
        info_guru_terpilih = next(g for g in st.session_state.database_guru if g['nama'] == pilihan_guru)
        mapel_otomatis = info_guru_terpilih['mapel']
        kelas_tersedia = info_guru_terpilih['kelas_diampu']
        
        # Tampilkan info otomatis agar user tahu
        st.info(f"📖 Mata Pelajaran: **{mapel_otomatis}**")
        
        # Pilihan kelas dibatasi HANYA kelas yang diampu guru tersebut
        pilihan_kelas = st.selectbox("Pilih Kelas (Hanya yang diampu guru ini):", kelas_tersedia)
        
        pilihan_hari = st.selectbox("Pilih Hari Pelajaran:", list_hari)
        
        max_jam = 6 if pilihan_hari == 'Jumat' else 9
        jam_mulai = st.number_input("Mulai dari Jam Ke-:", min_value=1, max_value=max_jam, value=1, key="jam_mulai_tab2")
        durasi_jam = st.number_input("Durasi (Jumlah Jam Pelajaran):", min_value=1, max_value=5, value=2, key="durasi_tab2")
        
        if st.button("🚀 Masukkan ke Jadwal Harian", type="primary"):
            # Validasi Aturan Waktu Sekolah
            if pilihan_hari == 'Jumat' and (jam_mulai + durasi_jam - 1) > 6:
                st.error("❌ Gagal! Hari Jumat maksimal hanya sampai jam ke-6.")
            elif pilihan_hari == 'Senin' and jam_mulai == 1:
                st.error("❌ Gagal! Hari Senin Jam ke-1 otomatis digunakan untuk Upacara.")
            else:
                bentrok = False
                slot_baru = []
                
                for i in range(durasi_jam):
                    j_sekarang = jam_mulai + i
                    
                    if j_sekarang > max_jam:
                        st.error(f"❌ Durasi melebihi batas jam sekolah hari {pilihan_hari}.")
                        bentrok = True
                        break
                        
                    # Validasi Anti-Bentrok Guru
                    b_guru = next((j for j in st.session_state.plotting_jadwal if j['hari'] == pilihan_hari and j['jam'] == j_sekarang and j['guru'] == pilihan_guru), None)
                    if b_guru:
                        st.error(f"❌ Guru {pilihan_guru} sudah mengajar di kelas {b_guru['kelas']} pada jam ke-{j_sekarang}!")
                        bentrok = True
                        break
                        
                    # Validasi Anti-Bentrok Kelas
                    b_kelas = next((j for j in st.session_state.plotting_jadwal if j['hari'] == pilihan_hari and j['jam'] == j_sekarang and j['kelas'] == pilihan_kelas), None)
                    if b_kelas:
                        st.error(f"❌ Kelas {pilihan_kelas} sudah terisi pelajaran lain pada jam ke-{j_sekarang}!")
                        bentrok = True
                        break
                        
                    slot_baru.append({
                        "guru": pilihan_guru, "mapel": mapel_otomatis, "kelas": pilihan_kelas, "hari": pilihan_hari, "jam": j_sekarang
                    })
                    
                if not bentrok:
                    st.session_state.plotting_jadwal.extend(slot_baru)
                    st.success(f"✔️ Berhasil memplot jadwal untuk {pilihan_guru} di kelas {pilihan_kelas}!")

    with col_info:
        st.markdown("""
        ### ⏱️ Aturan Jam & Istirahat Sekolah:
        * **Senin - Kamis (40 Menit/Jam)**
          * Jam 1: 07.20 - 08.00 *(Senin = Upacara)*
          * Istirahat 1: 09.20 - 09.40
          * Istirahat 2: 12.00 - 12.40
        * **Jumat (35 Menit/Jam)**
          * Maksimal hanya sampai Jam ke-6
          * Istirahat hanya 1 kali (09.40 - 10.00)
        """)

# ==================== TAB 3: HASIL JADWAL ====================
with tab3:
    st.subheader("📊 Cetak & Tinjau Jadwal Per Kelas")
    kelas_view = st.selectbox("Pilih Kelas untuk Dilihat:", list_kelas, key="view_kelas")
    
    matriks = pd.DataFrame(index=list_hari, columns=[f"Jam {i}" for i in range(1, 10)]).fillna("Kosong")
    
    for h in list_hari:
        max_j = 6 if h == 'Jumat' else 9
        for j in range(1, 10):
            if j > max_j:
                matriks.at[h, f"Jam {j}"] = "-"
                continue
            if h == 'Senin' and j == 1:
                matriks.at[h, f"Jam {j}"] = "🎗️ UPACARA"
                continue
                
            slot = next((d for d in st.session_state.plotting_jadwal if d['hari'] == h and d['jam'] == j and d['kelas'] == kelas_view), None)
            if slot:
                matriks.at[h, f"Jam {j}"] = f"{slot['mapel']}\n({slot['guru']})"
                
    st.dataframe(matriks, use_container_width=True)
    
    if st.button("🗑️ Reset Jadwal Harian"):
        st.session_state.plotting_jadwal = []
        st.rerun()
