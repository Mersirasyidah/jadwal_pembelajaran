import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(page_title="Smart Scheduler SMPN 2 Banguntapan v6", layout="wide")

list_hari = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat']
tingkatan = ['7', '8', '9']
abjad = ['A', 'B', 'C', 'D', 'E']
semua_kelas = [f"{t}{a}" for t in tingkatan for a in abjad]

def ambil_tiga_digit(nama_kode):
    if pd.isna(nama_kode) or not nama_kode:
        return ""
    nama_str = str(nama_kode).strip()
    return nama_str[:3].upper()

def get_default_data():
    raw_data = [
        {"No": 1, "Kode Guru": "Purwanto", "Nama Guru": "Purwanto, S.Pd.", "Mata Pelajaran": "IPA", "Kode Mapel": "IPA", "JP/Minggu": 5, "Mengajar Kelas": ["7A", "7B", "7C", "7D", "7E"], "Hari Libur/MGMP": ["Jumat"]},
        {"No": 2, "Kode Guru": "Nurkhasanah", "Nama Guru": "Nurkhasanah, S.Ag.", "Mata Pelajaran": "P.A. ISLAM", "Kode Mapel": "AGM", "JP/Minggu": 3, "Mengajar Kelas": ["9A", "9B", "9C", "9D", "9E"], "Hari Libur/MGMP": ["Rabu"]},
        {"No": 3, "Kode Guru": "Heni P", "Nama Guru": "Heni Purwaningsih, S.Pd.", "Mata Pelajaran": "B. INGGRIS", "Kode Mapel": "ING", "JP/Minggu": 5, "Mengajar Kelas": ["9A", "9B", "9C", "9D", "9E"], "Hari Libur/MGMP": ["Selasa"]},
        {"No": 4, "Kode Guru": "Sri Purwanti", "Nama Guru": "Sri Purwanti, S.Pd.", "Mata Pelajaran": "MATEMATIKA", "Kode Mapel": "MAT", "JP/Minggu": 5, "Mengajar Kelas": ["9A", "9B", "9C", "9D", "9E"], "Hari Libur/MGMP": ["Kamis"]},
        {"No": 5, "Kode Guru": "Rizka D", "Nama Guru": "Rizka Diestriana, S.Pd.", "Mata Pelajaran": "B. INGGRIS", "Kode Mapel": "ING", "JP/Minggu": 5, "Mengajar Kelas": ["8A", "8B", "8C", "8D", "8E"], "Hari Libur/MGMP": ["Selasa"]},
        {"No": 6, "Kode Guru": "Anik M", "Nama Guru": "Anik Mulyani, S.Ag.", "Mata Pelajaran": "P.A. ISLAM", "Kode Mapel": "AGM", "JP/Minggu": 3, "Mengajar Kelas": ["7A", "7B", "7C", "7D", "7E", "8C", "8D", "8E"], "Hari Libur/MGMP": ["Rabu"]},
        {"No": 7, "Kode Guru": "Mersi", "Nama Guru": "Mersi, S.T.", "Mata Pelajaran": "INFORMATIKA", "Kode Mapel": "INF", "JP/Minggu": 3, "Mengajar Kelas": ["8A", "8B", "8C", "9A", "9B", "9C", "9D", "9E"], "Hari Libur/MGMP": ["Kamis"]},
        {"No": 8, "Kode Guru": "Fitri L", "Nama Guru": "Fitri Lestari, S.S.", "Mata Pelajaran": "B. JAWA", "Kode Mapel": "JW", "JP/Minggu": 3, "Mengajar Kelas": ["7A", "7B", "7C", "7D", "7E", "8A", "8B", "9A", "9B", "9C", "9D", "9E"], "Hari Libur/MGMP": ["Selasa"]},
        {"No": 9, "Kode Guru": "Yoma S", "Nama Guru": "Yoma Septiantika, S.Pd.", "Mata Pelajaran": "SENI BUDAYA", "Kode Mapel": "SB", "JP/Minggu": 3, "Mengajar Kelas": ["8A", "8B", "8C", "8D", "8E"], "Hari Libur/MGMP": ["Kamis"]},
        {"No": 10, "Kode Guru": "Maftuhah", "Nama Guru": "Maftuhah Rahayu, S.Pd.", "Mata Pelajaran": "B. INDONESIA", "Kode Mapel": "IND", "JP/Minggu": 5, "Mengajar Kelas": ["9A", "9B", "9C", "9D", "9E"], "Hari Libur/MGMP": ["Selasa"]},
        {"No": 13, "Kode Guru": "Asti A", "Nama Guru": "Asti Am Rini, S.Pd.", "Mata Pelajaran": "B. INDONESIA", "Kode Mapel": "IND", "JP/Minggu": 5, "Mengajar Kelas": ["7A", "7B", "7C", "7D", "7E"], "Hari Libur/MGMP": ["Selasa"]},
        {"No": 14, "Kode Guru": "Cholid D", "Nama Guru": "Cholid Dalyanto, S.Pd.Kor.", "Mata Pelajaran": "PJOK", "Kode Mapel": "PJOK", "JP/Minggu": 3, "Mengajar Kelas": ["7A", "7B", "7C", "8A", "8B", "8C", "8D", "8E"], "Hari Libur/MGMP": ["Rabu"]},
        {"No": 15, "Kode Guru": "Feni D", "Nama Guru": "Feni Dwimartanti, S.Pd.", "Mata Pelajaran": "PEND. PANCASILA", "Kode Mapel": "PP", "JP/Minggu": 4, "Mengajar Kelas": ["7C", "7D", "7E", "9A", "9B", "9C", "9D", "9E"], "Hari Libur/MGMP": ["Senin"]},
        {"No": 17, "Kode Guru": "Rahmat M", "Nama Guru": "Rahmat Mas Said, S.Pd.", "Mata Pelajaran": "SENI BUDAYA", "Kode Mapel": "SB", "JP/Minggu": 3, "Mengajar Kelas": ["9A", "9B", "9C", "9D", "9E"], "Hari Libur/MGMP": ["Kamis"]},
        {"No": 18, "Kode Guru": "Ami R", "Nama Guru": "Ami Royati, S.Pd.", "Mata Pelajaran": "MATEMATIKA", "Kode Mapel": "MAT", "JP/Minggu": 5, "Mengajar Kelas": ["8A", "8B", "8C", "8D", "8E"], "Hari Libur/MGMP": ["Kamis"]},
        {"No": 19, "Kode Guru": "Umi K", "Nama Guru": "Umi Kulstum, S.Pd.", "Mata Pelajaran": "IPA", "Kode Mapel": "IPA", "JP/Minggu": 5, "Mengajar Kelas": ["9A", "9B", "9C", "9D", "9E"], "Hari Libur/MGMP": ["Jumat"]},
        {"No": 20, "Kode Guru": "Eny W", "Nama Guru": "Eny Widiyanti, S.Pd.", "Mata Pelajaran": "MATEMATIKA", "Kode Mapel": "MAT", "JP/Minggu": 5, "Mengajar Kelas": ["7A", "7B", "7C", "7D", "7E"], "Hari Libur/MGMP": ["Kamis"]},
        {"No": 21, "Kode Guru": "Budi P", "Nama Guru": "Budi Prasetya, S.Pd.", "Mata Pelajaran": "IPS", "Kode Mapel": "IPS", "JP/Minggu": 4, "Mengajar Kelas": ["9A", "9B", "9C", "9D", "9E", "8C", "8D", "8E"], "Hari Libur/MGMP": ["Rabu"]},
        {"No": 25, "Kode Guru": "Christina I", "Nama Guru": "Christina Dwi Ayu Wijaya, S.Pd.", "Mata Pelajaran": "PEND. PANCASILA", "Kode Mapel": "PP", "JP/Minggu": 4, "Mengajar Kelas": ["7A", "7B", "8A", "8B", "8C", "8D", "8E"], "Hari Libur/MGMP": ["Senin"]},
        {"No": 26, "Kode Guru": "Thiara M", "Nama Guru": "Thiara Maharani, S.Pd.", "Mata Pelajaran": "PRAKARYA", "Kode Mapel": "PRAK", "JP/Minggu": 3, "Mengajar Kelas": ["7A", "7B", "7C", "7D", "7E"], "Hari Libur/MGMP": ["Senin"]},
        {"No": 27, "Kode Guru": "Thiara M", "Nama Guru": "Thiara Maharani, S.Pd.", "Mata Pelajaran": "B.JAWA", "Kode Mapel": "JW", "JP/Minggu": 3, "Mengajar Kelas": ["8C", "8D", "8E"], "Hari Libur/MGMP": ["Senin"]},
        {"No": 28, "Kode Guru": "Luthfan Q", "Nama Guru": "Luthfan Qaedi Wicaksono, S.Pd.", "Mata Pelajaran": "PJOK", "Kode Mapel": "PJOK", "JP/Minggu": 3, "Mengajar Kelas": ["7D", "7E", "9A", "9B", "9C", "9D", "9E"], "Hari Libur/MGMP": ["Rabu"]},
        {"No": 29, "Kode Guru": "Anggiyani", "Nama Guru": "Anggiyani Fabilah Parwati, S.Pd.", "Mata Pelajaran": "INFORMATIKA", "Kode Mapel": "INF", "JP/Minggu": 3, "Mengajar Kelas": ["7A", "7B", "7C", "7D", "7E", "8D", "8E"], "Hari Libur/MGMP": ["Kamis"]},
        {"No": 30, "Kode Guru": "Wesda A", "Nama Guru": "Wesda Ayu Rahmadani", "Mata Pelajaran": "B. INGGRIS", "Kode Mapel": "ING", "JP/Minggu": 5, "Mengajar Kelas": ["7A", "7B", "7C", "7D", "7E"], "Hari Libur/MGMP": ["Selasa"]},
        {"No": 31, "Kode Guru": "Anisa S", "Nama Guru": "Anisa Safera Proborini, S.Pd.", "Mata Pelajaran": "IPA", "Kode Mapel": "IPA", "JP/Minggu": 5, "Mengajar Kelas": ["8A", "8B", "8C", "8D", "8E"], "Hari Libur/MGMP": ["Jumat"]},
        {"No": 32, "Kode Guru": "Krisma D", "Nama Guru": "Krisma Dewi, S.Pd.", "Mata Pelajaran": "B. INDONESIA", "Kode Mapel": "IND", "JP/Minggu": 5, "Mengajar Kelas": ["8A", "8B", "8C", "8D", "8E"], "Hari Libur/MGMP": ["Selasa"]},
        {"No": 34, "Kode Guru": "Rizal R", "Nama Guru": "Rizal Rahmanto, S.Pd.", "Mata Pelajaran": "IPS", "Kode Mapel": "IPS", "JP/Minggu": 4, "Mengajar Kelas": ["7A", "7B", "7C", "7D", "7E", "8A", "8B"], "Hari Libur/MGMP": ["Rabu"]}
    ]
    return pd.DataFrame(raw_data)

if 'data_beban_guru' not in st.session_state:
    st.session_state.data_beban_guru = get_default_data()
if 'jadwal_terplot' not in st.session_state:
    st.session_state.jadwal_terplot = []

st.title("🏫 AI Timetable Scheduler - ENGINE V6")

tab1, tab2, tab3, tab4 = st.tabs([
    "📋 Data Input & Engine AI", 
    "🗓️ Pratinjau & Cetak Jadwal", 
    "🛠️ Manual Editor & Anti-Bentrok",
    "📊 Laporan Audit & Download Excel"
])

# -------------------------------------------------------------
# TAB 1: INPUT DATA & ENGINE AI RUNNER
# -------------------------------------------------------------
with tab1:
    edited_df = st.data_editor(st.session_state.data_beban_guru, use_container_width=True)
    st.session_state.data_beban_guru = edited_df

    if st.button("⚡ Jalankan Engine AI Schedule Optimizer (V6)", type="primary"):
        plot_res = []
        def cek_ketersediaan_waktu_guru(hari, jam, libur_list):
            if hari in libur_list: return jam <= 4
            return True

        # Plot Otomatis Utama (Sama seperti logika V5 yang stabil)
        df_agama = edited_df[edited_df['Kode Mapel'] == 'AGM']
        hari_agama_map = {"7": ("Rabu", 4), "8": ("Kamis", 4), "9": ("Selasa", 4)}
        for kelas in semua_kelas:
            t = kelas[0]
            h_opt, j_opt = hari_agama_map[t]
            guru_terpilih = next((g['Kode Guru'] for _, g in df_agama.iterrows() if kelas in g['Mengajar Kelas']), None)
            if guru_terpilih:
                for s in range(3):
                    plot_res.append({"kode_guru": guru_terpilih, "kode_mapel": "AGM", "kelas": kelas, "hari": h_opt, "jam": j_opt + s})

        # Urutan prioritas plot sisa mapel
        mapel_order = ['PRAK', 'PJOK', 'MAT', 'IPA']
        for mp in mapel_order:
            df_sub = edited_df[edited_df['Kode Mapel'] == mp]
            for _, row in df_sub.iterrows():
                g_name = row['Kode Guru']
                libur_guru = row['Hari Libur/MGMP'] if isinstance(row['Hari Libur/MGMP'], list) else []
                for kelas in row['Mengajar Kelas']:
                    total_jp = int(row['JP/Minggu'])
                    sesi = [3, 2] if total_jp == 5 else [3] if total_jp == 3 else [total_jp]
                    for jp_sesi in sesi:
                        placed = False
                        for hari in list_hari:
                            for j_m in range(1, 7):
                                if hari == 'Senin' and j_m == 1: continue
                                if not all(cek_ketersediaan_waktu_guru(hari, j_m+s, libur_guru) for s in range(jp_sesi)): continue
                                if any(d['hari'] == hari and d['jam'] == (j_m+s) and (d['kode_guru'] == g_name or d['kelas'] == kelas) for s in range(jp_sesi) for d in plot_res): continue
                                for s in range(jp_sesi):
                                    plot_res.append({"kode_guru": g_name, "kode_mapel": mp, "kelas": kelas, "hari": hari, "jam": j_m + s})
                                placed = True
                                break
                            if placed: break

        # Sisa mapel umum lainnya
        df_umum = edited_df[~edited_df['Kode Mapel'].isin(['AGM', 'PRAK', 'PJOK', 'MAT', 'IPA'])]
        for _, row in df_umum.iterrows():
            g_name = row['Kode Guru']
            m_code = row['Kode Mapel']
            libur_guru = row['Hari Libur/MGMP'] if isinstance(row['Hari Libur/MGMP'], list) else []
            for kelas in row['Mengajar Kelas']:
                sisa = int(row['JP/Minggu'])
                for pola in [3, 2, 1]:
                    if sisa <= 0 or pola > sisa: continue
                    for hari in list_hari:
                        max_j = 6 if hari == 'Jumat' else 9
                        for j_m in range(1, max_j - pola + 2):
                            if hari == 'Senin' and j_m == 1: continue
                            if not all(cek_ketersediaan_waktu_guru(hari, j_m+s, libur_guru) for s in range(pola)): continue
                            if any(d['hari'] == hari and d['jam'] == (j_m+s) and (d['kode_guru'] == g_name or d['kelas'] == kelas) for s in range(pola) for d in plot_res): continue
                            for s in range(pola):
                                plot_res.append({"kode_guru": g_name, "kode_mapel": m_code, "kelas": kelas, "hari": hari, "jam": j_m + s})
                            sisa -= pola
                            break

        st.session_state.jadwal_terplot = plot_res
        st.success("✔️ Optimalisasi AI Berhasil Disinkronkan!")
        st.rerun()

# -------------------------------------------------------------
# TAB 2: PRATINJAU KELAS INDIVIDU
# -------------------------------------------------------------
with tab2:
    if not st.session_state.jadwal_terplot:
        st.warning("⚠️ Silakan jalankan optimasi terlebih dahulu di Tab 1.")
    else:
        pilihan_kelas = st.selectbox("Pilih Ruang Rombel:", semua_kelas, key="sb_preview")
        tabel_tampil = pd.DataFrame(index=list_hari, columns=[f"Jam {i}" for i in range(1, 10)]).fillna("Kosong")
        
        for h in list_hari:
            limit = 6 if h == 'Jumat' else 9
            for j in range(1, 10):
                if j > limit: tabel_tampil.at[h, f"Jam {j}"] = "-"
                elif h == 'Senin' and j == 1: tabel_tampil.at[h, f"Jam {j}"] = "🎗️ UPACARA"
                else:
                    m = [d for d in st.session_state.jadwal_terplot if d['hari'] == h and d['jam'] == j and d['kelas'] == pilihan_kelas]
                    if m: tabel_tampil.at[h, f"Jam {j}"] = f"{m[0]['kode_mapel']} ({ambil_tiga_digit(m[0]['kode_guru'])})"
        st.dataframe(tabel_tampil, use_container_width=True)

# -------------------------------------------------------------
# TAB 3: MANUAL EDITOR (ANTI-BENTROK & DROP-DOWN PINTAR)
# -------------------------------------------------------------
with tab3:
    if not st.session_state.jadwal_terplot:
        st.warning("⚠️ Jalankan Engine AI di Tab 1 terlebih dahulu.")
    else:
        st.subheader("🛠️ Pengubah Jadwal Manual & Detektor Tabrakan Guru")
        
        col_k, col_h, col_j = st.columns(3)
        with col_k: kelas_edit = st.selectbox("Pilih Kelas Yang Mau Diedit:", semua_kelas, key="ed_kelas")
        with col_h: hari_edit = st.selectbox("Pilih Hari:", list_hari, key="ed_hari")
        
        limit_jam = 6 if hari_edit == 'Jumat' else 9
        list_jam_pilih = [i for i in range(1, limit_jam + 1) if not (hari_edit == 'Senin' and i == 1)]
        with col_j: jam_edit = st.selectbox("Pilih Jam ke-:", list_jam_pilih, key="ed_jam")

        # Cek kondisi slot saat ini
        df_current = pd.DataFrame(st.session_state.jadwal_terplot)
        kondisi_sekarang = df_current[(df_current['kelas'] == kelas_edit) & (df_current['hari'] == hari_edit) & (df_current['jam'] == jam_edit)]
        
        if not kondisi_sekarang.empty:
            cur_item = kondisi_sekarang.iloc[0]
            st.info(f"📍 **Status Saat Ini:** Slot ini diisi oleh **{cur_item['kode_mapel']}** dengan Guru: **{cur_item['kode_guru']}**")
        else:
            st.warning("📍 **Status Saat Ini:** Slot ini masih **KOSONG**.")

        st.markdown("---")
        st.write("### ⬇️ Menu Perubahan")

        # HITUNG KEBUTUHAN: Drop-down pintar hanya menampilkan mapel yang masih kurang JP di kelas ini
        opsi_mapel_belum_masuk = ["-- HAPUS / KOSONGKAN SLOT --"]
        for _, row in edited_df.iterrows():
            if kelas_edit in row['Mengajar Kelas']:
                target = int(row['JP/Minggu'])
                aktual = len(df_current[(df_current['kode_guru'] == row['Kode Guru']) & (df_current['kode_mapel'] == row['Kode Mapel']) & (df_current['kelas'] == kelas_edit)])
                label_sisa = f"{row['Kode Mapel']} - {row['Kode Guru']} (Sisa Kebutuhan: {target - aktual} JP)"
                opsi_mapel_belum_masuk.append(label_sisa)

        mapel_dipilih_raw = st.selectbox("Pilih Mapel & Guru (Urut Sisa Beban):", opsi_mapel_belum_masuk)

        if st.button("💾 Simpan Perubahan Manual", type="primary"):
            # Aksi 1: Menghapus / Mengosongkan slot
            if mapel_dipilih_raw == "-- HAPUS / KOSONGKAN SLOT --":
                st.session_state.jadwal_terplot = [d for d in st.session_state.jadwal_terplot if not (d['kelas'] == kelas_edit and d['hari'] == hari_edit and d['jam'] == jam_edit)]
                st.success("🗑️ Slot berhasil dikosongkan!")
                st.rerun()
            else:
                # Ambil nama mapel dan guru asli dari teks string drop-down
                part_mapel = mapel_dipilih_raw.split(" - ")[0].strip()
                part_guru = mapel_dipilih_raw.split(" - ")[1].split(" (")[0].strip()

                # ANTI BENTROK CHECKER (Validasi Tabrakan Horizontal)
                bentrok_guru = df_current[
                    (df_current['kode_guru'] == part_guru) & 
                    (df_current['hari'] == hari_edit) & 
                    (df_current['jam'] == jam_edit) & 
                    (df_current['kelas'] != kelas_edit)
                ]

                if not bentrok_guru.empty:
                    kelas_bentrok = bentrok_guru.iloc[0]['kelas']
                    st.error(f"❌ **TIDAK DAPAT MENGINPUT!** Jam ini bentrok. Guru **{part_guru}** di hari **{hari_edit}** jam ke-**{jam_edit}** sudah ditugaskan mengajar di **Kelas {kelas_bentrok}**.")
                else:
                    # Hapus isi lama di slot tersebut (jika ada) lalu gantikan dengan yang baru
                    bersih_jadwal = [d for d in st.session_state.jadwal_terplot if not (d['kelas'] == kelas_edit and d['hari'] == hari_edit and d['jam'] == jam_edit)]
                    bersih_jadwal.append({
                        "kode_guru": part_guru,
                        "kode_mapel": part_mapel,
                        "kelas": kelas_edit,
                        "hari": hari_edit,
                        "jam": jam_edit
                    })
                    st.session_state.jadwal_terplot = bersih_jadwal
                    st.success(f"🎉 Sukses! Slot {hari_edit} Jam ke-{jam_edit} Kelas {kelas_edit} berhasil diubah menjadi {part_mapel} ({part_guru}).")
                    st.rerun()

# -------------------------------------------------------------
# TAB 4: AUDIT LAPORAN & DOWNLOAD MASTER EXCEL
# -------------------------------------------------------------
with tab4:
    if not st.session_state.jadwal_terplot:
        st.warning("⚠️ Tidak ada data audit.")
    else:
        df_p = pd.DataFrame(st.session_state.jadwal_terplot)
        
        st.subheader("📊 1. Laporan JP Per Kelas")
        jp_per_kelas = df_p.groupby('kelas').size().reset_index(name='Total JP Terplot')
        st.dataframe(jp_per_kelas, use_container_width=True)

        st.subheader("🚨 2. Laporan Mapel Kurang Jam / Belum Masuk")
        laporan_gagal = []
        for _, row in edited_df.iterrows():
            g_name = row['Kode Guru']
            m_code = row['Kode Mapel']
            target_jp = int(row['JP/Minggu'])
            for kelas in row['Mengajar Kelas']:
                aktual_jp = len(df_p[(df_p['kode_guru'] == g_name) & (df_p['kode_mapel'] == m_code) & (df_p['kelas'] == kelas)])
                if aktual_jp < target_jp:
                    laporan_gagal.append({
                        "Nama Guru": row['Nama Guru'], "Mapel": m_code, "Kelas Target": kelas,
                        "Kebutuhan JP": target_jp, "Terisi (Aktual)": aktual_jp, "Sisa": target_jp - aktual_jp
                    })
        if laporan_gagal: st.dataframe(pd.DataFrame(laporan_gagal), use_container_width=True)
        else: st.success("🎉 Sempurna! Seluruh target JP kelas terisi 100%.")

        st.subheader("📥 3. Download Master Jadwal Akhir")
        output = io.BytesIO()
        wb = openpyxl.Workbook()
        if "Sheet" in wb.sheetnames: wb.remove(wb["Sheet"])
            
        font_title = Font(name="Arial", size=14, bold=True)
        font_header = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        font_data = Font(name="Arial", size=11)
        fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        fill_upacara = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        thin_border = Border(left=Side(style='thin', color='D3D3D3'), right=Side(style='thin', color='D3D3D3'), top=Side(style='thin', color='D3D3D3'), bottom=Side(style='thin', color='D3D3D3'))

        for kls in semua_kelas:
            ws = wb.create_sheet(title=f"Kelas {kls}")
            ws.views.sheetView[0].showGridLines = True
            ws.merge_cells("A1:J1")
            ws["A1"] = f"JADWAL PELAJARAN KELAS {kls} - SMPN 2 BANGUNTAPAN"
            ws["A1"].font = font_title
            ws["A1"].alignment = align_center
            ws.row_dimensions[1].height = 35
            
            ws.append([])
            ws.append(["Hari", "Jam 1", "Jam 2", "Jam 3", "Jam 4", "Jam 5", "Jam 6", "Jam 7", "Jam 8", "Jam 9"])
            ws.row_dimensions[3].height = 25
            
            for c_idx in range(1, 11):
                cell = ws.cell(row=3, column=c_idx)
                cell.font = font_header
                cell.fill = fill_header
                cell.alignment = align_center
                cell.border = thin_border
            
            r_num = 4
            for h in list_hari:
                limit = 6 if h == 'Jumat' else 9
                row_data = [h]
                for j in range(1, 10):
                    if j > limit: row_data.append("-")
                    elif h == 'Senin' and j == 1: row_data.append("🎗️ UPACARA")
                    else:
                        m = [d for d in st.session_state.jadwal_terplot if d['hari'] == h and d['jam'] == j and d['kelas'] == kls]
                        row_data.append(f"{m[0]['kode_mapel']} ({ambil_tiga_digit(m[0]['kode_guru'])})" if m else "Kosong")
                
                ws.append(row_data)
                for c_idx in range(1, 11):
                    cell = ws.cell(row=r_num, column=c_idx)
                    cell.font = font_data
                    cell.alignment = align_center
                    cell.border = thin_border
                    if cell.value == "🎗️ UPACARA": cell.fill = fill_upacara
                r_num += 1
            ws.column_dimensions['A'].width = 15
            for c in ['B','C','D','E','F','G','H','I','J']: ws.column_dimensions[c].width = 14

        wb.save(output)
        output.seek(0)
        st.download_button(
            label="📥 Download File Master Jadwal Akhir (.xlsx)",
            data=output,
            file_name="Master_Jadwal_Akhir_SMPN2_Banguntapan.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
