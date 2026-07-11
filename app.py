import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io

st.set_page_config(page_title="Smart Scheduler SMPN 2 Banguntapan v3", layout="wide")

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
    df_init = pd.DataFrame(raw_data)
    return df_init[["No", "Kode Guru", "Nama Guru", "Mata Pelajaran", "Kode Mapel", "JP/Minggu", "Mengajar Kelas", "Hari Libur/MGMP"]]

if 'data_beban_guru' not in st.session_state:
    st.session_state.data_beban_guru = get_default_data()
if 'jadwal_terplot' not in st.session_state:
    st.session_state.jadwal_terplot = []

st.title("🏫 AI Timetable Scheduler - ENGINE V3 (Optimasi Hari Libur/MGMP)")
st.subheader("Aturan Baru: Hari Libur/MGMP diizinkan untuk mengajar di jam pagi (Jam 1 - 4)")

tab1, tab2 = st.tabs(["📋 Data Input & Engine AI", "🗓️ Cetak Master Jadwal"])

with tab1:
    st.info("💡 Logika Engine V3: Jika hari tersebut terdaftar di kolom 'Hari Libur/MGMP', Guru bersangkutan diperbolehkan mengajar HANYA pada jam 1 s.d 4 (tidak boleh ditaruh di jam 5 ke atas).")
    edited_df = st.data_editor(st.session_state.data_beban_guru, use_container_width=True)
    st.session_state.data_beban_guru = edited_df

    if st.button("⚡ Jalankan Engine AI Schedule Optimizer (V3)", type="primary"):
        plot_res = []
        
        # Fungsi pembantu mengecek batas jam mengajar guru (Rule Libur = Max jam 4)
        def cek_ketersediaan_waktu_guru(hari, jam, libur_list):
            if hari in libur_list:
                return jam <= 4  # Hanya boleh jam 1, 2, 3, 4 jika hari libur/MGMP
            return True

        # 1. PLOT AGAMA PARALEL
        df_agama = edited_df[edited_df['Mata Pelajaran'].str.contains("P.A.", case=False, na=False)]
        hari_agama_map = {"7": ("Rabu", 4), "8": ("Kamis", 4), "9": ("Selasa", 4)}
        for kelas in semua_kelas:
            t = kelas[0]
            h_opt, j_opt = hari_agama_map[t]
            guru_agama_rombel = df_agama[df_agama['Mengajar Kelas'].apply(lambda x: kelas in x if isinstance(x, list) else False)]
            for _, g_row in guru_agama_rombel.iterrows():
                libur_guru = g_row['Hari Libur/MGMP'] if isinstance(g_row['Hari Libur/MGMP'], list) else []
                # Cek apakah jam 4, 5, 6 diizinkan bagi guru agama
                aman = True
                for step in range(3):
                    if not cek_ketersediaan_waktu_guru(h_opt, j_opt + step, libur_guru):
                        aman = False
                if aman:
                    for step in range(3):
                        plot_res.append({"kode_guru": g_row['Kode Guru'], "kode_mapel": "AGM", "kelas": kelas, "hari": h_opt, "jam": j_opt + step})

        # 2. PJOK JAM PERTAMA PAGI
        df_pjok = edited_df[edited_df['Kode Mapel'] == 'PJOK']
        hari_idx = 0
        for _, row in df_pjok.iterrows():
            g_pjok = row['Kode Guru']
            libur_guru = row['Hari Libur/MGMP'] if isinstance(row['Hari Libur/MGMP'], list) else []
            for kelas in row['Mengajar Kelas']:
                placed = False
                for _ in range(25):
                    h_target = list_hari[hari_idx % len(list_hari)]
                    hari_idx += 1
                    j_start = 2 if h_target == 'Senin' else 1
                    
                    # Validasi aturan libur jam pagi
                    if not all(cek_ketersediaan_waktu_guru(h_target, j_start + s, libur_guru) for s in range(3)):
                        continue
                        
                    bentrok = False
                    for s in range(3):
                        chk_j = j_start + s
                        if any(d['hari'] == h_target and d['jam'] == chk_j and (d['kode_guru'] == g_pjok or d['kelas'] == kelas) for d in plot_res):
                            bentrok = True
                            break
                    if not bentrok:
                        for s in range(3):
                            plot_res.append({"kode_guru": g_pjok, "kode_mapel": "PJOK", "kelas": kelas, "hari": h_target, "jam": j_start + s})
                        placed = True
                        break

        # 3. CRITICAL: MATEMATIKA & IPA (MAKSIMAL JAM KE-5)
        df_critical = edited_df[edited_df['Kode Mapel'].isin(['MAT', 'IPA'])]
        for _, row in df_critical.iterrows():
            g_name = row['Kode Guru']
            m_code = row['Kode Mapel']
            libur_guru = row['Hari Libur/MGMP'] if isinstance(row['Hari Libur/MGMP'], list) else []
            total_jp = int(row['JP/Minggu']) if pd.notna(row['JP/Minggu']) else 0
            sesi = [3, 2] if total_jp == 5 else [total_jp]
            for kelas in row['Mengajar Kelas']:
                hari_terpakai = []
                for jp_sesi in sesi:
                    placed_sesi = False
                    for hari in list_hari:
                        if hari in hari_terpakai: continue
                        for j_mulai in range(1, 6):
                            if hari == 'Senin' and j_mulai == 1: continue
                            if not all(cek_ketersediaan_waktu_guru(hari, j_mulai + s, libur_guru) for s in range(jp_sesi)):
                                continue
                            if any(d['hari'] == hari and d['jam'] == (j_mulai + s) and (d['kode_guru'] == g_name or d['kelas'] == kelas) for s in range(jp_sesi) for d in plot_res):
                                continue
                            for s in range(jp_sesi):
                                plot_res.append({"kode_guru": g_name, "kode_mapel": m_code, "kelas": kelas, "hari": hari, "jam": j_mulai + s})
                            hari_terpakai.append(hari)
                            placed_sesi = True
                            break
                        if placed_sesi: break

        # 4. BOTTLENECK MAPEL UMUM & PRAKARYA
        df_umum = edited_df[(~edited_df['Mata Pelajaran'].str.contains("P.A.", case=False, na=False)) & (~edited_df['Kode Mapel'].isin(['PJOK', 'MAT', 'IPA']))]
        for _, row in df_umum.iterrows():
            g_name = row['Kode Guru']
            m_code = row['Kode Mapel']
            libur_guru = row['Hari Libur/MGMP'] if isinstance(row['Hari Libur/MGMP'], list) else []
            total_jp = int(row['JP/Minggu']) if pd.notna(row['JP/Minggu']) else 0
            if total_jp <= 0 or not isinstance(row['Mengajar Kelas'], list): continue
            
            for kelas in row['Mengajar Kelas']:
                sisa_jp = total_jp
                hari_terpakai = []
                
                # Iterasi pola pembagian jam 3, 2, 1
                for pola in [3, 2, 1]:
                    if sisa_jp <= 0: break
                    if pola > sisa_jp: continue
                    
                    for hari in list_hari:
                        if hari in hari_terpakai and pola > 1: continue
                        max_jam = 6 if hari == 'Jumat' else 9
                        placed_sub = False
                        for j_mulai in range(1, max_jam - pola + 2):
                            if hari == 'Senin' and j_mulai == 1: continue
                            if not all(cek_ketersediaan_waktu_guru(hari, j_mulai + s, libur_guru) for s in range(pola)):
                                continue
                            if any(d['hari'] == hari and d['jam'] == (j_mulai + s) and (d['kode_guru'] == g_name or d['kelas'] == kelas) for s in range(pola) for d in plot_res):
                                continue
                            for s in range(pola):
                                plot_res.append({"kode_guru": g_name, "kode_mapel": m_code, "kelas": kelas, "hari": hari, "jam": j_mulai + s})
                            hari_terpakai.append(hari)
                            sisa_jp -= pola
                            placed_sub = True
                            break
                        if placed_sub: break

                # Sapu bersih jam pencar sisa
                while sisa_jp > 0:
                    placed_1 = False
                    for hari in list_hari:
                        max_jam = 6 if hari == 'Jumat' else 9
                        for j_mulai in range(1, max_jam + 1):
                            if hari == 'Senin' and j_mulai == 1: continue
                            if not cek_ketersediaan_waktu_guru(hari, j_mulai, libur_guru): continue
                            if any(d['hari'] == hari and d['jam'] == j_mulai and (d['kode_guru'] == g_name or d['kelas'] == kelas) for d in plot_res):
                                continue
                            plot_res.append({"kode_guru": g_name, "kode_mapel": m_code, "kelas": kelas, "hari": hari, "jam": j_mulai})
                            sisa_jp -= 1
                            placed_1 = True
                            break
                        if placed_1: break
                    if not placed_1: break

        st.session_state.jadwal_terplot = plot_res
        st.success("✔️ Re-Optimasi V3 Selesai! Slot Jam Pagi Hari Libur Sukses Digunakan.")
        st.rerun()

with tab2:
    if not st.session_state.jadwal_terplot:
        st.warning("⚠️ Silakan jalankan optimasi terlebih dahulu di Tab 1.")
    else:
        pilihan_kelas = st.selectbox("Pilih Ruang Rombel:", semua_kelas)
        tabel_tampil = pd.DataFrame(index=list_hari, columns=[f"Jam {i}" for i in range(1, 10)]).fillna("Kosong")
        
        for h in list_hari:
            limit = 6 if h == 'Jumat' else 9
            for j in range(1, 10):
                if j > limit:
                    tabel_tampil.at[h, f"Jam {j}"] = "-"
                    continue
                if h == 'Senin' and j == 1:
                    tabel_tampil.at[h, f"Jam {j}"] = "🎗️ UPACARA"
                    continue
                matches = [d for d in st.session_state.jadwal_terplot if d['hari'] == h and d['jam'] == j and d['kelas'] == pilihan_kelas]
                if matches:
                    g_3d = ambil_tiga_digit(matches[0]['kode_guru'])
                    tabel_tampil.at[h, f"Jam {j}"] = f"{matches[0]['kode_mapel']} ({g_3d})"
                    
        st.dataframe(tabel_tampil, use_container_width=True)
