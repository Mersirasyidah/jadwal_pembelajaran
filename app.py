import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io

# 1. Konfigurasi Halaman & Tema Web
st.set_page_config(page_title="Smart Scheduler SMPN 2 Banguntapan", layout="wide")

# Konfigurasi Dimensi Waktu Sekolah
list_hari = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat']
tingkatan = ['7', '8', '9']
abjad = ['A', 'B', 'C', 'D', 'E']
semua_kelas = [f"{t}{a}" for t in tingkatan for a in abjad]

# Definisi Kategorisasi Kognitif Otak Siswa
MAPEL_SULIT = ["MATEMATIKA", "MAT", "IPA", "B. INGGRIS", "ING"]
MAPEL_PAGI = ["MATEMATIKA", "MAT", "IPA", "PJOK"]

# 2. Inisialisasi Data Master Guru Berdasarkan Lampiran Gambar
if 'data_beban_guru' not in st.session_state:
    st.session_state.data_beban_guru = pd.DataFrame([
        {"No": 1, "Kode Guru": "Purwanto", "Nama Guru": "Purwanto, S.Pd.", "Mata Pelajaran": "IPA", "Kode Mapel": "IPA", "JP/Minggu": 5, "Mengajar Kelas": ["7A", "7B", "7C", "7D", "7E"], "Hari Libur/MGMP": ["Kamis"]},
        {"No": 2, "Kode Guru": "Nurkhasanah", "Nama Guru": "Nurkhasanah, S.Ag.", "Mata Pelajaran": "P.A. ISLAM", "Kode Mapel": "AGM", "JP/Minggu": 3, "Mengajar Kelas": ["9A", "9B", "9C", "9D", "9E"], "Hari Libur/MGMP": []},
        {"No": 3, "Kode Guru": "Heni P", "Nama Guru": "Heni Purwaningsih, S.Pd.", "Mata Pelajaran": "B. INGGRIS", "Kode Mapel": "ING", "JP/Minggu": 4, "Mengajar Kelas": ["9A", "9B", "9C", "9D"], "Hari Libur/MGMP": []},
        {"No": 4, "Kode Guru": "Sri Purwanti", "Nama Guru": "Sri Purwanti, S.Pd.", "Mata Pelajaran": "MATEMATIKA", "Kode Mapel": "MAT", "JP/Minggu": 5, "Mengajar Kelas": ["9A", "9B", "9C", "9D", "9E"], "Hari Libur/MGMP": ["Selasa"]},
        {"No": 5, "Kode Guru": "Rizka D", "Nama Guru": "Rizka Diestriana, S.Pd.", "Mata Pelajaran": "B. INGGRIS", "Kode Mapel": "ING", "JP/Minggu": 4, "Mengajar Kelas": ["8A", "8B", "8C", "8D", "8E"], "Hari Libur/MGMP": []},
        {"No": 6, "Kode Guru": "Anik M", "Nama Guru": "Anik Mulyani, S.Ag.", "Mata Pelajaran": "P.A. ISLAM", "Kode Mapel": "AGM", "JP/Minggu": 3, "Mengajar Kelas": ["7A", "7B", "7C", "7D", "7E", "8C", "8D", "8E"], "Hari Libur/MGMP": []},
        {"No": 7, "Kode Guru": "Mersi", "Nama Guru": "Mersi, S.T.", "Mata Pelajaran": "INFORMATIKA", "Kode Mapel": "INF", "JP/Minggu": 2, "Mengajar Kelas": ["8A", "8B", "8C", "9A", "9B", "9C", "9D", "9E"], "Hari Libur/MGMP": []},
        {"No": 8, "Kode Guru": "Fitri L", "Nama Guru": "Fitri Lestari, S.S.", "Mata Pelajaran": "B. JAWA", "Kode Mapel": "JW", "JP/Minggu": 2, "Mengajar Kelas": ["7A", "7B", "7C", "7D", "7E", "8A", "8B", "9A", "9B", "9C", "9D", "9E"], "Hari Libur/MGMP": []},
        {"No": 9, "Kode Guru": "Yoma S", "Nama Guru": "Yoma Septiantika, S.Pd.", "Mata Pelajaran": "SENI BUDAYA", "Kode Mapel": "SB", "JP/Minggu": 3, "Mengajar Kelas": ["8A", "8B", "8C", "8D", "8E"], "Hari Libur/MGMP": []},
        {"No": 10, "Kode Guru": "Maftuhah", "Nama Guru": "Maftuhah Rahayu, S.Pd.", "Mata Pelajaran": "B. INDONESIA", "Kode Mapel": "IND", "JP/Minggu": 6, "Mengajar Kelas": ["9A", "9B", "9C", "9D", "9E"], "Hari Libur/MGMP": []},
        {"No": 11, "Kode Guru": "Nandar P", "Nama Guru": "Nandar Pamungkas Sari, S.Pd.", "Mata Pelajaran": "BK", "Kode Mapel": "BK", "JP/Minggu": 1, "Mengajar Kelas": ["9A", "9B", "9C", "9D", "9E"], "Hari Libur/MGMP": []},
        {"No": 12, "Kode Guru": "Darpito", "Nama Guru": "Darpito Nugroho, S.Pd.", "Mata Pelajaran": "BK", "Kode Mapel": "BK", "JP/Minggu": 1, "Mengajar Kelas": ["7A", "7B", "7C", "7D", "7E"], "Hari Libur/MGMP": []},
        {"No": 13, "Kode Guru": "Asti A", "Nama Guru": "Asti Am Rini, S.Pd.", "Mata Pelajaran": "B. INDONESIA", "Kode Mapel": "IND", "JP/Minggu": 6, "Mengajar Kelas": ["7A", "7B", "7C", "7D", "7E"], "Hari Libur/MGMP": []},
        {"No": 14, "Kode Guru": "Cholid D", "Nama Guru": "Cholid Dalyanto, S.Pd.Kor.", "Mata Pelajaran": "PJOK", "Kode Mapel": "PJOK", "JP/Minggu": 3, "Mengajar Kelas": ["7A", "7B", "7C", "8A", "8B", "8C", "8D", "8E"], "Hari Libur/MGMP": []},
        {"No": 15, "Kode Guru": "Feni D", "Nama Guru": "Feni Dwimartanti, S.Pd.", "Mata Pelajaran": "PEND. PANCASILA", "Kode Mapel": "PP", "JP/Minggu": 3, "Mengajar Kelas": ["7C", "7D", "7E", "9A", "9B", "9C", "9D", "9E"], "Hari Libur/MGMP": []},
        {"No": 16, "Kode Guru": "Ali S", "Nama Guru": "Ali Sudrajat, S.Pd.", "Mata Pelajaran": "BK", "Kode Mapel": "BK", "JP/Minggu": 1, "Mengajar Kelas": ["8A", "8B", "8C", "8D", "8E"], "Hari Libur/MGMP": []},
        {"No": 17, "Kode Guru": "Rahmat M", "Nama Guru": "Rahmat Mas Said, S.Pd.", "Mata Pelajaran": "SENI BUDAYA", "Kode Mapel": "SB", "JP/Minggu": 3, "Mengajar Kelas": ["9A", "9B", "9C", "9D", "9E"], "Hari Libur/MGMP": []},
        {"No": 18, "Kode Guru": "Ami R", "Nama Guru": "Ami Royati, S.Pd.", "Mata Pelajaran": "MATEMATIKA", "Kode Mapel": "MAT", "JP/Minggu": 5, "Mengajar Kelas": ["8A", "8B", "8C", "8D", "8E"], "Hari Libur/MGMP": []},
        {"No": 19, "Kode Guru": "Umi K", "Nama Guru": "Umi Kulstum, S.Pd.", "Mata Pelajaran": "IPA", "Kode Mapel": "IPA", "JP/Minggu": 5, "Mengajar Kelas": ["9A", "9B", "9C", "9D", "9E"], "Hari Libur/MGMP": []},
        {"No": 20, "Kode Guru": "Eny W", "Nama Guru": "Eny Widiyanti, S.Pd.", "Mata Pelajaran": "MATEMATIKA", "Kode Mapel": "MAT", "JP/Minggu": 5, "Mengajar Kelas": ["7A", "7B", "7C", "7D", "7E"], "Hari Libur/MGMP": []},
        {"No": 21, "Kode Guru": "Budi P", "Nama Guru": "Budi Prasetya, S.Pd.", "Mata Pelajaran": "IPS", "Kode Mapel": "IPS", "JP/Minggu": 4, "Mengajar Kelas": ["9A", "9B", "9C", "9D", "9E", "8C", "8D", "8E"], "Hari Libur/MGMP": []},
        {"No": 22, "Kode Guru": "Iswarasanti", "Nama Guru": "Iswarasanti, S.Ag.", "Mata Pelajaran": "P.A. Hindu", "Kode Mapel": "AGM", "JP/Minggu": 3, "Mengajar Kelas": ["7A", "8C", "9A"], "Hari Libur/MGMP": []},
        {"No": 23, "Kode Guru": "Sabti H", "Nama Guru": "Sabti Herma Nugraheni, S.Pd.", "Mata Pelajaran": "P.A. Katolik", "Kode Mapel": "AGM", "JP/Minggu": 3, "Mengajar Kelas": ["7A", "8A", "9A"], "Hari Libur/MGMP": []},
        {"No": 24, "Kode Guru": "Intan W", "Nama Guru": "Intan Widiastuti, S.Th.", "Mata Pelajaran": "P.A. Kristen", "Kode Mapel": "AGM", "JP/Minggu": 3, "Mengajar Kelas": ["7A", "8A", "9A"], "Hari Libur/MGMP": []},
        {"No": 25, "Kode Guru": "Christina D", "Nama Guru": "Christina Dwi Ayu Wijaya, S.Pd.", "Mata Pelajaran": "PEND. PANCASILA", "Kode Mapel": "PP", "JP/Minggu": 3, "Mengajar Kelas": ["7A", "7B", "8A", "8B", "8C", "8D", "8E"], "Hari Libur/MGMP": []},
        {"No": 26, "Kode Guru": "Thiara M", "Nama Guru": "Thiara Maharani, S.Pd.", "Mata Pelajaran": "PRAKARYA / B.JAWA", "Kode Mapel": "PRAK/JW", "JP/Minggu": 2, "Mengajar Kelas": ["7A", "7B", "7C", "7D", "7E", "8C", "8D", "8E"], "Hari Libur/MGMP": []},
        {"No": 27, "Kode Guru": "Luthfan Q", "Nama Guru": "Luthfan Qaedi Wicaksono, S.Pd.", "Mata Pelajaran": "PJOK", "Kode Mapel": "PJOK", "JP/Minggu": 3, "Mengajar Kelas": ["7D", "7E", "9A", "9B", "9C", "9D", "9E"], "Hari Libur/MGMP": []},
        {"No": 28, "Kode Guru": "Anggiyani F", "Nama Guru": "Anggiyani Fabilah Parwati, S.Pd.", "Mata Pelajaran": "INFORMATIKA", "Kode Mapel": "INF", "JP/Minggu": 2, "Mengajar Kelas": ["7A", "7B", "7C", "7D", "7E", "8D", "8E"], "Hari Libur/MGMP": []},
        {"No": 29, "Kode Guru": "Wesda A", "Nama Guru": "Wesda Ayu Rahmadani", "Mata Pelajaran": "B. INGGRIS", "Kode Mapel": "ING", "JP/Minggu": 4, "Mengajar Kelas": ["7A", "7B", "7C", "7D", "7E"], "Hari Libur/MGMP": []},
        {"No": 30, "Kode Guru": "Anisa S", "Nama Guru": "Anisa Safera Proborini, S.Pd.", "Mata Pelajaran": "IPA", "Kode Mapel": "IPA", "JP/Minggu": 5, "Mengajar Kelas": ["8A", "8B", "8C", "8D", "8E"], "Hari Libur/MGMP": []},
        {"No": 31, "Kode Guru": "Krisma D", "Nama Guru": "Krisma Dewi, S.Pd.", "Mata Pelajaran": "B. INDONESIA", "Kode Mapel": "IND", "JP/Minggu": 6, "Mengajar Kelas": ["8A", "8B", "8C", "8D", "8E"], "Hari Libur/MGMP": []},
        {"No": 32, "Kode Guru": "Anggi S", "Nama Guru": "Anggi Supriyadi, S.Hum", "Mata Pelajaran": "P.A. Islam", "Kode Mapel": "AGM", "JP/Minggu": 3, "Mengajar Kelas": ["8A", "8B"], "Hari Libur/MGMP": []},
        {"No": 33, "Kode Guru": "Rizal R", "Nama Guru": "Rizal Rahmanto, S.Pd.", "Mata Pelajaran": "IPS", "Kode Mapel": "IPS", "JP/Minggu": 4, "Mengajar Kelas": ["7A", "7B", "7C", "7D", "7E", "8A", "8B"], "Hari Libur/MGMP": []}
    ])

if 'jadwal_terplot' not in st.session_state:
    st.session_state.jadwal_terplot = []

# ==================== FUNGSI GENERATE EXCEL ====================
def export_excel_laporan(plotting_data, kelas_list, hari_list):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Jadwal Menyeluruh SMPN 2"
    ws.views.sheetView[0].showGridLines = True
    
    fill_header = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    fill_subhead = PatternFill(start_color="4A90E2", end_color="4A90E2", fill_type="solid")
    fill_day = PatternFill(start_color="F2F4F7", end_color="F2F4F7", fill_type="solid")
    
    thin_border = Border(
        left=Side(style='thin', color='D1D5DB'), right=Side(style='thin', color='D1D5DB'),
        top=Side(style='thin', color='D1D5DB'), bottom=Side(style='thin', color='D1D5DB')
    )
    
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(kelas_list) + 2)
    ws["A1"] = "JADWAL PEMBELAJARAN PARALEL AGAMA - SMP NEGERI 2 BANGUNTAPAN"
    ws["A1"].font = Font(name="Arial", size=13, bold=True, color="1B365D")
    ws["A1"].alignment = Alignment(horizontal="center")
    
    ws.cell(row=3, column=1, value="HARI").fill = fill_header
    ws.cell(row=3, column=2, value="JAM").fill = fill_header
    ws.cell(row=3, column=1).font = Font(color="FFFFFF", bold=True)
    ws.cell(row=3, column=2).font = Font(color="FFFFFF", bold=True)
    
    for c_idx, kelas in enumerate(kelas_list, 3):
        cell = ws.cell(row=3, column=c_idx, value=f"KLS {kelas}")
        cell.font = Font(color="FFFFFF", bold=True, size=10)
        cell.fill = fill_subhead
        cell.alignment = Alignment(horizontal="center")
        cell.border = thin_border
        
    current_row = 4
    for hari in hari_list:
        max_jam = 6 if hari == 'Jumat' else 9
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row + max_jam - 1, end_column=1)
        day_cell = ws.cell(row=current_row, column=1, value=hari.upper())
        day_cell.fill = fill_day
        day_cell.font = Font(bold=True)
        day_cell.alignment = Alignment(horizontal="center", vertical="center")
        
        for jam in range(1, max_jam + 1):
            r_num = current_row + jam - 1
            ws.cell(row=r_num, column=2, value=jam).border = thin_border
            ws.cell(row=r_num, column=1).border = thin_border
            
            for c_idx, kelas in enumerate(kelas_list, 3):
                cell_kbm = ws.cell(row=r_num, column=c_idx)
                cell_kbm.border = thin_border
                cell_kbm.alignment = Alignment(horizontal="center", wrap_text=True)
                
                if hari == 'Senin' and jam == 1:
                    cell_kbm.value = "UPACARA"
                    continue
                    
                slot = next((d for d in plotting_data if d['hari'] == hari and d['jam'] == jam and d['kelas'] == kelas), None)
                if slot:
                    cell_kbm.value = f"{slot['kode_mapel']}\n{slot['kode_guru']}"
                    
        current_row += max_jam
        
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_LEGAL
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ==================== INTERFACE UTAMA STREAMLIT ====================
st.title("🏫 AI Smart Timetable - SMPN 2 Banguntapan")
st.subheader("Modul Penjadwalan Agama Paralel Serentak (Islam, Kristen, Katolik, Hindu)")

tab1, tab2 = st.tabs(["📋 1. Beban Tugas Guru", "🗓️ 2. Hasil Peta Jadwal"])

with tab1:
    st.markdown("##### Atur / Edit Riwayat Beban Mengajar Guru")
    edited_df = st.data_editor(
        st.session_state.data_beban_guru,
        column_config={
            "No": {"disabled": True, "width": "small"},
            "Kode Guru": {"label": "Inisial Guru"},
            "Mata Pelajaran": {"disabled": True},
            "Kode Mapel": {"label": "Kelompok"},
            "JP/Minggu": {"label": "Beban JP", "min_value": 1},
            "Mengajar Kelas": st.column_config.MultiselectColumn(label="Rombel Diampu", options=semua_kelas),
            "Hari Libur/MGMP": st.column_config.MultiselectColumn(label="Libur Rutin", options=list_hari)
        },
        use_container_width=True,
        key="editor_paralel_v2"
    )
    
    if st.button("⚡ Terapkan & Generate Jadwal Serentak", type="primary"):
        st.session_state.data_beban_guru = edited_df
        st.session_state.jadwal_terplot = []
        
        pointer_kbm = {h: {k: (2 if h == 'Senin' else 1) for k in semua_kelas} for h in list_hari}
        
        # --- LANGKAH 1: PLOT MATRIKS AGAMA SECARA PARALEL (SERENTAK) ---
        df_agama = edited_df[edited_df['Mata Pelajaran'].str.contains("P.A.", case=False, na=False)]
        
        # Urutkan rombel dari 7 ke 9 untuk memastikan pengelompokan yang stabil
        for klswide in semua_kelas:
            hari_terpilih = None
            jam_mulai_terpilih = None
            
            # Cari slot jam kosong yang panjangnya 3 JP berturut-turut untuk rombel ini
            for hari in list_hari:
                max_jam = 6 if hari == 'Jumat' else 9
                current_p = pointer_kbm[hari][klswide]
                
                if current_p + 2 <= max_jam:
                    # Pastikan guru-guru agama yang mengajar rombel ini tidak bentrok pada slot waktu tersebut
                    bentrok_guru = False
                    guru_terlibat = df_agama[df_agama['Mengajar Kelas'].apply(lambda x: klswide in x if isinstance(x, list) else False)]
                    
                    for _, g_row in guru_terlibat.iterrows():
                        g_inisial = g_row['Kode Guru']
                        for step in range(3):
                            j_cek = current_p + step
                            if next((j for j in st.session_state.jadwal_terplot if j['hari'] == hari and j['jam'] == j_cek and j['kode_guru'] == g_inisial), None):
                                bentrok_guru = True
                                break
                        if bentrok_guru:
                            break
                            
                    if not bentrok_guru:
                        hari_terpilih = hari
                        jam_mulai_terpilih = current_p
                        break
            
            # Masukkan seluruh guru agama yang berhak mengajar rombel tersebut ke dalam slot paralel ini
            if hari_terpilih and jam_mulai_terpilih:
                guru_terlibat = df_agama[df_agama['Mengajar Kelas'].apply(lambda x: klswide in x if isinstance(x, list) else False)]
                for _, g_row in guru_terlibat.iterrows():
                    g_inisial = g_row['Kode Guru']
                    m_label = "AGM-" + g_row['Mata Pelajaran'].replace("P.A. ", "")
                    for step in range(3):
                        st.session_state.jadwal_terplot.append({
                            "kode_guru": g_inisial, "kode_mapel": m_label,
                            "kelas": klswide, "hari": hari_terpilih, "jam": jam_mulai_terpilih + step
                        })
                pointer_kbm[hari_terpilih][klswide] = jam_mulai_terpilih + 3

        # --- LANGKAH 2: PLOT MATA PELAJARAN UMUM LAINNYA ---
        df_umum = edited_df[~edited_df['Mata Pelajaran'].str.contains("P.A.", case=False, na=False)]
        
        def prioritas_umum(r):
            m_code = str(r["Kode Mapel"]).upper()
            if any(p in m_code for p in MAPEL_PAGI): return 1
            if any(s in m_code for s in MAPEL_SULIT): return 2
            return 3
            
        df_umum['prioritas'] = df_umum.apply(prioritas_umum, axis=1)
        df_umum_sorted = df_umum.sort_values(by='prioritas')
        
        for _, row in df_umum_sorted.iterrows():
            k_guru = row["Kode Guru"]
            k_mapel = str(row["Kode Mapel"]).upper()
            kelas_diampu = row["Mengajar Kelas"]
            jp_kontrak = int(row["JP/Minggu"]) if pd.notna(row["JP/Minggu"]) else 3
            hari_libur = row["Hari Libur/MGMP"] if isinstance(row["Hari Libur/MGMP"], list) else []
            
            if not isinstance(kelas_diampu, list): continue
            
            for kelas in kelas_diampu:
                jp_tersisa = jp_kontrak
                for hari in list_hari:
                    if jp_tersisa <= 0: break
                    if hari in hari_libur: continue
                    
                    max_jam = 6 if hari == 'Jumat' else 9
                    alokasi = min(jp_tersisa, 2 if k_mapel in MAPEL_PAGI else 3)
                    
                    # Cari slot kosong dari baris awal jam berjalan
                    jam_mulai = pointer_kbm[hari][kelas]
                    while jam_mulai + alokasi - 1 <= max_jam:
                        # Cek bentrok guru atau kelas
                        bentrok = False
                        for step in range(alokasi):
                            target_j = jam_mulai + step
                            bg = next((j for j in st.session_state.jadwal_terplot if j['hari'] == hari and j['jam'] == target_j and j['kode_guru'] == k_guru), None)
                            bk = next((j for j in st.session_state.jadwal_terplot if j['hari'] == hari and j['jam'] == target_j and j['kelas'] == kelas), None)
                            if bg or bk:
                                bentrok = True
                                break
                        
                        if not bentrok:
                            for step in range(alokasi):
                                st.session_state.jadwal_terplot.append({
                                    "kode_guru": k_guru, "kode_mapel": k_mapel,
                                    "kelas": kelas, "hari": hari, "jam": jam_mulai + step
                                })
                            if pointer_kbm[hari][kelas] == jam_mulai:
                                pointer_kbm[hari][kelas] += alokasi
                            jp_tersisa -= alokasi
                            break
                        else:
                            jam_mulai += 1 # Geser 1 jam ke depan untuk mencari celah kosong
                            
        st.success("✔️ Berhasil Mengoptimasi! Jadwal Seluruh Agama (Islam, Kristen, Katolik, Hindu) Kini Berada Pada Slot Hari & Jam yang Sama Secara Paralel.")
        st.rerun()

with tab2:
    if not st.session_state.jadwal_terplot:
        st.warning("⚠️ Jadwal belum di-generate. Silakan klik tombol di Tab 1.")
    else:
        col_v1, col_v2 = st.columns([2, 1])
        with col_v1:
            pilihan_kelas = st.selectbox("Pilih Ruang Kelas / Rombel:", semua_kelas)
        with col_v2:
            st.markdown("<br>", unsafe_allow_html=True)
            data_excel = export_excel_laporan(st.session_state.jadwal_terplot, semua_kelas, list_hari)
            st.download_button(
                label="📥 Download Master Jadwal Paralel (Excel)",
                data=data_excel,
                file_name="Jadwal_Paralel_Agama_SMPN2.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
        tabel_tampil = pd.DataFrame(index=list_hari, columns=[f"Jam {i}" for i in range(1, 10)]).fillna("Kosong")
        for h in list_hari:
            bts = 6 if h == 'Jumat' else 9
            for j in range(1, 10):
                if j > bts:
                    tabel_tampil.at[h, f"Jam {j}"] = "-"
                    continue
                if h == 'Senin' and j == 1:
                    tabel_tampil.at[h, f"Jam {j}"] = "🎗️ UPACARA"
                    continue
                    
                match = next((d for d in st.session_state.jadwal_terplot if d['hari'] == h and d['jam'] == j and d['kelas'] == pilihan_kelas), None)
                if match:
                    tabel_tampil.at[h, f"Jam {j}"] = f"{match['kode_mapel']} [{match['kode_guru']}]"
                    
        st.markdown(f"##### 📅 Preview Jadwal **Kelas {pilihan_kelas}**")
        st.dataframe(tabel_tampil, use_container_width=True)
