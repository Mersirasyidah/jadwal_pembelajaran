import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io

# 1. Konfigurasi Halaman & State
st.set_page_config(page_title="Sistem Peta Jadwal SMP Smart", layout="wide")

# Data Master
list_hari = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat']
tingkatan = ['7', '8', '9']
abjad = ['A', 'B', 'C', 'D', 'E']
semua_kelas = [f"{t}{a}" for t in tingkatan for a in abjad]

# Inisialisasi Session State untuk menampung Input Data Guru
if 'data_beban_guru' not in st.session_state:
    # Data awal sebagai contoh pengisian untuk user
    st.session_state.data_beban_guru = pd.DataFrame([
        {
            "No": 1,
            "Kode Guru": "G01",
            "Nama Guru": "Ani Pujiastuti, M.Hum",
            "Mata Pelajaran": "Bahasa Inggris",
            "Kode Mapel": "ING",
            "JP per Minggu": 5,
            "Jumlah Kelas": 2,
            "Total JP": 10,
            "Mengajar Kelas": ["7A", "7B"],
            "Hari Libur/MGMP": ["Jumat"]
        },
        {
            "No": 2,
            "Kode Guru": "G02",
            "Nama Guru": "Budi Santoso, S.Pd",
            "Mata Pelajaran": "Pendidikan Jasmani Olahraga & Kesehatan",
            "Kode Mapel": "PJOK",
            "JP per Minggu": 3,
            "Jumlah Kelas": 2,
            "Total JP": 6,
            "Mengajar Kelas": ["7A", "8A"],
            "Hari Libur/MGMP": ["Kamis"]
        }
    ])

if 'jadwal_terplot' not in st.session_state:
    st.session_state.jadwal_terplot = []

# ==================== FUNGSI GENERATE EXCEL (FORMAT HORIZONTAL LEGAL) ====================
def export_excel_laporan(plotting_data, kelas_list, hari_list):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Jadwal Menyeluruh"
    ws.views.sheetView[0].showGridLines = True
    
    fill_header = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    fill_subhead = PatternFill(start_color="4A90E2", end_color="4A90E2", fill_type="solid")
    fill_day = PatternFill(start_color="F2F4F7", end_color="F2F4F7", fill_type="solid")
    
    thin_border = Border(
        left=Side(style='thin', color='D1D5DB'), right=Side(style='thin', color='D1D5DB'),
        top=Side(style='thin', color='D1D5DB'), bottom=Side(style='thin', color='D1D5DB')
    )
    
    # Header Utama
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(kelas_list) + 2)
    ws["A1"] = "LAPORAN JADWAL PEMBELAJARAN MENYELURUH"
    ws["A1"].font = Font(name="Arial", size=14, bold=True, color="1B365D")
    ws["A1"].alignment = Alignment(horizontal="center")
    
    ws.cell(row=3, column=1, value="HARI").fill = fill_header
    ws.cell(row=3, column=2, value="JAM").fill = fill_header
    ws.cell(row=3, column=1).font = Font(color="FFFFFF", bold=True)
    ws.cell(row=3, column=2).font = Font(color="FFFFFF", bold=True)
    
    for c_idx, kelas in enumerate(kelas_list, 3):
        cell = ws.cell(row=3, column=c_idx, value=f"KELAS {kelas}")
        cell.font = Font(color="FFFFFF", bold=True)
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
                    cell_kbm.value = f"{slot['kode_mapel']}\n({slot['kode_guru']})"
                    
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

# ==================== MENU UTAMA STREAMLIT ====================
st.title("📋 Smart Generator Jadwal & Peta Hambatan Guru")
st.caption("Sistem optimasi jadwal sekolah dengan pembatas Hari Libur/MGMP, Jam Pagi PJOK, dan Pembagian Maksimal 3 JP/Hari.")

tab1, tab2 = st.tabs(["📝 1. Input Beban & Aturan Guru", "📅 2. Peta Jadwal Hasil Optimasi"])

# ==================== TAB 1: FORM INPUT BEBAN DATA GURU ====================
with tab1:
    st.subheader("Form Input Parameter & Aturan Khusus")
    st.info("💡 **Fitur Baru**: Kolom **Hari Libur/MGMP** digunakan agar sistem mengosongkan jadwal guru tersebut pada hari terpilih. Mapel **PJOK** otomatis dikunci pada jam pagi.")
    
    # MENGGUNAKAN DICTIONARY CONFIG DENGAN DUA MULTISELECT (Kelas & Hari Libur)
    edited_df = st.data_editor(
        st.session_state.data_beban_guru,
        column_config={
            "No": {"label": "No", "width": "small", "min_value": 1},
            "Kode Guru": {"label": "Kode Guru"},
            "Nama Guru": {"label": "Nama Guru"},
            "Mata Pelajaran": {"label": "Mata Pelajaran"},
            "Kode Mapel": {"label": "Kode Mapel", "help": "Ketik PJOK untuk aturan khusus olahraga pagi"},
            "JP per Minggu": {"label": "JP / Minggu", "min_value": 1, "default": 3},
            "Jumlah Kelas": {"label": "Jml Kelas", "min_value": 1, "default": 1},
            "Total JP": {"label": "Total JP (Otomatis)", "disabled": True},
            "Mengajar Kelas": st.column_config.MultiselectColumn(label="Mengajar Kelas Apa Saja", options=semua_kelas),
            "Hari Libur/MGMP": st.column_config.MultiselectColumn(label="Hari Libur / MGMP", options=list_hari)
        },
        num_rows="dynamic",
        use_container_width=True,
        key="editor_beban_v2"
    )
    
    # Tombol Aksi untuk hitung & kunci data
    if st.button("⚡ Simpan Data & Jalankan Smart Optimization Scheduler", type="primary"):
        # Hitung Total JP otomatis
        edited_df["Total JP"] = edited_df["JP per Minggu"] * edited_df["Jumlah Kelas"]
        st.session_state.data_beban_guru = edited_df
        
        # PROSES INTI AUTOMATIC PLOTTING JADWAL
        st.session_state.jadwal_terplot = [] 
        
        # Penunjuk posisi jam kosong saat ini untuk tiap kelas di masing-masing hari
        pointer_kbm = {h: {k: (2 if h == 'Senin' else 1) for k in semua_kelas} for h in list_hari}
        
        # Pengelompokan baris: Dahulukan PJOK agar mendapat prioritas jam pagi utama
        edited_df['is_pjok'] = edited_df['Kode Mapel'].str.upper() == 'PJOK'
        df_sorted = edited_df.sort_values(by='is_pjok', ascending=False)
        
        for _, row in df_sorted.iterrows():
            k_guru = row["Kode Guru"]
            k_mapel = str(row["Kode Mapel"]).upper()
            kelas_diampu = row["Mengajar Kelas"]
            jp_per_minggu = int(row["JP per Minggu"]) if pd.notna(row["JP per Minggu"]) else 0
            hari_libur = row["Hari Libur/MGMP"] if isinstance(row["Hari Libur/MGMP"], list) else []
            
            if not k_guru or not k_mapel or not isinstance(kelas_diampu, list):
                continue
                
            # Proses untuk setiap rombel kelas yang diajar oleh guru ini
            for kelas in kelas_diampu:
                jp_tersisa = jp_per_minggu
                
                for hari in list_hari:
                    if jp_tersisa <= 0:
                        break
                        
                    # Aturan 1: Lewati hari jika hari tersebut adalah hari libur/MGMP guru yang bersangkutan
                    if hari in hari_libur:
                        continue
                        
                    max_jam = 6 if hari == 'Jumat' else 9
                    
                    # Aturan 3: Batas Maksimal mengajar dalam sehari untuk setiap mapel dalam kelas adalah 3 JP
                    alokasi_hari_ini = min(jp_tersisa, 3)
                    
                    # Aturan 2: Aturan khusus penempatan PJOK di jam pagi
                    if k_mapel == "PJOK":
                        # Selain senin jam ke-1 (senin jam 1 upacara, jadi senin PJOK mulai jam ke-2)
                        jam_mulai_pjok = 2 if hari == 'Senin' else 1
                        
                        # Definisikan batas akhir jam pagi (misal sebelum istirahat pertama / maksimal selesai jam ke-4)
                        if jam_mulai_pjok + alokasi_hari_ini - 1 > 4:
                            continue # Skip jika tidak muat di pagi hari terpilih
                            
                        # Validasi tabrakan guru PJOK mengajar di kelas lain pada jam yang sama
                        bentrok = False
                        for step in range(alokasi_hari_ini):
                            target_j = jam_mulai_pjok + step
                            bg = next((j for j in st.session_state.jadwal_terplot if j['hari'] == hari and j['jam'] == target_j and j['kode_guru'] == k_guru), None)
                            if bg:
                                bentrok = True
                                break
                        
                        if not bentrok:
                            for step in range(alokasi_hari_ini):
                                st.session_state.jadwal_terplot.append({
                                    "kode_guru": k_guru,
                                    "kode_mapel": k_mapel,
                                    "kelas": kelas,
                                    "hari": hari,
                                    "jam": jam_mulai_pjok + step
                                })
                            # Jika jam pagi terpakai PJOK, majukan pointer kelas umum agar tidak menumpuk
                            if pointer_kbm[hari][kelas] == jam_mulai_pjok:
                                pointer_kbm[hari][kelas] += alokasi_hari_ini
                            jp_tersisa -= alokasi_hari_ini
                            
                    else:
                        # LOGIKA UNTUK MAPEL UMUM NON-PJOK
                        jam_mulai = pointer_kbm[hari][kelas]
                        
                        # Validasi kapasitas slot hari
                        if jam_mulai + alokasi_hari_ini - 1 <= max_jam:
                            bentrok = False
                            for step in range(alokasi_hari_ini):
                                target_j = jam_mulai + step
                                # Cek tabrakan mengajar guru lintas kelas
                                bg = next((j for j in st.session_state.jadwal_terplot if j['hari'] == hari and j['jam'] == target_j and j['kode_guru'] == k_guru), None)
                                if bg:
                                    bentrok = True
                                    break
                            
                            if not bentrok:
                                for step in range(alokasi_hari_ini):
                                    st.session_state.jadwal_terplot.append({
                                        "kode_guru": k_guru,
                                        "kode_mapel": k_mapel,
                                        "kelas": kelas,
                                        "hari": hari,
                                        "jam": jam_mulai + step
                                    })
                                pointer_kbm[hari][kelas] += alokasi_hari_ini
                                jp_tersisa -= alokasi_hari_ini
                                
        st.success("✔️ Optimalisasi Berhasil! Jadwal telah disesuaikan dengan aturan komparasi baru.")
        st.rerun()

# ==================== TAB 2: OUTPUT JADWAL PER KELAS ====================
with tab2:
    st.subheader("🔍 Peninjauan Jadwal Sementara")
    
    c_v1, c_v2 = st.columns([2, 1])
    with c_v1:
        pilihan_view_kelas = st.selectbox("Pilih Kelas untuk Melihat Jadwal:", semua_kelas)
    with c_v2:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.session_state.jadwal_terplot:
            data_file_excel = export_excel_laporan(st.session_state.jadwal_terplot, semua_kelas, list_hari)
            st.download_button(
                label="📥 Unduh Seluruh Jadwal (Format Excel Cetak)",
                data=data_file_excel,
                file_name="Jadwal_Optimasi_SMP.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
    # Membuat Grid Render Jadwal per Kelas Pilihan
    matriks_tabel = pd.DataFrame(index=list_hari, columns=[f"Jam {i}" for i in range(1, 10)]).fillna("Kosong")
    
    for h in list_hari:
        batas_jam = 6 if h == 'Jumat' else 9
        for j in range(1, 10):
            if j > batas_jam:
                matriks_tabel.at[h, f"Jam {j}"] = "-"
                continue
            if h == 'Senin' and j == 1:
                matriks_tabel.at[h, f"Jam {j}"] = "🎗️ UPACARA"
                continue
                
            # Ambil data dari state jadwal terplot
            match = next((d for d in st.session_state.jadwal_terplot if d['hari'] == h and d['jam'] == j and d['kelas'] == pilihan_view_kelas), None)
            if match:
                matriks_tabel.at[h, f"Jam {j}"] = f"{match['kode_mapel']} ({match['kode_guru']})"
                
    st.markdown(f"#### 🗓️ Tabel Jadwal Kelas: {pilihan_view_kelas}")
    st.dataframe(matriks_tabel, use_container_width=True)
