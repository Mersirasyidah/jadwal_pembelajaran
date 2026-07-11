import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io

# 1. Konfigurasi Halaman & State
st.set_page_config(page_title="Sistem Peta Jadwal SMP", layout="wide")

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
            "JP per Minggu": 3,
            "Jumlah Kelas": 3,
            "Total JP": 9,
            "Mengajar Kelas": ["7A", "7B", "7C"]
        },
        {
            "No": 2,
            "Kode Guru": "G02",
            "Nama Guru": "Agus Fuadi, M.Pd",
            "Mata Pelajaran": "Informatika",
            "Kode Mapel": "INF",
            "JP per Minggu": 3,
            "Jumlah Kelas": 2,
            "Total JP": 6,
            "Mengajar Kelas": ["7A", "8A"]
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
st.title("📋 Generator Jadwal Sementara & Distribusi Beban Guru")
st.caption("Masukkan detail beban mengajar guru, lalu biarkan sistem memetakan tabel jadwal per kelas secara otomatis.")

tab1, tab2 = st.tabs(["📝 1. Input Data Beban Mengajar", "📅 2. Peta Jadwal Sementara per Kelas"])

# ==================== TAB 1: FORM INPUT BEBAN DATA GURU ====================
with tab1:
    st.subheader("Form Input Detail Mengajar Guru")
    st.info("Silakan isi tabel di bawah ini. Anda bisa menambah baris baru dengan menekan tombol **(+) Add Row** di bagian bawah tabel.")
    
    # Menggunakan st.data_editor dengan kolom terkonfigurasi khusus
    edited_df = st.data_editor(
        st.session_state.data_beban_guru,
        column_config={
            "No": st.column_config.NumberColumn("No", width="small", min_value=1),
            "Kode Guru": st.column_config.TextColumn("Kode Guru", placeholder="Contoh: G01"),
            "Nama Guru": st.column_config.TextColumn("Nama Guru", placeholder="Nama Lengkap"),
            "Mata Pelajaran": st.column_config.TextColumn("Mata Pelajaran"),
            "Kode Mapel": st.column_config.TextColumn("Kode Mapel", placeholder="Contoh: INF / ING"),
            "JP per Minggu": st.column_config.NumberColumn("JP / Minggu", min_value=1, default=3),
            "Jumlah Kelas": st.column_config.NumberColumn("Jml Kelas", min_value=1, default=1),
            "Total JP": st.column_config.NumberColumn("Total JP (Otomatis)", disabled=True),
            "Mengajar Kelas": st.column_config.MultiselectColumn("Mengajar Kelas Apa Saja", options=semua_kelas)
        },
        num_rows="dynamic",
        use_container_width=True,
        key="editor_beban"
    )
    
    # Tombol Aksi untuk hitung & kunci data
    if st.button("💾 Simpan Data & Proses Peta Jadwal", type="primary"):
        # Hitung Total JP secara otomatis berdasarkan perkalian (JP per Minggu * Jumlah Kelas)
        edited_df["Total JP"] = edited_df["JP per Minggu"] * edited_df["Jumlah Kelas"]
        st.session_state.data_beban_guru = edited_df
        
        # PROSES AUTOMATIC PLOTTING KE JADWAL SEMENTARA
        st.session_state.jadwal_terplot = [] # Reset data lama
        
        # Dictionary untuk mencatat posisi jam terakhir yang terisi di tiap kelas & hari
        # Format: pointer_kbm[hari][kelas] = jam_ke
        pointer_kbm = {h: {k: (2 if h == 'Senin' else 1) for k in semua_kelas} for h in list_hari}
        
        for _, row in edited_df.iterrows():
            k_guru = row["Kode Guru"]
            k_mapel = row["Kode Mapel"]
            kelas_diampu = row["Mengajar Kelas"]
            jp_per_minggu = int(row["JP per Minggu"]) if pd.notna(row["JP per Minggu"]) else 0
            
            if not k_guru or not k_mapel or not isinstance(kelas_diampu, list):
                continue
                
            # Distribusikan JP ke kelas-kelas yang dipilih
            for kelas in kelas_diampu:
                jp_tersisa = jp_per_minggu
                
                # Coba plotting menyebar dari senin sampai jumat
                for hari in list_hari:
                    if jp_tersisa <= 0:
                        break
                        
                    max_jam = 6 if hari == 'Jumat' else 9
                    # Batasi maksimal jam per mapel dalam satu hari agar tidak bosan (misal maks 3 JP per hari)
                    alokasi_hari_ini = min(jp_tersisa, 3 if hari != 'Jumat' else 2)
                    
                    jam_mulai = pointer_kbm[hari][kelas]
                    
                    if jam_mulai + alokasi_hari_ini - 1 <= max_jam:
                        # Cek bentrok guru di jam tersebut
                        bentrok = False
                        for step in range(alokasi_hari_ini):
                            target_j = jam_mulai + step
                            bg = next((j for j in st.session_state.jadwal_terplot if j['hari'] == hari and j['jam'] == target_j and j['kode_guru'] == k_guru), None)
                            if bg:
                                bentrok = True
                                break
                        
                        if not bentrok:
                            # Masukkan ke jadwal terplot
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
                            
        st.success("✔️ Data berhasil disimpan dan Jadwal Sementara berhasil dibuat!")
        st.rerun()

# ==================== TAB 2: OUTPUT JADWAL PER KELAS ====================
with tab2:
    st.subheader("🔍 Peninjauan Jadwal Sementara")
    
    c_v1, c_v2 = st.columns([2, 1])
    with c_v1:
        pilihan_view_kelas = st.selectbox("Pilih Kelas untuk Melihat Jadwal:", semua_kelas)
    with c_v2:
        st.markdown("<br>", unsafe_allow_html=True)
        # Tombol Download Excel Laporan Akhir Keseluruhan
        if st.session_state.jadwal_terplot:
            data_file_excel = export_excel_laporan(st.session_state.jadwal_terplot, semua_kelas, list_hari)
            st.download_button(
                label="📥 Unduh Seluruh Jadwal (Format Excel Cetak)",
                data=data_file_excel,
                file_name="Jadwal_Sementara_SMP.xlsx",
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
