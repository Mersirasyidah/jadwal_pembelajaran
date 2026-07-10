import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# =====================================================================
# 1. KONFIGURASI HALAMAN & DATABASE STATE
# =====================================================================
st.set_page_config(page_title="Sistem Jadwal SMPN 1 Bambanglipuro", layout="wide")

# Database default (Bisa ditambah/ubah lewat aplikasi)
if 'database_guru' not in st.session_state:
    st.session_state.database_guru = [
        {"nama": "Sunarwi, S.Pd", "mapel": "PKN", "kelas_diampu": ["7E", "7F"]},
        {"nama": "Ani Pujiastuti, M.Hum", "mapel": "Bahasa Inggris", "kelas_diampu": ["7A", "7B", "7C"]},
        {"nama": "Agus Fuadi, M.Pd", "mapel": "IPA", "kelas_diampu": ["9E", "9A"]},
        {"nama": "Hartini, M.Pd", "mapel": "Matematika", "kelas_diampu": ["8A", "8B"]},
        {"nama": "Fitri Kurnia Pangestuti, S.Pd", "mapel": "Bahasa Indonesia", "kelas_diampu": ["7C", "7D", "9B"]}
    ]

if 'plotting_jadwal' not in st.session_state:
    st.session_state.plotting_jadwal = []

# Parameter Sekolah
tingkatan = ['7', '8', '9']
abjad = ['A', 'B', 'C', 'D', 'E']
list_kelas = [f"{t}{a}" for t in tingkatan for a in abjad]
list_hari = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat']

st.title("🗃️ Generator Jadwal SMP & Ekspor Excel Landscape")
st.caption("Sistem Aplikasi Anti-Bentrok berdasarkan aturan jam mengajar SMPN 1 Bambanglipuro.")

# =====================================================================
# 2. FUNGSI UNTUK GENERATE EXCEL LANDSCAPE (IN-MEMORY)
# =====================================================================
def buat_excel_landscape(database_plotting, kelas_terpilih):
    wb = Workbook()
    ws = wb.active
    ws.title = f"Kelas {kelas_terpilih}"
    ws.views.sheetView[0].showGridLines = True
    
    # Pengaturan Cetak Kertas Landscape A4 Pas 1 Halaman
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    
    # Header Dokumen
    ws.append([])
    ws.append([f"JADWAL PELAJARAN KELAS {kelas_terpilih}"])
    ws.append(["SMP NEGERI 1 BAMBANGLIPURO"])
    ws.append([])
    
    ws['A2'].font = Font(name="Arial", size=14, bold=True, color="1B365D")
    ws['A3'].font = Font(name="Arial", size=11, italic=True)
    
    # Membuat Struktur Matriks tabel
    headers = ["Hari"] + [f"Jam {i}" for i in range(1, 10)]
    ws.append(headers)
    
    for h in list_hari:
        row_data = [h]
        max_j = 6 if h == 'Jumat' else 9
        for j in range(1, 10):
            if j > max_j:
                row_data.append("-")
            elif h == 'Senin' and j == 1:
                row_data.append("🎗️ UPACARA")
            else:
                slot = next((d for d in database_plotting if d['hari'] == h and d['jam'] == j and d['kelas'] == kelas_terpilih), None)
                if slot:
                    row_data.append(f"{slot['mapel']}\n({slot['guru']})")
                else:
                    row_data.append("Kosong")
        ws.append(row_data)
        
    # Styling Tabel openpyxl
    font_header = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    fill_header = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    fill_upacara = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    fill_kosong = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    border_tipis = Border(left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'),
                          top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF'))
    
    # Format baris header (Baris ke-5)
    for col_idx in range(1, 12):
        cell = ws.cell(row=5, column=col_idx)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        
    # Format baris isi (Baris ke-6 s.d 10)
    for row_idx in range(6, 11):
        ws.row_dimensions[row_idx].height = 40
        for col_idx in range(1, 12):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = Font(name="Arial", size=10)
            cell.alignment = align_center
            cell.border = border_tipis
            
            if "UPACARA" in str(cell.value):
                cell.fill = fill_upacara
                cell.font = Font(name="Arial", size=10, bold=True, color="7F6000")
            elif cell.value == "Kosong":
                cell.fill = fill_kosong
                cell.font = Font(name="Arial", size=9, color="A6A6A6")
                
    ws.column_dimensions['A'].width = 12
    for col in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']:
        ws.column_dimensions[col].width = 16
        
    # Simpan ke bentuk byte stream agar bisa didownload langsung di web
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# =====================================================================
# 3. INTERFACE APLIKASI (TABS)
# =====================================================================
tab1, tab2, tab3 = st.tabs(["👥 1. Database Guru & Ampuan", "📅 2. Plotting Jadwal Harian", "📊 3. Hasil & Cetak Excel"])

# --- TAB 1 ---
with tab1:
    st.subheader("Manajemen Guru dan Kelas yang Diampu")
    with st.expander("➕ Tambah Data Guru Baru"):
        c1, c2, c3 = st.columns(3)
        with c1: new_nama = st.text_input("Nama Guru Baru:")
        with c2: new_mapel = st.text_input("Mata Pelajaran:")
        with c3: new_kelas = st.multiselect("Kelas yang Diampu:", list_kelas)
        if st.button("Simpan Data"):
            if new_nama and new_mapel and new_kelas:
                st.session_state.database_guru.append({"nama": new_nama, "mapel": new_mapel, "kelas_diampu": new_kelas})
                st.success("Data berhasil masuk database!")
                st.rerun()
                
    df_g = pd.DataFrame(st.session_state.database_guru)
    df_g['kelas_diampu'] = df_g['kelas_diampu'].apply(lambda x: ", ".join(x))
    st.dataframe(df_g, use_container_width=True)

# --- TAB 2 ---
with tab2:
    st.subheader("⏰ Input Jadwal Pelajaran (Sistem Validasi)")
    col_in, col_at = st.columns([1, 1])
    with col_in:
        list_nama_g = [g['nama'] for g in st.session_state.database_guru]
        p_guru = st.selectbox("Pilih Guru:", list_nama_g)
        
        info_g = next(g for g in st.session_state.database_guru if g['nama'] == p_guru)
        st.info(f"📖 Mapel: **{info_g['mapel']}**")
        
        p_kelas = st.selectbox("Pilih Kelas (Dibatasi sesuai hak ampu):", info_g['kelas_diampu'])
        p_hari = st.selectbox("Pilih Hari:", list_hari)
        
        max_j_hari = 6 if p_hari == 'Jumat' else 9
        j_mulai = st.number_input("Mulai Jam Ke-:", min_value=1, max_value=max_j_hari, value=1)
        durasi = st.number_input("Jumlah Jam Pelajaran:", min_value=1, max_value=5, value=2)
        
        if st.button("🚀 Plotting ke Jadwal Harian", type="primary"):
            if p_hari == 'Jumat' and (j_mulai + durasi - 1) > 6:
                st.error("❌ Hari Jumat maksimal hanya sampai jam ke-6!")
            elif p_hari == 'Senin' and j_mulai == 1:
                st.error("❌ Hari Senin jam ke-1 wajib digunakan untuk Upacara!")
            else:
                bentrok = False
                slot_baru = []
                for i in range(durasi):
                    js = j_mulai + i
                    if js > max_j_hari:
                        st.error("❌ Durasi melebihi batas waktu sekolah!")
                        bentrok = True; break
                        
                    bg = next((j for j in st.session_state.plotting_jadwal if j['hari'] == p_hari and j['jam'] == js and j['guru'] == p_guru), None)
                    if bg:
                        st.error(f"❌ Guru bentrok! Mengajar di kelas {bg['kelas']} pada jam yang sama."); bentrok = True; break
                        
                    bk = next((j for j in st.session_state.plotting_jadwal if j['hari'] == p_hari and j['jam'] == js and j['kelas'] == p_kelas), None)
                    if bk:
                        st.error(f"❌ Kelas bentrok! Sudah terisi mapel {bk['mapel']}."); bentrok = True; break
                        
                    slot_baru.append({"guru": p_guru, "mapel": info_g['mapel'], "kelas": p_kelas, "hari": p_hari, "jam": js})
                    
                if not bentrok:
                    st.session_state.plotting_jadwal.extend(slot_baru)
                    st.success("✔️ Jadwal berhasil di-plot!")

    with col_at:
        st.markdown("### ⏱️ Acuan Jam Sekolah:\n* **Senin-Kamis (40 Menit)**: Jam 1 (07.20), Ist1 (09.20), Ist2 (12.00), Pulang (14.20).\n* **Jumat (35 Menit)**: Maksimal Jam ke-6, Ist (09.40), Pulang setelah Jam 6.")

# --- TAB 3 ---
with tab3:
    st.subheader("📊 Visualisasi & Cetak File Excel")
    kelas_view = st.selectbox("Pilih Kelas Tinjauan:", list_kelas)
    
    matriks = pd.DataFrame(index=list_hari, columns=[f"Jam {i}" for i in range(1, 10)]).fillna("Kosong")
    for h in list_hari:
        max_j = 6 if h == 'Jumat' else 9
        for j in range(1, 10):
            if j > max_j: matriks.at[h, f"Jam {j}"] = "-"
            elif h == 'Senin' and j == 1: matriks.at[h, f"Jam {j}"] = "🎗️ UPACARA"
            else:
                s = next((d for d in st.session_state.plotting_jadwal if d['hari'] == h and d['jam'] == j and d['kelas'] == kelas_view), None)
                if s: matriks.at[h, f"Jam {j}"] = f"{s['mapel']} ({s['guru']})"
                
    st.dataframe(matriks, use_container_width=True)
    
    # TOMBOL HORE: UNDUH EXCEL LANDSCAPE
    excel_data = buat_excel_landscape(st.session_state.plotting_jadwal, kelas_view)
    st.download_button(
        label=f"📥 Unduh Jadwal Kelas {kelas_view} (Format Excel Landscape)",
        data=excel_data,
        file_name=f"Jadwal_Kelas_{kelas_view}_Landscape.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    if st.button("🗑️ Reset Seluruh Jadwal Hari Ini"):
        st.session_state.plotting_jadwal = []
        st.rerun()
