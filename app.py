import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io

# 1. Konfigurasi Halaman & State
st.set_page_config(page_title="Sistem Database Jadwal SMP", layout="wide")

# Struktur Kelas dari File Excel Anda (7A-7E, 8A-8D, 9A-9E)
list_kelas = [
    "7A", "7B", "7C", "7D", "7E",
    "8A", "8B", "8C", "8D",
    "9A", "9B", "9C", "9D", "9E"
]
list_hari = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat']

# Inisialisasi Database Guru Awal jika belum ada di Session State (Sesuai contoh Excel Anda)
if 'matrix_guru' not in st.session_state:
    raw_data = [
        {"KODE": "1", "NAMA GURU": "KIRNO WIDARSO, M.Pd., M.M", "7A":"", "7B":"", "7C":"", "7D":"", "7E":"", "8A":"", "8B":"", "8C":"", "8D":"", "9A":"", "9B":"", "9C":"", "9D":"", "9E":""},
        {"KODE": "2", "NAMA GURU": "Hartini, M.Pd", "7A":"", "7B":"", "7C":"", "7D":"", "7E":"", "8A":"PRA-2", "8B":"PRA-2", "8C":"PRA-2", "8D":"PRA-2", "7E":"", "8A":"", "8B":"", "8C":"", "8D":"", "9A":"", "9B":"", "9C":"", "9D":"", "9E":""},
        {"KODE": "3", "NAMA GURU": "Bartina, S.Pd", "7A":"", "7B":"", "7C":"", "7D":"", "7E":"", "8A":"", "8B":"", "8C":"", "8D":"IND-5", "9A":"", "9B":"", "9C":"", "9D":"", "9E":"IND-5"},
        {"KODE": "4", "NAMA GURU": "Ani Pujiastuti, M.Hum", "7A":"ING-3", "7B":"ING-3", "7C":"ING-3", "7D":"ING-3", "7E":"", "8A":"", "8B":"", "8C":"", "8D":"", "9A":"", "9B":"", "9C":"", "9D":"", "9E":""},
        {"KODE": "5", "NAMA GURU": "Agus Fuadi, M.Pd", "7A":"", "7B":"", "7C":"", "7D":"", "7E":"", "8A":"", "8B":"", "8C":"", "8D":"", "9A":"", "9B":"", "9C":"", "9D":"IPS-3", "9E":"IPS-3"},
    ]
    # Bungkus ke DataFrame agar siap diedit
    st.session_state.matrix_guru = pd.DataFrame(raw_data)

if 'plotting_jadwal' not in st.session_state:
    st.session_state.plotting_jadwal = []

# ==================== FUNGSI GENERATE EXCEL (HORIZONTAL, LEGAL, FIT 1 PAGE) ====================
def export_jadwal_komplit_excel(plotting_data, kelas_list, hari_list):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Jadwal Menyeluruh"
    ws.views.sheetView[0].showGridLines = True
    
    navy_dark = "1B365D"
    blue_accent = "4A90E2"
    ice_blue = "E6F0FA"
    white = "FFFFFF"
    
    font_title = Font(name="Arial", size=14, bold=True, color=navy_dark)
    font_header = Font(name="Arial", size=10, bold=True, color=white)
    font_data = Font(name="Arial", size=9)
    font_bold = Font(name="Arial", size=9, bold=True)
    
    fill_header = PatternFill(start_color=navy_dark, end_color=navy_dark, fill_type="solid")
    fill_subhead = PatternFill(start_color=blue_accent, end_color=blue_accent, fill_type="solid")
    fill_day = PatternFill(start_color="F2F4F7", end_color="F2F4F7", fill_type="solid")
    fill_special = PatternFill(start_color=ice_blue, end_color=ice_blue, fill_type="solid")
    
    thin_border = Border(
        left=Side(style='thin', color='D1D5DB'), right=Side(style='thin', color='D1D5DB'),
        top=Side(style='thin', color='D1D5DB'), bottom=Side(style='thin', color='D1D5DB')
    )
    
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(kelas_list) + 2)
    ws["A1"] = "LAPORAN JADWAL PEMBELAJARAN MENYELURUH KELAS 7A - 9E"
    ws["A1"].font = font_title
    ws["A1"].alignment = Alignment(horizontal="center")
    
    ws.cell(row=3, column=1, value="HARI").fill = fill_header
    ws.cell(row=3, column=2, value="JAM").fill = fill_header
    for c in [1, 2]:
        ws.cell(row=3, column=c).font = font_header
        ws.cell(row=3, column=c).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=3, column=c).border = thin_border
    
    for c_idx, kelas in enumerate(kelas_list, 3):
        cell = ws.cell(row=3, column=c_idx, value=f"KELAS {kelas}")
        cell.font = font_header
        cell.fill = fill_subhead
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
        
    current_row = 4
    for hari in hari_list:
        max_jam = 6 if hari == 'Jumat' else 9
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row + max_jam - 1, end_column=1)
        day_cell = ws.cell(row=current_row, column=1, value=hari.upper())
        day_cell.font = font_bold
        day_cell.fill = fill_day
        day_cell.alignment = Alignment(horizontal="center", vertical="center")
        
        for jam in range(1, max_jam + 1):
            r_num = current_row + jam - 1
            jam_cell = ws.cell(row=r_num, column=2, value=jam)
            jam_cell.font = font_bold
            jam_cell.alignment = Alignment(horizontal="center")
            jam_cell.border = thin_border
            ws.cell(row=r_num, column=1).border = thin_border
            
            for c_idx, kelas in enumerate(kelas_list, 3):
                cell_kbm = ws.cell(row=r_num, column=c_idx)
                cell_kbm.border = thin_border
                cell_kbm.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell_kbm.font = font_data
                
                if hari == 'Senin' and jam == 1:
                    cell_kbm.value = "UPACARA"
                    cell_kbm.fill = fill_special
                    continue
                    
                slot = next((d for d in plotting_data if d['hari'] == hari and d['jam'] == jam and d['kelas'] == kelas), None)
                if slot:
                    cell_kbm.value = f"{slot['mapel']}\n({slot['guru'].split(',')[0]})"
                else:
                    cell_kbm.value = ""
                    
        current_row += max_jam
        
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 6
    for col in range(3, len(kelas_list) + 3):
        ws.column_dimensions[get_column_letter(col)].width = 14
        
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_LEGAL
    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25
    ws.page_margins.top = 0.25
    ws.page_margins.bottom = 0.25
    
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ==================== MAIN LAYOUT STREAMLIT ====================
st.title("🗃️ Sistem Matriks Distribusi Guru & Generator Jadwal")
st.caption("Input data beban mengajar menggunakan tabel interaktif layaknya Microsoft Excel.")

tab1, tab2, tab3 = st.tabs(["📊 1. Input Matriks Guru & Kelas", "⚡ 2. Auto-Plotting Sistem", "📋 3. Hasil Laporan Akhir"])

# ==================== TAB 1: INPUT TABEL MATRIKS (SEPERTI EXCEL) ====================
with tab1:
    st.subheader("Tabel Distribusi Mengajar (Bisa Diedit Langsung)")
    st.info("💡 **Petunjuk Pengisian:** Isi sel kelas dengan format `[Singkatan Mapel]-[Jumlah Jam]`. Contoh: **IND-5** (Bahasa Indonesia, 5 Jam) atau **ING-3** (Bahasa Inggris, 3 Jam). Kosongkan jika guru tidak mengampu di kelas tersebut.")
    
    # Menampilkan Data Editor agar user bisa langsung ketik di tabel layaknya Excel
    edited_df = st.data_editor(
        st.session_state.matrix_guru,
        use_container_width=True,
        num_rows="dynamic", # Memungkinkan user menambah/menghapus baris guru baru
        key="data_guru_editor"
    )
    
    # Tombol untuk mengunci perubahan ke dalam sistem
    if st.button("💾 Simpan & Perbarui Database Matriks"):
        st.session_state.matrix_guru = edited_df
        st.success("Database berhasil diperbarui!")
        st.rerun()

# ==================== TAB 2: OTOMATISASI GENERATOR JADWAL HARIAN ====================
with tab2:
    st.subheader("🤖 Generator Jadwal Otomatis Lintas Kelas")
    st.caption("Sistem akan membaca seluruh kode mapel & durasi jam dari Matriks Tab 1, lalu menyusun jadwal harian secara otomatis tanpa bentrok.")
    
    c_p1, c_p2 = st.columns(2)
    with c_p1:
        auto_hari = st.selectbox("Pilih Hari Eksekusi Plotting:", list_hari)
    with c_p2:
        st.markdown("<br>", unsafe_allow_html=True)
        btn_generate = st.button("⚡ Jalankan Smart Auto-Scheduler untuk Hari Terpilih", type="primary", use_container_width=True)
        
    if btn_generate:
        # Menghapus plotting hari terpilih terlebih dahulu (supaya tidak menumpuk saat di-generate ulang)
        st.session_state.plotting_jadwal = [j for j in st.session_state.plotting_jadwal if j['hari'] != auto_hari]
        
        max_jam = 6 if auto_hari == 'Jumat' else 9
        
        # Dictionary untuk melacak jam ke berapa kelas tersebut saat ini siap diisi
        # Jika hari senin, kelas dimulai dari jam ke-2 (karena jam 1 dipakai upacara)
        pointer_jam_kelas = {kelas: (2 if auto_hari == 'Senin' else 1) for kelas in list_kelas}
        
        sukses_count = 0
        bentrok_count = 0
        
        # Iterasi setiap baris guru dari hasil input Matriks Excel di Tab 1
        for _, row in st.session_state.matrix_guru.iterrows():
            guru_nama = row["NAMA GURU"]
            
            if pd.isna(guru_nama) or str(guru_nama).strip() == "":
                continue
                
            # Cek setiap kolom kelas untuk melihat apakah guru ini memiliki tugas mengajar
            for kelas in list_kelas:
                cell_value = str(row[kelas]).strip() if pd.notna(row[kelas]) else ""
                
                # Memastikan format benar (ada tanda minus '-', misal IND-5)
                if "-" in cell_value:
                    try:
                        mapel, total_jp = cell_value.split("-")
                        total_jp = int(total_jp)
                    except:
                        continue # Lewati jika format penulisan salah
                        
                    # Tentukan durasi mengajar per hari (misal jika beban 5 jam, dipecah maksimal 2-3 jam per hari pertemuan)
                    durasi_hari_ini = min(total_jp, 2 if auto_hari == 'Jumat' else 3)
                    
                    jam_mulai = pointer_jam_kelas[kelas]
                    
                    # Validasi apakah slot jam kelas melebihi batas KBM sekolah
                    if jam_mulai + durasi_hari_ini - 1 > max_jam:
                        continue
                        
                    # Validasi Anti-Bentrok Guru di kelas lain pada jam yang sama
                    bentrok = False
                    slot_sementara = []
                    
                    for step in range(durasi_hari_ini):
                        jam_target = jam_mulai + step
                        
                        # Cek apakah guru sedang mengajar di tempat lain
                        is_bentrok_guru = next((j for j in st.session_state.plotting_jadwal if j['hari'] == auto_hari and j['jam'] == jam_target and j['guru'] == guru_nama), None)
                        
                        if is_bentrok_guru:
                            bentrok = True
                            break
                            
                        slot_sementara.append({
                            "guru": guru_nama,
                            "mapel": mapel,
                            "kelas": kelas,
                            "hari": auto_hari,
                            "jam": jam_target
                        })
                        
                    if not bentrok:
                        st.session_state.plotting_jadwal.extend(slot_sementara)
                        pointer_jam_kelas[kelas] += durasi_hari_ini
                        sukses_count += len(slot_sementara)
                    else:
                        bentrok_count += 1
                        
        st.success(f"✔️ Pemrosesan Selesai! Berhasil menempatkan {sukses_count} slot jam mengajar pada hari {auto_hari}.")
        if bentrok_count > 0:
            st.warning(f"⚠️ Terdapat {bentrok_count} jadwal guru yang ditunda/dilewati karena potensi bentrok mengajar lintas kelas di jam yang sama.")

# ==================== TAB 3: UNDUH LAPORAN AKHIR FORMAT EXCEL ====================
with tab3:
    st.subheader("📊 Cetak Laporan Akhir Jadwal Pembelajaran")
    st.write("Dapatkan file Excel (.xlsx) dengan tata letak horizontal penuh dari kelas 7A s.d 9E yang siap dicetak di kertas ukuran **Legal** secara mendatar:")
    
    # Generate file Excel dari state plotting saat ini
    excel_file = export_jadwal_komplit_excel(st.session_state.plotting_jadwal, list_kelas, list_hari)
    
    st.download_button(
        label="📥 Download Excel Laporan Akhir (Format Cetak Kertas Legal)",
        data=excel_file,
        file_name="Laporan_Jadwal_Horizontal_SMP.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
    
    st.markdown("---")
    st.subheader("🔍 Preview Jadwal Terplot Aktual (Per Kelas)")
    kelas_view = st.selectbox("Pilih Rombel Kelas untuk Peninjauan:", list_kelas)
    
    # Matriks Grid Preview untuk Streamlit UI
    preview_df = pd.DataFrame(index=list_hari, columns=[f"Jam {i}" for i in range(1, 10)]).fillna("Kosong")
    for h in list_hari:
        lim_j = 6 if h == 'Jumat' else 9
        for j in range(1, 10):
            if j > lim_j:
                preview_df.at[h, f"Jam {j}"] = "-"
                continue
            if h == 'Senin' and j == 1:
                preview_df.at[h, f"Jam {j}"] = "🎗️ UPACARA"
                continue
                
            match = next((d for d in st.session_state.plotting_jadwal if d['hari'] == h and d['jam'] == j and d['kelas'] == kelas_view), None)
            if match:
                preview_df.at[h, f"Jam {j}"] = f"{match['mapel']} ({match['guru'].split(',')[0]})"
                
    st.dataframe(preview_df, use_container_width=True)
    
    if st.button("🗑️ Reset Seluruh Plotting Jadwal"):
        st.session_state.plotting_jadwal = []
        st.rerun()
