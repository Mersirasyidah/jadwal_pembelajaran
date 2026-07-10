import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io

# 1. Konfigurasi Halaman & State
st.set_page_config(page_title="Sistem Database Jadwal SMP", layout="wide")

# Database awal (Mapping Guru, Mapel, dan Kelas yang diampu)
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

st.title("🗃️ Sistem Database Pengajar & Generator Jadwal")
st.caption("Aplikasi otomatis mencocokkan guru dengan kelas yang diampunya untuk menghindari salah input.")

tab1, tab2, tab3 = st.tabs(["👥 1. Database Guru & Ampuan", "📅 2. Plotting Jadwal Harian", "📊 3. Hasil Jadwal & Cetak Laporan"])

# ==================== TAB 1: DATABASE GURU ====================
with tab1:
    st.subheader("Manajemen Guru dan Kelas yang Diampu")
    
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
                    "nama": new_nama, "mapel": new_mapel, "kelas_diampu": new_kelas
                })
                st.success(f"Berhasil menambahkan {new_nama} ke database!")
                st.rerun()
            else:
                st.error("Semua data form wajib diisi/dipilih!")

    df_guru = pd.DataFrame(st.session_state.database_guru)
    df_guru['kelas_diampu'] = df_guru['kelas_diampu'].apply(lambda x: ", ".join(x))
    st.dataframe(df_guru, use_container_width=True)

# ==================== TAB 2: PLOTTING JADWAL ====================
with tab2:
    st.subheader("⏰ Input Jadwal Pembelajaran")
    
    # Kategori Metode Input
    metode_input = st.radio("Pilih Metode Plotting:", ["⚡ Otomatis (Satu Kelas Sekaligus)", "✍️ Manual (Per Guru)"], horizontal=True)
    
    st.markdown("---")
    
    if metode_input == "⚡ Otomatis (Satu Kelas Sekaligus)":
        st.markdown("### 🤖 Smart Auto-Scheduler per Kelas")
        st.caption("Fitur ini akan mencari seluruh guru yang mengampu kelas terpilih, lalu menata jadwalnya secara otomatis berurutan ke bawah.")
        
        c_auto1, c_auto2, c_auto3 = st.columns(3)
        with c_auto1:
            auto_kelas = st.selectbox("Pilih Kelas Sasaran:", list_kelas, key="auto_kelas")
        with c_auto2:
            auto_hari = st.selectbox("Pilih Hari:", list_hari, key="auto_hari")
        with c_auto3:
            durasi_default = st.number_input("Durasi per Mapel (Jam):", min_value=1, max_value=4, value=2)
            
        # Tampilkan daftar guru yang akan dimasukkan berdasarkan database
        guru_terkait = [g for g in st.session_state.database_guru if auto_kelas in g['kelas_diampu']]
        
        if guru_terkait:
            st.write("**Daftar Guru & Mapel yang terdeteksi mengampu kelas ini:**")
            df_terkait = pd.DataFrame(guru_terkait)[['nama', 'mapel']]
            st.dataframe(df_terkait, use_container_width=True)
            
            if st.button("⚡ Generate Jadwal Otomatis untuk Kelas Ini", type="primary"):
                max_jam = 6 if auto_hari == 'Jumat' else 9
                jam_sekarang = 2 if (auto_hari == 'Senin') else 1 # Jika senin, jam 1 dilewati untuk upacara
                
                sukses_plot = 0
                gagal_plot = 0
                
                for guru_info in guru_terkait:
                    # Cek apakah jam sekolah sudah penuh
                    if jam_sekarang > max_jam:
                        break
                        
                    # Cek potensi bentrok untuk alokasi jam ke depan
                    bentrok = False
                    slot_auto_baru = []
                    
                    for i in range(durasi_default):
                        target_jam = jam_sekarang + i
                        if target_jam > max_jam:
                            bentrok = True
                            break
                        
                        # 1. Cek apakah gurunya sedang mengajar di kelas lain pada jam ini
                        b_guru = next((j for j in st.session_state.plotting_jadwal if j['hari'] == auto_hari and j['jam'] == target_jam and j['guru'] == guru_info['nama']), None)
                        # 2. Cek apakah kelas ini sudah terisi secara manual sebelumnya
                        b_kelas = next((j for j in st.session_state.plotting_jadwal if j['hari'] == auto_hari and j['jam'] == target_jam and j['kelas'] == auto_kelas), None)
                        
                        if b_guru or b_kelas:
                            bentrok = True
                            break
                            
                        slot_auto_baru.append({
                            "guru": guru_info['nama'], "mapel": guru_info['mapel'], "kelas": auto_kelas, "hari": auto_hari, "jam": target_jam
                        })
                    
                    if not bentrok and slot_auto_baru:
                        st.session_state.plotting_jadwal.extend(slot_auto_baru)
                        jam_sekarang += durasi_default
                        sukses_plot += 1
                    else:
                        gagal_plot += 1
                        
                st.success(f"✔️ Berhasil memplot otomatis {sukses_plot} Mata Pelajaran ke dalam Hari {auto_hari}!")
                if gagal_plot > 0:
                    st.warning(f"⚠️ {gagal_plot} Guru dilewati karena terdeteksi bentrok jam mengajar di kelas lain.")
                st.rerun()
        else:
            st.warning(f"Belum ada data guru di Tab 1 yang di-setting untuk mengampu kelas {auto_kelas}.")

    else:
        # ==================== METODE MANUAL (KODE ASLI ANDA) ====================
        col_input, col_info = st.columns([1, 1])
        with col_input:
            list_nama_guru = [g['nama'] for g in st.session_state.database_guru]
            pilihan_guru = st.selectbox("Pilih Guru pengajar:", list_nama_guru)
            
            info_guru_terpilih = next(g for g in st.session_state.database_guru if g['nama'] == pilihan_guru)
            mapel_otomatis = info_guru_terpilih['mapel']
            kelas_tersedia = info_guru_terpilih['kelas_diampu']
            
            st.info(f"📖 Mata Pelajaran: **{mapel_otomatis}**")
            
            pilihan_kelas = st.selectbox("Pilih Kelas (Hanya yang diampu guru ini):", kelas_tersedia)
            pilihan_hari = st.selectbox("Pilih Hari Pelajaran:", list_hari)
            
            max_jam = 6 if pilihan_hari == 'Jumat' else 9
            jam_mulai = st.number_input("Mulai dari Jam Ke-:", min_value=1, max_value=max_jam, value=1, key="jam_mulai_tab2")
            durasi_jam = st.number_input("Durasi (Jumlah Jam Pelajaran):", min_value=1, max_value=5, value=2, key="durasi_tab2")
            
            if st.button("🚀 Masukkan ke Jadwal Harian", type="primary"):
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
                            
                        b_guru = next((j for j in st.session_state.plotting_jadwal if j['hari'] == pilihan_hari and j['jam'] == j_sekarang and j['guru'] == pilihan_guru), None)
                        if b_guru:
                            st.error(f"❌ Guru {pilihan_guru} sudah mengajar di kelas {b_guru['kelas']} pada jam ke-{j_sekarang}!")
                            bentrok = True
                            break
                            
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

# ==================== TAB 3: HASIL JADWAL & DOWNLOAD EXCEL ====================
with tab3:
    st.subheader("📊 Laporan Jadwal Pembelajaran Menyeluruh")
    st.write("Unduh berkas laporan akhir berformat matriks horizontal terpadu untuk kebutuhan cetak dinding / arsip kurikulum:")
    
    excel_data = export_jadwal_komplit_excel(st.session_state.plotting_jadwal, list_kelas, list_hari)
    
    st.download_button(
        label="📥 Unduh Laporan Akhir Jadwal (Excel - Kertas Legal Horizontal)",
        data=excel_data,
        file_name="Laporan_Jadwal_Menyeluruh_SMP_7A_9E.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
    
    st.markdown("---")
    st.subheader("🔍 Peninjauan Internal per Rombel Kelas")
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
                matriks.at[h, f"Jam {j}"] = f"{slot['mapel']}\n({slot['guru'].split(',')[0]})"
                
    st.dataframe(matriks, use_container_width=True)
    
    if st.button("🗑️ Reset Jadwal Harian"):
        st.session_state.plotting_jadwal = []
        st.rerun()
