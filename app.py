import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def create_aplikasi_jadwal():
    # 1. Inisialisasi Workbook
    wb = openpyxl.Workbook()
    
    # Desain & Tema Warna (Classic Corporate Navy)
    navy_dark = "1B365D"      # Warna utama header tabel
    blue_accent = "4A90E2"    # Warna sub-header / aksen tingkat kelas
    ice_blue = "E6F0FA"       # Warna highlight panel pelacak (Sidebar)
    zebra_light = "F7FAFC"    # Warna abu-abu tipis untuk baris data genap
    white = "FFFFFF"
    gray_border = "D1D5DB"
    
    # Pengaturan Gaya Huruf (Typography)
    font_title = Font(name="Segoe UI", size=15, bold=True, color=navy_dark)
    font_subtitle = Font(name="Segoe UI", size=10, italic=True, color="555555")
    font_header = Font(name="Segoe UI", size=10, bold=True, color=white)
    font_subhead = Font(name="Segoe UI", size=10, bold=True, color=white)
    font_data = Font(name="Segoe UI", size=10)
    font_bold = Font(name="Segoe UI", size=10, bold=True)
    
    # Pengaturan Isian Warna (Fills)
    fill_header = PatternFill(start_color=navy_dark, end_color=navy_dark, fill_type="solid")
    fill_subhead = PatternFill(start_color=blue_accent, end_color=blue_accent, fill_type="solid")
    fill_zebra = PatternFill(start_color=zebra_light, end_color=zebra_light, fill_type="solid")
    fill_accent = PatternFill(start_color=ice_blue, end_color=ice_blue, fill_type="solid")
    
    # Garis Tepi (Borders)
    thin_side = Side(border_style="thin", color=gray_border)
    border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    
    # =========================================================================
    # TAB 1: DATA MASTER GURU
    # =========================================================================
    ws1 = wb.active
    ws1.title = "1. Data Master Guru"
    ws1.views.sheetView[0].showGridLines = True  # Memastikan garis kisi Excel aktif
    
    # Judul Lembar Kerja
    ws1["A1"] = "DATABASE GURU & BEBAN MENGAJAR (INPUT MASTER)"
    ws1["A1"].font = font_title
    ws1["A2"] = "Isi daftar guru, kode singkatan, mata pelajaran, dan jumlah total Jam Pelajaran (JP) per minggu."
    ws1["A2"].font = font_subtitle
    
    headers_ws1 = ["KODE GURU", "NAMA GURU", "MATA PELAJARAN", "TOTAL JP / MINGGU"]
    for col_idx, h in enumerate(headers_ws1, 1):
        cell = ws1.cell(row=4, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_cell
        
    # Data Sampel Guru (Struktur Pengajar Rujukan)
    sample_teachers = [
        ("KRN", "KIRNO WIDARSO, M.Pd., M.M", "Kepala Sekolah / PAI", 6),
        ("HRT", "Hartini, M.Pd", "Bahasa Indonesia", 24),
        ("BRT", "Bartina, S.Pd", "Matematika", 24),
        ("ANI", "Ani Pujiastuti, M.Hum", "Bahasa Inggris", 20),
        ("AGS", "Agus Fuadi, M.Pd", "IPA", 20),
        ("IST", "Dra. Isti Widayanti", "IPS", 18),
        ("SRW", "Sunarwi, S.Pd", "PPKn", 16),
        ("HLT", "Herlita Dewi Setyawati, S.Pd", "Informatika / TIK", 16),
        ("TNA", "Martina Supraptini, S.Pd", "Seni Budaya", 12),
        ("DWI", "Dwi Indriyani, S.Pd", "Bahasa Jawa", 14),
    ]
    
    for idx, data in enumerate(sample_teachers):
        r = 5 + idx
        for c_idx, val in enumerate(data, 1):
            cell = ws1.cell(row=r, column=c_idx, value=val)
            cell.font = font_data
            cell.border = border_cell
            if c_idx in [1, 4]:
                cell.alignment = Alignment(horizontal="center")
            if r % 2 == 0:
                cell.fill = fill_zebra
                
    # Menyesuaikan Lebar Kolom secara Proporsional
    ws1.column_dimensions['A'].width = 15
    ws1.column_dimensions['B'].width = 35
    ws1.column_dimensions['C'].width = 25
    ws1.column_dimensions['D'].width = 22

    # =========================================================================
    # TAB 2: ENGINE GENERATOR JADWAL & MONITOR BENTROK
    # =========================================================================
    ws2 = wb.create_sheet(title="2. Generator & Cek Bentrok")
    ws2.views.sheetView[0].showGridLines = True
    
    ws2["A1"] = "PANEL GENERATOR JADWAL PELAJARAN LINTAS KELAS"
    ws2["A1"].font = font_title
    ws2["A2"] = "Masukkan KODE GURU di bawah kolom kelas. Cek panel kanan untuk memantau jam terpakai (Live Tracker)."
    ws2["A2"].font = font_subtitle
    
    # Inisialisasi Kolom Rombongan Belajar (Paralel Kelas 7, 8, dan 9)
    classes = [
        "7A", "7B", "7C", "7D", "7E", "7F", "7G",
        "8A", "8B", "8C", "8D", "8E", "8F", "8G",
        "9A", "9B", "9C", "9D", "9E", "9F", "9G"
    ]
    
    # Penataan Struktur Judul Matriks
    ws2.merge_cells("A4:A5")
    ws2["A4"] = "HARI"
    ws2.merge_cells("B4:B5")
    ws2["B4"] = "JAM KE"
    
    # Mengisi baris kelas paralel
    for idx, cls in enumerate(classes):
        col_num = 3 + idx
        cell = ws2.cell(row=5, column=col_num, value=cls)
        cell.font = font_bold
        cell.alignment = Alignment(horizontal="center")
        cell.border = border_cell
        
    # Pengelompokan baris tingkat rombel (VII, VIII, IX)
    ws2.merge_cells("C4:I4")
    ws2["C4"] = "TINGKAT VII"
    ws2.merge_cells("J4:P4")
    ws2["J4"] = "TINGKAT VIII"
    ws2.merge_cells("Q4:W4")
    ws2["Q4"] = "TINGKAT IX"
    
    # Memberikan pewarnaan terpadu pada area Header Matriks
    for r_idx in [4, 5]:
        for c_idx in range(1, 24):
            cell = ws2.cell(row=r_idx, column=c_idx)
            cell.font = font_header
            cell.fill = fill_header if r_idx == 4 else fill_subhead
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border_cell

    # =========================================================================
    # SIDEBAR PANEL: LIVE TRACKER (VERIFIKASI ALOKASI JAM GURU)
    # =========================================================================
    start_sb_col = 25  # Berada di Kolom Y ke kanan
    ws2.cell(row=4, column=start_sb_col, value="KODE").fill = fill_header
    ws2.cell(row=4, column=start_sb_col+1, value="TERPLOT (Live)").fill = fill_header
    ws2.cell(row=4, column=start_sb_col+2, value="TARGET JP").fill = fill_header
    
    for c_offset in range(3):
        target_cell = ws2.cell(row=4, column=start_sb_col + c_offset)
        target_cell.font = font_header
        target_cell.alignment = Alignment(horizontal="center")

    # Penyusunan Formula Otomatis untuk Mengawasi Jadwal Bentrok
    for idx, data in enumerate(sample_teachers):
        r_sb = 5 + idx
        # Kolom Kode Guru
        cell_kd = ws2.cell(row=r_sb, column=start_sb_col, value=data[0])
        cell_kd.font = font_bold
        cell_kd.alignment = Alignment(horizontal="center")
        cell_kd.border = border_cell
        cell_kd.fill = fill_accent
        
        # Formula COUNTIF: Menghitung kemunculan kode pengajar secara waktu-nyata pada seluruh grid KBM
        cell_count = ws2.cell(row=r_sb, column=start_sb_col+1, value=f"=COUNTIF(C$6:W$50, Y{r_sb})")
        cell_count.font = font_data
        cell_count.alignment = Alignment(horizontal="center")
        cell_count.border = border_cell
        
        # Formula VLOOKUP: Mengambil batas beban mengajar wajib dari database master di tab pertama
        cell_target = ws2.cell(row=r_sb, column=start_sb_col+2, value=f"=VLOOKUP(Y{r_sb}, '1. Data Master Guru'!A$5:D$50, 4, FALSE)")
        cell_target.font = font_data
        cell_target.alignment = Alignment(horizontal="center")
        cell_target.border = border_cell

    # =========================================================================
    # MATRIKS ALOKASI WAKTU HARIAN (SENIN - JUMAT)
    # =========================================================================
    hari_list = ["SENIN", "SELASA", "RABU", "KAMIS", "JUMAT"]
    row_kbm = 6
    
    for hari in hari_list:
        # Menentukan rentang jam belajar (misal: Jumat 5 jam pelajaran, hari lain 7 jam pelajaran)
        total_jam = 5 if hari == "JUMAT" else 7  
        
        # Menggabungkan blok hari secara vertikal
        ws2.merge_cells(start_row=row_kbm, start_column=1, end_row=row_kbm + total_jam - 1, end_column=1)
        cell_hari = ws2.cell(row=row_kbm, column=1, value=hari)
        cell_hari.font = font_bold
        cell_hari.alignment = Alignment(horizontal="center", vertical="center")
        cell_hari.fill = PatternFill(start_color="EAEDF1", end_color="EAEDF1", fill_type="solid")
        cell_hari.border = border_cell
        
        for jam in range(total_jam):
            r_current = row_kbm + jam
            
            # Kolom Nomor Urut Jam Pelajaran
            cell_jam = ws2.cell(row=r_current, column=2, value=jam + 1)
            cell_jam.font = font_bold
            cell_jam.alignment = Alignment(horizontal="center")
            cell_jam.border = border_cell
            
            # Pembentukan Garis Kisi Grid Utama Pengisian Jadwal (Kolom C s.d W)
            for c_kelas in range(3, 24):
                grid_cell = ws2.cell(row=r_current, column=c_kelas)
                grid_cell.border = border_cell
                grid_cell.font = font_data
                grid_cell.alignment = Alignment(horizontal="center")
                
                # Pengisian Contoh Data Pengisi Jadwal (Dapat dihapus/diubah secara bebas di Excel)
                if hari == "SENIN" and jam in [0, 1] and c_kelas == 3: 
                    grid_cell.value = "HRT"  # Mengajar di 7A pada jam 1-2 Senin
                if hari == "SENIN" and jam in [0, 1] and c_kelas == 4: 
                    grid_cell.value = "BRT"  # Mengajar di 7B pada jam 1-2 Senin
                    
        row_kbm += total_jam

    # Menyesuaikan Lebar Kolom Matriks Otomatis Agar Presisi dan Rapi
    for col in range(1, 24):
        col_letter = get_column_letter(col)
        ws2.column_dimensions[col_letter].width = 8 if col > 2 else 10
        
    ws2.column_dimensions['Y'].width = 12
    ws2.column_dimensions['Z'].width = 18
    ws2.column_dimensions['AA'].width = 15

    # Menyimpan dan Menghasilkan File Spreadsheet Jadi
    output_filename = "Aplikasi_Sistem_Jadwal_Pelajaran_Otomatis.xlsx"
    wb.save(output_filename)
    return output_filename

if __name__ == "__main__":
    create_aplikasi_jadwal()
