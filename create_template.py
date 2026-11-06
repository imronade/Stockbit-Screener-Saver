import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()

# Typography
font_title = Font(name="Segoe UI", size=16, bold=True, color="1B5E20")
font_subtitle = Font(name="Segoe UI", size=10, italic=True, color="4A5568")
font_section = Font(name="Segoe UI", size=11, bold=True, color="1A202C")
font_header = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
font_bold = Font(name="Segoe UI", size=10, bold=True, color="2D3748")
font_regular = Font(name="Segoe UI", size=10, color="2D3748")
font_muted = Font(name="Segoe UI", size=9, italic=True, color="718096")

# Fills
fill_green_main = PatternFill(start_color="00875A", end_color="00875A", fill_type="solid") # Hijau Stockbit
fill_slate_main = PatternFill(start_color="334155", end_color="334155", fill_type="solid") # Slate
fill_step_badge = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")

thin_border = Border(
    left=Side(style='thin', color="E2E8F0"),
    right=Side(style='thin', color="E2E8F0"),
    top=Side(style='thin', color="E2E8F0"),
    bottom=Side(style='thin', color="E2E8F0")
)

align_center = Alignment(horizontal="center", vertical="center")
align_left = Alignment(horizontal="left", vertical="center")
align_right = Alignment(horizontal="right", vertical="center")

# =============================================================
# 1. SHEET: PANDUAN_PENGGUNAAN (Sheet Pertama)
# =============================================================
ws_guide = wb.active
ws_guide.title = "panduan_penggunaan"
ws_guide.views.sheetView[0].showGridLines = True

ws_guide.cell(row=1, column=1, value="📖 PANDUAN TEMPLATE STOCKBIT SCREENER SAVER").font = font_title
ws_guide.cell(row=2, column=1, value="Template otomatis untuk mencatat, merekap frekuensi kemunculan emiten, dan melihat irisan screener saham.").font = font_subtitle
ws_guide.row_dimensions[1].height = 26
ws_guide.row_dimensions[2].height = 18

guide_steps = [
    ("Langkah 1", "Buat Salinan Pribadi (Make a Copy)", "Klik menu 'File' > 'Make a copy' (Buat salinan) agar template ini tersimpan di Google Drive pribadi Anda dan bisa diedit sepenuhnya."),
    ("Langkah 2", "Atur Nama Screener & Kategori", "Buka sheet 'setting_kategori'. Masukkan nama screener Stockbit Anda dan tentukan nama kategorinya (misal: Buy on Weakness, Breakout, Trend Following)."),
    ("Langkah 3", "Hubungkan dengan Browser Extension", "Buka menu 'Extensions' > 'Apps Script' > klik 'Deploy' > 'New deployment' (Type: Web app, Who has access: Anyone) > Salin Web App URL ke pengaturan Extension browser Anda."),
    ("Langkah 4", "Kirim Data dari Stockbit Screener", "Buka screener di stockbit.com/screener, lalu klik tombol 'Sync to Sheets' (atau 'Download Excel'). Data mentah akan otomatis masuk ke sheet 'record_data'."),
    ("Langkah 5", "Pantau Hasil di Sheet Dashboard", "Buka sheet 'dashboard'. Tabel-tabel per kategori akan otomatis terisi, menghitung frekuensi emiten muncul, dan memetakan irisan screener tanpa perlu hitung manual!"),
]

# Table Header Panduan
guide_headers = ["Tahap", "Aktivitas", "Penjelasan Lengkap"]
for c_idx, h in enumerate(guide_headers, 1):
    cell = ws_guide.cell(row=4, column=c_idx, value=h)
    cell.font = font_header
    cell.fill = fill_slate_main
    cell.alignment = align_center
    cell.border = thin_border
ws_guide.row_dimensions[4].height = 24

for r_idx, step in enumerate(guide_steps, 5):
    ws_guide.row_dimensions[r_idx].height = 28
    for c_idx, val in enumerate(step, 1):
        cell = ws_guide.cell(row=r_idx, column=c_idx, value=val)
        cell.border = thin_border
        if c_idx == 1:
            cell.font = font_bold
            cell.alignment = align_center
            cell.fill = fill_step_badge
        elif c_idx == 2:
            cell.font = font_bold
            cell.alignment = align_left
        else:
            cell.font = font_regular
            cell.alignment = align_left

# Catatan Tambahan di Panduan
note_row = 12
ws_guide.cell(row=note_row, column=1, value="💡 Catatan Penting:").font = font_bold
ws_guide.cell(row=note_row+1, column=1, value="• Jangan mengubah nama sheet ('dashboard', 'record_data', 'setting_kategori') agar otomatisasi skrip berjalan lancar.").font = font_muted
ws_guide.cell(row=note_row+2, column=1, value="• Sheet 'dashboard' akan otomatis menggambar tabel baru jika Anda menambahkan kategori baru di 'setting_kategori'.").font = font_muted

# =============================================================
# 2. SHEET: dashboard (Template Kosong)
# =============================================================
ws_dash = wb.create_sheet(title="dashboard")
ws_dash.views.sheetView[0].showGridLines = True

ws_dash.cell(row=1, column=1, value="STOCKBIT SCREENER DASHBOARD").font = font_title
ws_dash.cell(row=2, column=1, value="Rekap otomatis per kategori: Menyesuaikan jumlah emiten dinamis dari database").font = font_subtitle
ws_dash.row_dimensions[1].height = 24
ws_dash.row_dimensions[2].height = 18

dash_headers = ["No.", "Emiten", "Total Muncul", "Irisan Screener", "Terakhir Muncul", "Harga Terakhir"]

# Placeholder Table Contoh (Kategori A)
ws_dash.merge_cells(start_row=4, start_column=1, end_row=4, end_column=6)
sec_cell = ws_dash.cell(row=4, column=1, value='List Hasil Screener "Nama Kategori"')
sec_cell.font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
sec_cell.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
sec_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws_dash.row_dimensions[4].height = 24

for c_idx, h in enumerate(dash_headers, 1):
    th_cell = ws_dash.cell(row=5, column=c_idx, value=h)
    th_cell.font = font_header
    th_cell.fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
    th_cell.alignment = align_center
    th_cell.border = thin_border
ws_dash.row_dimensions[5].height = 22

# Baris pesan kosong
ws_dash.merge_cells(start_row=6, start_column=1, end_row=6, end_column=6)
empty_msg = ws_dash.cell(row=6, column=1, value="(Belum ada data. Tabel akan terisi otomatis saat Anda mengirim hasil screener dari Stockbit)")
empty_msg.font = font_muted
empty_msg.alignment = align_center
empty_msg.border = thin_border
ws_dash.row_dimensions[6].height = 24

# =============================================================
# 3. SHEET: record_data (Template Kosong)
# =============================================================
ws_rec = wb.create_sheet(title="record_data")
ws_rec.views.sheetView[0].showGridLines = True

headers_rec = [
    "No", "Tanggal", "Hari", "Nama Screener", "Kategori", 
    "Emiten", "Harga Penutupan", "1D Return (%)", "Bandar Value", 
    "Net Foreign (NBSA)", "Foreign Streak"
]

for col_num, h in enumerate(headers_rec, 1):
    cell = ws_rec.cell(row=1, column=col_num, value=h)
    cell.font = font_header
    cell.fill = fill_green_main
    cell.alignment = align_center
    cell.border = thin_border
ws_rec.row_dimensions[1].height = 26

# =============================================================
# 4. SHEET: setting_kategori (Template Kosong Siap Diisi)
# =============================================================
ws_kat = wb.create_sheet(title="setting_kategori")
ws_kat.views.sheetView[0].showGridLines = True

headers_kat = ["Nama Screener Stockbit", "Kategori / Strategi", "Deskripsi Strategi (Opsional)"]
for col_num, h in enumerate(headers_kat, 1):
    cell = ws_kat.cell(row=1, column=col_num, value=h)
    cell.font = font_header
    cell.fill = fill_slate_main
    cell.alignment = align_center
    cell.border = thin_border
ws_kat.row_dimensions[1].height = 26

# =============================================================
# Auto Column Widths
# =============================================================
ws_guide.column_dimensions['A'].width = 14
ws_guide.column_dimensions['B'].width = 32
ws_guide.column_dimensions['C'].width = 75

ws_dash.column_dimensions['A'].width = 8
ws_dash.column_dimensions['B'].width = 14
ws_dash.column_dimensions['C'].width = 16
ws_dash.column_dimensions['D'].width = 36
ws_dash.column_dimensions['E'].width = 18
ws_dash.column_dimensions['F'].width = 16

ws_rec.column_dimensions['A'].width = 8
ws_rec.column_dimensions['B'].width = 14
ws_rec.column_dimensions['C'].width = 12
ws_rec.column_dimensions['D'].width = 32
ws_rec.column_dimensions['E'].width = 24
ws_rec.column_dimensions['F'].width = 12
ws_rec.column_dimensions['G'].width = 16
ws_rec.column_dimensions['H'].width = 16
ws_rec.column_dimensions['I'].width = 18
ws_rec.column_dimensions['J'].width = 20
ws_rec.column_dimensions['K'].width = 16

ws_kat.column_dimensions['A'].width = 32
ws_kat.column_dimensions['B'].width = 24
ws_kat.column_dimensions['C'].width = 50

out_path = r"G:\My Drive\Saham\Screener_Stockbit\Hasil_Screener_Stockbit.xlsx"
wb.save(out_path)
print("Template successfully saved at:", out_path)
