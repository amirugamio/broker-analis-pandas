import pandas as pd
import re
import openpyxl
from pathlib import Path

# 1. Baca File Excel Broker dan Kepemilikan >5% (gunakan header=1 untuk membaca baris judul gabungan)
ROOT = Path(__file__).resolve().parents[1]
broker_file = ROOT / 'Ringkasan Broker Terbaru.xlsx'
ownership_files = sorted(ROOT.glob('peng-*-lima-persen.xlsx'), key=lambda p: p.stat().st_mtime, reverse=True)
if not broker_file.exists():
    raise FileNotFoundError('File broker belum diunduh. Jalankan scripts/download_idx_sources.py terlebih dahulu.')
if not ownership_files:
    raise FileNotFoundError('File kepemilikan >5% IDX belum diunduh. Jalankan scripts/download_idx_sources.py terlebih dahulu.')
df_broker = pd.read_excel(broker_file, sheet_name=0)
df_raw = pd.read_excel(ownership_files[0], sheet_name=0, header=1)

# 2. Daftar kata umum untuk membersihkan nama broker
kata_umum = [
    'pt', 'tbk', 'sekuritas', 'indonesia', 'asia', 'digital', 'capital', 'asset', 
    'investama', 'investment', 'vickers', 'international', 'and', 'global', 
    'securities', 'maju', 'bank', 'but', 'n', 'a', 'ag', 'branch', 'cabang', 
    'persero', 'custody', 'rekening', 'tampungan', 'ksei', 'untuk', 'closed', 
    'member', 'saham', 'efek', 'tamp', 'tbk.', 'pt.'
]

def get_kata_unik(nama):
    if pd.isna(nama):
        return ""
    words = re.sub(r'[^a-zA-Z0-9\s]', ' ', str(nama)).lower().split()
    filtered = [w for w in words if w not in kata_umum]
    return " ".join(filtered).strip()

# 3. Buat Kamus Pemetaan Kata Unik -> Kode Broker
keyword_to_code = {}
for idx, row in df_broker.iterrows():
    kw = get_kata_unik(row['Nama Perusahaan'])
    code = row['Kode Perusahaan']
    if kw:
        keyword_to_code[kw] = code
        kata_pertama = kw.split()[0]
        if kata_pertama not in keyword_to_code:
            keyword_to_code[kata_pertama] = code

def map_kode_broker(nama_pemegang):
    if pd.isna(nama_pemegang):
        return None
    kw = get_kata_unik(nama_pemegang)
    if not kw:
        return None
    if kw in keyword_to_code:
        return keyword_to_code[kw]
    for word in kw.split():
        if word in keyword_to_code:
            return keyword_to_code[word]
    return kw.upper()

# Terapkan pemetaan kode broker pada Rekening Efek (kolom index 3)
df_raw['Kode Broker'] = df_raw.iloc[:, 3].apply(map_kode_broker)

# Forward fill kolom Kode Efek (index 1) dan Nama Pemegang Saham (index 4)
df_raw.iloc[:, 1] = df_raw.iloc[:, 1].ffill()
df_raw.iloc[:, 4] = df_raw.iloc[:, 4].ffill()

# Bersihkan baris header atau baris kosong yang tidak penting
df_clean = df_raw.dropna(subset=[df_raw.columns[1]]).copy()
df_clean = df_clean[df_clean.iloc[:, 1] != 'Kode Efek']

# Ambil kolom Saham Gabungan Per Investor (Index 12 untuk 16-OCT, Index 15 untuk 19-OCT, Index 17 untuk Perubahan)
# Kita gunakan nilai pertama yang muncul di setiap kelompok pemegang saham (karena nilai gabungan ada di baris pertama)
summary_rows = []

for (saham, pemegang), group in df_clean.groupby([df_clean.columns[1], df_clean.columns[4]]):
    # Ambil nilai Saham Gabungan Per Investor (kolom index 15) dari baris pertama kelompok ini
    saham_gabungan_val = group.iloc[0, 15]
    saham_lalu_val = group.iloc[0, 12]
    
    # Jumlahkan total perubahan di semua sub-rekening (kolom index 17)
    total_perubahan = pd.to_numeric(group.iloc[:, 17].astype(str).str.replace(',', ''), errors='coerce').fillna(0).sum()
    
    total_saham = pd.to_numeric(str(saham_gabungan_val).replace(',', ''), errors='coerce') or 0
    total_saham_lalu = pd.to_numeric(str(saham_lalu_val).replace(',', ''), errors='coerce') or 0
    
    # Gabungkan semua kode broker unik dalam kelompok ini
    brokers = ", ".join(sorted(set(str(i) for i in group['Kode Broker'] if pd.notna(i) and str(i).strip() != '')))
    
    summary_rows.append({
        'SAHAM': saham,
        'PEMEGANG SAHAM': pemegang,
        'TOTAL_SAHAM': total_saham,
        'TOTAL_SAHAM_LALU': total_saham_lalu,
        'TOTAL_PERUBAHAN': total_perubahan,
        'BROKERS': brokers
    })

df_summary = pd.DataFrame(summary_rows)

# Fungsi Format Angka Porto (K, M, B, T)
def format_number_custom(val):
    try:
        num = float(val)
        if num >= 1e12:
            return f"{num/1e12:.1f}T".replace('.0', '')
        elif num >= 1e9:
            return f"{num/1e9:.1f}B".replace('.0', '')
        elif num >= 1e6:
            return f"{num/1e6:.1f}M".replace('.0', '')
        elif num >= 1e3:
            return f"{num/1e3:.1f}K".replace('.0', '')
        else:
            return str(int(num))
    except:
        return "0"

# 4. Susun dan Format ke File Excel Baru (Baris Bertingkat)
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Rekap 5 Persen"

headers = ["SAHAM", "PEMEGANG SAHAM", "PORTO", "CHANGE %"]
ws.append(headers)

for idx, row in df_summary.iterrows():
    saham = row['SAHAM']
    pemegang = row['PEMEGANG SAHAM']
    total_saham = row['TOTAL_SAHAM']
    total_saham_lalu = row['TOTAL_SAHAM_LALU']
    total_perubahan = row['TOTAL_PERUBAHAN']
    brokers = row['BROKERS']
    
    porto_str = format_number_custom(total_saham)
    
    # Hitung CHANGE % = (perubahan / porto sebelum) * 100 dengan 4 angka di belakang koma
    if total_saham_lalu > 0:
        change_pct = (total_perubahan / total_saham_lalu) * 100
        change_str = f"{change_pct:.4f}%".replace('.', ',')
    else:
        change_str = "0,0000%"
        
    # Baris 1: Kode Saham, Nama Pemegang Saham, Jumlah Porto Gabungan, Persentase Perubahan (4 desimal)
    ws.append([saham, pemegang, porto_str, change_str])
    
    # Baris 2: Kosong di kolom saham, daftar kode broker di bawah nama pemegang saham
    ws.append(["", brokers, "", ""])

output_filename = ROOT / "Laporan_Format_Baris_Garis.xlsx"
wb.save(output_filename)
print(f"Selesai! File berhasil diperbarui dan disimpan sebagai: {output_filename}")