"""Download the newest IDX ownership >5% workbook and all broker codes from IDX."""
from datetime import datetime, timedelta, timezone
from pathlib import Path
import re, sys, time
import requests
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
OWNERSHIP_PAGE = "https://www.idx.co.id/id/perusahaan-tercatat/data-kepemilikan-saham/"
BROKER_API = "https://www.idx.co.id/primary/TradingSummary/GetBrokerSummary"
SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/125 Safari/537.36",
    "Accept": "text/html,application/json,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,*/*",
    "Referer": "https://www.idx.co.id/",
})

def download_latest_ownership():
    # Page IDX lists the actual downloadable Excel file under /Media/...
    r = SESSION.get(OWNERSHIP_PAGE, timeout=45)
    r.raise_for_status()
    html = r.text.replace("\\u0026", "&").replace("\\/", "/")
    urls = re.findall(r'https?://[^\"\'\s<>]+/Media/[^\"\'\s<>]*?lima-persen\.xlsx|/Media/[^\"\'\s<>]*?lima-persen\.xlsx', html, re.I)
    # Some IDX pages embed JSON-escaped URLs.
    if not urls:
        urls = re.findall(r'[^\"\'\s<>]*?/Media/[^\"\'\s<>]*?lima-persen\.xlsx', html, re.I)
    urls = [u.replace('\\/', '/') for u in urls]
    # Deduplicate and prefer the newest date/sequence visible in filename.
    urls = list(dict.fromkeys(urls))
    if not urls:
        raise RuntimeError('Tidak menemukan link Excel kategori lima-persen di halaman IDX. Struktur halaman mungkin berubah.')
    def key(u):
        m = re.search(r'peng-(\d{4}-\d{2}-\d{2})-(\d+)-lima-persen\.xlsx', u, re.I)
        return (m.group(1), int(m.group(2))) if m else ('', 0)
    url = max(urls, key=key)
    if url.startswith('/'):
        url = 'https://www.idx.co.id' + url
    elif url.startswith('//'):
        url = 'https:' + url
    print('Ownership source:', url)
    data = SESSION.get(url, timeout=90)
    data.raise_for_status()
    if not data.content[:2] == b'PK':
        raise RuntimeError('Unduhan kepemilikan bukan file XLSX yang valid.')
    name = url.rsplit('/', 1)[-1].split('?', 1)[0]
    path = ROOT / name
    path.write_bytes(data.content)
    print('Saved ownership workbook:', path.name, len(data.content), 'bytes')

def download_broker_directory():
    # IDX's public GetBrokerSummary returns IDFirm and FirmName. Try recent
    # calendar dates because IDX may not publish a record on weekends/holidays.
    now = datetime.now(timezone.utc) + timedelta(hours=7)
    last_error = None
    payload = None
    for offset in range(0, 15):
        date = (now - timedelta(days=offset)).strftime('%Y%m%d')
        try:
            r = SESSION.get(BROKER_API, params={'length': 9999, 'start': 0, 'date': date}, timeout=45)
            if r.status_code != 200:
                last_error = f'HTTP {r.status_code} untuk tanggal {date}'
                continue
            body = r.json()
            rows = body.get('data') if isinstance(body, dict) else None
            if isinstance(rows, list) and rows:
                payload = rows
                print('Broker directory source date:', date, 'records:', len(rows))
                break
            last_error = f'Tidak ada data broker untuk tanggal {date}'
        except Exception as e:
            last_error = str(e)
    if not payload:
        raise RuntimeError('Gagal mengambil daftar broker dari IDX: ' + str(last_error))
    df = pd.DataFrame(payload)
    required = {'IDFirm', 'FirmName'}
    if not required.issubset(df.columns):
        raise RuntimeError('Respons IDX berubah; kolom IDFirm/FirmName tidak ditemukan.')
    out = pd.DataFrame({
        'Kode Perusahaan': df['IDFirm'].astype(str).str.strip(),
        'Nama Perusahaan': df['FirmName'].astype(str).str.strip(),
    }).drop_duplicates(subset=['Kode Perusahaan']).sort_values('Kode Perusahaan')
    if out.empty:
        raise RuntimeError('Daftar broker IDX kosong.')
    path = ROOT / 'Ringkasan Broker Terbaru.xlsx'
    out.to_excel(path, index=False)
    print('Saved broker directory:', path.name, 'records:', len(out))

if __name__ == '__main__':
    download_latest_ownership()
    download_broker_directory()
