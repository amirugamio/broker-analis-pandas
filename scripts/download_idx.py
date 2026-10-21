from __future__ import annotations

import re
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
from curl_cffi import requests

ROOT = Path(__file__).resolve().parents[1]
OWNERSHIP_OUT = ROOT / 'peng-lima-persen.xlsx'
BROKER_OUT = ROOT / 'Ringkasan Broker.xlsx'
OWNERSHIP_PAGE = 'https://www.idx.co.id/id/perusahaan-tercatat/data-kepemilikan-saham/'
BROKER_API = 'https://www.idx.co.id/primary/TradingSummary/GetBrokerSummary'
HEADERS = {
    'accept': 'text/html,application/xhtml+xml,application/json,text/plain,*/*',
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/130.0.0.0 Safari/537.36',
    'referer': 'https://www.idx.co.id/',
}


def download_ownership() -> None:
    print('Mencari file kepemilikan saham >5% terbaru dari halaman IDX...')
    response = requests.get(OWNERSHIP_PAGE, headers=HEADERS, impersonate='chrome', timeout=45)
    response.raise_for_status()
    html = response.text
    # IDX publishes direct Excel links with names such as peng-2026-10-19-00098-lima-persen.xlsx.
    matches = set(re.findall(r'https?://[^\"\'<>\s]+/Media/[^\"\'<>\s]+?lima-persen\.xlsx', html, flags=re.I))
    matches.update('https://www.idx.co.id' + x for x in re.findall(r'(/Media/[^\"\'<>\s]+?lima-persen\.xlsx)', html, flags=re.I))
    if not matches:
        # Some IDX pages only expose media URLs in embedded JSON; inspect all encoded URL fragments too.
        matches.update(re.findall(r'(?:https?:\\?/\\?/www\.idx\.co\.id)?\\?/Media\\?/[^\"\'<>\s]+?lima-persen\.xlsx', html, flags=re.I))
        matches = {x.replace('\\/', '/') for x in matches}
        matches = {x if x.startswith('http') else 'https://www.idx.co.id' + (x if x.startswith('/') else '/' + x) for x in matches}
    if not matches:
        raise RuntimeError('Tidak menemukan tautan Excel lima-persen.xlsx di halaman IDX. Kemungkinan struktur halaman berubah atau perlu endpoint API baru.')

    def extract_date(url: str) -> str:
        m = re.search(r'peng-(\d{4}-\d{2}-\d{2})-', url, flags=re.I)
        return m.group(1) if m else '0000-00-00'

    latest_url = max(matches, key=extract_date)
    print(f'File kepemilikan terbaru: {latest_url}')
    file_response = requests.get(latest_url, headers={**HEADERS, 'referer': OWNERSHIP_PAGE}, impersonate='chrome', timeout=90)
    file_response.raise_for_status()
    if not file_response.content[:2] == b'PK':
        raise RuntimeError(f'Respons unduhan kepemilikan bukan file XLSX/ZIP (HTTP {file_response.status_code}).')
    OWNERSHIP_OUT.write_bytes(file_response.content)
    print(f'Tersimpan: {OWNERSHIP_OUT.name} ({len(file_response.content):,} bytes)')


def download_broker_codes() -> None:
    print('Mengambil daftar kode dan nama broker dari endpoint Ringkasan Broker IDX...')
    today = date.today()
    # Try today first, then the previous 10 calendar days to handle weekends/holidays.
    for offset in range(0, 11):
        day = today - timedelta(days=offset)
        params = {'length': 9999, 'start': 0, 'date': day.strftime('%Y%m%d')}
        response = requests.get(BROKER_API, params=params, headers={**HEADERS, 'referer': 'https://www.idx.co.id/id/data-pasar/ringkasan-perdagangan/ringkasan-broker'}, impersonate='chrome', timeout=45)
        response.raise_for_status()
        try:
            payload = response.json()
        except Exception:
            continue
        rows = payload.get('data') if isinstance(payload, dict) else None
        if isinstance(rows, dict):
            rows = rows.get('data')
        if not isinstance(rows, list) or not rows:
            continue
        frame = pd.DataFrame(rows)
        code_col = next((c for c in ['IDFirm', 'Kode Perusahaan', 'BrokerCode'] if c in frame.columns), None)
        name_col = next((c for c in ['FirmName', 'Nama Perusahaan', 'BrokerName'] if c in frame.columns), None)
        if not code_col or not name_col:
            raise RuntimeError('Respons IDX broker tidak memiliki kolom kode dan nama broker yang dikenali.')
        brokers = frame[[code_col, name_col]].copy()
        brokers.columns = ['Kode Perusahaan', 'Nama Perusahaan']
        brokers = brokers.dropna().drop_duplicates(subset=['Kode Perusahaan']).sort_values('Kode Perusahaan')
        if brokers.empty:
            continue
        brokers.to_excel(BROKER_OUT, index=False)
        print(f'Daftar broker diambil untuk tanggal {day.isoformat()}: {len(brokers)} broker -> {BROKER_OUT.name}')
        return
    raise RuntimeError('IDX tidak mengembalikan data broker untuk hari ini maupun 10 hari sebelumnya.')


if __name__ == '__main__':
    try:
        download_ownership()
        download_broker_codes()
    except Exception as exc:
        print(f'GAGAL DOWNLOAD IDX: {exc}', file=sys.stderr)
        raise
