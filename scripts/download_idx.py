from __future__ import annotations

import re
import sys
import time
from datetime import date, timedelta
from pathlib import Path
from urllib.parse import urljoin

import pandas as pd
from curl_cffi import requests

ROOT = Path(__file__).resolve().parents[1]
OWNERSHIP_OUT = ROOT / "peng-lima-persen.xlsx"
BROKER_OUT = ROOT / "Ringkasan Broker.xlsx"

BASE_URL = "https://www.idx.co.id/"
OWNERSHIP_PAGE = "https://www.idx.co.id/id/perusahaan-tercatat/data-kepemilikan-saham/"
BROKER_PAGE = "https://www.idx.co.id/id/data-pasar/ringkasan-perdagangan/ringkasan-broker"
BROKER_API = "https://www.idx.co.id/primary/TradingSummary/GetBrokerSummary"

HEADERS = {
    "accept": "text/html,application/xhtml+xml,application/json,text/plain,*/*",
    "accept-language": "id-ID,id;q=0.9,en-US;q=0.8,en;q=0.7",
    "user-agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/130.0.0.0 Safari/537.36"
    ),
    "referer": BASE_URL,
    "sec-fetch-dest": "document",
    "sec-fetch-mode": "navigate",
    "sec-fetch-site": "same-origin",
}


def make_session():
    """Gunakan satu sesi untuk mempertahankan cookie yang diberikan IDX secara normal."""
    session = requests.Session(impersonate="chrome")
    session.headers.update(HEADERS)
    return session


def warm_session(session) -> None:
    """Buka halaman utama terlebih dahulu agar sesi/cookie normal tersedia."""
    response = session.get(BASE_URL, timeout=30)
    if response.status_code == 403:
        raise RuntimeError(
            "IDX mengembalikan HTTP 403 pada halaman utama. "
            "Akses otomatis dari runner GitHub Actions sedang ditolak; "
            "ini bukan kesalahan parsing Excel."
        )
    response.raise_for_status()


def extract_ownership_links(html: str) -> set[str]:
    # URL dapat muncul sebagai URL absolut, /Media/..., atau JSON dengan slash ter-escape.
    normalized = html.replace("\\/", "/").replace("&amp;", "&")
    patterns = [
        r'https?://[^"\'<>\s]+?/Media/[^"\'<>\s]+?lima-persen\.xlsx(?:\?[^"\'<>\s]*)?',
        r'(?<![A-Za-z0-9])(/Media/[^"\'<>\s]+?lima-persen\.xlsx(?:\?[^"\'<>\s]*)?)',
    ]
    found: set[str] = set()
    for pattern in patterns:
        for match in re.findall(pattern, normalized, flags=re.IGNORECASE):
            found.add(urljoin(BASE_URL, match))
    return found


def extract_date(url: str) -> str:
    match = re.search(r"peng-(\d{4}-\d{2}-\d{2})-", url, flags=re.IGNORECASE)
    return match.group(1) if match else "0000-00-00"


def download_ownership(session) -> None:
    print("Membuka halaman resmi Data Kepemilikan Saham IDX...")
    response = session.get(
        OWNERSHIP_PAGE,
        headers={"referer": BASE_URL},
        timeout=45,
    )
    if response.status_code == 403:
        raise RuntimeError(
            "HTTP 403 saat membuka halaman kepemilikan IDX. "
            "IDX menolak request dari runner GitHub Actions. "
            "Skrip tidak akan mengganti data lama atau mengunduh dari sumber tidak resmi."
        )
    response.raise_for_status()

    links = extract_ownership_links(response.text)
    if not links:
        raise RuntimeError(
            "Halaman IDX berhasil dibuka, tetapi tautan XLSX kategori lima-persen "
            "tidak ditemukan pada HTML. Periksa apakah IDX memuat daftar melalui "
            "endpoint JavaScript/API yang berbeda."
        )

    latest_url = max(links, key=extract_date)
    print(f"Tautan berkas terbaru yang ditemukan: {latest_url}")

    file_response = session.get(
        latest_url,
        headers={"referer": OWNERSHIP_PAGE, "accept": "*/*"},
        timeout=90,
    )
    if file_response.status_code == 403:
        raise RuntimeError(
            "HTTP 403 saat mengunduh file XLSX kepemilikan dari IDX. "
            "Link ditemukan, tetapi server menolak unduhan dari runner GitHub Actions."
        )
    file_response.raise_for_status()
    if not file_response.content.startswith(b"PK"):
        raise RuntimeError(
            "Respons unduhan kepemilikan bukan file XLSX yang valid; "
            "file lama dipertahankan."
        )

    # Validasi workbook sebelum mengganti sumber terakhir yang berhasil.
    import io
    pd.ExcelFile(io.BytesIO(file_response.content), engine="openpyxl")
    OWNERSHIP_OUT.write_bytes(file_response.content)
    print(f"OK: {OWNERSHIP_OUT.name} ({len(file_response.content):,} bytes)")


def download_broker_codes(session) -> None:
    print("Mengambil data tabel Ringkasan Broker IDX (data yang dipakai tombol Unduh)...")
    session.headers.update({"referer": BROKER_PAGE, "accept": "application/json, text/plain, */*"})

    today = date.today()
    last_error = None
    for offset in range(0, 11):
        day = today - timedelta(days=offset)
        params = {"length": 9999, "start": 0, "date": day.strftime("%Y%m%d")}
        try:
            response = session.get(BROKER_API, params=params, timeout=45)
            if response.status_code == 403:
                raise RuntimeError(
                    "HTTP 403 dari endpoint tabel Ringkasan Broker IDX. "
                    "IDX menolak request dari runner GitHub Actions."
                )
            response.raise_for_status()
            payload = response.json()
        except Exception as exc:
            last_error = exc
            # Do not repeatedly hammer an endpoint that is actively returning 403.
            if "403" in str(exc):
                raise
            print(f"Tidak dapat membaca data broker tanggal {day}: {exc}")
            time.sleep(1)
            continue

        rows = payload.get("data") if isinstance(payload, dict) else None
        if isinstance(rows, dict):
            rows = rows.get("data")
        if not isinstance(rows, list) or not rows:
            continue

        frame = pd.DataFrame(rows)
        code_col = next((c for c in ["IDFirm", "Kode Perusahaan", "BrokerCode"] if c in frame.columns), None)
        name_col = next((c for c in ["FirmName", "Nama Perusahaan", "BrokerName"] if c in frame.columns), None)
        if not code_col or not name_col:
            raise RuntimeError(
                "Respons tabel broker tidak memiliki kolom kode/nama broker yang dikenali. "
                f"Kolom yang diterima: {list(frame.columns)}"
            )

        brokers = frame.copy()
        # Keep the same columns as the IDX Download button's exported workbook.
        output_cols = [
            (code_col, "Kode Perusahaan"),
            (name_col, "Nama Perusahaan"),
        ]
        for possible, output_name in [
            ("Volume", "Volume"),
            ("Value", "Nilai"),
            ("Frequency", "Frekuensi"),
        ]:
            if possible in brokers.columns:
                output_cols.append((possible, output_name))
        brokers = brokers[[source for source, _ in output_cols]].copy()
        brokers.columns = [dest for _, dest in output_cols]
        brokers = brokers.dropna(subset=["Kode Perusahaan", "Nama Perusahaan"])
        brokers["Kode Perusahaan"] = brokers["Kode Perusahaan"].astype(str).str.strip()
        brokers["Nama Perusahaan"] = brokers["Nama Perusahaan"].astype(str).str.strip()
        brokers = brokers[brokers["Kode Perusahaan"].ne("")].drop_duplicates(subset=["Kode Perusahaan"])
        if brokers.empty:
            continue

        # Write to a temporary file first so a failed export never destroys the last good file.
        temp = BROKER_OUT.with_suffix(".tmp.xlsx")
        brokers.to_excel(temp, index=False)
        temp.replace(BROKER_OUT)
        print(f"OK: data broker tanggal {day.isoformat()}, {len(brokers)} broker -> {BROKER_OUT.name}")
        return

    raise RuntimeError(
        "IDX tidak mengembalikan data broker yang valid untuk hari ini hingga 10 hari sebelumnya. "
        f"Detail terakhir: {last_error}"
    )


if __name__ == "__main__":
    try:
        session = make_session()
        warm_session(session)
        download_ownership(session)
        download_broker_codes(session)
    except Exception as exc:
        print(f"GAGAL DOWNLOAD IDX: {exc}", file=sys.stderr)
        raise
