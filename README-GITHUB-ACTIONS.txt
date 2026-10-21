BROKER ANALIS — GITHUB ACTIONS SAJA (PAKET RINGKAS)

FILE YANG PERLU DI-UPLOAD
- .github/workflows/process-pandas.yml
- scripts/download_idx.py
- scripts/proses_broker.py
- scripts/export_json.py
- README-GITHUB-ACTIONS.txt

Tidak perlu upload file Excel IDX. File sumber diunduh otomatis oleh workflow setiap hari.
Tidak perlu upload Laporan_Format_Baris_Garis.xlsx; file ini dibuat ulang oleh proses pandas.

JADWAL
- Workflow berjalan setiap hari pukul 18.00 WIB (11.00 UTC).
- Bisa dijalankan manual lewat GitHub > Actions > Download IDX and process ownership > Run workflow.
- Push perubahan ke scripts/** atau workflow juga memicu proses.

ALUR
1. scripts/download_idx.py membuka halaman Data Kepemilikan Saham IDX, mencari tautan Excel kategori lima-persen terbaru, lalu mengunduhnya.
2. Skrip yang sama mengambil daftar kode/nama broker dari endpoint Ringkasan Broker IDX. Jika tanggal hari ini belum tersedia (misalnya hari libur), dicoba mundur sampai 10 hari kalender.
3. scripts/proses_broker.py menghitung ringkasan dengan pandas dan menghasilkan Laporan_Format_Baris_Garis.xlsx.
4. scripts/export_json.py mengubah workbook menjadi data/ownership.json.
5. GitHub Actions hanya melakukan commit data/ownership.json; file Excel sumber dan hasil proses tidak ikut di-commit.

PENTING
- Jadwal cron GitHub Actions memakai UTC dan bisa terlambat beberapa menit; 11.00 UTC setara 18.00 WIB.
- IDX dapat mengubah HTML, endpoint, format file, atau proteksi akses. Jika itu terjadi, job akan gagal dengan pesan yang jelas di log Actions; cek log langkah "Download latest source files from IDX".
- Endpoint broker yang dipakai adalah endpoint publik halaman Ringkasan Broker IDX. Data kode/nama yang dikembalikan bergantung pada respons IDX untuk tanggal yang tersedia.
- Workflow ini menghasilkan JSON di repository. Agar dashboard mengambil JSON tersebut, Cloudflare Worker harus dikonfigurasi untuk membaca URL raw GitHub dan route ownership harus menggunakan JSON itu. Paket ringkas ini tidak mengubah atau men-deploy Worker secara otomatis.
- Jika repository private, Worker tidak bisa membaca raw JSON tanpa autentikasi tambahan. Gunakan repository publik atau tambahkan mekanisme token pada Worker.
