BROKER ANALIS — GITHUB ACTIONS SAJA (PAKET RINGKAS v14)

FILE YANG PERLU DI-UPLOAD KE ROOT REPOSITORY
- .github/workflows/process-pandas.yml
- scripts/download_idx.py
- scripts/proses_broker.py
- scripts/export_json.py
- README-GITHUB-ACTIONS.txt
- .gitignore

Jangan upload file Excel sumber IDX. File sumber diunduh saat workflow berjalan.
Laporan_Format_Baris_Garis.xlsx dibuat ulang oleh proses pandas dan tidak perlu di-commit.

JADWAL
- Setiap hari pukul 18.00 WIB (11.00 UTC).
- Bisa dijalankan manual lewat GitHub > Actions > Download IDX and process ownership > Run workflow.
- Push perubahan ke scripts/** atau workflow juga memicu proses.

PERBAIKAN PADA v14
- Menggunakan satu sesi HTTP untuk mempertahankan cookie yang diberikan IDX saat membuka halaman utama dan halaman data.
- Mengambil tautan file kepemilikan dari halaman resmi IDX, lalu memvalidasi file XLSX sebelum mengganti file lokal.
- Mengambil tabel ringkasan broker dari endpoint data yang digunakan halaman Ringkasan Broker, dengan tanggal hari ini lalu mundur maksimal 10 hari kalender.
- Jika request ditolak, data sumber lama tidak ditimpa oleh halaman error/HTML.
- Pesan error HTTP 403 dibuat lebih jelas.

BATASAN PENTING
IDX dapat menolak akses otomatis dari alamat IP runner GitHub Actions dengan HTTP 403. Penggunaan sesi/cookie normal membantu jika masalahnya berasal dari sesi, tetapi TIDAK menjamin 403 teratasi bila IP runner diblokir oleh IDX. Skrip ini tidak mencoba melewati CAPTCHA atau kontrol akses. Jika tetap 403, unduhan otomatis langsung dari GitHub-hosted runner tidak dapat diandalkan; opsi berikutnya adalah memakai runner komputer sendiri yang menjalankan browser/sesi resmi, atau memakai kanal data resmi yang diizinkan.

ALUR
1. Download source IDX.
2. Proses sumber dengan scripts/proses_broker.py (logika perhitungan dipertahankan).
3. Ekspor Laporan_Format_Baris_Garis.xlsx menjadi data/ownership.json.
4. Commit hanya data/ownership.json.

JADWAL CRON GITHUB ACTIONS
Cron memakai UTC dan bisa terlambat beberapa menit. 11:00 UTC = 18:00 WIB.

KONEKSI DASHBOARD
Paket ini menghasilkan JSON di repository. Agar dashboard membaca data baru, Cloudflare Worker tetap perlu dikonfigurasi membaca URL raw GitHub. Paket ini tidak mengubah atau men-deploy Worker.
