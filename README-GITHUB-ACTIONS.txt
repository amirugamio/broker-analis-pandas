BROKER ANALIS — GITHUB ACTIONS SAJA (SUMBER DIUNDUH OTOMATIS DARI IDX)

APA YANG DIUPLOAD KE GITHUB
Upload hanya isi folder git_only ini: index.html, worker.js, folder scripts, folder .github, dan folder data (boleh kosong). File Excel sumber TIDAK perlu diupload.

SUMBER DATA OTOMATIS
1. Kepemilikan saham >5%: mengambil link XLSX terbaru dari halaman resmi IDX:
   https://www.idx.co.id/id/perusahaan-tercatat/data-kepemilikan-saham/
2. Kode/nama broker: mengambil daftar IDFirm/FirmName dari endpoint publik IDX GetBrokerSummary.
3. GitHub Actions menjalankan proses_broker.py (pandas) dan export_json.py. Hasil akhir disimpan sebagai data/ownership.json.

LANGKAH
1. Upload isi folder ini ke repository GitHub proyek Anda (bukan folder pembungkus git_only).
2. Buka Actions > Download IDX and process ownership with pandas > Run workflow.
3. Pastikan workflow selesai hijau dan file data/ownership.json terbentuk.
4. Ambil URL raw JSON, format:
   https://raw.githubusercontent.com/USERNAME/REPOSITORY/BRANCH/data/ownership.json
5. Di Cloudflare Worker broker-summary, tambahkan variable OWNERSHIP_JSON_URL dengan URL tersebut lalu deploy Worker.
   Worker yang disertakan sudah memiliki handler untuk membaca URL itu pada endpoint /api/idx-ownership?stock=KODE.

CATATAN
- Tidak perlu Render/server Python terpisah.
- Workflow berjalan manual dan otomatis hari kerja pukul 11:30 WIB. IDX dapat membatasi akses/merubah struktur; jika gagal, buka tab Actions untuk melihat log.
- Hanya data/ownership.json yang di-commit sebagai hasil proses. File Excel yang diunduh disimpan sementara di runner GitHub Actions, bukan ikut diupload ke repository.
- Skrip tetap menggunakan rumus CHANGE % dan pemetaan broker dari proses_broker.py berbasis pandas.
- Jika repository private, Worker tidak bisa mengambil raw JSON tanpa autentikasi. Gunakan repository public khusus data non-rahasia atau atur mekanisme akses yang sesuai.
