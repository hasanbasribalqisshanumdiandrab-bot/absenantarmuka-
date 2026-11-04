# Dashboard Jadwal & Absen Toko

Web app Google Apps Script (jadwal shift 19 toko + foto absen ke Drive).

## File
- `Kode.gs` — backend (baca jadwal, upload foto, log)
- `Index.html` — tampilan
- `appsscript.json` — pengaturan proyek

## Deploy
1. Push ketiga file ke proyek Apps Script (clasp: `clasp push`, atau salin manual).
2. Jalankan fungsi `setupSekali` sekali di editor (otorisasi + folder foto).
3. Deploy > Manage deployments > Edit > Version: New version > Deploy.
4. Jika foto baru tidak tampil, ubah `SHARE_PER_FILE` ke `true` di `CONFIG` (Kode.gs).

Catatan: buka dari link Apps Script `/exec`, bukan dari GitHub Pages
(`google.script.run` hanya jalan di domain Apps Script).
