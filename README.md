# Dashboard Status Relokasi APJATEL — Frontend (Vercel)

Frontend statis. Backend = Google Apps Script (Web App), database = Google Spreadsheet.

## Struktur
- `index.html`      — aplikasi (hasil konversi dari Dashboard.html)
- `js/config.js`    — **satu-satunya file yang perlu diedit**: isi `API_URL`
- `js/api-shim.js`  — adapter: mengubah `google.script.run.*` menjadi request HTTP ber-token
- `vercel.json`     — header keamanan & cache

## Deploy singkat
1. Edit `js/config.js` -> isi `API_URL` dengan URL Web App Apps Script (berakhiran `/exec`).
2. Push ke GitHub.
3. Vercel -> Add New -> Project -> import repo -> Framework Preset: **Other** -> Deploy.

Update berikutnya: ubah file -> commit/push -> Vercel otomatis deploy ulang.
