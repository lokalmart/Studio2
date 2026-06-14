# Lokalmart Studio v2 - Netlify Direct Upload

Paket ini dibuat khusus agar Netlify tidak 404 ketika diupload langsung: `index.html` ada di root ZIP.

Catatan penting:
- Drag/drop ZIP biasanya hanya aman untuk static frontend.
- Fitur Odoo membutuhkan Netlify Functions di `netlify/functions/odoo.js` dan dependency `xlsx`.
- Agar semua fungsi API berjalan stabil, deploy lewat GitHub integration atau Netlify CLI.

Tes halaman:
- `/`
- `/assistant`
- `/#assistant`

Tes API setelah functions benar-benar terdeploy:
- `/api/odoo`
- `/.netlify/functions/odoo`
