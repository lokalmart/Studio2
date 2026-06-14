# Studio2 Vercel Fix

Patch ini mengubah repo Studio2 agar cocok untuk Vercel.

Masalah yang diperbaiki:
- Repo sebelumnya masih paket Netlify direct upload.
- Tidak ada `api/odoo.js`, sehingga `/api/odoo` di Vercel tidak hidup.
- Tidak ada `vercel.json`, sehingga route `/assistant` tidak di-rewrite.
- `netlify/functions/odoo.js` dan `_redirects` tidak dipakai oleh Vercel.

Cara pakai:
1. Copy semua file/folder patch ini ke root repo `lokalmart/Studio2`.
2. Commit ke branch main.
3. Di Vercel, hubungkan ke repo ini.
4. Build command kosong saja.
5. Output directory kosong saja.
6. Deploy ulang.

Tes:
- `/`
- `/assistant`
- `/api/odoo`

Untuk Odoo Online, field Password/API Key sebaiknya diisi API Key Odoo, bukan password login web.
