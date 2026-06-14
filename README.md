# Lokalmart Studio v3 — Clean Vercel Rebuild

Studio ini dibuat ulang dari nol khusus untuk Vercel. Tujuannya: ringan, mudah dipahami, dan tidak bergantung pada banyak file `lib` yang bisa membuat seluruh API mati kalau satu file rusak.

## Struktur

```text
/
├─ index.html              # UI utama mobile-first
├─ assistant.html          # route /assistant -> /#assistant
├─ assistant/index.html    # fallback route fisik
├─ api/odoo.js             # satu endpoint backend Vercel
├─ package.json
├─ vercel.json
└─ scripts/check-html.js
```

## Fitur utama

- Target Odoo tersimpan di browser/localStorage.
- Tes koneksi memakai isi form yang sedang dibuka, bukan target lama.
- Import XLSX generic ke Odoo:
  - `_model`
  - `__action`
  - `_external_id`
  - Many2one: `field_external_id`
  - Many2many: `field_external_ids`
  - `image_url` otomatis masuk ke `image_1920` jika field tersedia.
- Scan schema model dan custom field `x_`.
- Audit data model inti.
- Export context umum untuk ChatGPT.
- Export satu project terpilih:
  - `project.project`
  - `project.task`
  - task/subtask hierarchy
  - `project.milestone`
  - `project.update`
  - chatter ringkas dari `mail.message`
  - external IDs
  - prompt siap copy ke ChatGPT.
- Full/partial export JSON atau XLSX.
- Barcode lookup.

## Deploy ke Vercel

1. Replace isi repo dengan file-file ini.
2. Commit ke GitHub.
3. Hubungkan Vercel ke repo.
4. Setting Vercel:

```text
Root Directory: kosong / root repo
Build Command: kosong
Output Directory: kosong
Install Command: npm install
```

5. Redeploy.
6. Tes:

```text
/
/assistant
/api/odoo
```

`/api/odoo` harus menampilkan JSON `Lokalmart Studio v3 Vercel API` ketika dibuka via browser.

## Odoo Online Login

Untuk Odoo Online, sering kali password login web tidak bisa dipakai oleh external API. Gunakan API Key dari Odoo sebagai pengganti password.

Target contoh:

```text
URL: https://edu-lokalmart.odoo.com
Database: edu-lokalmart
Username: email user admin di database Odoo
Password / API Key: API key Odoo
```

Jangan pakai URL yang berakhir `/web`.

## Catatan penting

Versi ini sengaja memakai satu API file: `api/odoo.js`. Jangan campur dengan folder `lib/` lama atau file Netlify lama jika tujuannya deploy di Vercel.
