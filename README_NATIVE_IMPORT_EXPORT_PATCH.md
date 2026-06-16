# Studio2 Native Import Export Patch v5

Patch ini mengganti Studio2 menjadi workflow yang lebih fokus dan lebih “native” untuk Odoo:

- Menu utama hanya **Import** dan **Export**.
- Semua proses melewati: **Preview → Editor → Validasi → Eksekusi/Download**.
- Panel/card yang tidak sedang difokuskan otomatis mengecil setelah file/export terbuka.
- Sheet context/helper seperti `task_hierarchy`, `chatter_project`, `chatter_tasks`, `prompt`, `README_AI`, `Dashboard`, dan `validation_report` tidak dianggap error import.
- Sheet hasil export project sekarang memakai sheet model yang lebih aman:
  - `project.project`
  - `project.task`
  - `project.milestone`
  - `project.update`
  - `ir.model.data`
  - context sheet tetap dipisah.
- Editor membaca schema Odoo via action `model_fields`, sehingga field dapat ditampilkan menurut tipe:
  - boolean checkbox
  - selection dropdown
  - date/datetime input
  - text/html textarea
  - many2one/many2many dengan hint relation
- Editor khusus tetap ada untuk:
  - Product
  - Contact
  - Project/Task/Milestone/Update
  - Knowledge
  - Sales
  - Generic dynamic model
- Backend import tetap memakai cache login/model/fields/xmlid per request agar lebih stabil di Vercel.
- `image_url` tetap tidak otomatis diupload ke `image_1920` kecuali `import_images=true`, supaya import tidak berat.
- Import backend sekarang bisa membaca nilai Many2one/M2M hasil export Odoo seperti `[73,"Project Name"]` dan mengambil ID-nya.

## File yang diganti

Replace file berikut di repo `lokalmart/Studio2`:

```text
index.html
api/odoo.js
package.json
vercel.json
assistant.html
assistant/index.html
README_NATIVE_IMPORT_EXPORT_PATCH.md
```

## Cara pasang

1. Extract ZIP patch.
2. Replace file repo `lokalmart/Studio2` dengan isi patch ini.
3. Commit ke GitHub.
4. Redeploy Vercel.
5. Buka `/api/odoo` dan pastikan action baru muncul: `model_fields` dan `name_search`.
6. Buka `/` dan coba flow:
   - Import XLSX → Preview → Editor → Validasi → Import.
   - Export Project → pilih project → Export ke editor → Download hasil edit / Buat import patch.

## Catatan penting

- Untuk export project, sheet chatter dan hierarchy adalah konteks untuk ChatGPT, bukan data import.
- Untuk membuat XLSX yang siap diimport ulang, gunakan tombol **Buat import patch**.
- Kalau Odoo Online lambat, import tetap lebih aman per 120–250 row/request.
- Jangan import foto massal lewat Vercel kecuali benar-benar perlu; foto produk lebih baik divalidasi di editor lalu diupload bertahap.
