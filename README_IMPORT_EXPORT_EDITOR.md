# Studio2 Import Export Editor Patch

Patch ini menyederhanakan Studio2 menjadi dua fitur utama saja:

1. **Import XLSX**
2. **Export XLSX**

Semua proses masuk ke alur:

**Preview → Editor → Validasi → Eksekusi / Download**

## File yang diganti / ditambahkan

Replace atau upload file berikut ke repo `lokalmart/Studio2`:

- `index.html`
- `api/odoo.js`
- `vercel.json`
- `package.json`
- `assistant.html`
- `assistant/index.html`

`assistant.html` dan `assistant/index.html` hanya redirect ke homepage, supaya route lama tidak menampilkan fitur Asisten lagi.

## Perubahan UI

- Homepage hanya punya dua tab bawah: **Import** dan **Export**.
- Target Odoo tetap ada di tombol gear, disimpan di `localStorage` browser.
- Import XLSX wajib dibaca dulu di browser sebelum dikirim ke Odoo.
- Export dari Odoo langsung dibuka sebagai workbook XLSX editor sebelum didownload.

## Editor per jenis sheet

Auto-detect berdasarkan nama sheet, `_model`, dan kolom:

- `product.template` / `product.product` → **Product Editor**
  - validasi nama produk
  - harga
  - foto / URL foto
  - vendor / supplier
  - barcode
  - deskripsi
- `res.partner` / contacts → **Contact Editor**
  - nama kontak
  - telepon / WA
  - email
  - alamat
  - role Lokalmart
  - supplier/customer rank
- `project.project` / `project.task` → **Project Editor**
  - judul project/task
  - `_external_id`
  - `__action`
  - `project_id_external_id`
  - `parent_id_external_id`
  - `stage_id_external_id`
  - deadline
  - deskripsi
- `knowledge.article` → **Knowledge Editor**
- Sheet lain → **Generic XLSX Editor**

## Catatan import foto

Backend tetap dibuat aman: `image_url` tidak otomatis diupload ke `image_1920` kecuali `options.import_images=true`. UI default mengirim `import_images:false` supaya import tidak lambat/timeout.

Untuk katalog produk pihak lain, simpan foto sebagai `x_source_image_url` atau `image_url` untuk referensi internal dulu. Jangan publish foto tanpa izin.

## Rekomendasi deploy Vercel

1. Replace file di repo GitHub.
2. Commit.
3. Redeploy Vercel.
4. Buka `/api/odoo`; harus muncul JSON `Lokalmart Studio2 Import Export API`.
5. Buka `/`; UI harus hanya menampilkan Import dan Export.

## Catatan backend

Backend masih mempertahankan beberapa action lama untuk kompatibilitas internal, tetapi UI tidak lagi menampilkan Scan, Barcode, atau Asisten. Ini sengaja agar file lama tidak langsung rusak kalau ada fungsi yang masih dipakai secara tidak langsung.
