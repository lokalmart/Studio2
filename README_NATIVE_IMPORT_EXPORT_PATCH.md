# Studio2 Native Import Export Patch v6

Patch ini menyempurnakan Studio2 agar lebih dekat ke kebutuhan all-in-one import/export Odoo Lokalmart.

## Perbaikan utama v6

1. **Bug editor res.partner diperbaiki**
   - Sheet/model `res.partner` sekarang selalu masuk **Contact Editor**.
   - Deteksi editor sekarang memprioritaskan model Odoo terlebih dahulu, baru membaca kolom XLSX.
   - Jadi kolom seperti `default_code` atau field produk lain tidak akan membuat `res.partner` keliru masuk Product Editor.

2. **Error Excel 32767 karakter diperbaiki**
   - Export XLSX sekarang memotong cell panjang ke batas aman.
   - Kolom `_studio2_truncated_fields` otomatis ditambahkan jika ada field yang dipotong.
   - Ini mengatasi error seperti: `Text length must not exceed 32767 characters`.
   - Biasanya sumbernya dari `body`, `body_html`, `description`, chatter, atau HTML panjang dari Odoo.

3. **Export model sekarang bisa scan dan pilih record**
   - Buka Export → Model.
   - Isi model, contoh: `res.partner`, `product.template`, `project.task`.
   - Klik **Scan record**.
   - Pilih record satu per satu, pilih semua terlihat, atau kosongkan pilihan.
   - Klik **Export record terpilih**.
   - Hasilnya tetap masuk editor XLSX dulu sebelum didownload.

4. **Backend action baru**
   - `record_scan`: membaca daftar record ringan untuk dipilih.
   - `selected_export`: export hanya record yang dipilih.

5. **Tetap mempertahankan workflow utama**
   - Import: Upload → Preview → Editor → Validasi → Import.
   - Export: Pilih sumber → Scan/Pilih → Editor → Download/Patch.

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
5. Buka `/api/odoo` dan pastikan action baru muncul:
   - `model_fields`
   - `name_search`
   - `record_scan`
   - `selected_export`
6. Buka `/`, lalu tes:
   - Export → Model → model `res.partner` → Scan record → pilih beberapa record → Export record terpilih.
   - Pastikan sheet `res.partner` terbuka sebagai Contact Editor, bukan Product Editor.

## Catatan teknis

- Export cepat model tetap ada untuk sample cepat, tetapi workflow yang disarankan adalah scan dan pilih record.
- Cell panjang tidak dibuang total, hanya dipotong agar XLSX valid. Untuk backup full HTML/body yang sangat panjang, lebih baik export JSON terpisah.
- Import foto massal tetap tidak direkomendasikan lewat Vercel. Validasi foto boleh di editor, upload foto sebaiknya bertahap.
