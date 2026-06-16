# Studio2 Fast Import Patch

Patch ini dibuat untuk mengurangi gagal/lemot ketika import XLSX Lokalmart ke Odoo lewat Vercel.

## File yang diganti

1. Replace `api/odoo.js` dengan file `api/odoo.js` dari folder ini.
2. Replace `vercel.json` dengan file `vercel.json` dari folder ini.
3. Commit ke GitHub dan redeploy di Vercel.

## Perubahan penting

- Import XLSX tidak lagi melakukan `modelExists`, `fields_get`, dan pencarian external ID berulang-ulang tanpa cache.
- Sheet bantuan seperti `README_IMPORT`, `Dashboard`, `validation_report`, `ai_memory_index`, dan sheet non-model otomatis dilewati.
- `image_url` tidak otomatis diupload ke `image_1920`, karena ini sering membuat import sangat lambat. Simpan URL foto di `x_source_image_url`, atau aktifkan `options.import_images=true` jika benar-benar ingin upload foto.
- Default import aman: 120 row per request, dengan cap 250 row. Untuk import besar bisa gunakan `force_large=true`, tapi lebih rawan timeout.
- `vercel.json` menambahkan `maxDuration: 60` untuk `api/odoo.js`. Batas sebenarnya tetap mengikuti plan/project setting Vercel.

## Rekomendasi cara import

Urutan aman untuk Odoo Lokalmart:

1. Import `ir.model` dan `ir.model.fields` dulu jika ada custom model/field.
2. Import parent/master data dulu: `project.project`, `project.task.type`, `product.category`, `res.partner`.
3. Import child data: `project.task`, `product.template`, `knowledge.article`.
4. Foto dipisahkan belakangan, jangan dicampur dengan data inti.

## Format XLSX aman

- Sheet boleh bernama model, misalnya `project.task`.
- Atau setiap row wajib punya `_model`.
- Gunakan `__action`: `upsert`, `update`, `delete`, `archive`, `skip`.
- Gunakan `_external_id` untuk record yang ingin bisa di-update ulang.
- Many2one pakai `field_external_id`.
- Many2many pakai `field_external_ids`.
- Custom field Odoo Online pakai awalan `x_`.
