/*
 Lokalmart Studio2 - fast import patch for Vercel
 Drop-in replacement: /api/odoo.js
 Frontend contract is preserved: POST /api/odoo { action, target, payload }

 Main fixes:
 - Cache Odoo login, model existence, fields_get, and external IDs inside one request.
 - Skip helper/non-import sheets automatically.
 - Make image_url import opt-in to avoid slow imports and large writes.
 - Cap safe rows per request unless force_large is enabled.
 - Return clearer per-sheet/per-row report.
*/
'use strict';

const XLSX = require('xlsx');

const MAX_BODY_BYTES = 22 * 1024 * 1024;
const DEFAULT_MODULE = 'lokalmart_studio';
const DEFAULT_IMPORT_ROWS = 120;
const SAFE_IMPORT_CAP = 250;
const LARGE_IMPORT_CAP = 3000;
const DEFAULT_LIMIT = 250;
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

const HELPER_SHEETS = new Set([
  'README', 'README_IMPORT', 'README IMPORT', 'DASHBOARD', 'VALIDATION', 'VALIDATION_REPORT',
  'AI_MEMORY_INDEX', 'PROMPT_HANDBOOK', 'RELATIONSHIP_MAP', 'TASK_DATABASE', 'NOTES', 'INFO'
]);

module.exports = async function handler(req, res) {
  applyCors(res);

  if (req.method === 'OPTIONS') return json(res, 200, { ok: true });

  if (req.method === 'GET') {
    return json(res, 200, {
      ok: true,
      name: 'Lokalmart Studio2 Fast Import API',
      endpoint: '/api/odoo',
      usage: 'POST JSON { action, target, payload }',
      actions: [
        'health', 'test_connection', 'schema_scan', 'data_audit', 'context_export',
        'xlsx_preview', 'import_xlsx', 'full_export', 'project_list',
        'project_context_export', 'project_xlsx_export', 'barcode_lookup', 'read_records'
      ],
      import_notes: [
        'import_xlsx now caches fields/model/xmlid per request',
        'image_url is not imported unless payload.options.import_images=true',
        `default max rows per request: ${DEFAULT_IMPORT_ROWS}`
      ]
    });
  }

  if (req.method !== 'POST') return json(res, 405, { ok: false, error: 'Method tidak didukung.' });

  try {
    const body = await readJsonBody(req);
    const action = String(body.action || '').trim();
    const target = normalizeTarget(body.target || {});
    const payload = body.payload || {};

    if (!action) throw new UserError('Action kosong.');
    if (action === 'health') return json(res, 200, { ok: true, time: new Date().toISOString() });

    const requireTargetActions = new Set([
      'test_connection', 'schema_scan', 'data_audit', 'context_export', 'xlsx_preview',
      'import_xlsx', 'full_export', 'project_list', 'project_context_export',
      'project_xlsx_export', 'barcode_lookup', 'read_records'
    ]);
    if (requireTargetActions.has(action)) validateTarget(target);

    const ctx = createOdooClient(target);

    if (action === 'test_connection') {
      const session = await ctx.login();
      const user = await ctx.execute('res.users', 'read', [[session.uid], ['id', 'name', 'login', 'company_id', 'groups_id']]);
      return json(res, 200, {
        ok: true,
        uid: session.uid,
        db: target.db,
        username: target.username,
        url: target.url,
        user: Array.isArray(user) ? user[0] : user,
        message: 'Koneksi Odoo berhasil.'
      });
    }

    if (action === 'schema_scan') return json(res, 200, { ok: true, schema: await schemaScan(ctx, payload) });
    if (action === 'data_audit') return json(res, 200, { ok: true, audit: await dataAudit(ctx, payload) });
    if (action === 'context_export') return json(res, 200, { ok: true, context: await contextExport(ctx, payload) });
    if (action === 'xlsx_preview') return json(res, 200, { ok: true, preview: xlsxPreview(payload) });
    if (action === 'import_xlsx') return json(res, 200, { ok: true, result: await importXlsx(ctx, payload) });
    if (action === 'full_export') return json(res, 200, { ok: true, export: await fullExport(ctx, payload) });
    if (action === 'project_list') return json(res, 200, { ok: true, projects: await projectList(ctx, payload) });
    if (action === 'project_context_export') return json(res, 200, { ok: true, context: await projectContextExport(ctx, payload) });
    if (action === 'project_xlsx_export') return json(res, 200, { ok: true, export: await projectXlsxExport(ctx, payload) });
    if (action === 'barcode_lookup') return json(res, 200, { ok: true, result: await barcodeLookup(ctx, payload) });
    if (action === 'read_records') return json(res, 200, { ok: true, result: await readRecords(ctx, payload) });

    throw new UserError(`Action tidak dikenal: ${action}`);
  } catch (err) {
    const status = err instanceof UserError ? 400 : 500;
    return json(res, status, {
      ok: false,
      error: err.message || String(err),
      code: err.code || err.name || 'ERROR',
      hint: err.hint || undefined,
      details: err.details || undefined
    });
  }
};

class UserError extends Error {
  constructor(message, hint, details) {
    super(message);
    this.name = 'UserError';
    this.hint = hint;
    this.details = details;
  }
}

function applyCors(res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

function json(res, status, body) {
  res.statusCode = status;
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.setHeader('Cache-Control', 'no-store');
  res.end(JSON.stringify(body));
}

function readJsonBody(req) {
  if (req.body && typeof req.body === 'object' && !Buffer.isBuffer(req.body)) return Promise.resolve(req.body);
  if (typeof req.body === 'string' && req.body.trim()) {
    try { return Promise.resolve(JSON.parse(req.body)); }
    catch (_) { return Promise.reject(new UserError('Body request bukan JSON valid.')); }
  }

  return new Promise((resolve, reject) => {
    let raw = '';
    let total = 0;
    req.on('data', chunk => {
      total += chunk.length;
      if (total > MAX_BODY_BYTES) {
        reject(new UserError('Payload terlalu besar.', 'Kurangi jumlah sheet/baris atau gunakan file XLSX yang lebih kecil.'));
        req.destroy();
        return;
      }
      raw += chunk.toString('utf8');
    });
    req.on('end', () => {
      try {
        if (!raw.trim()) return resolve({});
        resolve(JSON.parse(raw));
      } catch (e) {
        reject(new UserError('Body request bukan JSON valid.', 'Pastikan frontend mengirim Content-Type application/json.'));
      }
    });
    req.on('error', reject);
  });
}

function normalizeTarget(target) {
  let url = String(target.url || target.host || '').trim();
  url = url.replace(/\/web\/?$/i, '').replace(/\/+$/, '');
  return {
    url,
    db: String(target.db || target.database || '').trim(),
    username: String(target.username || target.email || target.login || '').trim(),
    password: String(target.password || target.apiKey || target.api_key || target.key || '').trim()
  };
}

function validateTarget(target) {
  if (!target.url || !/^https?:\/\//i.test(target.url)) {
    throw new UserError('URL Odoo tidak valid.', 'Contoh: https://edu-lokalmart.odoo.com. Jangan pakai /web di belakang URL.');
  }
  if (!target.db) throw new UserError('Database Odoo kosong.', 'Contoh: edu-lokalmart');
  if (!target.username) throw new UserError('Username/email Odoo kosong.');
  if (!target.password) throw new UserError('Password/API key kosong.', 'Untuk Odoo Online, gunakan API Key sebagai pengganti password login web.');
}

function createOdooClient(target) {
  let uidCache = null;
  let loginPromise = null;
  const modelCache = new Map();
  const fieldsCache = new Map();
  const xmlidCache = new Map();

  async function jsonRpc(service, method, args) {
    const id = Date.now() + Math.random();
    const response = await fetch(`${target.url}/jsonrpc`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ jsonrpc: '2.0', method: 'call', params: { service, method, args }, id })
    });
    const text = await response.text();
    let data;
    try { data = JSON.parse(text); }
    catch (_) {
      throw new UserError(
        `Odoo mengembalikan non-JSON HTTP ${response.status}.`,
        'Cek URL Odoo. Jangan pakai /web di belakang URL.',
        text.slice(0, 1200)
      );
    }
    if (!response.ok) throw new UserError(`Odoo HTTP ${response.status}`, undefined, data);
    if (data.error) {
      const message = data.error?.data?.message || data.error?.message || 'Odoo JSON-RPC error';
      const debug = data.error?.data?.debug || data.error;
      throw new UserError(message, 'Cek permission user, model, field, dan format data.', debug);
    }
    return data.result;
  }

  async function login() {
    if (uidCache) return { uid: uidCache };
    if (!loginPromise) {
      loginPromise = (async () => {
        let uid = await jsonRpc('common', 'authenticate', [target.db, target.username, target.password, {}]);
        if (!uid) uid = await jsonRpc('common', 'login', [target.db, target.username, target.password]);
        if (!uid) {
          throw new UserError(
            `Login Odoo gagal untuk database "${target.db}" dan user "${target.username}".`,
            'Gunakan API Key Odoo sebagai password. Pastikan database dan email user benar.'
          );
        }
        uidCache = uid;
        return { uid };
      })();
    }
    return loginPromise;
  }

  async function execute(model, method, args = [], kwargs = {}) {
    const session = await login();
    return jsonRpc('object', 'execute_kw', [target.db, session.uid, target.password, model, method, args, kwargs || {}]);
  }

  return { target, login, execute, modelCache, fieldsCache, xmlidCache };
}

async function safeExecute(ctx, model, method, args = [], kwargs = {}, fallback = null) {
  try { return await ctx.execute(model, method, args, kwargs); }
  catch (_) { return fallback; }
}

async function modelExists(ctx, model) {
  if (ctx.modelCache.has(model)) return ctx.modelCache.get(model);
  const ids = await safeExecute(ctx, 'ir.model', 'search', [[[ 'model', '=', model ]]], { limit: 1 }, []);
  const ok = Array.isArray(ids) && ids.length > 0;
  ctx.modelCache.set(model, ok);
  return ok;
}

async function fieldsGet(ctx, model) {
  if (ctx.fieldsCache.has(model)) return ctx.fieldsCache.get(model);
  const fields = await ctx.execute(model, 'fields_get', [], { attributes: ['string', 'type', 'relation', 'required', 'readonly', 'selection'] });
  ctx.fieldsCache.set(model, fields || {});
  return fields || {};
}

function parseCsvList(value, fallback = []) {
  if (Array.isArray(value)) return value.map(v => String(v).trim()).filter(Boolean);
  if (value === null || value === undefined || value === '') return fallback;
  return String(value).split(/[\n,;|]+/).map(v => v.trim()).filter(Boolean);
}

function normalizeLimit(value, fallback = DEFAULT_LIMIT, max = 2000) {
  const n = Number(value || fallback);
  if (!Number.isFinite(n) || n <= 0) return fallback;
  return Math.min(Math.floor(n), max);
}

async function schemaScan(ctx, payload = {}) {
  const defaultModels = [
    'project.project', 'project.task', 'project.milestone', 'project.update', 'knowledge.article',
    'product.template', 'product.product', 'product.category', 'res.partner', 'sale.order',
    'sale.order.line', 'ir.model', 'ir.model.fields'
  ];
  const models = parseCsvList(payload.models, defaultModels);
  const out = [];
  for (const model of models) {
    const exists = await modelExists(ctx, model);
    if (!exists) {
      out.push({ model, exists: false, fields: [], warning: 'Model tidak ditemukan atau tidak bisa diakses.' });
      continue;
    }
    const fg = await fieldsGet(ctx, model);
    const fields = Object.entries(fg)
      .map(([name, meta]) => ({ name, ...meta }))
      .sort((a, b) => a.name.localeCompare(b.name));
    out.push({ model, exists: true, field_count: fields.length, fields });
  }
  const customModels = await safeExecute(ctx, 'ir.model', 'search_read', [[[ 'model', 'like', 'x_' ]]], { fields: ['id', 'name', 'model', 'state'], limit: 200 }, []);
  return { scanned_at: new Date().toISOString(), models: out, custom_models: customModels };
}

async function dataAudit(ctx, payload = {}) {
  const models = parseCsvList(payload.models, [
    'project.project', 'project.task', 'project.milestone', 'project.update', 'knowledge.article',
    'product.template', 'product.product', 'res.partner', 'sale.order', 'sale.order.line',
    'ir.model.fields', 'ir.ui.view'
  ]);
  const counts = [];
  for (const model of models) {
    if (!(await modelExists(ctx, model))) {
      counts.push({ model, exists: false, count: null });
      continue;
    }
    const count = await safeExecute(ctx, model, 'search_count', [[]], {}, null);
    counts.push({ model, exists: true, count });
  }
  const taskFields = await safeExecute(ctx, 'project.task', 'fields_get', [], { attributes: ['type', 'relation'] }, {});
  const hasParent = !!taskFields?.parent_id;
  const topTasks = await safeExecute(
    ctx,
    'project.task',
    'search_read',
    hasParent ? [[[ 'project_id', '!=', false ], [ 'parent_id', '=', false ]]] : [[[ 'project_id', '!=', false ]]],
    { fields: ['id', 'name', 'project_id', 'stage_id', 'parent_id'], limit: 100, order: 'project_id,name' },
    []
  );
  return {
    audited_at: new Date().toISOString(),
    counts,
    task_parent_field_exists: hasParent,
    top_level_tasks_sample: topTasks,
    notes: [
      'Top-level task bukan selalu error; tetapi untuk Ground Zero biasanya ide harus berada dalam hierarki parent.',
      'Gunakan Project Export untuk membaca konteks satu project secara mendalam.'
    ]
  };
}

async function contextExport(ctx, payload = {}) {
  const limit = normalizeLimit(payload.limit, 80, 500);
  const models = parseCsvList(payload.models, ['project.project', 'project.task', 'knowledge.article', 'product.template', 'res.partner']);
  const schema = await schemaScan(ctx, { models });
  const samples = {};
  for (const model of models) {
    if (!(await modelExists(ctx, model))) {
      samples[model] = { exists: false, records: [] };
      continue;
    }
    const fields = await defaultReadableFields(ctx, model);
    const records = await safeExecute(ctx, model, 'search_read', [[]], { fields, limit, order: 'write_date desc' }, []);
    samples[model] = { exists: true, fields, records };
  }
  return {
    kind: 'lokalmart_studio_context',
    generated_at: new Date().toISOString(),
    target: { url: ctx.target.url, db: ctx.target.db, username: ctx.target.username },
    instruction_for_chatgpt: 'Baca konteks Odoo Lokalmart ini. Identifikasi struktur terakhir, gap, orphan task, relasi yang kurang, lalu buat rekomendasi dan XLSX patch import-safe jika diminta.',
    schema,
    samples
  };
}

async function defaultReadableFields(ctx, model) {
  const fg = await fieldsGet(ctx, model);
  const preferred = [
    'id', 'display_name', 'name', 'active', 'sequence', 'create_date', 'write_date',
    'user_id', 'project_id', 'parent_id', 'stage_id', 'partner_id', 'company_id',
    'description', 'date_start', 'date_end', 'deadline', 'date_deadline', 'barcode',
    'list_price', 'standard_price', 'categ_id', 'type', 'model', 'state'
  ];
  const fields = preferred.filter(f => fg[f]);
  Object.keys(fg).filter(f => f.startsWith('x_')).slice(0, 25).forEach(f => fields.push(f));
  return [...new Set(fields)].slice(0, 80);
}

function workbookFromBase64(payload) {
  const b64 = payload.file_base64 || payload.base64 || payload.file || '';
  if (!b64) throw new UserError('File XLSX kosong.');
  const clean = String(b64).includes(',') ? String(b64).split(',').pop() : String(b64);
  const buf = Buffer.from(clean, 'base64');
  return XLSX.read(buf, { type: 'buffer', cellDates: true, cellNF: false, cellText: false });
}

function sheetToRows(workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false, blankrows: false });
  return rows.filter(row => Object.values(row).some(v => !isBlank(v)));
}

function sheetShouldImport(sheetName, rows) {
  const normalized = String(sheetName || '').trim().toUpperCase().replace(/[\s-]+/g, '_');
  if (HELPER_SHEETS.has(normalized)) return false;
  if (String(sheetName || '').includes('.')) return true;
  return rows.some(r => String(r._model || r.__model || '').trim());
}

function xlsxPreview(payload = {}) {
  const workbook = workbookFromBase64(payload);
  const sheets = workbook.SheetNames.map(name => {
    const rows = sheetToRows(workbook, name);
    const columns = rows.length ? Object.keys(rows[0]) : [];
    const models = [...new Set(rows.map(r => String(r._model || r.__model || '').trim()).filter(Boolean))];
    return {
      sheet: name,
      rows: rows.length,
      columns,
      models,
      importable: sheetShouldImport(name, rows),
      sample: rows.slice(0, 3)
    };
  });
  return { file: payload.file_name || payload.filename || 'upload.xlsx', sheets };
}

async function importXlsx(ctx, payload = {}) {
  const workbook = workbookFromBase64(payload);
  const options = payload.options || {};
  const hardMax = options.force_large ? LARGE_IMPORT_CAP : SAFE_IMPORT_CAP;
  const maxRows = normalizeLimit(options.max_rows || payload.max_rows, DEFAULT_IMPORT_ROWS, hardMax);
  const continueOnError = options.continue_on_error !== false;
  const onlySheets = parseCsvList(options.sheets || payload.sheets, []);
  const importImages = options.import_images === true || payload.import_images === true;
  const dryRun = options.dry_run === true || payload.dry_run === true;

  const result = {
    started_at: new Date().toISOString(),
    processed: 0,
    created: 0,
    updated: 0,
    deleted: 0,
    archived: 0,
    skipped: 0,
    errors: [],
    warnings: [],
    sheets: [],
    truncated: false,
    settings: { maxRows, continueOnError, importImages, dryRun, forceLarge: !!options.force_large }
  };

  for (const sheetName of workbook.SheetNames) {
    if (onlySheets.length && !onlySheets.includes(sheetName)) continue;
    const rows = sheetToRows(workbook, sheetName);
    const sheetResult = { sheet: sheetName, importable: true, processed: 0, created: 0, updated: 0, deleted: 0, archived: 0, skipped: 0, errors: [], warnings: [] };

    if (!sheetShouldImport(sheetName, rows)) {
      sheetResult.importable = false;
      sheetResult.skipped = rows.length;
      sheetResult.warnings.push('Sheet dilewati karena bukan sheet model Odoo dan tidak punya kolom _model.');
      result.sheets.push(sheetResult);
      result.skipped += rows.length;
      continue;
    }

    for (let i = 0; i < rows.length; i++) {
      if (result.processed >= maxRows) {
        result.truncated = true;
        sheetResult.warnings.push(`Import dihentikan aman di batas ${maxRows} row/request. Jalankan lagi dengan sheet/row berikutnya atau force_large=true jika benar-benar perlu.`);
        break;
      }
      const rowNumber = i + 2;
      const row = rows[i];
      try {
        const one = await importOneRow(ctx, row, sheetName, { importImages, dryRun });
        sheetResult.processed++;
        result.processed++;
        countStatus(one.status, sheetResult, result);
        for (const warning of one.warnings || []) {
          const msg = `${sheetName} row ${rowNumber}: ${warning}`;
          pushLimited(sheetResult.warnings, msg, 80);
          pushLimited(result.warnings, msg, 250);
        }
      } catch (e) {
        const msg = `${sheetName} row ${rowNumber}: ${e.message}`;
        sheetResult.errors.push(msg);
        result.errors.push(msg);
        if (!continueOnError) throw e;
      }
    }

    result.sheets.push(sheetResult);
    if (result.truncated) break;
  }

  result.finished_at = new Date().toISOString();
  result.cache_stats = {
    models: ctx.modelCache.size,
    fields: ctx.fieldsCache.size,
    xmlids: ctx.xmlidCache.size
  };
  return result;
}

function countStatus(status, sheetResult, result) {
  if (status === 'created') { sheetResult.created++; result.created++; }
  else if (status === 'updated') { sheetResult.updated++; result.updated++; }
  else if (status === 'deleted') { sheetResult.deleted++; result.deleted++; }
  else if (status === 'archived') { sheetResult.archived++; result.archived++; }
  else { sheetResult.skipped++; result.skipped++; }
}

function pushLimited(arr, value, max) {
  if (arr.length < max) arr.push(value);
  else if (arr.length === max) arr.push(`... warning lain disembunyikan setelah ${max} item.`);
}

async function importOneRow(ctx, row, sheetName, options = {}) {
  const clean = cleanRow(row);
  const warnings = [];
  const model = String(clean._model || clean.__model || guessModelFromSheet(sheetName) || '').trim();
  const action = String(clean.__action || clean._action || clean.action || 'upsert').trim().toLowerCase();
  const externalId = String(clean._external_id || clean.external_id || '').trim();

  if (!model) throw new UserError('Model kosong. Isi _model di sheet atau pakai nama sheet sebagai model.');
  if (action === 'skip' || action === 'noop') return { status: 'skipped', warnings };

  if (!(await modelExists(ctx, model))) throw new UserError(`Model tidak ditemukan atau tidak bisa diakses: ${model}`);
  const fieldsMeta = await fieldsGet(ctx, model);

  let existingId = null;
  if (externalId) existingId = await resolveExternalId(ctx, externalId, model);
  if (!existingId && clean.id && ['update', 'upsert', 'delete', 'archive'].includes(action)) existingId = Number(clean.id);

  if (action === 'delete') {
    if (!existingId) return { status: 'skipped', warnings: ['Delete dilewati karena record tidak ditemukan.'] };
    if (!options.dryRun) await ctx.execute(model, 'unlink', [[existingId]]);
    return { status: 'deleted', id: existingId, warnings };
  }

  if (action === 'archive') {
    if (!existingId) return { status: 'skipped', warnings: ['Archive dilewati karena record tidak ditemukan.'] };
    if (!fieldsMeta.active) throw new UserError(`Model ${model} tidak punya field active untuk archive.`);
    if (!options.dryRun) await ctx.execute(model, 'write', [[existingId], { active: false }]);
    return { status: 'archived', id: existingId, warnings };
  }

  const converted = await rowToVals(ctx, clean, model, fieldsMeta, options);
  const vals = converted.vals;
  warnings.push(...converted.warnings);

  if (Object.keys(vals).length === 0) return { status: 'skipped', warnings: warnings.concat('Tidak ada value yang bisa diwrite/create.') };

  if (existingId) {
    if (!options.dryRun) await ctx.execute(model, 'write', [[existingId], vals]);
    return { status: 'updated', id: existingId, warnings };
  }

  if (action === 'update' || action === 'write') {
    throw new UserError(`Record untuk update tidak ditemukan: ${externalId || clean.id || '(tanpa id)'}`);
  }

  let newId = null;
  if (!options.dryRun) {
    newId = await ctx.execute(model, 'create', [vals]);
    if (externalId) await createExternalId(ctx, externalId, model, newId);
  }
  return { status: 'created', id: newId, warnings };
}

function cleanRow(row) {
  const out = {};
  for (const [k, v] of Object.entries(row || {})) {
    const key = String(k || '').trim();
    if (!key) continue;
    out[key] = typeof v === 'string' ? v.trim() : v;
  }
  return out;
}

function guessModelFromSheet(sheetName) {
  const name = String(sheetName || '').trim();
  if (name.includes('.')) return name;
  return '';
}

async function rowToVals(ctx, row, model, fieldsMeta, options = {}) {
  const vals = {};
  const warnings = [];
  const handled = new Set();

  for (const [key, value] of Object.entries(row)) {
    if (key.endsWith('_external_id') && key !== '_external_id') {
      const field = key.slice(0, -'_external_id'.length);
      handled.add(key);
      handled.add(field);
      if (!fieldsMeta[field]) {
        warnings.push(`Kolom ${key} dilewati: field ${field} tidak ada di ${model}.`);
        continue;
      }
      if (isBlank(value)) continue;
      const id = await resolveExternalId(ctx, String(value).trim(), fieldsMeta[field].relation || null);
      if (!id) throw new UserError(`External ID tidak ditemukan untuk ${key}: ${value}`);
      vals[field] = id;
    }

    if (key.endsWith('_external_ids')) {
      const field = key.slice(0, -'_external_ids'.length);
      handled.add(key);
      handled.add(field);
      if (!fieldsMeta[field]) {
        warnings.push(`Kolom ${key} dilewati: field ${field} tidak ada di ${model}.`);
        continue;
      }
      const ids = [];
      for (const xmlid of parseCsvList(value, [])) {
        const id = await resolveExternalId(ctx, xmlid, fieldsMeta[field].relation || null);
        if (!id) throw new UserError(`External ID tidak ditemukan untuk ${key}: ${xmlid}`);
        ids.push(id);
      }
      vals[field] = [[6, 0, ids]];
    }
  }

  if (row.image_url && fieldsMeta.image_1920) {
    handled.add('image_url');
    if (options.importImages === true) {
      const image = await fetchImageAsBase64(row.image_url);
      if (image) vals.image_1920 = image;
      else warnings.push('image_url tidak berhasil diambil atau terlalu besar; foto dilewati.');
    } else {
      warnings.push('image_url tidak diimport otomatis. Gunakan options.import_images=true atau simpan sebagai x_source_image_url.');
    }
  }

  for (const [key, value] of Object.entries(row)) {
    if (handled.has(key)) continue;
    if (key.startsWith('_') || ['__action', '_action', 'action', 'external_id', 'id'].includes(key)) continue;
    if (key.endsWith('_external_id') || key.endsWith('_external_ids')) continue;
    if (key === 'image_url') continue;

    const meta = fieldsMeta[key];
    if (!meta) {
      // Banyak file ChatGPT punya kolom bantuan. Jangan bikin warning banjir kecuali strict.
      if (options.strictFields) warnings.push(`Kolom ${key} dilewati: field tidak ada di ${model}.`);
      continue;
    }
    if (meta.readonly && !meta.required) continue;
    if (isBlank(value)) continue;

    const casted = castValue(value, meta);
    if (casted === undefined) {
      warnings.push(`Kolom ${key} dilewati: format tidak cocok untuk field ${meta.type}.`);
      continue;
    }
    vals[key] = casted;
  }

  return { vals, warnings };
}

function isBlank(value) {
  return value === null || value === undefined || (typeof value === 'string' && value.trim() === '');
}

function castValue(value, meta) {
  const type = meta.type;
  if (type === 'boolean') {
    const s = String(value).toLowerCase().trim();
    if (['1', 'true', 'yes', 'y', 'ya', 'iya', 'aktif', 'published'].includes(s)) return true;
    if (['0', 'false', 'no', 'n', 'tidak', 'nonaktif', 'draft'].includes(s)) return false;
    return Boolean(value);
  }
  if (type === 'integer') return Number.parseInt(value, 10) || 0;
  if (type === 'float' || type === 'monetary') return Number.parseFloat(String(value).replace(',', '.')) || 0;
  if (type === 'date') return toOdooDate(value);
  if (type === 'datetime') return toOdooDatetime(value);
  if (type === 'many2one') {
    if (typeof value === 'number') return value;
    if (/^\d+$/.test(String(value))) return Number.parseInt(value, 10);
    return undefined;
  }
  if (type === 'many2many') {
    const ids = parseCsvList(value, []).map(v => Number.parseInt(v, 10)).filter(Number.isFinite);
    return [[6, 0, ids]];
  }
  return value;
}

function toOdooDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0, 10);
  const s = String(value).trim();
  return s ? s.slice(0, 10) : false;
}

function toOdooDatetime(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().replace('T', ' ').slice(0, 19);
  return String(value).trim();
}

function splitXmlId(xmlid) {
  const raw = String(xmlid || '').trim();
  if (!raw) return null;
  if (raw.includes('.')) {
    const [module, ...rest] = raw.split('.');
    return { module: safeXmlPart(module), name: safeXmlPart(rest.join('.')) };
  }
  return { module: DEFAULT_MODULE, name: safeXmlPart(raw) };
}

function safeXmlPart(value) {
  return String(value || '').trim().replace(/[^a-zA-Z0-9_\-.]/g, '_').replace(/\.+/g, '_');
}

async function resolveExternalId(ctx, xmlid, expectedModel = null) {
  const parts = splitXmlId(xmlid);
  if (!parts) return null;
  const cacheKey = `${parts.module}.${parts.name}|${expectedModel || ''}`;
  if (ctx.xmlidCache.has(cacheKey)) return ctx.xmlidCache.get(cacheKey);

  const rows = await ctx.execute('ir.model.data', 'search_read', [[[ 'module', '=', parts.module ], [ 'name', '=', parts.name ]]], { fields: ['id', 'module', 'name', 'model', 'res_id'], limit: 1 });
  let result = null;
  if (rows.length && (!expectedModel || rows[0].model === expectedModel)) result = rows[0].res_id;
  ctx.xmlidCache.set(cacheKey, result);
  return result;
}

async function createExternalId(ctx, xmlid, model, resId) {
  const parts = splitXmlId(xmlid);
  if (!parts) return null;
  const rawRows = await ctx.execute('ir.model.data', 'search_read', [[[ 'module', '=', parts.module ], [ 'name', '=', parts.name ]]], { fields: ['id', 'model', 'res_id'], limit: 1 });
  if (rawRows.length) {
    if (rawRows[0].model !== model || rawRows[0].res_id !== resId) {
      throw new UserError(`External ID sudah dipakai oleh record lain: ${parts.module}.${parts.name}`);
    }
    return rawRows[0].res_id;
  }
  await ctx.execute('ir.model.data', 'create', [{ module: parts.module, name: parts.name, model, res_id: resId, noupdate: true }]);
  ctx.xmlidCache.set(`${parts.module}.${parts.name}|${model}`, resId);
  ctx.xmlidCache.set(`${parts.module}.${parts.name}|`, resId);
  return resId;
}

async function fetchImageAsBase64(url) {
  try {
    const response = await fetch(String(url), { redirect: 'follow' });
    if (!response.ok) return null;
    const ab = await response.arrayBuffer();
    if (ab.byteLength > 5 * 1024 * 1024) return null;
    return Buffer.from(ab).toString('base64');
  } catch (_) {
    return null;
  }
}

async function fullExport(ctx, payload = {}) {
  const models = parseCsvList(payload.models, ['project.project', 'project.task', 'knowledge.article', 'product.template']);
  const limit = normalizeLimit(payload.limit, 300, 3000);
  const format = String(payload.format || 'json').toLowerCase();
  const result = { generated_at: new Date().toISOString(), target: { url: ctx.target.url, db: ctx.target.db }, models: {} };

  for (const model of models) {
    if (!(await modelExists(ctx, model))) {
      result.models[model] = { exists: false, records: [] };
      continue;
    }
    const fields = payload.fields ? parseCsvList(payload.fields, []) : await defaultReadableFields(ctx, model);
    const records = await ctx.execute(model, 'search_read', [[]], { fields, limit, order: 'write_date desc' });
    result.models[model] = { exists: true, fields, records };
  }

  if (format === 'xlsx') {
    return { file_name: `lokalmart_full_export_${dateStamp()}.xlsx`, mime: XLSX_MIME, base64: objectToWorkbookBase64(result.models) };
  }
  return result;
}

async function projectList(ctx, payload = {}) {
  const limit = normalizeLimit(payload.limit, 200, 1000);
  const domain = [];
  if (payload.search) domain.push(['name', 'ilike', String(payload.search)]);
  const fields = await safeFields(ctx, 'project.project', ['id', 'name', 'display_name', 'active', 'user_id', 'partner_id', 'company_id', 'create_date', 'write_date']);
  return await ctx.execute('project.project', 'search_read', [domain], { fields, limit, order: 'write_date desc' });
}

async function projectContextExport(ctx, payload = {}) {
  const projectId = Number(payload.project_id || payload.projectId || payload.id || 0);
  if (!projectId) throw new UserError('project_id kosong. Pilih project dulu.');

  const projectFields = await safeFields(ctx, 'project.project', ['id', 'name', 'display_name', 'active', 'user_id', 'partner_id', 'company_id', 'create_date', 'write_date', 'description']);
  const projectArr = await ctx.execute('project.project', 'read', [[projectId], projectFields]);
  if (!projectArr.length) throw new UserError(`Project ID ${projectId} tidak ditemukan.`);
  const project = projectArr[0];

  const taskFields = await safeFields(ctx, 'project.task', ['id', 'name', 'display_name', 'active', 'project_id', 'parent_id', 'child_ids', 'stage_id', 'user_ids', 'partner_id', 'priority', 'sequence', 'date_deadline', 'create_date', 'write_date', 'description']);
  const tasks = await ctx.execute('project.task', 'search_read', [[[ 'project_id', '=', projectId ]]], { fields: taskFields, limit: 3000, order: 'parent_id,sequence,id' });
  const hierarchy = buildTaskHierarchy(tasks);

  const milestones = await readIfModel(ctx, 'project.milestone', [[[ 'project_id', '=', projectId ]]], ['id', 'name', 'project_id', 'deadline', 'is_reached', 'create_date', 'write_date'], 1000);
  const updates = await readIfModel(ctx, 'project.update', [[[ 'project_id', '=', projectId ]]], ['id', 'name', 'project_id', 'status', 'progress', 'description', 'create_date', 'write_date'], 500);
  const messagesProject = await readIfModel(ctx, 'mail.message', [[[ 'model', '=', 'project.project' ], [ 'res_id', '=', projectId ]]], ['id', 'subject', 'body', 'date', 'author_id', 'message_type'], 200);
  const taskIds = tasks.map(t => t.id);
  const messagesTasks = taskIds.length ? await readIfModel(ctx, 'mail.message', [[[ 'model', '=', 'project.task' ], [ 'res_id', 'in', taskIds.slice(0, 800) ]]], ['id', 'subject', 'body', 'date', 'author_id', 'message_type', 'res_id'], 500) : [];
  const xmlids = await exportXmlIds(ctx, ['project.project', 'project.task', 'project.milestone'], [projectId, ...taskIds, ...milestones.map(m => m.id)]);

  return {
    kind: 'lokalmart_project_context',
    generated_at: new Date().toISOString(),
    target: { url: ctx.target.url, db: ctx.target.db, username: ctx.target.username },
    project,
    counts: { tasks: tasks.length, top_level_tasks: hierarchy.length, milestones: milestones.length, updates: updates.length, chatter_project: messagesProject.length, chatter_tasks: messagesTasks.length },
    task_hierarchy: hierarchy,
    tasks,
    milestones,
    updates,
    chatter: { project: sanitizeMessages(messagesProject), tasks: sanitizeMessages(messagesTasks) },
    external_ids: xmlids,
    prompt_for_chatgpt: buildProjectPrompt(project, tasks, milestones, updates)
  };
}

async function projectXlsxExport(ctx, payload = {}) {
  const context = await projectContextExport(ctx, payload);
  const sheets = {
    project: [context.project],
    tasks: context.tasks,
    task_hierarchy: flattenHierarchy(context.task_hierarchy),
    milestones: context.milestones,
    updates: context.updates,
    chatter_project: context.chatter.project,
    chatter_tasks: context.chatter.tasks,
    external_ids: context.external_ids,
    prompt: [{ prompt: context.prompt_for_chatgpt }]
  };
  const name = sanitizeFilename(`project_${context.project.name || context.project.id}_${dateStamp()}.xlsx`);
  return { file_name: name, mime: XLSX_MIME, base64: objectToWorkbookBase64(sheets) };
}

async function safeFields(ctx, model, wanted) {
  const fg = await fieldsGet(ctx, model);
  const fields = wanted.filter(f => fg[f]);
  Object.keys(fg).filter(f => f.startsWith('x_')).slice(0, 30).forEach(f => fields.push(f));
  return [...new Set(fields)];
}

async function readIfModel(ctx, model, domain, fields, limit = 500) {
  if (!(await modelExists(ctx, model))) return [];
  const safe = await safeFields(ctx, model, fields);
  return await safeExecute(ctx, model, 'search_read', domain, { fields: safe, limit, order: 'write_date desc' }, []);
}

async function exportXmlIds(ctx, models, resIds) {
  const ids = resIds.filter(Boolean);
  if (!ids.length) return [];
  return await safeExecute(ctx, 'ir.model.data', 'search_read', [[[ 'model', 'in', models ], [ 'res_id', 'in', ids ]]], { fields: ['module', 'name', 'model', 'res_id'], limit: 5000 }, []);
}

function buildTaskHierarchy(tasks) {
  const byId = new Map();
  for (const t of tasks) byId.set(t.id, { ...t, children: [] });
  const roots = [];
  for (const node of byId.values()) {
    const parentId = Array.isArray(node.parent_id) ? node.parent_id[0] : node.parent_id;
    if (parentId && byId.has(parentId)) byId.get(parentId).children.push(node);
    else roots.push(node);
  }
  const sortTree = nodes => {
    nodes.sort((a, b) => (Number(a.sequence || 0) - Number(b.sequence || 0)) || String(a.name || '').localeCompare(String(b.name || '')));
    for (const n of nodes) sortTree(n.children || []);
    return nodes;
  };
  return sortTree(roots);
}

function flattenHierarchy(nodes, level = 0, parent = '') {
  const out = [];
  for (const node of nodes || []) {
    out.push({ level, parent, id: node.id, name: node.name, stage_id: pairName(node.stage_id), write_date: node.write_date });
    out.push(...flattenHierarchy(node.children || [], level + 1, node.name));
  }
  return out;
}

function pairName(value) { return Array.isArray(value) ? value[1] : value; }

function sanitizeMessages(messages) {
  return (messages || []).map(m => ({ ...m, body: stripHtml(m.body || '').slice(0, 4000) }));
}

function stripHtml(html) {
  return String(html || '')
    .replace(/<br\s*\/?\s*>/gi, ' ')
    .replace(/<\/p>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/\s+/g, ' ')
    .trim();
}

function buildProjectPrompt(project, tasks, milestones, updates) {
  return [
    'Saya memberikan export konteks satu project Odoo Lokalmart.',
    `Project: ${project.name || project.display_name || project.id}`,
    `Jumlah task: ${tasks.length}. Jumlah milestone: ${milestones.length}. Jumlah update: ${updates.length}.`,
    'Tolong baca perkembangan terakhir, bandingkan dengan diskusi Lokalmart sebelumnya, lalu kembangkan struktur ide/task agar lebih matang, hierarkis, import-safe, dan tidak membuat task liar tanpa parent.',
    'Jika membuat XLSX patch, gunakan aturan: __action, _external_id, _model, Many2one pakai field_external_id, Many2many pakai field_external_ids, custom field pakai x_.'
  ].join('\n');
}

async function barcodeLookup(ctx, payload = {}) {
  const barcode = String(payload.barcode || '').trim();
  if (!barcode) throw new UserError('Barcode kosong.');
  const productProduct = await readIfModel(ctx, 'product.product', [[[ 'barcode', '=', barcode ]]], ['id', 'display_name', 'barcode', 'product_tmpl_id', 'lst_price', 'default_code'], 20);
  const productTemplate = await readIfModel(ctx, 'product.template', [[[ 'barcode', '=', barcode ]]], ['id', 'name', 'barcode', 'list_price', 'default_code', 'categ_id'], 20);
  return { barcode, product_product: productProduct, product_template: productTemplate, found: productProduct.length + productTemplate.length };
}

async function readRecords(ctx, payload = {}) {
  const model = String(payload.model || '').trim();
  if (!model) throw new UserError('Model kosong.');
  const limit = normalizeLimit(payload.limit, 100, 2000);
  const domain = Array.isArray(payload.domain) ? payload.domain : [];
  const fields = parseCsvList(payload.fields, await defaultReadableFields(ctx, model));
  return await ctx.execute(model, 'search_read', [domain], { fields, limit, order: payload.order || 'write_date desc' });
}

function objectToWorkbookBase64(sheets) {
  const wb = XLSX.utils.book_new();
  for (const [name, value] of Object.entries(sheets || {})) {
    const rows = Array.isArray(value) ? value : [value];
    const normalized = rows.map(row => flattenRecord(row));
    const ws = XLSX.utils.json_to_sheet(normalized.length ? normalized : [{ empty: '' }]);
    XLSX.utils.book_append_sheet(wb, ws, safeSheetName(name));
  }
  const buf = XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' });
  return buf.toString('base64');
}

function flattenRecord(row, prefix = '', out = {}) {
  if (row === null || row === undefined) return out;
  if (typeof row !== 'object' || row instanceof Date) {
    out[prefix || 'value'] = row;
    return out;
  }
  if (Array.isArray(row)) {
    out[prefix || 'value'] = JSON.stringify(row);
    return out;
  }
  for (const [key, value] of Object.entries(row)) {
    const next = prefix ? `${prefix}.${key}` : key;
    if (value && typeof value === 'object' && !Array.isArray(value) && !(value instanceof Date)) flattenRecord(value, next, out);
    else out[next] = Array.isArray(value) ? JSON.stringify(value) : value;
  }
  return out;
}

function safeSheetName(name) { return String(name || 'sheet').replace(/[\\/?*\[\]:]/g, '_').slice(0, 31) || 'sheet'; }
function sanitizeFilename(name) { return String(name || 'export.xlsx').replace(/[^a-zA-Z0-9_.-]+/g, '_').slice(0, 120); }
function dateStamp() { return new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19); }
