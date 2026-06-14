/*
  Lokalmart Studio v3 - single Vercel API
  ------------------------------------------------------------
  One file, one endpoint: /api/odoo
  Frontend sends: { action, target, payload }

  Target shape:
  {
    url: 'https://edu-lokalmart.odoo.com',
    db: 'edu-lokalmart',
    username: 'admin@example.com',
    password: 'Odoo API key or local password'
  }
*/

const XLSX = require('xlsx');

const MAX_BODY_BYTES = 18 * 1024 * 1024;
const DEFAULT_MODULE = 'lokalmart_studio';
const DEFAULT_LIMIT = 250;

module.exports = async function handler(req, res) {
  applyCors(res);
  if (req.method === 'OPTIONS') return json(res, 200, { ok: true });
  if (req.method === 'GET') {
    return json(res, 200, {
      ok: true,
      name: 'Lokalmart Studio v3 Vercel API',
      endpoint: '/api/odoo',
      usage: 'POST JSON { action, target, payload }',
      actions: [
        'health', 'test_connection', 'schema_scan', 'data_audit', 'context_export',
        'xlsx_preview', 'import_xlsx', 'full_export',
        'project_list', 'project_context_export', 'project_xlsx_export',
        'barcode_lookup', 'read_records'
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
      'test_connection', 'schema_scan', 'data_audit', 'context_export',
      'xlsx_preview', 'import_xlsx', 'full_export', 'project_list',
      'project_context_export', 'project_xlsx_export', 'barcode_lookup', 'read_records'
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
  res.end(JSON.stringify(body));
}

function readJsonBody(req) {
  return new Promise((resolve, reject) => {
    let raw = '';
    let total = 0;
    req.on('data', chunk => {
      total += chunk.length;
      if (total > MAX_BODY_BYTES) {
        reject(new UserError('Payload terlalu besar. Kurangi jumlah sheet/baris atau gunakan batch lebih kecil.'));
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
  let url = String(target.url || '').trim();
  url = url.replace(/\/+$/, '').replace(/\/web$/, '');
  return {
    url,
    db: String(target.db || target.database || '').trim(),
    username: String(target.username || target.email || target.login || '').trim(),
    password: String(target.password || target.apiKey || target.api_key || '').trim()
  };
}

function validateTarget(target) {
  if (!target.url || !/^https?:\/\//i.test(target.url)) throw new UserError('URL Odoo tidak valid.', 'Contoh: https://edu-lokalmart.odoo.com');
  if (!target.db) throw new UserError('Database Odoo kosong.', 'Contoh: edu-lokalmart');
  if (!target.username) throw new UserError('Username/email Odoo kosong.');
  if (!target.password) throw new UserError('Password/API key kosong.', 'Untuk Odoo Online, gunakan API Key sebagai pengganti password.');
}

function createOdooClient(target) {
  let uidCache = null;
  let loginPromise = null;

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
    catch (e) {
      throw new UserError(`Odoo mengembalikan non-JSON HTTP ${response.status}.`, 'Cek URL Odoo. Jangan pakai /web di belakang URL.', text.slice(0, 1200));
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
      loginPromise = jsonRpc('common', 'login', [target.db, target.username, target.password])
        .then(uid => {
          if (!uid) {
            throw new UserError(
              `Login Odoo gagal untuk database "${target.db}" dan user "${target.username}".`,
              'Gunakan API Key Odoo sebagai password. Pastikan database dan email user benar.'
            );
          }
          uidCache = uid;
          return { uid };
        });
    }
    return loginPromise;
  }

  async function execute(model, method, args = [], kwargs = {}) {
    const session = await login();
    return jsonRpc('object', 'execute_kw', [target.db, session.uid, target.password, model, method, args, kwargs]);
  }

  return { target, login, execute };
}

async function safeExecute(ctx, model, method, args = [], kwargs = {}, fallback = null) {
  try { return await ctx.execute(model, method, args, kwargs); }
  catch (e) { return fallback; }
}

async function modelExists(ctx, model) {
  const ids = await safeExecute(ctx, 'ir.model', 'search', [[[ 'model', '=', model ]]], { limit: 1 }, []);
  return Array.isArray(ids) && ids.length > 0;
}

async function fieldsGet(ctx, model) {
  return await ctx.execute(model, 'fields_get', [], { attributes: ['string', 'type', 'relation', 'required', 'readonly'] });
}

function parseCsvList(value, fallback = []) {
  if (Array.isArray(value)) return value.map(v => String(v).trim()).filter(Boolean);
  if (!value) return fallback;
  return String(value).split(/[\n,]+/).map(v => v.trim()).filter(Boolean);
}

function normalizeLimit(value, fallback = DEFAULT_LIMIT, max = 2000) {
  const n = Number(value || fallback);
  if (!Number.isFinite(n) || n <= 0) return fallback;
  return Math.min(Math.floor(n), max);
}

async function schemaScan(ctx, payload = {}) {
  const defaultModels = [
    'project.project', 'project.task', 'project.milestone', 'project.update',
    'knowledge.article', 'product.template', 'product.product', 'product.category',
    'res.partner', 'sale.order', 'sale.order.line', 'ir.model', 'ir.model.fields'
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
    const fields = Object.entries(fg).map(([name, meta]) => ({ name, ...meta })).sort((a, b) => a.name.localeCompare(b.name));
    out.push({ model, exists: true, field_count: fields.length, fields });
  }
  const customModels = await safeExecute(ctx, 'ir.model', 'search_read', [[[ 'model', 'like', 'x_' ]]], { fields: ['id', 'name', 'model', 'state'], limit: 200 }, []);
  return { scanned_at: new Date().toISOString(), models: out, custom_models: customModels };
}

async function dataAudit(ctx, payload = {}) {
  const models = parseCsvList(payload.models, [
    'project.project', 'project.task', 'project.milestone', 'project.update',
    'knowledge.article', 'product.template', 'product.product', 'res.partner',
    'sale.order', 'sale.order.line', 'ir.model.fields', 'ir.ui.view'
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
  const orphanDomain = hasParent ? [[['project_id', '!=', false], ['parent_id', '=', false]]] : [[['project_id', '!=', false]]];
  const topTasks = await safeExecute(ctx, 'project.task', 'search_read', orphanDomain, {
    fields: ['id', 'name', 'project_id', 'stage_id', 'parent_id'],
    limit: 100,
    order: 'project_id,name'
  }, []);

  return {
    audited_at: new Date().toISOString(),
    counts,
    task_parent_field_exists: hasParent,
    top_level_tasks_sample: topTasks,
    notes: [
      'Top-level task bukan selalu error; tetapi untuk Ground Zero user biasanya ingin semua ide punya parent hierarki.',
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
  const preferred = ['id', 'display_name', 'name', 'active', 'sequence', 'create_date', 'write_date', 'user_id', 'project_id', 'parent_id', 'stage_id', 'partner_id', 'company_id', 'description', 'date_start', 'date_end', 'deadline', 'barcode', 'list_price', 'standard_price', 'categ_id', 'type', 'model', 'state'];
  const fields = preferred.filter(f => fg[f]);
  Object.keys(fg).filter(f => f.startsWith('x_')).slice(0, 25).forEach(f => fields.push(f));
  return [...new Set(fields)].slice(0, 80);
}

function workbookFromBase64(payload) {
  const b64 = payload.file_base64 || payload.base64 || '';
  if (!b64) throw new UserError('File XLSX kosong.');
  const clean = String(b64).includes(',') ? String(b64).split(',').pop() : String(b64);
  const buf = Buffer.from(clean, 'base64');
  return XLSX.read(buf, { type: 'buffer', cellDates: true });
}

function sheetToRows(workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });
}

function xlsxPreview(payload = {}) {
  const workbook = workbookFromBase64(payload);
  const sheets = workbook.SheetNames.map(name => {
    const rows = sheetToRows(workbook, name);
    const columns = rows.length ? Object.keys(rows[0]) : [];
    const models = [...new Set(rows.map(r => String(r._model || '').trim()).filter(Boolean))];
    return {
      sheet: name,
      rows: rows.length,
      columns,
      models,
      sample: rows.slice(0, 3)
    };
  });
  return { file: payload.file_name || payload.filename || 'upload.xlsx', sheets };
}

async function importXlsx(ctx, payload = {}) {
  const workbook = workbookFromBase64(payload);
  const options = payload.options || {};
  const maxRows = normalizeLimit(options.max_rows || payload.max_rows, 2000, 10000);
  const continueOnError = options.continue_on_error !== false;
  const onlySheets = parseCsvList(options.sheets || payload.sheets, []);
  const result = { started_at: new Date().toISOString(), processed: 0, created: 0, updated: 0, skipped: 0, errors: [], warnings: [], sheets: [] };

  for (const sheetName of workbook.SheetNames) {
    if (onlySheets.length && !onlySheets.includes(sheetName)) continue;
    const rows = sheetToRows(workbook, sheetName);
    const sheetResult = { sheet: sheetName, processed: 0, created: 0, updated: 0, skipped: 0, errors: [] };
    for (let i = 0; i < rows.length && result.processed < maxRows; i++) {
      const rowNumber = i + 2;
      const row = rows[i];
      try {
        const one = await importOneRow(ctx, row, sheetName);
        sheetResult.processed++;
        result.processed++;
        if (one.status === 'created') { sheetResult.created++; result.created++; }
        else if (one.status === 'updated') { sheetResult.updated++; result.updated++; }
        else { sheetResult.skipped++; result.skipped++; }
      } catch (e) {
        const msg = `${sheetName} row ${rowNumber}: ${e.message}`;
        sheetResult.errors.push(msg);
        result.errors.push(msg);
        if (!continueOnError) throw e;
      }
    }
    result.sheets.push(sheetResult);
  }
  result.finished_at = new Date().toISOString();
  return result;
}

async function importOneRow(ctx, row, sheetName) {
  const clean = cleanRow(row);
  const model = String(clean._model || guessModelFromSheet(sheetName) || '').trim();
  const action = String(clean.__action || clean.action || 'upsert').trim().toLowerCase();
  const externalId = String(clean._external_id || clean.external_id || '').trim();
  if (!model) throw new UserError('Model kosong. Isi _model di sheet atau pakai nama sheet sebagai model.');
  if (action === 'skip' || action === 'noop') return { status: 'skipped' };
  if (!(await modelExists(ctx, model))) throw new UserError(`Model tidak ditemukan atau tidak bisa diakses: ${model}`);

  const fieldsMeta = await fieldsGet(ctx, model);
  let existingId = null;
  if (externalId) existingId = await resolveExternalId(ctx, externalId, model);
  if (!existingId && clean.id && ['update', 'upsert', 'delete', 'archive'].includes(action)) existingId = Number(clean.id);

  if (action === 'delete') {
    if (!existingId) return { status: 'skipped' };
    await ctx.execute(model, 'unlink', [[existingId]]);
    return { status: 'deleted', id: existingId };
  }

  if (action === 'archive') {
    if (!existingId) return { status: 'skipped' };
    if (!fieldsMeta.active) throw new UserError(`Model ${model} tidak punya field active untuk archive.`);
    await ctx.execute(model, 'write', [[existingId], { active: false }]);
    return { status: 'archived', id: existingId };
  }

  const vals = await rowToVals(ctx, clean, model, fieldsMeta);
  if (Object.keys(vals).length === 0) return { status: 'skipped' };

  if (existingId) {
    await ctx.execute(model, 'write', [[existingId], vals]);
    return { status: 'updated', id: existingId };
  }

  if (action === 'update') throw new UserError(`Record untuk update tidak ditemukan: ${externalId || clean.id || '(tanpa id)'}`);
  const newId = await ctx.execute(model, 'create', [vals]);
  if (externalId) await createExternalId(ctx, externalId, model, newId);
  return { status: 'created', id: newId };
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

async function rowToVals(ctx, row, model, fieldsMeta) {
  const vals = {};
  const handled = new Set();

  for (const [key, value] of Object.entries(row)) {
    if (key.endsWith('_external_id') && key !== '_external_id') {
      const field = key.slice(0, -'_external_id'.length);
      if (!fieldsMeta[field]) continue;
      if (isBlank(value)) continue;
      const id = await resolveExternalId(ctx, String(value).trim(), fieldsMeta[field].relation || null);
      if (!id) throw new UserError(`External ID tidak ditemukan untuk ${key}: ${value}`);
      vals[field] = id;
      handled.add(key);
      handled.add(field);
    }
    if (key.endsWith('_external_ids')) {
      const field = key.slice(0, -'_external_ids'.length);
      if (!fieldsMeta[field]) continue;
      const ids = [];
      for (const xmlid of parseCsvList(value, [])) {
        const id = await resolveExternalId(ctx, xmlid, fieldsMeta[field].relation || null);
        if (!id) throw new UserError(`External ID tidak ditemukan untuk ${key}: ${xmlid}`);
        ids.push(id);
      }
      vals[field] = [[6, 0, ids]];
      handled.add(key);
      handled.add(field);
    }
  }

  if (row.image_url && fieldsMeta.image_1920) {
    const image = await fetchImageAsBase64(row.image_url);
    if (image) vals.image_1920 = image;
    handled.add('image_url');
  }

  for (const [key, value] of Object.entries(row)) {
    if (handled.has(key)) continue;
    if (key.startsWith('_') || key === '__action' || key === 'action' || key === 'external_id' || key === 'id') continue;
    if (key.endsWith('_external_id') || key.endsWith('_external_ids')) continue;
    if (key === 'image_url') continue;
    const meta = fieldsMeta[key];
    if (!meta) continue;
    if (meta.readonly && !meta.required) continue;
    if (isBlank(value)) continue;
    vals[key] = castValue(value, meta.type);
  }
  return vals;
}

function isBlank(value) {
  return value === null || value === undefined || (typeof value === 'string' && value.trim() === '');
}

function castValue(value, type) {
  if (type === 'boolean') {
    const s = String(value).toLowerCase().trim();
    return ['1', 'true', 'yes', 'y', 'ya', 'iya', 'aktif'].includes(s);
  }
  if (type === 'integer') return Number.parseInt(value, 10) || 0;
  if (type === 'float' || type === 'monetary') return Number.parseFloat(String(value).replace(',', '.')) || 0;
  if (type === 'date' || type === 'datetime') return String(value).trim();
  return value;
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
  const rows = await ctx.execute('ir.model.data', 'search_read', [[[ 'module', '=', parts.module ], [ 'name', '=', parts.name ]]], { fields: ['id', 'model', 'res_id'], limit: 1 });
  if (!rows.length) return null;
  if (expectedModel && rows[0].model !== expectedModel) return null;
  return rows[0].res_id;
}

async function createExternalId(ctx, xmlid, model, resId) {
  const parts = splitXmlId(xmlid);
  if (!parts) return null;
  const existing = await resolveExternalId(ctx, xmlid, model);
  if (existing) return existing;
  return ctx.execute('ir.model.data', 'create', [{ module: parts.module, name: parts.name, model, res_id: resId, noupdate: true }]);
}

async function fetchImageAsBase64(url) {
  try {
    const response = await fetch(String(url), { redirect: 'follow' });
    if (!response.ok) return null;
    const ab = await response.arrayBuffer();
    if (ab.byteLength > 5 * 1024 * 1024) return null;
    return Buffer.from(ab).toString('base64');
  } catch (e) { return null; }
}

async function fullExport(ctx, payload = {}) {
  const models = parseCsvList(payload.models, ['project.project', 'project.task', 'knowledge.article', 'product.template']);
  const limit = normalizeLimit(payload.limit, 500, 5000);
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
  return await ctx.execute('project.project', 'search_read', [domain], {
    fields: ['id', 'name', 'display_name', 'active', 'user_id', 'partner_id', 'company_id', 'create_date', 'write_date'],
    limit,
    order: 'write_date desc'
  });
}

async function projectContextExport(ctx, payload = {}) {
  const projectId = Number(payload.project_id || payload.id || 0);
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

  const context = {
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
  return context;
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

function pairName(value) {
  return Array.isArray(value) ? value[1] : value;
}

function sanitizeMessages(messages) {
  return (messages || []).map(m => ({ ...m, body: stripHtml(m.body || '').slice(0, 4000) }));
}

function stripHtml(html) {
  return String(html || '').replace(/<style[\s\S]*?<\/style>/gi, ' ').replace(/<script[\s\S]*?<\/script>/gi, ' ').replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ').replace(/\s+/g, ' ').trim();
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

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

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

function safeSheetName(name) {
  return String(name || 'sheet').replace(/[\\/?*\[\]:]/g, '_').slice(0, 31) || 'sheet';
}

function sanitizeFilename(name) {
  return String(name || 'export.xlsx').replace(/[^a-zA-Z0-9_.-]+/g, '_').slice(0, 120);
}

function dateStamp() {
  return new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
}
