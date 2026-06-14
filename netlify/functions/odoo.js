'use strict';

const XLSX = require('xlsx');

const DEFAULT_MODELS = [
  'project.project',
  'project.task',
  'project.milestone',
  'project.update',
  'product.template',
  'product.product',
  'product.category',
  'res.partner',
  'sale.order',
  'sale.order.line',
  'account.move',
  'stock.quant',
  'website.page',
  'ir.ui.view',
  'ir.model',
  'ir.model.fields',
  'ir.model.data'
];

const SAFE_READ_METHODS = new Set(['search_read', 'search_count', 'read', 'fields_get', 'search']);
const RESERVED_COLUMNS = new Set(['__action', '_action', '_external_id', 'external_id', '_model', '__model', '_id', 'id']);
const MAX_BODY_SIZE = 35 * 1024 * 1024;

function sendJson(res, status, payload) {
  res.statusCode = status;
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.setHeader('Cache-Control', 'no-store');
  res.end(JSON.stringify(payload));
}

function normalizeError(err) {
  const message = err && err.message ? String(err.message) : String(err || 'Unknown error');
  const detail = err && err.stack ? String(err.stack).slice(0, 8000) : undefined;
  return { message, detail };
}

async function readJsonBody(req) {
  const chunks = [];
  let total = 0;
  for await (const chunk of req) {
    total += chunk.length;
    if (total > MAX_BODY_SIZE) throw new Error('Payload terlalu besar. Kurangi batch atau file.');
    chunks.push(chunk);
  }
  const raw = Buffer.concat(chunks).toString('utf8');
  if (!raw.trim()) return {};
  try {
    return JSON.parse(raw);
  } catch (err) {
    throw new Error('Body bukan JSON valid: ' + err.message);
  }
}

function cleanUrl(url) {
  const value = String(url || '').trim().replace(/\/+$/, '');
  if (!value) throw new Error('URL Odoo belum diisi.');
  if (!/^https?:\/\//i.test(value)) throw new Error('URL Odoo harus diawali http:// atau https://');
  return value;
}

function assertTarget(target) {
  const t = target || {};
  return {
    url: cleanUrl(t.url || t.host),
    db: String(t.db || t.database || '').trim(),
    username: String(t.username || t.login || '').trim(),
    password: String(t.password || t.apiKey || t.key || '').trim()
  };
}

function boolish(value) {
  if (typeof value === 'boolean') return value;
  if (typeof value === 'number') return value !== 0;
  const s = String(value || '').trim().toLowerCase();
  if (['true', '1', 'yes', 'y', 'iya', 'ya', 'aktif', 'published'].includes(s)) return true;
  if (['false', '0', 'no', 'n', 'tidak', 'nonaktif', 'draft'].includes(s)) return false;
  return value;
}

function isBlank(value) {
  return value === null || value === undefined || String(value).trim() === '';
}

function asArray(value) {
  if (Array.isArray(value)) return value;
  if (isBlank(value)) return [];
  return String(value)
    .split(/[;,|]/)
    .map(v => v.trim())
    .filter(Boolean);
}

function sanitizeExternalName(name) {
  return String(name || '')
    .trim()
    .replace(/\s+/g, '_')
    .replace(/[^A-Za-z0-9_.-]/g, '_')
    .replace(/_+/g, '_')
    .slice(0, 180) || ('x_' + Date.now());
}

function parseExternalId(value) {
  const raw = String(value || '').trim();
  if (!raw) return null;
  const idx = raw.indexOf('.');
  if (idx > 0) {
    return {
      module: sanitizeExternalName(raw.slice(0, idx)),
      name: sanitizeExternalName(raw.slice(idx + 1)),
      complete: sanitizeExternalName(raw.slice(0, idx)) + '.' + sanitizeExternalName(raw.slice(idx + 1))
    };
  }
  return {
    module: '__import__',
    name: sanitizeExternalName(raw),
    complete: '__import__.' + sanitizeExternalName(raw)
  };
}

function stripHtml(html) {
  return String(html || '')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<br\s*\/?\s*>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/\s+\n/g, '\n')
    .replace(/\n\s+/g, '\n')
    .replace(/[ \t]+/g, ' ')
    .trim();
}

function pickFields(fieldsMap, wanted) {
  return wanted.filter(f => fieldsMap && Object.prototype.hasOwnProperty.call(fieldsMap, f));
}

function flattenForSheet(obj, prefix = '', out = {}) {
  if (obj === null || obj === undefined) {
    out[prefix || 'value'] = '';
    return out;
  }
  if (Array.isArray(obj)) {
    if (obj.length && typeof obj[0] === 'object') out[prefix || 'items'] = JSON.stringify(obj);
    else out[prefix || 'items'] = obj.join(', ');
    return out;
  }
  if (typeof obj !== 'object') {
    out[prefix || 'value'] = obj;
    return out;
  }
  for (const [k, v] of Object.entries(obj)) {
    const key = prefix ? `${prefix}.${k}` : k;
    if (v && typeof v === 'object') {
      if (Array.isArray(v) || Object.keys(v).length > 8) out[key] = JSON.stringify(v);
      else flattenForSheet(v, key, out);
    } else {
      out[key] = v;
    }
  }
  return out;
}

class OdooJsonRpc {
  constructor(target) {
    this.target = assertTarget(target);
    this.uid = null;
    this.fieldCache = new Map();
    this.modelOkCache = new Map();
  }

  async rpc(service, method, args) {
    const body = {
      jsonrpc: '2.0',
      method: 'call',
      params: { service, method, args },
      id: Date.now() + Math.floor(Math.random() * 100000)
    };
    const response = await fetch(this.target.url + '/jsonrpc', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
    const text = await response.text();
    let json;
    try {
      json = JSON.parse(text);
    } catch (_) {
      throw new Error(`Odoo mengembalikan non-JSON HTTP ${response.status}: ${text.slice(0, 600)}`);
    }
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}: ${JSON.stringify(json).slice(0, 1200)}`);
    }
    if (json.error) {
      const data = json.error.data || {};
      const msg = data.message || json.error.message || JSON.stringify(json.error);
      const debug = data.debug ? '\n' + String(data.debug).slice(0, 1800) : '';
      throw new Error(msg + debug);
    }
    return json.result;
  }

  async authenticate() {
    if (!this.target.db || !this.target.username || !this.target.password) {
      throw new Error('Database, username, dan password/API key harus diisi.');
    }
    const uid = await this.rpc('common', 'authenticate', [
      this.target.db,
      this.target.username,
      this.target.password,
      {}
    ]);
    if (!uid) throw new Error(`Login Odoo gagal untuk database "${this.target.db}" dengan user "${this.target.username}". Periksa database, username, dan password/API key. Catatan: Odoo Online sering membutuhkan API Key, bukan password login web.`);
    this.uid = uid;
    return uid;
  }

  async version() {
    return this.rpc('common', 'version', []);
  }

  async executeKw(model, method, args = [], kwargs = {}) {
    if (!this.uid) await this.authenticate();
    return this.rpc('object', 'execute_kw', [
      this.target.db,
      this.uid,
      this.target.password,
      model,
      method,
      args,
      kwargs || {}
    ]);
  }

  async fieldsGet(model) {
    if (this.fieldCache.has(model)) return this.fieldCache.get(model);
    const fields = await this.executeKw(model, 'fields_get', [], {
      attributes: ['string', 'type', 'relation', 'required', 'readonly', 'selection']
    });
    this.fieldCache.set(model, fields || {});
    return fields || {};
  }

  async hasModel(model) {
    if (this.modelOkCache.has(model)) return this.modelOkCache.get(model);
    try {
      await this.executeKw(model, 'search_count', [[]], {});
      this.modelOkCache.set(model, true);
      return true;
    } catch (_) {
      this.modelOkCache.set(model, false);
      return false;
    }
  }

  async searchRead(model, domain = [], fields = [], limit = 80, order = 'id desc') {
    const kwargs = { fields, limit };
    if (order) kwargs.order = order;
    return this.executeKw(model, 'search_read', [domain], kwargs);
  }

  async search(model, domain = [], limit = 0, order = '') {
    const kwargs = {};
    if (limit) kwargs.limit = limit;
    if (order) kwargs.order = order;
    return this.executeKw(model, 'search', [domain], kwargs);
  }

  async count(model, domain = []) {
    return this.executeKw(model, 'search_count', [domain], {});
  }

  async read(model, ids, fields = []) {
    if (!ids || !ids.length) return [];
    const kwargs = fields && fields.length ? { fields } : {};
    return this.executeKw(model, 'read', [ids], kwargs);
  }
}

async function findExternal(odoo, externalId) {
  const parsed = parseExternalId(externalId);
  if (!parsed) return null;
  const rows = await odoo.searchRead(
    'ir.model.data',
    [['module', '=', parsed.module], ['name', '=', parsed.name]],
    ['module', 'name', 'model', 'res_id'],
    1,
    'id desc'
  );
  return rows && rows[0] ? rows[0] : null;
}

async function bindExternal(odoo, externalId, model, resId) {
  const parsed = parseExternalId(externalId);
  if (!parsed || !model || !resId) return null;
  const existing = await findExternal(odoo, externalId);
  if (existing) return existing;
  try {
    const id = await odoo.executeKw('ir.model.data', 'create', [{
      module: parsed.module,
      name: parsed.name,
      model,
      res_id: resId,
      noupdate: true
    }], {});
    return { id, module: parsed.module, name: parsed.name, model, res_id: resId };
  } catch (err) {
    return { warning: 'Gagal membuat external ID: ' + err.message };
  }
}

async function resolveExternalIdToId(odoo, externalId, expectedModel) {
  if (isBlank(externalId)) return false;
  const row = await findExternal(odoo, externalId);
  if (!row) throw new Error(`External ID tidak ditemukan: ${externalId}`);
  if (expectedModel && row.model && row.model !== expectedModel) {
    throw new Error(`External ID ${externalId} mengarah ke ${row.model}, bukan ${expectedModel}`);
  }
  return row.res_id;
}

async function prepareValues(odoo, model, row, fieldsMap) {
  const values = {};
  const warnings = [];
  for (const [rawKey, rawValue] of Object.entries(row || {})) {
    const key = String(rawKey || '').trim();
    if (!key || RESERVED_COLUMNS.has(key)) continue;
    if (key.startsWith('__')) continue;
    if (isBlank(rawValue)) continue;

    if (key.endsWith('_external_id')) {
      const base = key.replace(/_external_id$/, '');
      if (!fieldsMap[base]) {
        warnings.push(`Kolom ${key} dilewati: field ${base} tidak ada di ${model}.`);
        continue;
      }
      const field = fieldsMap[base];
      const id = await resolveExternalIdToId(odoo, rawValue, field.relation);
      values[base] = id;
      continue;
    }

    if (key.endsWith('_external_ids')) {
      const base = key.replace(/_external_ids$/, '');
      if (!fieldsMap[base]) {
        warnings.push(`Kolom ${key} dilewati: field ${base} tidak ada di ${model}.`);
        continue;
      }
      const field = fieldsMap[base];
      const ids = [];
      for (const ext of asArray(rawValue)) {
        ids.push(await resolveExternalIdToId(odoo, ext, field.relation));
      }
      if (field.type === 'many2many') values[base] = [[6, 0, ids]];
      else values[base] = ids;
      continue;
    }

    if (!fieldsMap[key]) {
      warnings.push(`Kolom ${key} dilewati: field tidak ada di ${model}.`);
      continue;
    }
    const field = fieldsMap[key];
    if (field.readonly && !['ir.ui.view', 'website.page'].includes(model)) {
      warnings.push(`Kolom ${key} dilewati: readonly.`);
      continue;
    }

    let value = rawValue;
    if (field.type === 'boolean') value = boolish(rawValue);
    else if (['float', 'monetary'].includes(field.type)) value = Number(rawValue);
    else if (field.type === 'integer') value = parseInt(rawValue, 10);
    else if (field.type === 'many2one') {
      if (typeof rawValue === 'number') value = rawValue;
      else if (/^\d+$/.test(String(rawValue))) value = parseInt(rawValue, 10);
      else warnings.push(`Kolom ${key}: many2one sebaiknya memakai ${key}_external_id. Nilai dibiarkan apa adanya.`);
    } else if (field.type === 'many2many') {
      value = [[6, 0, asArray(rawValue).map(v => parseInt(v, 10)).filter(Number.isFinite)]];
    }

    values[key] = value;
  }
  return { values, warnings };
}

async function importRows(odoo, payload) {
  const rows = Array.isArray(payload.rows) ? payload.rows : [];
  const fallbackModel = payload.model || '';
  const allowDelete = !!payload.allowDelete;
  const report = { processed: 0, created: 0, updated: 0, deleted: 0, skipped: 0, errors: [], warnings: [], rows: [] };
  const fieldsByModel = new Map();

  for (let i = 0; i < rows.length; i++) {
    const rowNumber = Number(payload.offset || 0) + i + 2;
    const row = rows[i] || {};
    const model = String(row._model || row.__model || fallbackModel || '').trim();
    const action = String(row.__action || row._action || 'upsert').trim().toLowerCase();
    const externalId = row._external_id || row.external_id || '';
    report.processed += 1;

    try {
      if (!model) throw new Error('Model kosong. Isi _model di sheet atau pilih model target.');
      if (!(await odoo.hasModel(model))) throw new Error(`Model tidak tersedia / tidak bisa diakses: ${model}`);
      if (!fieldsByModel.has(model)) fieldsByModel.set(model, await odoo.fieldsGet(model));
      const fieldsMap = fieldsByModel.get(model);

      let existing = null;
      if (!isBlank(externalId)) existing = await findExternal(odoo, externalId);
      if (!existing && row.id && /^\d+$/.test(String(row.id))) existing = { model, res_id: parseInt(row.id, 10) };

      if (action === 'skip') {
        report.skipped += 1;
        report.rows.push({ row: rowNumber, status: 'skipped', model });
        continue;
      }
      if (['delete', 'unlink'].includes(action)) {
        if (!allowDelete) throw new Error('Delete/unlink diblokir. Aktifkan allowDelete jika benar-benar diperlukan.');
        if (!existing) throw new Error('Tidak ada record/external ID untuk dihapus.');
        await odoo.executeKw(model, 'unlink', [[existing.res_id]], {});
        report.deleted += 1;
        report.rows.push({ row: rowNumber, status: 'deleted', model, id: existing.res_id });
        continue;
      }

      const prepared = await prepareValues(odoo, model, row, fieldsMap);
      for (const w of prepared.warnings) report.warnings.push(`row ${rowNumber}: ${w}`);
      const values = prepared.values;
      if (!Object.keys(values).length) {
        report.skipped += 1;
        report.rows.push({ row: rowNumber, status: 'skipped-empty-values', model });
        continue;
      }

      if (existing && ['upsert', 'update', 'write'].includes(action)) {
        await odoo.executeKw(model, 'write', [[existing.res_id], values], {});
        report.updated += 1;
        report.rows.push({ row: rowNumber, status: 'updated', model, id: existing.res_id, external_id: externalId || null });
      } else {
        if (['update', 'write'].includes(action)) throw new Error('Action update/write tetapi external ID/id tidak ditemukan.');
        const id = await odoo.executeKw(model, 'create', [values], {});
        let bind = null;
        if (!isBlank(externalId)) bind = await bindExternal(odoo, externalId, model, id);
        if (bind && bind.warning) report.warnings.push(`row ${rowNumber}: ${bind.warning}`);
        report.created += 1;
        report.rows.push({ row: rowNumber, status: 'created', model, id, external_id: externalId || null });
      }
    } catch (err) {
      report.errors.push(`row ${rowNumber}: ${err.message}`);
      report.rows.push({ row: rowNumber, status: 'error', model, error: err.message });
      if (payload.stopOnError) break;
    }
  }
  return report;
}

async function fetchAsBase64(url) {
  const response = await fetch(url, { redirect: 'follow' });
  if (!response.ok) throw new Error(`Gagal fetch foto HTTP ${response.status}`);
  const arrayBuffer = await response.arrayBuffer();
  return Buffer.from(arrayBuffer).toString('base64');
}

async function importPhotos(odoo, payload) {
  const rows = Array.isArray(payload.rows) ? payload.rows : [];
  const model = payload.model || 'product.template';
  const imageField = payload.imageField || 'image_1920';
  const report = { processed: 0, done: 0, missing: 0, failed: 0, errors: [], rows: [] };
  const fieldsMap = await odoo.fieldsGet(model);
  if (!fieldsMap[imageField]) throw new Error(`Field foto ${imageField} tidak ada di ${model}.`);

  for (let i = 0; i < rows.length; i++) {
    const rowNumber = Number(payload.offset || 0) + i + 2;
    const row = rows[i] || {};
    report.processed += 1;
    try {
      const ext = row._external_id || row.external_id || row.product_external_id || row.product_tmpl_external_id;
      let id = row.id && /^\d+$/.test(String(row.id)) ? parseInt(row.id, 10) : null;
      if (!id && ext) id = await resolveExternalIdToId(odoo, ext, model);
      if (!id) {
        report.missing += 1;
        report.rows.push({ row: rowNumber, status: 'missing-target' });
        continue;
      }
      let image = row[imageField] || row.image_base64 || row.image || '';
      if (!image && row.image_url) image = await fetchAsBase64(row.image_url);
      if (!image) {
        report.missing += 1;
        report.rows.push({ row: rowNumber, status: 'missing-image', id });
        continue;
      }
      await odoo.executeKw(model, 'write', [[id], { [imageField]: image }], {});
      report.done += 1;
      report.rows.push({ row: rowNumber, status: 'done', id });
    } catch (err) {
      report.failed += 1;
      report.errors.push(`row ${rowNumber}: ${err.message}`);
      report.rows.push({ row: rowNumber, status: 'error', error: err.message });
    }
  }
  return report;
}

async function assistantHealth(odoo) {
  const uid = await odoo.authenticate();
  let version = null;
  try { version = await odoo.version(); } catch (_) {}
  const checks = [];
  for (const model of ['project.project', 'project.task', 'product.template', 'res.partner', 'ir.model', 'ir.model.fields', 'ir.model.data']) {
    const ok = await odoo.hasModel(model);
    checks.push({ model, ok });
  }
  return { uid, version, checks, generated_at: new Date().toISOString() };
}

async function schemaScan(odoo, payload = {}) {
  const models = Array.isArray(payload.models) && payload.models.length ? payload.models : DEFAULT_MODELS;
  const result = { generated_at: new Date().toISOString(), models: [], missing: [] };
  for (const model of models) {
    try {
      if (!(await odoo.hasModel(model))) {
        result.missing.push(model);
        continue;
      }
      const fields = await odoo.fieldsGet(model);
      let count = null;
      try { count = await odoo.count(model, []); } catch (_) {}
      const custom = Object.entries(fields).filter(([name]) => name.startsWith('x_')).map(([name, f]) => ({ name, ...f }));
      result.models.push({
        model,
        count,
        field_count: Object.keys(fields).length,
        custom_field_count: custom.length,
        custom_fields: custom,
        fields: Object.entries(fields).map(([name, f]) => ({ name, ...f }))
      });
    } catch (err) {
      result.models.push({ model, error: err.message });
    }
  }
  return result;
}

async function dataAudit(odoo, payload = {}) {
  const limit = Math.min(Number(payload.limit || 80), 300);
  const audit = { generated_at: new Date().toISOString(), counts: {}, warnings: [], samples: {} };
  for (const model of ['project.project', 'project.task', 'product.template', 'product.product', 'res.partner', 'sale.order', 'account.move']) {
    try {
      if (await odoo.hasModel(model)) audit.counts[model] = await odoo.count(model, []);
    } catch (err) {
      audit.warnings.push(`${model}: ${err.message}`);
    }
  }

  async function sample(model, domain, fields, key) {
    try {
      if (await odoo.hasModel(model)) audit.samples[key || model] = await odoo.searchRead(model, domain, fields, limit, 'write_date desc');
    } catch (err) {
      audit.warnings.push(`${model}: ${err.message}`);
    }
  }

  await sample('product.template', ['|', ['list_price', '=', 0], ['categ_id', '=', false]], ['name', 'default_code', 'list_price', 'categ_id', 'write_date'], 'products_need_attention');
  await sample('res.partner', ['|', ['phone', '=', false], ['mobile', '=', false]], ['name', 'phone', 'mobile', 'email', 'write_date'], 'partners_missing_phone');
  await sample('project.task', [['parent_id', '=', false]], ['name', 'project_id', 'stage_id', 'parent_id', 'write_date'], 'root_tasks');
  await sample('project.task', [['project_id', '=', false]], ['name', 'stage_id', 'parent_id', 'write_date'], 'tasks_without_project');
  return audit;
}

async function contextExport(odoo, payload = {}) {
  const models = Array.isArray(payload.models) && payload.models.length ? payload.models : ['project.project', 'project.task', 'product.template', 'product.category', 'res.partner'];
  const limit = Math.min(Number(payload.limit || 60), 500);
  const out = { generated_at: new Date().toISOString(), target: { url: odoo.target.url, db: odoo.target.db }, models: [] };
  for (const model of models) {
    try {
      if (!(await odoo.hasModel(model))) {
        out.models.push({ model, available: false });
        continue;
      }
      const fieldsMap = await odoo.fieldsGet(model);
      const preferred = pickFields(fieldsMap, [
        'id', 'display_name', 'name', 'active', 'sequence', 'default_code', 'barcode', 'list_price', 'standard_price', 'categ_id',
        'project_id', 'parent_id', 'stage_id', 'user_id', 'user_ids', 'partner_id', 'date_start', 'date', 'date_deadline',
        'description', 'write_date', 'create_date'
      ]).concat(Object.keys(fieldsMap).filter(f => f.startsWith('x_')).slice(0, 30));
      const rows = await odoo.searchRead(model, [], [...new Set(preferred)], limit, 'write_date desc');
      out.models.push({ model, available: true, field_count: Object.keys(fieldsMap).length, rows });
    } catch (err) {
      out.models.push({ model, error: err.message });
    }
  }
  return out;
}

async function projectList(odoo, payload = {}) {
  if (!(await odoo.hasModel('project.project'))) throw new Error('Model project.project tidak tersedia.');
  const fieldsMap = await odoo.fieldsGet('project.project');
  const fields = pickFields(fieldsMap, ['id', 'display_name', 'name', 'active', 'stage_id', 'user_id', 'partner_id', 'date_start', 'date', 'write_date']);
  const domain = payload.includeArchived ? [] : (fieldsMap.active ? [['active', 'in', [true, false]]] : []);
  return odoo.searchRead('project.project', domain, fields, Math.min(Number(payload.limit || 200), 1000), 'write_date desc');
}

function buildTaskTree(tasks) {
  const byId = new Map();
  for (const task of tasks) byId.set(task.id, { ...task, children: [] });
  const roots = [];
  for (const task of byId.values()) {
    const parent = Array.isArray(task.parent_id) ? task.parent_id[0] : task.parent_id;
    if (parent && byId.has(parent)) byId.get(parent).children.push(task);
    else roots.push(task);
  }
  return roots;
}

async function externalIdsFor(odoo, pairs) {
  const result = [];
  const grouped = new Map();
  for (const p of pairs) {
    if (!p.model || !p.id) continue;
    if (!grouped.has(p.model)) grouped.set(p.model, []);
    grouped.get(p.model).push(p.id);
  }
  for (const [model, ids] of grouped.entries()) {
    try {
      const rows = await odoo.searchRead('ir.model.data', [['model', '=', model], ['res_id', 'in', [...new Set(ids)]]], ['module', 'name', 'model', 'res_id'], 1000, 'id desc');
      for (const r of rows) result.push({ ...r, complete_name: `${r.module}.${r.name}` });
    } catch (_) {}
  }
  return result;
}

async function projectContextExport(odoo, payload = {}) {
  const projectId = Number(payload.projectId || payload.project_id || 0);
  if (!projectId) throw new Error('Pilih project terlebih dahulu.');

  const out = {
    generated_at: new Date().toISOString(),
    target: { url: odoo.target.url, db: odoo.target.db },
    project_id: projectId,
    project: null,
    tasks: [],
    task_tree: [],
    milestones: [],
    updates: [],
    messages: [],
    external_ids: [],
    related: {},
    chatgpt_prompt: ''
  };

  const projectFieldsMap = await odoo.fieldsGet('project.project');
  const projectFields = pickFields(projectFieldsMap, [
    'id', 'display_name', 'name', 'active', 'stage_id', 'user_id', 'partner_id', 'date_start', 'date', 'write_date',
    'description', 'label_tasks', 'alias_name', 'analytic_account_id', 'privacy_visibility'
  ]).concat(Object.keys(projectFieldsMap).filter(f => f.startsWith('x_')));
  const projects = await odoo.read('project.project', [projectId], [...new Set(projectFields)]);
  if (!projects.length) throw new Error('Project tidak ditemukan atau tidak bisa dibaca.');
  out.project = projects[0];

  if (await odoo.hasModel('project.task')) {
    const taskFieldsMap = await odoo.fieldsGet('project.task');
    const taskFields = pickFields(taskFieldsMap, [
      'id', 'display_name', 'name', 'active', 'project_id', 'parent_id', 'child_ids', 'stage_id', 'user_id', 'user_ids',
      'partner_id', 'priority', 'kanban_state', 'date_deadline', 'planned_hours', 'remaining_hours', 'effective_hours',
      'progress', 'description', 'write_date', 'create_date', 'sequence'
    ]).concat(Object.keys(taskFieldsMap).filter(f => f.startsWith('x_')));
    out.tasks = await odoo.searchRead('project.task', [['project_id', '=', projectId]], [...new Set(taskFields)], Math.min(Number(payload.taskLimit || 1500), 3000), 'parent_id asc, sequence asc, id asc');
    out.task_tree = buildTaskTree(out.tasks);
  }

  if (await odoo.hasModel('project.milestone')) {
    const mf = await odoo.fieldsGet('project.milestone');
    const fields = pickFields(mf, ['id', 'name', 'project_id', 'deadline', 'is_reached', 'reached_date', 'write_date']).concat(Object.keys(mf).filter(f => f.startsWith('x_')));
    out.milestones = await odoo.searchRead('project.milestone', [['project_id', '=', projectId]], [...new Set(fields)], 300, 'deadline asc, id asc');
  }

  if (await odoo.hasModel('project.update')) {
    const uf = await odoo.fieldsGet('project.update');
    const fields = pickFields(uf, ['id', 'name', 'project_id', 'status', 'progress', 'date', 'description', 'write_date']).concat(Object.keys(uf).filter(f => f.startsWith('x_')));
    out.updates = await odoo.searchRead('project.update', [['project_id', '=', projectId]], [...new Set(fields)], 200, 'write_date desc');
    for (const u of out.updates) if (u.description) u.description_text = stripHtml(u.description);
  }

  const pairs = [{ model: 'project.project', id: projectId }, ...out.tasks.map(t => ({ model: 'project.task', id: t.id }))];
  out.external_ids = await externalIdsFor(odoo, pairs);

  if (await odoo.hasModel('mail.message')) {
    const taskIds = out.tasks.map(t => t.id);
    const domains = [[['model', '=', 'project.project'], ['res_id', '=', projectId]]];
    if (taskIds.length) domains.push([['model', '=', 'project.task'], ['res_id', 'in', taskIds]]);
    const messages = [];
    for (const domain of domains) {
      const rows = await odoo.searchRead('mail.message', domain, ['id', 'model', 'res_id', 'date', 'author_id', 'subject', 'body', 'message_type', 'subtype_id'], Math.min(Number(payload.messageLimit || 120), 300), 'date desc');
      messages.push(...rows.map(r => ({ ...r, body_text: stripHtml(r.body).slice(0, 3000), body: undefined })));
    }
    out.messages = messages;
  }

  const analytic = out.project && out.project.analytic_account_id;
  if (Array.isArray(analytic) && analytic[0] && await odoo.hasModel('account.analytic.account')) {
    try {
      out.related.analytic_account = await odoo.read('account.analytic.account', [analytic[0]], ['id', 'name', 'code', 'partner_id', 'write_date']);
    } catch (_) {}
  }

  const projectName = out.project.display_name || out.project.name || `Project ${projectId}`;
  out.chatgpt_prompt = [
    `Saya mengekspor konteks project Odoo Lokalmart: ${projectName}.`,
    'Tolong baca JSON ini, pahami perkembangan terakhir project, hierarki task/subtask, update, milestone, chatter ringkas, dan external ID.',
    'Setelah itu kembangkan isi project agar sesuai dengan diskusi Lokalmart terbaru: Ground Zero sebagai manajemen ide, Lokal ID sebagai portal user Odoo dengan role bertambah, Agen/Dropshipper, POS/Kasir Agen, katalog, dan alur koloni/koperasi.',
    'Berikan usulan struktur task/subtask yang lebih rapi, alasan bisnis, dan bila diminta buatkan XLSX import-ready yang tidak merusak record lama.',
    '',
    'JSON PROJECT CONTEXT:',
    JSON.stringify(out, null, 2)
  ].join('\n');

  return out;
}

function workbookToBase64(sheets) {
  const wb = XLSX.utils.book_new();
  for (const sheet of sheets) {
    const rows = Array.isArray(sheet.rows) ? sheet.rows : [];
    const ws = XLSX.utils.json_to_sheet(rows.length ? rows : [{ kosong: 'Tidak ada data' }]);
    XLSX.utils.book_append_sheet(wb, ws, String(sheet.name || 'Sheet').slice(0, 31));
  }
  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  return buf.toString('base64');
}

function asXlsxExport(name, sheets) {
  return {
    filename: name,
    mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    base64: workbookToBase64(sheets)
  };
}

function contextToWorkbook(context, filename = 'lokalmart_context_export.xlsx') {
  const sheets = [];
  if (context.project) sheets.push({ name: 'project', rows: [flattenForSheet(context.project)] });
  if (context.tasks) sheets.push({ name: 'tasks', rows: context.tasks.map(r => flattenForSheet(r)) });
  if (context.milestones) sheets.push({ name: 'milestones', rows: context.milestones.map(r => flattenForSheet(r)) });
  if (context.updates) sheets.push({ name: 'updates', rows: context.updates.map(r => flattenForSheet(r)) });
  if (context.messages) sheets.push({ name: 'messages', rows: context.messages.map(r => flattenForSheet(r)) });
  if (context.external_ids) sheets.push({ name: 'external_ids', rows: context.external_ids.map(r => flattenForSheet(r)) });
  if (context.models) {
    sheets.push({ name: 'models', rows: context.models.map(m => ({ model: m.model, available: m.available, row_count: (m.rows || []).length, field_count: m.field_count, error: m.error || '' })) });
    for (const m of context.models) if (m.rows && m.rows.length) sheets.push({ name: m.model.replace(/[^A-Za-z0-9]/g, '_').slice(0, 31), rows: m.rows.map(r => flattenForSheet(r)) });
  }
  if (!sheets.length) sheets.push({ name: 'context', rows: [flattenForSheet(context)] });
  return asXlsxExport(filename, sheets);
}

async function fullExportXlsx(odoo, payload = {}) {
  const models = Array.isArray(payload.models) && payload.models.length ? payload.models : ['project.project', 'project.task', 'product.template', 'product.category', 'res.partner'];
  const limit = Math.min(Number(payload.limit || 500), 3000);
  const sheets = [];
  const manifest = [];
  for (const model of models) {
    try {
      if (!(await odoo.hasModel(model))) {
        manifest.push({ model, status: 'missing' });
        continue;
      }
      const fieldsMap = await odoo.fieldsGet(model);
      const fields = Object.keys(fieldsMap).filter(f => !['image_1920', 'image_1024', 'image_512', 'image_256', 'image_128'].includes(f)).slice(0, 120);
      const rows = await odoo.searchRead(model, [], fields, limit, 'id asc');
      sheets.push({ name: model.replace(/[^A-Za-z0-9]/g, '_').slice(0, 31), rows: rows.map(r => flattenForSheet(r)) });
      manifest.push({ model, status: 'ok', rows: rows.length, fields: fields.length });
    } catch (err) {
      manifest.push({ model, status: 'error', error: err.message });
    }
  }
  sheets.unshift({ name: 'manifest', rows: manifest });
  return asXlsxExport(`lokalmart_full_export_${new Date().toISOString().slice(0, 10)}.xlsx`, sheets);
}

async function handleAction(body) {
  const action = body.action;
  const target = body.target || {};
  const payload = body.payload || {};
  if (!action) throw new Error('Action kosong.');

  const noAuthActions = new Set([]);
  const odoo = new OdooJsonRpc(target);

  if (!noAuthActions.has(action)) await odoo.authenticate();

  if (action === 'test_connection') {
    let version = null;
    try { version = await odoo.version(); } catch (_) {}
    return { ok: true, uid: odoo.uid, version };
  }

  if (action === 'import_rows') return { ok: true, report: await importRows(odoo, payload) };
  if (action === 'import_photos') return { ok: true, report: await importPhotos(odoo, payload) };

  if (action === 'ai_health') return { ok: true, assistant: await assistantHealth(odoo) };
  if (action === 'ai_schema_scan') return { ok: true, scan: await schemaScan(odoo, payload) };
  if (action === 'ai_data_audit') return { ok: true, audit: await dataAudit(odoo, payload) };
  if (action === 'ai_context_export') return { ok: true, context: await contextExport(odoo, payload) };
  if (action === 'ai_export_xlsx') {
    const context = await contextExport(odoo, payload);
    return { ok: true, export: contextToWorkbook(context, `lokalmart_context_${new Date().toISOString().slice(0, 10)}.xlsx`) };
  }
  if (action === 'ai_project_list') return { ok: true, projects: await projectList(odoo, payload) };
  if (action === 'ai_project_context_export') return { ok: true, context: await projectContextExport(odoo, payload) };
  if (action === 'ai_project_xlsx_export') {
    const context = await projectContextExport(odoo, payload);
    const safeName = String((context.project && (context.project.name || context.project.display_name)) || 'project').replace(/[^A-Za-z0-9_-]/g, '_').slice(0, 50);
    return { ok: true, context, export: contextToWorkbook(context, `lokalmart_project_${safeName}_${new Date().toISOString().slice(0, 10)}.xlsx`) };
  }
  if (action === 'full_export_xlsx') return { ok: true, export: await fullExportXlsx(odoo, payload) };

  if (action === 'read_rpc') {
    const model = String(payload.model || '').trim();
    const method = String(payload.method || '').trim();
    if (!model || !method) throw new Error('model dan method wajib diisi.');
    if (!SAFE_READ_METHODS.has(method)) throw new Error('Method ini diblokir. Gunakan method read-only saja.');
    const result = await odoo.executeKw(model, method, payload.args || [], payload.kwargs || {});
    return { ok: true, result };
  }

  throw new Error(`Action tidak dikenal: ${action}`);
}


function corsHeaders(extra = {}) {
  return {
    'Content-Type': 'application/json; charset=utf-8',
    'Cache-Control': 'no-store',
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
    ...extra
  };
}

function response(statusCode, payload, extraHeaders) {
  return {
    statusCode,
    headers: corsHeaders(extraHeaders),
    body: JSON.stringify(payload)
  };
}

exports.handler = async function handler(event) {
  if (event.httpMethod === 'OPTIONS') {
    return {
      statusCode: 204,
      headers: corsHeaders(),
      body: ''
    };
  }

  if (event.httpMethod !== 'POST') {
    return response(405, { ok: false, error: 'Gunakan POST ke /api/odoo.' });
  }

  try {
    let body = {};
    const rawBody = event.isBase64Encoded
      ? Buffer.from(event.body || '', 'base64').toString('utf8')
      : (event.body || '');

    if (rawBody.trim()) {
      if (Buffer.byteLength(rawBody, 'utf8') > MAX_BODY_SIZE) {
        throw new Error('Payload terlalu besar. Kurangi batch atau file.');
      }
      try {
        body = JSON.parse(rawBody);
      } catch (err) {
        throw new Error('Body bukan JSON valid: ' + err.message);
      }
    }

    const result = await handleAction(body);
    return response(200, result);
  } catch (err) {
    const simple = normalizeError(err);
    return response(500, { ok: false, error: simple.message, detail: simple.detail });
  }
};
