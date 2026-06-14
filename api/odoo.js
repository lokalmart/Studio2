'use strict';

const XLSX = require('xlsx');

const DEFAULT_MODELS = [
  'project.project',
  'project.task',
  'project.milestone',
  'project.update',
  'knowledge.article',
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
  if (req.body && typeof req.body === 'object') return req.body;
  if (typeof req.body === 'string' && req.body.trim()) return JSON.parse(req.body);

  const chunks = [];
  let total = 0;
  for await (const chunk of req) {
    total += chunk.length;
    if (total > MAX_BODY_SIZE) {
      throw new Error('Payload terlalu besar. Kurangi batch atau file.');
    }
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
  if (!/^https?:\/\//i.test(value)) {
    throw new Error('URL Odoo harus diawali http:// atau https://');
  }
  return value;
}

function assertTarget(target) {
  const t = target || {};
  const clean = {
    url: cleanUrl(t.url || t.host),
    db: String(t.db || t.database || '').trim(),
    username: String(t.username || t.login || t.email || '').trim(),
    password: String(t.password || t.apiKey || t.key || '').trim()
  };
  if (!clean.db) throw new Error('Database Odoo belum diisi.');
  if (!clean.username) throw new Error('Username/email Odoo belum diisi.');
  if (!clean.password) throw new Error('Password/API Key Odoo belum diisi.');
  return clean;
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
  return String(value).split(/[;,|]/).map(v => v.trim()).filter(Boolean);
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
    const module = sanitizeExternalName(raw.slice(0, idx));
    const name = sanitizeExternalName(raw.slice(idx + 1));
    return { module, name, complete: module + '.' + name };
  }
  return { module: '__import__', name: sanitizeExternalName(raw), complete: '__import__.' + sanitizeExternalName(raw) };
}

function stripHtml(html) {
  return String(html || '')
    .replace(/<br\s*\/?>/gi, '\n')
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

function pickFields(fieldsMap, wanted) {
  return wanted.filter(f => fieldsMap && Object.prototype.hasOwnProperty.call(fieldsMap, f));
}

function workbookResponse(workbook, filename) {
  const buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });
  return {
    filename,
    fileName: filename,
    mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    base64: buffer.toString('base64')
  };
}

function addSheet(workbook, name, rows) {
  const safeName = String(name || 'Sheet').replace(/[\\/?*\[\]:]/g, ' ').slice(0, 31) || 'Sheet';
  const ws = XLSX.utils.json_to_sheet(rows && rows.length ? rows : [{ note: 'Tidak ada data' }]);
  XLSX.utils.book_append_sheet(workbook, ws, safeName);
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
    if (!uid) {
      throw new Error(
        `Login Odoo gagal untuk database "${this.target.db}" dengan user "${this.target.username}". ` +
        'Periksa database, username, dan password/API key. Untuk Odoo Online, gunakan API Key sebagai pengganti password login web.'
      );
    }
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
        if (!allowDelete) throw new Error('Delete/unlink diblokir. Aktifkan allowDelete bila benar-benar perlu.');
        if (!existing) throw new Error('Record delete tidak ditemukan.');
        await odoo.executeKw(model, 'unlink', [[existing.res_id]], {});
        report.deleted += 1;
        report.rows.push({ row: rowNumber, status: 'deleted', model, id: existing.res_id });
        continue;
      }

      const prepared = await prepareValues(odoo, model, row, fieldsMap);
      report.warnings.push(...prepared.warnings.map(w => `row ${rowNumber}: ${w}`));
      if (!Object.keys(prepared.values).length) {
        report.skipped += 1;
        report.rows.push({ row: rowNumber, status: 'skipped-empty', model });
        continue;
      }

      if (existing && existing.res_id) {
        await odoo.executeKw(model, 'write', [[existing.res_id], prepared.values], {});
        report.updated += 1;
        report.rows.push({ row: rowNumber, status: 'updated', model, id: existing.res_id });
      } else {
        const id = await odoo.executeKw(model, 'create', [prepared.values], {});
        if (!isBlank(externalId)) await bindExternal(odoo, externalId, model, id);
        report.created += 1;
        report.rows.push({ row: rowNumber, status: 'created', model, id });
      }
    } catch (err) {
      report.errors.push(`row ${rowNumber}: ${err.message}`);
      report.rows.push({ row: rowNumber, status: 'error', model, error: err.message });
      if (payload.stopOnError) break;
    }
  }
  return report;
}

async function testConnection(odoo) {
  const [uid, version] = await Promise.all([odoo.authenticate(), odoo.version()]);
  return {
    ok: true,
    uid,
    db: odoo.target.db,
    username: odoo.target.username,
    server_version: version && (version.server_version || version.server_serie),
    version
  };
}

async function schemaScan(odoo, payload = {}) {
  const models = Array.isArray(payload.models) && payload.models.length ? payload.models : DEFAULT_MODELS;
  const result = { generated_at: new Date().toISOString(), models: [] };
  for (const model of models) {
    try {
      if (!(await odoo.hasModel(model))) {
        result.models.push({ model, available: false, error: 'Tidak tersedia / tidak bisa diakses' });
        continue;
      }
      const fields = await odoo.fieldsGet(model);
      const count = await odoo.count(model, []);
      const fieldRows = Object.entries(fields).map(([name, f]) => ({
        name,
        label: f.string || name,
        type: f.type,
        relation: f.relation || '',
        required: !!f.required,
        readonly: !!f.readonly
      })).sort((a, b) => a.name.localeCompare(b.name));
      result.models.push({ model, available: true, count, fields: fieldRows, field_count: fieldRows.length });
    } catch (err) {
      result.models.push({ model, available: false, error: err.message });
    }
  }
  return result;
}

async function dataAudit(odoo, payload = {}) {
  const scan = await schemaScan(odoo, payload);
  const audit = {
    generated_at: scan.generated_at,
    counts: {},
    warnings: [],
    models: scan.models.map(m => ({ model: m.model, available: m.available, count: m.count || 0, field_count: m.field_count || 0, error: m.error || '' }))
  };
  for (const m of scan.models) {
    if (m.available) audit.counts[m.model] = m.count || 0;
    else audit.warnings.push(`${m.model}: ${m.error || 'tidak tersedia'}`);
  }

  try {
    if (await odoo.hasModel('project.task')) {
      const fields = await odoo.fieldsGet('project.task');
      if (fields.parent_id && fields.project_id) {
        audit.project_task_without_project = await odoo.count('project.task', [['project_id', '=', false]]);
      }
    }
  } catch (err) {
    audit.warnings.push('Audit task gagal: ' + err.message);
  }
  return audit;
}

async function contextExport(odoo, payload = {}) {
  const models = Array.isArray(payload.models) && payload.models.length ? payload.models : DEFAULT_MODELS;
  const limit = Math.min(Number(payload.limit || 25), 100);
  const out = {
    type: 'lokalmart_studio_context_export',
    generated_at: new Date().toISOString(),
    target: { url: odoo.target.url, db: odoo.target.db, username: odoo.target.username },
    models: []
  };
  for (const model of models) {
    try {
      if (!(await odoo.hasModel(model))) {
        out.models.push({ model, available: false, error: 'Tidak tersedia / tidak bisa diakses' });
        continue;
      }
      const fields = await odoo.fieldsGet(model);
      const fieldNames = Object.keys(fields).slice(0, 80);
      const sampleFields = pickFields(fields, ['id', 'name', 'display_name', 'active', 'create_date', 'write_date', 'parent_id', 'project_id', 'stage_id', 'user_ids', 'partner_id']);
      const rows = await odoo.searchRead(model, [], sampleFields.length ? sampleFields : fieldNames.slice(0, 12), limit, 'id desc');
      out.models.push({
        model,
        available: true,
        count: await odoo.count(model, []),
        fields: Object.entries(fields).map(([name, f]) => ({ name, label: f.string || name, type: f.type, relation: f.relation || '' })),
        sample: rows
      });
    } catch (err) {
      out.models.push({ model, available: false, error: err.message });
    }
  }
  return out;
}

async function contextWorkbook(context, filename) {
  const wb = XLSX.utils.book_new();
  addSheet(wb, 'summary', [{
    type: context.type,
    generated_at: context.generated_at,
    db: context.target && context.target.db,
    url: context.target && context.target.url,
    models: context.models ? context.models.length : 0
  }]);
  for (const m of context.models || []) {
    addSheet(wb, (m.model || 'model').replace(/\./g, '_').slice(0, 28), (m.sample || []).map(r => flattenForSheet(r)));
  }
  addSheet(wb, 'fields', (context.models || []).flatMap(m => (m.fields || []).map(f => ({ model: m.model, ...f }))));
  return workbookResponse(wb, filename);
}

async function listProjects(odoo) {
  const fieldsMap = await odoo.fieldsGet('project.project');
  const fields = pickFields(fieldsMap, ['id', 'name', 'display_name', 'active', 'partner_id', 'user_id', 'date_start', 'date', 'stage_id', 'write_date', 'create_date']);
  return {
    projects: await odoo.searchRead('project.project', [], fields.length ? fields : ['id', 'name'], 200, 'write_date desc')
  };
}

function m2oId(value) {
  if (Array.isArray(value)) return value[0];
  if (typeof value === 'number') return value;
  if (/^\d+$/.test(String(value || ''))) return parseInt(value, 10);
  return null;
}

function makeTaskTree(tasks) {
  const byId = new Map();
  const roots = [];
  for (const task of tasks || []) {
    byId.set(task.id, { ...task, children: [] });
  }
  for (const task of byId.values()) {
    const pid = m2oId(task.parent_id);
    if (pid && byId.has(pid)) byId.get(pid).children.push(task);
    else roots.push(task);
  }
  return roots;
}

async function externalIdsFor(odoo, model, ids) {
  if (!ids.length) return [];
  try {
    return await odoo.searchRead('ir.model.data', [['model', '=', model], ['res_id', 'in', ids]], ['module', 'name', 'model', 'res_id'], 1000, 'id asc');
  } catch (_) {
    return [];
  }
}

async function projectContextExport(odoo, payload = {}) {
  const projectId = Number(payload.projectId || payload.project_id || payload.id);
  if (!projectId) throw new Error('Pilih project terlebih dahulu. projectId kosong.');

  const projectFieldsMap = await odoo.fieldsGet('project.project');
  const projectFields = pickFields(projectFieldsMap, [
    'id', 'name', 'display_name', 'active', 'description', 'partner_id', 'user_id', 'stage_id',
    'date_start', 'date', 'create_date', 'write_date', 'company_id', 'privacy_visibility'
  ]);
  const projectRows = await odoo.read('project.project', [projectId], projectFields.length ? projectFields : ['id', 'name']);
  const project = projectRows[0];
  if (!project) throw new Error('Project tidak ditemukan: ' + projectId);

  const taskFieldsMap = await odoo.fieldsGet('project.task');
  const taskFields = pickFields(taskFieldsMap, [
    'id', 'name', 'display_name', 'active', 'project_id', 'parent_id', 'stage_id', 'milestone_id',
    'user_ids', 'partner_id', 'priority', 'sequence', 'description', 'date_deadline', 'planned_date_begin',
    'create_date', 'write_date', 'x_studio_chapter', 'x_lm_reason', 'x_lm_category', 'x_lm_priority'
  ]);
  const tasks = await odoo.searchRead('project.task', [['project_id', '=', projectId]], taskFields.length ? taskFields : ['id', 'name', 'parent_id', 'project_id'], 2000, 'sequence asc, id asc');

  let milestones = [];
  if (await odoo.hasModel('project.milestone')) {
    const fields = await odoo.fieldsGet('project.milestone');
    const mFields = pickFields(fields, ['id', 'name', 'project_id', 'deadline', 'is_reached', 'create_date', 'write_date']);
    if (fields.project_id) {
      milestones = await odoo.searchRead('project.milestone', [['project_id', '=', projectId]], mFields.length ? mFields : ['id', 'name'], 500, 'id asc');
    }
  }

  let updates = [];
  if (await odoo.hasModel('project.update')) {
    const fields = await odoo.fieldsGet('project.update');
    const uFields = pickFields(fields, ['id', 'name', 'project_id', 'description', 'status', 'progress', 'date', 'create_date', 'write_date']);
    if (fields.project_id) {
      updates = await odoo.searchRead('project.update', [['project_id', '=', projectId]], uFields.length ? uFields : ['id', 'name'], 200, 'id desc');
      updates = updates.map(u => ({ ...u, description_text: stripHtml(u.description) }));
    }
  }

  let messages = [];
  if (await odoo.hasModel('mail.message')) {
    const ids = tasks.map(t => t.id);
    const domains = [['model', '=', 'project.project'], ['res_id', '=', projectId]];
    try {
      const projectMessages = await odoo.searchRead('mail.message', domains, ['id', 'subject', 'body', 'date', 'author_id', 'model', 'res_id'], 80, 'date desc');
      let taskMessages = [];
      if (ids.length) {
        taskMessages = await odoo.searchRead('mail.message', [['model', '=', 'project.task'], ['res_id', 'in', ids.slice(0, 500)]], ['id', 'subject', 'body', 'date', 'author_id', 'model', 'res_id'], 160, 'date desc');
      }
      messages = [...projectMessages, ...taskMessages]
        .sort((a, b) => String(b.date || '').localeCompare(String(a.date || '')))
        .slice(0, 200)
        .map(m => ({ ...m, body_text: stripHtml(m.body).slice(0, 2000) }));
    } catch (err) {
      messages = [{ error: 'Gagal membaca chatter: ' + err.message }];
    }
  }

  const taskExternal = await externalIdsFor(odoo, 'project.task', tasks.map(t => t.id));
  const projectExternal = await externalIdsFor(odoo, 'project.project', [projectId]);
  const milestoneExternal = await externalIdsFor(odoo, 'project.milestone', milestones.map(m => m.id));

  const context = {
    type: 'lokalmart_project_context_export',
    generated_at: new Date().toISOString(),
    target: { url: odoo.target.url, db: odoo.target.db, username: odoo.target.username },
    project,
    counts: {
      tasks: tasks.length,
      root_tasks: tasks.filter(t => !m2oId(t.parent_id)).length,
      milestones: milestones.length,
      updates: updates.length,
      messages: messages.length
    },
    tasks,
    task_tree: makeTaskTree(tasks),
    milestones,
    updates,
    messages,
    external_ids: {
      project: projectExternal,
      tasks: taskExternal,
      milestones: milestoneExternal
    }
  };

  context.prompt = [
    'Saya akan memberikan JSON konteks project Odoo Lokalmart hasil export dari Studio2.',
    'Tolong baca perkembangan terakhir project ini, pahami struktur task/subtask, chatter/update, milestone, dan external ID-nya.',
    'Setelah itu bantu kembangkan isi project agar sesuai dengan diskusi Lokalmart terbaru, tanpa membuat task liar tanpa parent, dan siapkan rekomendasi patch XLSX/import yang aman.',
    '',
    'JSON_CONTEXT:',
    JSON.stringify(context, null, 2)
  ].join('\n');

  return context;
}

async function projectWorkbook(context) {
  const wb = XLSX.utils.book_new();
  addSheet(wb, 'project', [flattenForSheet(context.project)]);
  addSheet(wb, 'counts', [context.counts]);
  addSheet(wb, 'tasks', (context.tasks || []).map(t => ({ ...flattenForSheet(t), description_text: stripHtml(t.description) })));
  addSheet(wb, 'milestones', (context.milestones || []).map(m => flattenForSheet(m)));
  addSheet(wb, 'updates', (context.updates || []).map(u => flattenForSheet(u)));
  addSheet(wb, 'messages', (context.messages || []).map(m => flattenForSheet(m)));
  addSheet(wb, 'external_ids_tasks', context.external_ids && context.external_ids.tasks ? context.external_ids.tasks : []);
  addSheet(wb, 'prompt', [{ prompt: context.prompt }]);
  const name = sanitizeExternalName((context.project && (context.project.name || context.project.display_name)) || 'project');
  return workbookResponse(wb, `lokalmart_project_context_${name}.xlsx`);
}

async function genericExport(odoo, payload = {}) {
  const models = Array.isArray(payload.models) && payload.models.length ? payload.models : DEFAULT_MODELS;
  const limit = Math.min(Number(payload.limit || 1000), 5000);
  const wb = XLSX.utils.book_new();
  const summary = [];
  for (const model of models) {
    try {
      if (!(await odoo.hasModel(model))) {
        summary.push({ model, available: false, error: 'Tidak tersedia / tidak bisa diakses' });
        continue;
      }
      const fields = await odoo.fieldsGet(model);
      const wanted = payload.fields && payload.fields[model] ? payload.fields[model] : ['id', 'display_name', 'name', 'active', 'create_date', 'write_date'];
      const safeFields = pickFields(fields, wanted);
      if (!safeFields.includes('id')) safeFields.unshift('id');
      const rows = await odoo.searchRead(model, payload.domain || [], safeFields, limit, 'id desc');
      summary.push({ model, available: true, rows: rows.length, exported_fields: safeFields.join(', ') });
      addSheet(wb, model.replace(/\./g, '_').slice(0, 31), rows.map(r => flattenForSheet(r)));
    } catch (err) {
      summary.push({ model, available: false, error: err.message });
    }
  }
  addSheet(wb, 'summary', summary);
  return workbookResponse(wb, `lokalmart_export_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

async function readRpc(odoo, payload = {}) {
  const model = String(payload.model || '').trim();
  const method = String(payload.method || '').trim();
  if (!model || !method) throw new Error('model dan method wajib diisi.');
  if (!SAFE_READ_METHODS.has(method)) throw new Error('Method tidak diizinkan untuk readOnlyRpc: ' + method);
  const args = Array.isArray(payload.args) ? payload.args : [];
  const kwargs = payload.kwargs || {};
  return { result: await odoo.executeKw(model, method, args, kwargs) };
}

async function barcodeLookup(odoo, payload = {}) {
  const barcode = String(payload.barcode || payload.code || '').trim();
  if (!barcode) throw new Error('Barcode kosong.');
  const fields = ['id', 'display_name', 'name', 'barcode', 'list_price', 'standard_price', 'default_code'];
  const products = await odoo.searchRead('product.product', [['barcode', '=', barcode]], fields, 20, 'id desc');
  return { barcode, products };
}

async function handleAction(body) {
  const action = String(body.action || body.__action || '').trim();
  const payload = body.payload && typeof body.payload === 'object' ? body.payload : body;
  const odoo = new OdooJsonRpc(body.target || payload.target || {});

  switch (action) {
    case 'health':
    case 'ai_health':
      return { ok: true, service: 'Lokalmart Studio2 Vercel API', time: new Date().toISOString() };

    case 'test_connection':
    case 'test':
      return testConnection(odoo);

    case 'ai_schema_scan':
    case 'schema_scan':
    case 'scan_schema':
      return { schema: await schemaScan(odoo, payload) };

    case 'ai_data_audit':
    case 'data_audit':
      return { audit: await dataAudit(odoo, payload) };

    case 'ai_context_export':
    case 'context_export':
      return { context: await contextExport(odoo, payload) };

    case 'ai_export_xlsx': {
      const context = await contextExport(odoo, payload);
      return await contextWorkbook(context, 'lokalmart_context_export.xlsx');
    }

    case 'ai_project_list':
    case 'project_list':
      return await listProjects(odoo);

    case 'ai_project_context_export':
    case 'project_context_export':
      return { context: await projectContextExport(odoo, payload) };

    case 'ai_project_xlsx_export': {
      const context = await projectContextExport(odoo, payload);
      return await projectWorkbook(context);
    }

    case 'export_records':
    case 'full_export':
    case 'partial_export':
      return await genericExport(odoo, payload);

    case 'import_rows':
    case 'import_xlsx_rows':
    case 'import_batch':
    case 'import':
      return { report: await importRows(odoo, payload) };

    case 'barcode_lookup':
    case 'scan_barcode':
      return await barcodeLookup(odoo, payload);

    case 'ai_read_rpc':
    case 'read_rpc':
      return await readRpc(odoo, payload);

    default:
      throw new Error('Action tidak dikenal: ' + (action || '(kosong)'));
  }
}

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');

  if (req.method === 'OPTIONS') {
    res.statusCode = 204;
    res.end();
    return;
  }

  if (req.method === 'GET') {
    sendJson(res, 200, {
      ok: true,
      service: 'Lokalmart Studio2 Vercel API',
      usage: 'POST JSON ke /api/odoo dengan { action, target, payload }',
      actions: [
        'test_connection',
        'ai_schema_scan',
        'ai_data_audit',
        'ai_context_export',
        'ai_export_xlsx',
        'ai_project_list',
        'ai_project_context_export',
        'ai_project_xlsx_export',
        'export_records',
        'import_rows',
        'barcode_lookup'
      ]
    });
    return;
  }

  if (req.method !== 'POST') {
    sendJson(res, 405, { ok: false, error: 'Method tidak diizinkan. Gunakan POST.' });
    return;
  }

  try {
    const body = await readJsonBody(req);
    const data = await handleAction(body);
    sendJson(res, 200, { ok: true, ...data });
  } catch (err) {
    sendJson(res, 500, { ok: false, error: normalizeError(err) });
  }
};
