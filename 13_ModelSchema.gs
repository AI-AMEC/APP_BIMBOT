// ==========================================
// BIMBOT v41.0 - 13_ModelSchema
// "סכמת מודל" (MDC step 5): per-project 3D layout of lots / buildings / basements.
// Only geometry + colour live here. Names, floors and relationships stay in the
// step-1 tabs (Lots / Buildings) - the schema is keyed by those codes.
// Sheet "Model_Schema": one readable row per entity.
// ==========================================

const MODEL_SCHEMA_SHEET_ = 'Model_Schema';
const MODEL_SCHEMA_HEADERS_ = ['Type', 'Code', 'PosX', 'PosZ', 'RelX', 'RelZ', 'Width', 'Depth', 'Color'];
const MODEL_SCHEMA_TYPES_ = { LOT: 'lots', BUILDING: 'buildings', BASEMENT: 'basements' };
const MODEL_SCHEMA_MAX_ROWS_ = 2000;

function modelSchemaNumber_(value, label) {
  if (value === '' || value == null) return null;
  const n = Number(value);
  if (!isFinite(n) || Math.abs(n) > 100000) throw new Error('Invalid model schema value: ' + label);
  return Math.round(n * 10) / 10;
}

// Normalizes client input to {version, lots:{code:geom}, buildings:{...}, basements:{...}}.
function normalizeModelSchema_(schema) {
  if (schema == null) return null;
  if (typeof schema !== 'object' || Array.isArray(schema)) throw new Error('Invalid project table: modelSchema');
  const out = { version: 1, lots: {}, buildings: {}, basements: {} };
  let count = 0;
  Object.keys(MODEL_SCHEMA_TYPES_).forEach(type => {
    const key = MODEL_SCHEMA_TYPES_[type];
    const group = schema[key] || {};
    if (typeof group !== 'object' || Array.isArray(group)) throw new Error('Invalid model schema section: ' + key);
    Object.keys(group).forEach(code => {
      if (++count > MODEL_SCHEMA_MAX_ROWS_) throw new Error('Model schema is too large.');
      const g = group[code];
      if (!code || code.length > 200 || !g || typeof g !== 'object') throw new Error('Invalid model schema entry in ' + key);
      const color = String(g.color || '');
      out[key][code] = {
        posX: modelSchemaNumber_(g.posX, code), posZ: modelSchemaNumber_(g.posZ, code),
        relX: modelSchemaNumber_(g.relX, code), relZ: modelSchemaNumber_(g.relZ, code),
        width: modelSchemaNumber_(g.width, code), depth: modelSchemaNumber_(g.depth, code),
        color: /^#[0-9a-f]{6}$/i.test(color) ? color : ''
      };
    });
  });
  return out;
}

function writeModelSchema_(ss, schema) {
  const normalized = normalizeModelSchema_(schema);
  if (!normalized) return;
  let sheet = getSheetIgnoreCase(ss, MODEL_SCHEMA_SHEET_);
  if (!sheet) sheet = ss.insertSheet(MODEL_SCHEMA_SHEET_);
  localWipe(sheet);
  const rows = [];
  Object.keys(MODEL_SCHEMA_TYPES_).forEach(type => {
    const section = normalized[MODEL_SCHEMA_TYPES_[type]];
    Object.keys(section).sort().forEach(code => {
      const g = section[code];
      rows.push([type, code, g.posX, g.posZ, g.relX, g.relZ, g.width, g.depth, g.color].map(v => v == null ? '' : v));
    });
  });
  sheet.getRange(1, 1, 1, MODEL_SCHEMA_HEADERS_.length).setValues([MODEL_SCHEMA_HEADERS_]).setFontWeight('bold').setHorizontalAlignment('center');
  if (rows.length) sheet.getRange(2, 1, rows.length, MODEL_SCHEMA_HEADERS_.length).setValues(rows).setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
}

// Finds the project's BEP folder under BIM (any folder whose name contains "BEP"); creates "BEP" if none.
function modelSchemaBepFolder_(bimFolder) {
  const matches = [];
  const folders = bimFolder.getFolders();
  while (folders.hasNext()) {
    const f = folders.next();
    if (/(^|[^A-Za-z])BEP([^A-Za-z]|$)/i.test(f.getName())) matches.push(f); // "BEP", "02. BEP", "BEP_2025"
  }
  if (matches.length > 1) {
    const exact = matches.filter(f => f.getName().trim().toUpperCase() === 'BEP');
    if (exact.length === 1) return exact[0];
    throw new Error('נמצאו כמה תיקיות BEP בתיקיית BIM (' + matches.map(f => f.getName()).join(', ') + '). השאר אחת בלבד.');
  }
  return matches[0] || bimFolder.createFolder('BEP');
}

// Client: saves the step-5 PNG (data URL) into the project's BIM/BEP folder. Returns the Drive link.
function saveModelSchemaImage(folderId, projectNo, projectName, dataUrl) {
  try {
    const m = /^data:image\/png;base64,([A-Za-z0-9+/=]+)$/.exec(String(dataUrl || ''));
    if (!m) throw new Error('התמונה אינה PNG תקין.');
    if (m[1].length > 12 * 1024 * 1024) throw new Error('התמונה גדולה מדי לשמירה.');
    const db = getLocalProjectDbAndBimFolder_(folderId, projectNo, projectName, true);
    const folder = modelSchemaBepFolder_(db.bimFolder);
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HHmm');
    const name = db.project.no + ' - ' + db.project.name + ' - סכמת מודל - ' + stamp + '.png';
    const blob = Utilities.newBlob(Utilities.base64Decode(m[1]), 'image/png', name);
    const file = folder.createFile(blob);
    return { success: true, url: file.getUrl(), name: name, folderUrl: folder.getUrl(), folderName: folder.getName() };
  } catch (e) { return { success: false, error: String(e.message || e) }; }
}

function readModelSchema_(ss) {
  const sheet = ss ? getSheetIgnoreCase(ss, MODEL_SCHEMA_SHEET_) : null;
  if (!sheet || sheet.getLastRow() <= 1) return null;
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h || '').trim());
  if (headers.slice(0, MODEL_SCHEMA_HEADERS_.length).join('|') !== MODEL_SCHEMA_HEADERS_.join('|')) {
    console.log('Model_Schema headers changed; ignoring saved layout.');
    return null;
  }
  const raw = { lots: {}, buildings: {}, basements: {} };
  values.slice(1).forEach(r => {
    const key = MODEL_SCHEMA_TYPES_[String(r[0] || '').trim().toUpperCase()];
    const code = String(r[1] == null ? '' : r[1]).trim();
    if (!key || !code) return;
    raw[key][code] = { posX: r[2], posZ: r[3], relX: r[4], relZ: r[5], width: r[6], depth: r[7], color: r[8] };
  });
  try { return normalizeModelSchema_(raw); } catch (e) { console.log('Model_Schema unreadable: ' + e); return null; }
}
