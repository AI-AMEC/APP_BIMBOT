// ==========================================
// BIMBOT v41.0 - 12_AutodeskService
// Autodesk Platform Services (APS) Data Management: hubs -> projects for BIM 360 / ACC (Forma).
//
// Setup (once, by an editor of this script):
//   1. Project Settings -> Script properties: APS_CLIENT_ID, APS_CLIENT_SECRET (your APS app).
//   2. In each ACC/BIM 360 account (hub) you want to see: Account Admin -> Custom Integrations ->
//      add the APS app's Client ID. Hubs without it will not be listed.
//   3. Run diagnoseAutodesk() from the editor.
// App-level (2-legged) token: BIMBOT's own user permissions still gate every call.
//
// Speed: one index of every hub + project is built with parallel requests (UrlFetchApp.fetchAll)
// and cached for 6 hours; the picker, link resolve and auto-match all read from it.
// ==========================================

const APS_BASE_URL_ = 'https://developer.api.autodesk.com';
const APS_INDEX_CACHE_SECONDS_ = 300; // CacheService maximum (6h)
const APS_PAGE_LIMIT_ = 200;            // Data Management maximum page size
const APS_PROJECT_ID_RE_ = /projects\/(?:b\.)?([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})/i;

function apsCredentials_() {
  const props = PropertiesService.getScriptProperties();
  const id = String(props.getProperty('APS_CLIENT_ID') || '').trim();
  const secret = String(props.getProperty('APS_CLIENT_SECRET') || '').trim();
  return id && secret ? { id: id, secret: secret } : null;
}

function apsRequireUser_() {
  const user = getAuthenticatedUserContext_();
  if (!user.success) throw new Error(user.error || 'User is not authorized.');
  return user;
}

const APS_NOT_CONFIGURED_ = 'החיבור ל-Autodesk טרם הוגדר. מנהל מערכת יכול להזין את פרטי אפליקציית ה-APS בחלון "בחירה מ-Autodesk".';

// Requests a 2-legged token. Returns {token, expiresIn} or throws with Autodesk's reason.
function apsRequestToken_(cred) {
  const res = UrlFetchApp.fetch(APS_BASE_URL_ + '/authentication/v2/token', {
    method: 'post',
    headers: { Authorization: 'Basic ' + Utilities.base64Encode(cred.id + ':' + cred.secret) },
    contentType: 'application/x-www-form-urlencoded',
    payload: { grant_type: 'client_credentials', scope: 'data:read account:read' },
    muteHttpExceptions: true
  });
  let body = {};
  try { body = JSON.parse(res.getContentText()); } catch (e) { body = {}; }
  if (res.getResponseCode() !== 200 || !body.access_token) {
    const reason = body.errorCode || body.error || body.developerMessage || body.error_description || res.getContentText().slice(0, 200);
    throw new Error('Autodesk סירב לפרטי האפליקציה (' + res.getResponseCode() + ': ' + reason + '). בדוק את ה-Client ID וה-Client Secret.');
  }
  return { token: body.access_token, expiresIn: Number(body.expires_in || 3600) };
}

function apsToken_() { return apsUserToken_(); }

function apsNotConfigured_(user) {
  return { success: false, notConfigured: true, canConfigure: !!(user && user.canViewDatabase), error: APS_NOT_CONFIGURED_ };
}

function apsHttpError_(code, text) {
  if (code === 401) throw new Error('חיבור Autodesk פג או בוטל. יש להתחבר מחדש לחשבון האישי.');
  if (code === 401 || code === 403) {
    return new Error('אין הרשאה ב-Autodesk (' + code + '). ודא שאפליקציית ה-APS נוספה כ-Custom Integration בחשבון ה-ACC/BIM 360.');
  }
  return new Error('שגיאת Autodesk ' + code + ': ' + String(text || '').slice(0, 300));
}

// Parallel GETs. Returns [{code, body}] in request order; never throws for HTTP errors.
function apsFetchAll_(paths) {
  if (!paths.length) return [];
  const token = apsToken_();
  const responses = UrlFetchApp.fetchAll(paths.map(p => ({
    url: p.indexOf('http') === 0 ? p : APS_BASE_URL_ + p,
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  })));
  return responses.map(r => {
    const code = r.getResponseCode();
    let body = null;
    try { body = JSON.parse(r.getContentText()); } catch (e) { body = null; }
    return { code: code, body: body, text: code === 200 ? '' : r.getContentText() };
  });
}

// Cache values may exceed CacheService's 100KB-per-key limit; store in chunks.
function apsCachePut_(key, value, seconds) {
  key=apsUserCacheEpoch_()+':'+key;
  const json = JSON.stringify(value), size = 90000, entries = {};
  const count = Math.ceil(json.length / size) || 1;
  for (let i = 0; i < count; i++) entries[key + ':' + i] = json.slice(i * size, (i + 1) * size);
  entries[key + ':n'] = String(count);
  CacheService.getUserCache().putAll(entries, seconds);
}
function apsCacheGet_(key) {
  key=apsUserCacheEpoch_()+':'+key;
  const cache = CacheService.getUserCache(), n = Number(cache.get(key + ':n') || 0);
  if (!n) return null;
  const keys = []; for (let i = 0; i < n; i++) keys.push(key + ':' + i);
  const parts = cache.getAll(keys);
  if (keys.some(k => parts[k] == null)) return null;
  try { return JSON.parse(keys.map(k => parts[k]).join('')); } catch (e) { return null; }
}

function apsProjectUrl_(project, type) {
  const href = project.links && project.links.webView && project.links.webView.href;
  if (href) return href;
  const pid = String(project.id || '').replace(/^b\./, '');
  return type === 'ACC' ? 'https://acc.autodesk.com/docs/files/projects/' + pid : 'https://docs.b360.autodesk.com/projects/' + pid;
}

function apsMapProject_(p, hubId) {
  const attrs = p.attributes || {};
  const data = attrs.extension && attrs.extension.data || {};
  const type = data.projectType === 'ACC' ? 'ACC' : (data.projectType === 'BIM360' ? 'BIM360' : String(data.projectType || ''));
  return { hubId: hubId, id: p.id, name: attrs.name || '', type: type, url: apsProjectUrl_(p, type) };
}

function apsProjectsPath_(hubId, page) {
  return '/project/v1/hubs/' + encodeURIComponent(hubId) + '/projects?page%5Bnumber%5D=' + page + '&page%5Blimit%5D=' + APS_PAGE_LIMIT_;
}

// Every hub the APS app can see + all their projects. Pages are fetched 5 at a time in parallel.
function apsBuildIndex_() {
  const hubRes = apsFetchAll_(['/project/v1/hubs'])[0];
  if (hubRes.code !== 200) throw apsHttpError_(hubRes.code, hubRes.text);
  const hubs = ((hubRes.body && hubRes.body.data) || [])
    .filter(h => h.attributes && h.attributes.extension && h.attributes.extension.type === 'hubs:autodesk.bim360:Account')
    .map(h => ({ id: h.id, name: h.attributes.name || h.id, region: h.attributes.region || '', error: '' }))
    .sort((a, b) => a.name.localeCompare(b.name));

  const projects = [], seen = {};
  let queue = hubs.map(h => ({ hub: h, page: 0 }));
  for (let round = 0; queue.length && round < 20; round++) {
    const results = apsFetchAll_(queue.map(q => apsProjectsPath_(q.hub.id, q.page)));
    const maxQueued = {};
    queue.forEach(q => { maxQueued[q.hub.id] = Math.max(maxQueued[q.hub.id] == null ? -1 : maxQueued[q.hub.id], q.page); });
    const next = [];
    results.forEach((r, i) => {
      const q = queue[i];
      if (r.code !== 200) { if (q.page === 0) q.hub.error = apsHttpError_(r.code, r.text).message; return; } // later pages: past the end
      const data = (r.body && r.body.data) || [];
      data.forEach(p => { if (!seen[p.id]) { seen[p.id] = true; projects.push(apsMapProject_(p, q.hub.id)); } });
      // Continue only from the last page of this batch; a short/empty page means the hub is done.
      const full = data.length >= APS_PAGE_LIMIT_;
      if (full && q.page === maxQueued[q.hub.id]) {
        for (let k = 1; k <= 5; k++) next.push({ hub: q.hub, page: q.page + k });
      }
    });
    queue = next;
  }
  projects.sort((a, b) => a.name.localeCompare(b.name));
  return { builtAt: new Date().toISOString(), hubs: hubs, projects: projects };
}

function apsIndex_(refresh) {
  apsUserToken_();
  if (!refresh) { const cached = apsCacheGet_('APS_INDEX'); if (cached) return cached; }
  const lock = LockService.getScriptLock();
  lock.waitLock(60000);
  try {
    if (!refresh) { const cached = apsCacheGet_('APS_INDEX'); if (cached) return cached; } // built by a parallel call
    const index = apsBuildIndex_();
    apsCachePut_('APS_INDEX', index, APS_INDEX_CACHE_SECONDS_);
    return index;
  } finally { lock.releaseLock(); }
}

function apsHubById_(index, hubId) { return index.hubs.filter(h => h.id === hubId)[0] || null; }

// Scores an Autodesk project name against a BIMBOT project (code / number / Hebrew name).
function apsMatchScore_(apsName, project) {
  const n = String(apsName || '').toLowerCase();
  const esc = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const code = String(project.projectCode || '').toLowerCase().trim();
  const no = String(project.no || '').trim();
  const name = String(project.name || '').toLowerCase().trim();
  let score = 0;
  if (code && code !== '—' && new RegExp('(^|[^a-z0-9])' + esc(code) + '([^a-z0-9]|$)').test(n)) score += 3;
  if (no && new RegExp('(^|\\D)' + esc(no) + '(\\D|$)').test(n)) score += 3;
  if (name && name.length >= 3 && n.indexOf(name) >= 0) score += 2;
  return score;
}

function apsAuthorizedProject_(projectNo) {
  const list = getCachedProjectData();
  if (!list.success) throw new Error(list.error || 'Cannot verify project access.');
  const project = list.data.filter(p => p.no === String(projectNo || '').trim())[0];
  if (!project) throw new Error('הפרויקט אינו מורשה למשתמש.');
  return project;
}

function apsWithHub_(index, project) {
  const hub = apsHubById_(index, project.hubId) || { id: project.hubId, name: project.hubId };
  return { hub: { id: hub.id, name: hub.name, region: hub.region || '' }, project: project };
}

// ---------- Client API (google.script.run) ----------

function apsGetStatus() {
  try {
    const user = apsRequireUser_();
    const cred = apsCredentials_();
    return { success: true, configured: !!cred, canConfigure: !!user.canViewDatabase,
      clientIdHint: cred ? cred.id.slice(0, 6) + '…' : '' };
  } catch (e) { return { success: false, error: String(e.message || e) }; }
}

// Admin-only: tests the APS app credentials against Autodesk and saves them only if they work.
// The secret is stored in Script Properties and never sent back to the browser.
function apsSaveCredentials(clientId, clientSecret) {
  try {
    const user = apsRequireUser_();
    if (!user.canViewDatabase) throw new Error('רק מנהל מערכת (הרשאת מסד נתונים) יכול להגדיר את החיבור ל-Autodesk.');
    const cred = { id: String(clientId || '').trim(), secret: String(clientSecret || '').trim() };
    if (!/^[A-Za-z0-9_-]{16,128}$/.test(cred.id)) throw new Error('ה-Client ID אינו תקין (העתק אותו מ-aps.autodesk.com → My Apps).');
    if (!/^[\x21-\x7E]{8,256}$/.test(cred.secret)) throw new Error('ה-Client Secret אינו תקין.');

    const t = apsRequestToken_(cred);
    const res = UrlFetchApp.fetch(APS_BASE_URL_ + '/project/v1/hubs', { headers: { Authorization: 'Bearer ' + t.token }, muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) throw apsHttpError_(res.getResponseCode(), res.getContentText());
    const hubs = (JSON.parse(res.getContentText()).data || [])
      .filter(h => h.attributes && h.attributes.extension && h.attributes.extension.type === 'hubs:autodesk.bim360:Account')
      .map(h => h.attributes.name);

    PropertiesService.getScriptProperties().setProperties({ APS_CLIENT_ID: cred.id, APS_CLIENT_SECRET: cred.secret });
    const cache = CacheService.getUserCache();
    cache.remove('APS_TOKEN_2L');
    cache.remove('APS_INDEX:n');
    console.log('APS credentials updated by ' + user.email + '; hubs visible: ' + hubs.length);
    return { success: true, hubs: hubs };
  } catch (e) { return { success: false, error: String(e.message || e) }; }
}

function apsListHubs(refresh) {
  try {
    const user = apsRequireUser_();
    if (!apsCredentials_()) return apsNotConfigured_(user);
    const index = apsIndex_(refresh === true);
    return { success: true, hubs: index.hubs, builtAt: index.builtAt };
  } catch (e) { return { success: false, error: String(e.message || e) }; }
}

function apsListProjects(hubId, refresh) {
  try {
    apsRequireUser_();
    const index = apsIndex_(refresh === true);
    const hub = apsHubById_(index, String(hubId || ''));
    if (!hub) throw new Error('ה-HUB לא נמצא. לחץ רענון.');
    if (hub.error) throw new Error(hub.name + ': ' + hub.error);
    return { success: true, projects: index.projects.filter(p => p.hubId === hub.id) };
  } catch (e) { return { success: false, error: String(e.message || e) }; }
}

// Finds which hub owns the project in a pasted BIM 360 / ACC link.
function apsResolveProjectLink(url) {
  try {
    apsRequireUser_();
    const m = APS_PROJECT_ID_RE_.exec(String(url || ''));
    if (!m) throw new Error('לא נמצא מזהה פרויקט Autodesk בקישור. הדבק קישור של BIM 360 / ACC שמכיל /projects/<id>.');
    const projectId = 'b.' + m[1].toLowerCase();
    let index = apsIndex_(false);
    let found = index.projects.filter(p => p.id.toLowerCase() === projectId)[0];
    if (!found) { index = apsIndex_(true); found = index.projects.filter(p => p.id.toLowerCase() === projectId)[0]; }
    if (!found) throw new Error('הפרויקט מהקישור לא נמצא באף HUB שמחובר לאפליקציה (' + index.hubs.length + ' HUBs נבדקו). ייתכן שה-HUB של הלקוח לא הוסיף את אפליקציית ה-APS.');
    return Object.assign({ success: true }, apsWithHub_(index, found));
  } catch (e) { return { success: false, error: String(e.message || e) }; }
}

// Automatic: finds the Autodesk project for a BIMBOT project by code / number / name across all hubs.
// auto = the single best match when it is unambiguous (code or number matched, no tie).
function apsAutoMatchProject(projectNo) {
  try {
    const user = apsRequireUser_();
    if (!apsCredentials_()) return { success: true, configured: false, canConfigure: !!user.canViewDatabase, matches: [], auto: null };
    const project = apsAuthorizedProject_(projectNo);
    const index = apsIndex_(false);
    const scored = index.projects
      .map(p => ({ p: p, score: apsMatchScore_(p.name, project) }))
      .filter(x => x.score > 0)
      .sort((a, b) => b.score - a.score || a.p.name.localeCompare(b.p.name));
    const matches = scored.slice(0, 15).map(x => Object.assign({ score: x.score }, apsWithHub_(index, x.p)));
    const top = scored[0];
    const unique = top && top.score >= 3 && !(scored[1] && scored[1].score === top.score);
    return { success: true, configured: true, matches: matches, auto: unique ? matches[0] : null, builtAt: index.builtAt };
  } catch (e) { return { success: false, error: String(e.message || e) }; }
}

// Editor-only check: Run -> diagnoseAutodesk
function diagnoseAutodesk() {
  const cred = apsCredentials_();
  Logger.log('APS credentials configured: ' + !!cred);
  if (!cred) return;
  try {
    const started = Date.now();
    const index = apsIndex_(true);
    Logger.log('Index built in ' + ((Date.now() - started) / 1000).toFixed(1) + 's: ' + index.hubs.length + ' hubs, ' + index.projects.length + ' projects');
    index.hubs.forEach(h => Logger.log('  ' + h.name + '  ' + h.id + '  ' + h.region + '  projects=' +
      index.projects.filter(p => p.hubId === h.id).length + (h.error ? '  ERROR: ' + h.error : '')));
  } catch (e) { Logger.log('FAILED: ' + e.message); }
}

