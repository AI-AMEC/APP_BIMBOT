// ==========================================
// BIMBOT v41.0 - 04_ProjectIdentity
// Resolves a project to its REAL Drive folder -> BIM folder -> "BIM INFO" spreadsheet.
// All project reads/writes go through getLocalProjectDbAndBimFolder_().
// ==========================================

// Single switch for every write path (sheet sync, workflow, kickoff, publish review, governance).
function canWriteProject_(project) {
  return !APP_READ_ONLY && !!project && !project.ambiguousIdentity;
}

function openDriveFolder_(folderId, label) {
  try {
    return DriveApp.getFolderById(folderId);
  } catch (e) {
    throw new Error('לא ניתן לגשת ל' + label + ' (ID: ' + folderId + '). ודא שלחשבון שמריץ את האפליקציה יש הרשאה לתיקייה. ' + e.message);
  }
}

function readProjectIdentifier_(ss) {
  const info = getSheetIgnoreCase(ss, 'BIM_Info');
  if (!info) return { sheet: null, rows: [], value: '' };
  const rows = info.getDataRange().getDisplayValues();
  const matches = rows.filter(r => String(r[0]).trim().toLowerCase() === 'project identifier');
  if (matches.length > 1) throw new Error('Multiple "Project Identifier" rows in BIM_Info. Recovery is required.');
  return { sheet: info, rows: rows, value: matches.length ? String(matches[0][1]).trim() : '' };
}

// Legacy sheets (created by the original app) may lack the identifier; stamp it on first write.
function stampProjectIdentifier_(ss, projectNo) {
  let info = getSheetIgnoreCase(ss, 'BIM_Info');
  if (!info) info = ss.insertSheet('BIM_Info');
  const rows = info.getDataRange().getDisplayValues();
  const index = rows.findIndex(r => String(r[0]).trim().toLowerCase() === 'project identifier');
  if (index >= 0) info.getRange(index + 1, 2).setValue(projectNo);
  else info.appendRow(['Project Identifier', projectNo]);
  SpreadsheetApp.flush();
}

// Returns a registered spreadsheet only if it still exists inside the project's BIM folder.
function registeredProjectFile_(record, bimFolder) {
  if (!record || !record.spreadsheetId || !bimFolder || record.bimFolderId !== bimFolder.getId()) return null;
  try {
    const file = DriveApp.getFileById(record.spreadsheetId);
    if (file.isTrashed()) return null;
    const parents = file.getParents();
    while (parents.hasNext()) if (parents.next().getId() === bimFolder.getId()) return file;
  } catch (e) { /* deleted or no access: fall back to discovery */ }
  return null;
}

function discoverProjectFile_(bimFolder, project) {
  const exactName = project.no + ' - ' + project.name + ' - BIM INFO';
  const candidates = [];
  const files = bimFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (files.hasNext()) {
    const f = files.next();
    const name = f.getName();
    if (name.startsWith(project.no + ' - ') && name.endsWith(' - BIM INFO')) candidates.push(f);
  }
  if (candidates.length <= 1) return candidates[0] || null;
  const exact = candidates.filter(f => f.getName() === exactName);
  if (exact.length === 1) return exact[0];
  throw new Error('נמצאו ' + candidates.length + ' גיליונות BIM INFO לפרויקט ' + project.no + ' בתיקיית BIM. יש להשאיר גיליון אחד בלבד.');
}

function getLocalProjectDbAndBimFolder_(folderId, projectNo, projectName, createIfMissing) {
  const create = createIfMissing !== false;
  const result = getCachedProjectData();
  if (!result.success) throw new Error(result.error || 'Cannot verify project access.');
  const no = String(projectNo == null ? '' : projectNo).trim();
  const matches = result.data.filter(p => p.no === no);
  if (!no || matches.length !== 1) throw new Error('הפרויקט ' + no + ' לא נמצא ברשימת הפרויקטים המורשים למשתמש.');
  const project = matches[0];
  if (project.ambiguousIdentity) throw new Error('זהות הפרויקט ' + no + ' אינה חד-משמעית (מספר/תיקייה/קוד כפולים במקור). יש לתקן את גיליונות המקור.');
  if (!project.folderId) throw new Error('לפרויקט ' + no + ' אין תיקיית Drive באינדקס הפרויקטים (INDEX_DB_ID).');
  if (String(folderId || '').trim() !== project.folderId) {
    throw new Error('Project folder does not match the canonical index. Refresh the project.');
  }
  if (create && !canWriteProject_(project)) {
    throw new Error('השמירה חסומה: האפליקציה במצב קריאה בלבד (APP_READ_ONLY).');
  }

  const lock = create ? LockService.getScriptLock() : null;
  if (lock) lock.waitLock(30000);
  try {
    const root = openDriveFolder_(project.folderId, 'תיקיית הפרויקט');
    const folders = root.getFoldersByName('BIM');
    let bimFolder = folders.hasNext() ? folders.next() : null;
    if (folders.hasNext()) throw new Error('בתיקיית הפרויקט יש יותר מתיקיית BIM אחת. יש להשאיר אחת בלבד.');
    if (!bimFolder && !create) return { ss: null, bimFolder: null, project: project };
    if (!bimFolder) bimFolder = root.createFolder('BIM');

    const properties = PropertiesService.getScriptProperties();
    const key = 'BIM_PROJECT_DB:' + no;
    const registered = properties.getProperty(key);
    let record = null;
    try { record = registered ? JSON.parse(registered) : null; } catch (e) { record = null; }

    let file = registeredProjectFile_(record, bimFolder);
    if (!file) {
      if (record) properties.deleteProperty(key); // stale: moved, trashed, or folder changed
      record = null;
      file = discoverProjectFile_(bimFolder, project);
    }
    if (!file && !create) return { ss: null, bimFolder: bimFolder, project: project };

    let isNew = false;
    if (!file) {
      file = DriveApp.getFileById(TEMPLATE_DB_ID).makeCopy(no + ' - ' + project.name + ' - BIM INFO', bimFolder);
      isNew = true;
    }
    const ss = SpreadsheetApp.openById(file.getId());

    const identifier = readProjectIdentifier_(ss);
    if (!isNew && identifier.value && identifier.value !== no) {
      throw new Error('הגיליון ' + file.getName() + ' מסומן כפרויקט ' + identifier.value + ' ולא ' + no + '. לא בוצעה פעולה.');
    }
    if (create && (isNew || !identifier.value)) stampProjectIdentifier_(ss, no);
    if (create && !record) {
      properties.setProperty(key, JSON.stringify({ folderId: project.folderId, bimFolderId: bimFolder.getId(), spreadsheetId: ss.getId() }));
    }
    if (isNew) safeProtect_(ss.getId());
    return { ss: ss, bimFolder: bimFolder, project: project };
  } finally {
    if (lock) lock.releaseLock();
  }
}
