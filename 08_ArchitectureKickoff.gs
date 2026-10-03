// ==========================================
// BIMBOT v41.0 - 08_ArchitectureKickoff
// ==========================================
// Architecture kickoff drafts, generated versions, and reviews are append-only.
// This history is separate from BIM workflow milestones and Autodesk publication.
const ARCH_KICKOFF_LOG_SHEET_ = 'BIM_Architecture_Kickoff_Log';
const ARCH_KICKOFF_LOG_HEADERS_ = ['Revision','Event Type','Version ID','At','Actor','Event JSON'];
const ARCH_KICKOFF_MAX_EVENT_CHARS_ = 46000; // under Sheets' per-cell text limit

function architectureKickoffClone_(value) {
  return JSON.parse(JSON.stringify(value == null ? null : value));
}
function architectureKickoffStableJson_(value) {
  if (Array.isArray(value)) return '[' + value.map(architectureKickoffStableJson_).join(',') + ']';
  if (value && typeof value === 'object') return '{' + Object.keys(value).sort().map(function(key) { return JSON.stringify(key) + ':' + architectureKickoffStableJson_(value[key]); }).join(',') + '}';
  return JSON.stringify(value);
}
// Missing / null / "" are the same "empty" value; text is compared trimmed.
// (Without this, a field the browser never set - undefined - never equals "" from the sheet.)
function architectureKickoffComparable_(value) {
  if (value === undefined || value === null) return '';
  if (typeof value === 'string') return value.trim();
  if (typeof value === 'number' || typeof value === 'boolean') return String(value);
  return architectureKickoffStableJson_(value);
}
function architectureKickoffSourceMatches_(submitted, source) {
  if (!submitted || !source || typeof submitted !== 'object' || typeof source !== 'object') return false;
  return Object.keys(source).every(function(key) { return architectureKickoffComparable_(submitted[key]) === architectureKickoffComparable_(source[key]); });
}function architectureKickoffValidatePayload_(payload) {
  if (!payload || typeof payload !== 'object' || Array.isArray(payload)) throw new Error('יש לשמור נתוני התנעה תקינים.');
  if (!payload.bimInfo || typeof payload.bimInfo !== 'object' || !payload.kickoff || typeof payload.kickoff !== 'object') {
    throw new Error('נתוני הפרויקט והתנעת האדריכלות חסרים.');
  }
  const meeting = payload.meeting || {};
  ['date','attendees','notes','surveyLink'].forEach(function(key) {
    if (meeting[key] != null && String(meeting[key]).length > 5000) throw new Error('אחד משדות הפגישה ארוך מדי.');
  });
  if (meeting.consultantWorkViews != null && (!Array.isArray(meeting.consultantWorkViews) || meeting.consultantWorkViews.length > 100 || meeting.consultantWorkViews.some(function(v) { return typeof v !== 'string' || v.length > 500; }))) {
    throw new Error('רשימת מבטי היועצים אינה תקינה.');
  }
  if (meeting.sourceRefs != null && (!Array.isArray(meeting.sourceRefs) || meeting.sourceRefs.length > 100)) {
    throw new Error('רשימת האסמכתאות אינה תקינה.');
  }
  const safe = architectureKickoffClone_(payload);
  if (JSON.stringify(safe).length > 20000) throw new Error('נתוני ההתנעה חורגים מגודל הגרסה המותר.');
  return safe;
}
function architectureKickoffActor_() {
  const context = getAuthenticatedUserContext_();
  if (!context || !context.success || !context.email) throw new Error((context && context.error) || 'לא ניתן לאמת משתמש.');
  return String(context.email).toLowerCase().trim();
}
function architectureKickoffResolve_(folderId, projectNo, projectName, forWrite) {
  const actor = architectureKickoffActor_();
  const resolved = getLocalProjectDbAndBimFolder_(folderId, projectNo, projectName, !!forWrite);
  return {resolved:resolved,actor:actor};
}
function architectureKickoffSheet_(ss, create) {
  let sheet = ss.getSheetByName(ARCH_KICKOFF_LOG_SHEET_);
  if (!sheet && create) {
    sheet = ss.insertSheet(ARCH_KICKOFF_LOG_SHEET_);
    sheet.getRange(1,1,1,ARCH_KICKOFF_LOG_HEADERS_.length).setValues([ARCH_KICKOFF_LOG_HEADERS_]);
    sheet.setFrozenRows(1);
    if (typeof safeProtect_ === 'function') safeProtect_(ss.getId());
  }
  if (sheet && sheet.getLastRow() > 0) {
    const headers = sheet.getRange(1,1,1,ARCH_KICKOFF_LOG_HEADERS_.length).getDisplayValues()[0];
    if (headers.join('|') !== ARCH_KICKOFF_LOG_HEADERS_.join('|')) throw new Error('יומן התנעת האדריכלות אינו תואם לסכמה; נדרשת בדיקה.');
  }
  return sheet;
}
function architectureKickoffRead_(ss) {
  const state = {revision:0,draft:null,draftRevision:0,draftUpdatedAt:'',draftUpdatedBy:'',versions:[]};
  const sheet = architectureKickoffSheet_(ss,false);
  if (!sheet || sheet.getLastRow() <= 1) return state;
  const rows = sheet.getRange(2,1,sheet.getLastRow()-1,ARCH_KICKOFF_LOG_HEADERS_.length).getValues();
  const byId = Object.create(null);
  rows.forEach(function(row) {
    const revision = Number(row[0]), type = String(row[1] || ''), id = String(row[2] || ''), at = String(row[3] || ''), actor = String(row[4] || '');
    let event;
    try { event = JSON.parse(String(row[5] || '')); } catch(e) { throw new Error('יומן התנעת האדריכלות פגום; נשמרו הנתונים, נדרשת התאוששות.'); }
    if (!Number.isInteger(revision) || revision !== state.revision + 1 || !event || typeof event !== 'object' || !at || !actor) {
      throw new Error('רצף גרסאות התנעת האדריכלות אינו תקין; נדרשת התאוששות.');
    }
    state.revision = revision;
    if (type === 'DRAFT_SAVED') {
      if (!event.payload || typeof event.payload !== 'object') throw new Error('טיוטת התנעה פגומה; נדרשת התאוששות.');
      state.draft = event.payload;
      state.draftRevision = revision;
      state.draftUpdatedAt = at;
      state.draftUpdatedBy = actor;
    } else if (type === 'GENERATION_STARTED') {
      if (!id || byId[id]) throw new Error('יומן גרסאות מכיל מזהה כפול; נדרשת התאוששות.');
      if (state.versions.some(function(v) { return v.status === 'GENERATING'; }) || !event.version || !event.version.sourceSnapshot || !event.version.inputSnapshot) throw new Error('אירוע הפקה פגום או מקביל; נדרשת התאוששות.');
      const version = event.version;
      version.id = id;
      version.createdAt = at;
      version.createdBy = actor;
      version.status = 'GENERATING';
      delete version.documentUrl;
      version.reviewNote = '';
      version.reviewedAt = '';
      version.reviewedBy = '';
      byId[id] = version;
      state.versions.push(version);
    } else if (type === 'GENERATION_FINISHED' || type === 'GENERATION_FAILED' || type === 'REVIEWED') {
      const version = byId[id];
      if (!version) throw new Error('אירוע היסטוריה מפנה לגרסה חסרה; נדרשת התאוששות.');
      if (type === 'GENERATION_FINISHED') { if (version.status !== 'GENERATING' || !String(event.documentUrl || '')) throw new Error('תוצאת הפקה פגומה; נדרשת התאוששות.'); version.documentUrl = String(event.documentUrl); version.status = 'GENERATED'; }
      if (type === 'GENERATION_FAILED') { if (version.status !== 'GENERATING') throw new Error('אירוע כשל הפקה פגום; נדרשת התאוששות.'); version.status = 'GENERATION_FAILED'; version.failure = String(event.error || ''); version.orphanUrl = String(event.orphanUrl || ''); }
      if (type === 'REVIEWED') {
        if (['GENERATED','APPROVED','CHANGES_REQUESTED'].indexOf(version.status) < 0 || !version.documentUrl || ['APPROVED','CHANGES_REQUESTED'].indexOf(event.decision) < 0 || (event.decision === 'CHANGES_REQUESTED' && !String(event.note || '').trim())) throw new Error('אירוע סקירה פגום או מוקדם מדי; נדרשת התאוששות.');
        version.status = event.decision;
        version.reviewNote = String(event.note || '');
        version.reviewedAt = at;
        version.reviewedBy = actor;
      }
    } else throw new Error('סוג אירוע לא מוכר ביומן התנעת האדריכלות.');
  });
  return state;
}
function architectureKickoffAppend_(ss, state, eventType, versionId, actor, event) {
  const json = JSON.stringify(event);
  if (json.length > ARCH_KICKOFF_MAX_EVENT_CHARS_) throw new Error('צילום הנתונים גדול מדי לשמירה בטוחה ביומן; לא נשמרה גרסה.');
  const sheet = architectureKickoffSheet_(ss,true);
  const revision = state.revision + 1;
  sheet.getRange(sheet.getLastRow()+1,1,1,ARCH_KICKOFF_LOG_HEADERS_.length)
    .setValues([[revision,eventType,versionId || '',new Date().toISOString(),actor,json]]);
  SpreadsheetApp.flush();
  return revision;
}
function architectureKickoffResponse_(state, project) {
  return {success:true,revision:state.revision,draft:state.draft,draftRevision:state.draftRevision,
    draftUpdatedAt:state.draftUpdatedAt,draftUpdatedBy:state.draftUpdatedBy,
    writable:canWriteProject_(project),
    versions:state.versions.map(function(v) { return architectureKickoffClone_(v); })};
}
function architectureKickoffSaveLegacyFields_(resolved, payload) {
  const sheet = typeof getSheetIgnoreCase === 'function' ? getSheetIgnoreCase(resolved.ss,'Kick_Off') : resolved.ss.getSheetByName('Kick_Off');
  if (!sheet) throw new Error('גיליון Kick_Off חסר; לא נשמרו נתוני ההתנעה.');
  const values = sheet.getDataRange().getDisplayValues();
  const kickoff = payload.kickoff;
  const fields = [
    {key:'arcModel',label:'Arc Model'},
    {key:'grids',label:'Grids',value:function() { return kickoff.grids === 'No' ? 'No|' + String(kickoff.gridsX || '') + '|' + String(kickoff.gridsY || '') : String(kickoff.grids || ''); }},
    {key:'arLevel',label:'AR Level'},
    {key:'urs',label:'URS'},
    {key:'licensing',label:'Licensing'}
  ];
  fields.forEach(function(field) {
    const value = kickoff[field.key] == null ? '' : String(kickoff[field.key]).trim();
    if (['arcModel','grids','arLevel','urs','licensing'].indexOf(field.key) >= 0 && value && value !== 'Yes' && value !== 'No') throw new Error('ערך לא תקין בשדה ' + field.label + '.');
    field.nextValue = field.value ? field.value() : value;
    const matches = [];
    for (let i=0;i<values.length;i++) if (String(values[i][0] || '').toLowerCase().includes(field.label.toLowerCase())) matches.push(i+1);
    if (matches.length !== 1) throw new Error('לא ניתן לזהות שורת ' + field.label + ' יחידה בגיליון Kick_Off.');
    field.row = matches[0];
    field.previousValue = String(values[field.row-1][1] == null ? '' : values[field.row-1][1]).trim();
  });
  const changed = fields.some(function(field) { return field.previousValue !== field.nextValue; });
  let backupUrl = '';
  if (changed) {
    if (typeof backupProjectBeforeSave_ !== 'function') throw new Error('שירות הגיבוי אינו זמין; לא נשמרו נתוני ההתנעה.');
    backupUrl = backupProjectBeforeSave_(resolved);
    try {
      fields.forEach(function(field) { sheet.getRange(field.row,2).setValue(field.nextValue); });
      SpreadsheetApp.flush();
    } catch(error) { error.backupUrl = backupUrl; error.partialSave = true; throw error; }
  }
  return {changed:changed,backupUrl:backupUrl};
}
function architectureKickoffNonEditableFieldsMatch_(submitted, source) {
  const editable = {arcModel:true,grids:true,gridsX:true,gridsY:true,arLevel:true,urs:true,licensing:true};
  const remaining = {};
  Object.keys(source || {}).forEach(function(key) { if (!editable[key]) remaining[key] = source[key]; });
  return Object.keys(remaining).every(function(key) {
    return architectureKickoffComparable_(submitted ? submitted[key] : undefined) === architectureKickoffComparable_(remaining[key]);
  });
}function architectureKickoffSourceSnapshot_(folderId, projectNo, projectName) {
  const result = pullProjectDataFromSheet(folderId,projectNo,projectName);
  if (!result || !result.success) throw new Error((result && result.error) || 'לא ניתן לקרוא את נתוני הפרויקט הקנוניים.');
  const source = result.data || {};
  // Capture every workbook-backed input the kickoff document generator may consume.
  return {bimInfo:source.bimInfo || {},kickoff:source.kickoff || {},lots:source.lots || [],buildings:source.buildings || [],
    basements:source.basements || [],genModels:source.genModels || [],disciplines:source.disciplines || [],matrix:source.matrix || {},
    contacts:source.contacts || [],models:source.models || [],siteGroups:source.siteGroups || [],designOptions:source.designOptions || [],
    sheetUrl:source.sheetUrl || ''};
}
function getArchitectureKickoff(folderId, projectNo, projectName) {
  try {
    const access = architectureKickoffResolve_(folderId,projectNo,projectName,false);
    const state = access.resolved.ss ? architectureKickoffRead_(access.resolved.ss) : {revision:0,draft:null,draftRevision:0,draftUpdatedAt:"",draftUpdatedBy:"",versions:[]};
    return architectureKickoffResponse_(state,access.resolved.project);
  } catch(error) { return {success:false,error:String(error.message || error)}; }
}
function saveArchitectureKickoffDraft(folderId, projectNo, projectName, input) {
  let lock = null, backupUrl = "";
  try {
    if (!input || !Number.isInteger(input.revision) || input.revision < 0) throw new Error('חסר מספר גרסה. טען מחדש לפני שמירה.');
    const payload = architectureKickoffValidatePayload_(input.payload);
    const access = architectureKickoffResolve_(folderId,projectNo,projectName,true);
    lock = LockService.getScriptLock(); lock.waitLock(30000);
    const state = architectureKickoffRead_(access.resolved.ss);
    if (Number(input.revision) !== state.revision) throw new Error('הנתונים עודכנו בחלון אחר. טען מחדש לפני שמירה; השינויים לא נדרסו.');
    // Only the meeting fields and the questionnaire answers come from the form. Shared project facts
    // are taken from the project sheet itself, so Sheets' re-formatting (10:30 -> 10:30:00, 0.00 -> 0)
    // can never make the draft "differ" from its own source.
    const saveResult = architectureKickoffSaveLegacyFields_(access.resolved,payload);
    backupUrl = saveResult.backupUrl;
    const source = architectureKickoffSourceSnapshot_(folderId,projectNo,projectName);
    const canonical = {bimInfo:source.bimInfo,kickoff:source.kickoff,meeting:payload.meeting || {}};
    architectureKickoffAppend_(access.resolved.ss,state,'DRAFT_SAVED','',access.actor,{payload:canonical,backupUrl:saveResult.backupUrl});
    const next = architectureKickoffRead_(access.resolved.ss);
    return Object.assign(architectureKickoffResponse_(next,access.resolved.project),{backupUrl:saveResult.backupUrl});
  } catch(error) { return {success:false,error:String(error.message || error),backupUrl:String(error.backupUrl || backupUrl),partialSave:!!(error.partialSave || backupUrl)}; }
  finally { if (lock) lock.releaseLock(); }
}
function generateArchitectureKickoff(folderId, projectNo, projectName, input) {
  let lock = null;
  let versionId = '';
  let resolved = null;
  let actor = '';
  let generatedArtifactUrl = '';
  try {
    if (!input || !Number.isInteger(input.revision) || input.revision < 0) throw new Error('חסר מספר גרסה. טען מחדש לפני הפקה.');
    const payload = architectureKickoffValidatePayload_(input.payload);
    const access = architectureKickoffResolve_(folderId,projectNo,projectName,true);
    resolved = access.resolved; actor = access.actor;
    lock = LockService.getScriptLock(); lock.waitLock(30000);
    let state = architectureKickoffRead_(resolved.ss);
    if (Number(input.revision) !== state.revision) throw new Error('הנתונים עודכנו בחלון אחר. טען מחדש לפני הפקה; השינויים לא נדרסו.');
    // The saved draft is authoritative. The request only has to carry the same form-owned input
    // (meeting + questionnaire answers) - i.e. nothing unsaved.
    const draft = state.draft;
    const owned = function(p) {
      const k = (p && p.kickoff) || {};
      return {meeting:(p && p.meeting) || {},kickoff:{arcModel:k.arcModel,grids:k.grids,gridsX:k.grids === 'No' ? k.gridsX : '',gridsY:k.grids === 'No' ? k.gridsY : '',arLevel:k.arLevel,urs:k.urs,licensing:k.licensing}};
    };
    if (!draft || !architectureKickoffSourceMatches_(owned(payload).kickoff,owned(draft).kickoff) ||
        architectureKickoffStableJson_(owned(payload).meeting) !== architectureKickoffStableJson_(owned(draft).meeting)) {
      throw new Error('יש לשמור את קלט ההתנעה הנוכחי לפני הפקת מסמך.');
    }
    if (state.versions.some(function(v) { return v.status === 'GENERATING'; })) throw new Error('קיימת הפקה שלא הוכרעה. יש לבדוק את יומן המקור לפני ניסיון נוסף.');
    const sourceSnapshot = architectureKickoffSourceSnapshot_(folderId,projectNo,projectName);
    if (!architectureKickoffSourceMatches_(draft.bimInfo,sourceSnapshot.bimInfo) || !architectureKickoffSourceMatches_(draft.kickoff,sourceSnapshot.kickoff)) throw new Error("נתוני הפרויקט השתנו מאז שמירת הקלט. שמור שוב והפק.");
    const generatorPayload = Object.assign({},sourceSnapshot,{meeting:draft.meeting || {}});
    const version = {inputSnapshot:{meeting:draft.meeting || {}},sourceSnapshot:sourceSnapshot};
    const sizeCheck = JSON.stringify(version);
    if (sizeCheck.length > ARCH_KICKOFF_MAX_EVENT_CHARS_ - 256) throw new Error('צילום הקלט והמקור גדול מדי לשמירה; לא נוצר מסמך.');
    versionId = Utilities.getUuid();
    architectureKickoffAppend_(resolved.ss,state,'GENERATION_STARTED',versionId,actor,{version:version});
    state = architectureKickoffRead_(resolved.ss);
    const generated = generateKickoffDocWithResolvedDb_(folderId,projectNo,projectName,generatorPayload,resolved);
    generatedArtifactUrl = String((generated && (generated.url || generated.orphanUrl)) || '');
    if (!generated || !generated.success || !generated.url) {
      architectureKickoffAppend_(resolved.ss,state,'GENERATION_FAILED',versionId,actor,{error:String((generated && generated.error) || 'Document generator returned no URL.'),orphanUrl:String((generated && generated.orphanUrl) || '')});
      const failed = architectureKickoffRead_(resolved.ss);
      return Object.assign(architectureKickoffResponse_(failed,resolved.project),{generationError:String((generated && generated.error) || 'Document generator returned no URL.')});
    }
    architectureKickoffAppend_(resolved.ss,state,'GENERATION_FINISHED',versionId,actor,{documentUrl:generated.url});
    const next = architectureKickoffRead_(resolved.ss);
    return architectureKickoffResponse_(next,resolved.project);
  } catch(error) {
    return {success:false,error:String(error.message || error),orphanUrl:String(error.orphanUrl || generatedArtifactUrl)};
  } finally { if (lock) lock.releaseLock(); }
}
function reviewArchitectureKickoff(folderId, projectNo, projectName, input) {
  let lock = null;
  try {
    if (!input || !Number.isInteger(input.revision) || input.revision < 0) throw new Error('חסר מספר גרסה. טען מחדש לפני סקירה.');
    if (['APPROVED','CHANGES_REQUESTED'].indexOf(input.decision) < 0) throw new Error('החלטת סקירה לא תקינה.');
    const note = String(input.note || '').trim();
    if (note.length > 5000) throw new Error('הערת הסקירה ארוכה מדי.');
    if (input.decision === 'CHANGES_REQUESTED' && !note) throw new Error('יש לפרט מה נדרש לתיקון.');
    const access = architectureKickoffResolve_(folderId,projectNo,projectName,true);
    lock = LockService.getScriptLock(); lock.waitLock(30000);
    const state = architectureKickoffRead_(access.resolved.ss);
    if (Number(input.revision) !== state.revision) throw new Error('הנתונים עודכנו בחלון אחר. טען מחדש לפני סקירה; השינויים לא נדרסו.');
    const version = state.versions.filter(function(v) { return v.id === String(input.versionId || ''); })[0];
    if (!version || ['GENERATED','APPROVED','CHANGES_REQUESTED'].indexOf(version.status) < 0 || !version.documentUrl) throw new Error('לא נמצאה גרסה שניתן לסקור.');
    architectureKickoffAppend_(access.resolved.ss,state,'REVIEWED',version.id,access.actor,{decision:input.decision,note:note});
    return architectureKickoffResponse_(architectureKickoffRead_(access.resolved.ss),access.resolved.project);
  } catch(error) { return {success:false,error:String(error.message || error)}; }
  finally { if (lock) lock.releaseLock(); }
}
