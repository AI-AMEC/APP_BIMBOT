// ==========================================
// BIMBOT v41.0 - 07_Workflow
// ==========================================
// Model-manager decisions are kept separately from imported project facts.
// Append-only revisions preserve the previous state and reject stale editors.
function workflowDefinitions_() {
  return [
    ['architect_meeting','INITIAL','פגישה מקדימה עם האדריכל'],
    ['decomposition','INITIAL','הגדרת מבנה הפרויקט ורשימת המודלים'],
    ['cde_urs','INITIAL','הכנת סביבת הענן ומודל URS'],
    ['architect_check','INITIAL','בדיקת מודל האדריכלות'],
    ['bep','INITIAL','הכנת BEP ועדכונו'],
    ['consultant_kickoff','INITIAL','התנעת יועצים והעלאת מודלים'],
    ['first_review','INITIAL','בקרה ראשונה ומעקב'],
    ['tender_review','FINAL','בקרת מכרז וסגירת הערות'],
    ['detailed_review','DETAILED','בקרת תכנון מפורט'],
    ['stage_summary','DETAILED','סיכום שלב והנחיות BIM לביצוע']
  ];
}
function validateWorkflowPayload_(payload) {
  if (!payload || !Number.isInteger(payload.revision) || payload.revision < 0) throw new Error('Missing workflow revision. Reload before saving.');
  if (['UNSET','INITIAL','FINAL','DETAILED'].indexOf(payload.stage) < 0) throw new Error('Invalid project stage.');
  const definitions = workflowDefinitions_();
  if (!Array.isArray(payload.tasks) || payload.tasks.length !== definitions.length) throw new Error('The complete workflow is required.');
  const ids = Object.create(null);
  const tasks = payload.tasks.map(function(task) {
    if (!task || !definitions.some(function(d) { return d[0] === task.id; }) || ids[task.id]) throw new Error('Unknown or duplicate workflow task.');
    ids[task.id] = true;
    if (['NOT_STARTED','IN_PROGRESS','BLOCKED','DONE'].indexOf(task.status) < 0) throw new Error('Invalid task status.');
    if (typeof task.note !== 'string' || task.note.length > 2000) throw new Error('Task note must contain at most 2000 characters.');
    if (task.status === 'DONE' && !task.note.trim()) throw new Error('יש לציין אסמכתה או הסבר לפני סימון משימה כהושלמה.');
    return {id:task.id,status:task.status,note:task.note.trim()};
  });
  return {revision:payload.revision,stage:payload.stage,tasks:tasks};
}
function readWorkflowState_(ss) {
  const sheet = ss ? ss.getSheetByName('BIM_Workflow_Log') : null;
  if (!sheet || sheet.getLastRow() <= 1) return {revision:0,stage:'UNSET',tasks:[]};
  const row = sheet.getRange(sheet.getLastRow(),1,1,4).getValues()[0];
  const state = JSON.parse(String(row[3]));
  validateWorkflowPayload_(state);
  if (state.revision !== Number(row[0])) throw new Error('Workflow revision mismatch. Recovery is required.');
  return state;
}
function workflowResponse_(saved, project) {
  return {
    success:true, revision:saved.revision, stage:saved.stage,
    updatedAt:saved.updatedAt || '', updatedBy:saved.updatedBy || '',
    writable:canWriteProject_(project),
    tasks:workflowDefinitions_().map(function(d) {
      const stored = saved.tasks.filter(function(t) { return t.id === d[0]; })[0] || {};
      return {id:d[0],stage:d[1],label:d[2],status:stored.status || 'NOT_STARTED',note:stored.note || '',updatedAt:stored.updatedAt || '',updatedBy:stored.updatedBy || ''};
    })
  };
}
function getProjectWorkflow(folderId, projectNo, projectName) {
  try {
    const resolved = getLocalProjectDbAndBimFolder_(folderId,projectNo,projectName,false);
    return workflowResponse_(readWorkflowState_(resolved.ss),resolved.project);
  } catch(error) { return {success:false,error:String(error.message || error)}; }
}
function saveProjectWorkflow(folderId, projectNo, projectName, payload) {
  let lock = null;
  try {
    const input = validateWorkflowPayload_(payload);
    // Resolving with create=true enforces write permission and project identity checks.
    const resolved = getLocalProjectDbAndBimFolder_(folderId,projectNo,projectName,true);
    lock = LockService.getScriptLock();
    lock.waitLock(30000);
    const previous = readWorkflowState_(resolved.ss);
    if (input.revision !== previous.revision) throw new Error('הנתונים עודכנו בחלון אחר. טען מחדש לפני שמירה; השינויים שלך לא נדרסו.');
    const at = new Date().toISOString();
    const actor = Session.getActiveUser().getEmail();
    if (!actor) throw new Error('Cannot verify workflow editor.');
    const next = {revision:previous.revision+1,stage:input.stage,updatedAt:at,updatedBy:actor,tasks:input.tasks.map(function(task) {
      const old = previous.tasks.filter(function(t) { return t.id === task.id; })[0];
      const changed = !old || old.status !== task.status || old.note !== task.note;
      return {id:task.id,status:task.status,note:task.note,updatedAt:changed ? at : old.updatedAt,updatedBy:changed ? actor : old.updatedBy};
    })};
    let sheet = resolved.ss.getSheetByName('BIM_Workflow_Log');
    if (!sheet) {
      sheet = resolved.ss.insertSheet('BIM_Workflow_Log');
      sheet.getRange(1,1,1,4).setValues([['Revision','Updated At','Updated By','State JSON']]);
      sheet.setFrozenRows(1);
    }
    sheet.getRange(sheet.getLastRow()+1,1,1,4).setValues([[next.revision,at,actor,JSON.stringify(next)]]);
    SpreadsheetApp.flush();
    return workflowResponse_(next,resolved.project);
  } catch(error) { return {success:false,error:String(error.message || error)}; }
  finally { if (lock) lock.releaseLock(); }
}
