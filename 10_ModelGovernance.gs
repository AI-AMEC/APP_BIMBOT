// ==========================================
// BIMBOT v41.0 - 10_ModelGovernance
// ==========================================
// Central human model classification and control decisions. Kept append-only
// and separate from Autodesk publication facts and the project Model_Table.
const MODEL_GOVERNANCE_LOG_SHEET_ = 'Model_Governance_Log';
const MODEL_GOVERNANCE_HEADERS_ = ['Model ID','Revision','At','Actor','Decision JSON'];
const MODEL_GOVERNANCE_PURPOSES_ = ['PLANNING','COORDINATION','REFERENCE','AUXILIARY','ARCHIVE'];
const MODEL_GOVERNANCE_MODES_ = ['ALL','SELECTED','NONE'];
const MODEL_GOVERNANCE_CONTROLS_ = ['PUBLISH','COORDINATION','MODEL_QUALITY'];

function modelGovernanceSheet_(ss,create) {
  let sheet=ss?ss.getSheetByName(MODEL_GOVERNANCE_LOG_SHEET_):null;
  if(!sheet&&create){sheet=ss.insertSheet(MODEL_GOVERNANCE_LOG_SHEET_);sheet.getRange(1,1,1,MODEL_GOVERNANCE_HEADERS_.length).setValues([MODEL_GOVERNANCE_HEADERS_]);sheet.setFrozenRows(1);if(typeof safeProtect_==='function')safeProtect_(ss.getId());}
  if(sheet&&sheet.getLastRow()>0){const h=sheet.getRange(1,1,1,MODEL_GOVERNANCE_HEADERS_.length).getDisplayValues()[0];if(h.join('|')!==MODEL_GOVERNANCE_HEADERS_.join('|'))throw Error('מבנה יומן ממשל המודלים השתנה; נדרשת התאוששות.');}
  return sheet;
}
function modelGovernanceRead_(ss) {
  const latest=Object.create(null),sheet=modelGovernanceSheet_(ss,false);if(!sheet||sheet.getLastRow()<=1)return latest;
  sheet.getRange(2,1,sheet.getLastRow()-1,MODEL_GOVERNANCE_HEADERS_.length).getValues().forEach(function(row){
    const id=String(row[0]||''),revision=Number(row[1]),at=String(row[2]||''),actor=String(row[3]||'');let d;
    try{d=JSON.parse(String(row[4]||''));}catch(e){throw Error('יומן ממשל המודלים פגום; נדרשת התאוששות.');}
    const expected=latest[id]?latest[id].revision+1:1;
    if(!id||!Number.isInteger(revision)||revision!==expected||!at||!actor||!d||d.modelId!==id||d.revision!==revision||d.updatedAt!==at||d.updatedBy!==actor)throw Error('היסטוריית גרסאות ממשל המודלים אינה עקבית; נדרשת התאוששות.');
    modelGovernanceValidateDecision_(d,false);latest[id]=d;
  });return latest;
}
function modelGovernanceValidateDecision_(p,withExpected) {
  if(!p||typeof p!=='object'||Array.isArray(p))throw Error('נדרשת החלטת ממשל למודל.');
  if(withExpected&&(!Number.isInteger(p.expectedRevision)||p.expectedRevision<0))throw Error('חסרה גרסת החלטה צפויה. יש לטעון מחדש לפני השמירה.');
  if(typeof p.classified!=='boolean')throw Error('יש לבחור אם המודל מסווג.');
  const mode=p.controlMode;
  if(MODEL_GOVERNANCE_MODES_.indexOf(mode)<0)throw Error('אופן הבקרות שנבחר אינו תקין.');
  if(!Array.isArray(p.controls)||p.controls.some(function(c){return MODEL_GOVERNANCE_CONTROLS_.indexOf(c)<0;})||new Set(p.controls).size!==p.controls.length)throw Error('רשימת הבקרות אינה תקינה.');
  let controls=mode==='ALL'?MODEL_GOVERNANCE_CONTROLS_.slice():(mode==='NONE'?[]:p.controls.slice());
  const reason=typeof p.exclusionReason==='string'?p.exclusionReason.trim():'';
  if(p.classified&&(mode==='NONE'||(mode==='SELECTED'&&controls.length<MODEL_GOVERNANCE_CONTROLS_.length))&&!reason)throw Error('יש לנמק החרגת מודל מבקרה אחת או יותר.');
  if(reason.length>2000)throw Error('סיבת ההחרגה ארוכה מדי (עד 2,000 תווים).');
  const notes=typeof p.notes==='string'?p.notes:'';if(notes.length>4000)throw Error('ההערות ארוכות מדי (עד 4,000 תווים).');
  const d={classified:p.classified,registerCode:'',disciplineCode:'',buildingCode:'',lotCode:'',purpose:'',notes:notes.trim(),controlMode:mode,controls:controls,exclusionReason:reason};
  if(p.classified){
    if(typeof p.disciplineCode!=='string'||!p.disciplineCode.trim())throw Error('קוד תחום נדרש למודל מסווג.');
    if(typeof p.purpose!=='string'||MODEL_GOVERNANCE_PURPOSES_.indexOf(p.purpose)<0)throw Error('מטרת המודל אינה תקינה.');
    [['registerCode','קוד רשומה'],['disciplineCode','קוד תחום'],['buildingCode','קוד בניין'],['lotCode','קוד מגרש']].forEach(function(x){if(p[x[0]]!=null&&(typeof p[x[0]]!=='string'||p[x[0]].length>100))throw Error(x[1]+' אינו תקין.');});
    d.registerCode=(p.registerCode||'').trim();d.disciplineCode=p.disciplineCode.trim();d.buildingCode=(p.buildingCode||'').trim();d.lotCode=(p.lotCode||'').trim();d.purpose=p.purpose;
  }
  if(withExpected)d.expectedRevision=p.expectedRevision;
  return d;
}
function modelGovernanceSnapshot_(projectNo){if(typeof getPublishSnapshot!=='function')throw Error('Publish snapshot service is unavailable.');const r=getPublishSnapshot(projectNo);if(!r||r.available!==true||!r.snapshot||!Array.isArray(r.snapshot.models))throw Error('A current Publish model snapshot is unavailable.');return r.snapshot;}
function modelGovernanceAssertModel_(snapshot,id){const matches=snapshot.models.filter(function(m){return m&&String(m.id)===String(id);});if(matches.length!==1)throw Error('Model does not match the current Publish snapshot.');return matches[0];}
function modelGovernanceValidateRegisterCode_(ss,code){if(!code)return;const sh=ss.getSheetByName('Model_Table');if(!sh||sh.getLastRow()<=1)throw Error('קוד הרשומה אינו קיים בטבלת המודלים של הפרויקט.');const values=sh.getDataRange().getDisplayValues(),headers=values[0].map(function(v){return String(v||'').trim().toLowerCase();}),col=headers.indexOf('model name');if(col<0||headers.filter(function(h){return h==='model name';}).length!==1)throw Error('לא ניתן לאמת את קוד הרשומה בטבלת המודלים.');const matches=values.slice(1).filter(function(row){return String(row[col]||'').trim()===code;});if(matches.length!==1)throw Error(matches.length?'קוד הרשומה מופיע יותר מפעם אחת בטבלת המודלים.':'קוד הרשומה אינו קיים בטבלת המודלים של הפרויקט.');}
function modelGovernanceIncluded_(decision,control){if(!decision||decision.classified!==true)return false;if(MODEL_GOVERNANCE_CONTROLS_.indexOf(control)<0)return false;if(decision.controlMode==='ALL')return true;if(decision.controlMode==='SELECTED')return Array.isArray(decision.controls)&&decision.controls.indexOf(control)>=0;return false;}
function getModelGovernance(folderId,projectNo,projectName){try{const c=publishReviewResolve_(folderId,projectNo,projectName,false),all=modelGovernanceRead_(c.resolved.ss),writable=canWriteProject_(c.resolved.project);return {success:true,writable:writable,decisions:all};}catch(e){return {success:false,error:String(e.message||e)};}}
function saveModelGovernance(folderId,projectNo,projectName,modelId,payload){let lock=null;try{
  if(typeof modelId!=='string'||!modelId.trim()||modelId.length>500)throw Error('מזהה המודל חסר או אינו תקין.');
  const input=modelGovernanceValidateDecision_(payload,true),c=publishReviewResolve_(folderId,projectNo,projectName,true),snapshot=modelGovernanceSnapshot_(projectNo);modelGovernanceAssertModel_(snapshot,modelId);
  lock=LockService.getScriptLock();lock.waitLock(30000);const latest=modelGovernanceRead_(c.resolved.ss),previous=latest[modelId],revision=previous?previous.revision:0;
  if(input.expectedRevision!==revision)throw Error('החלטת הממשל השתנתה בחלון אחר. יש לטעון מחדש לפני השמירה.');
  modelGovernanceValidateRegisterCode_(c.resolved.ss,input.registerCode);
  if(input.registerCode&&Object.keys(latest).some(function(id){return id!==modelId&&latest[id].classified===true&&latest[id].registerCode===input.registerCode;}))throw Error('קוד הרשומה כבר משויך למודל פעיל אחר.');
  const at=new Date().toISOString(),d={modelId:modelId,revision:revision+1,updatedAt:at,updatedBy:c.actor,classified:input.classified,registerCode:input.registerCode,disciplineCode:input.disciplineCode,buildingCode:input.buildingCode,lotCode:input.lotCode,purpose:input.purpose,notes:input.notes,controlMode:input.controlMode,controls:input.controls,exclusionReason:input.exclusionReason};
  const sheet=modelGovernanceSheet_(c.resolved.ss,true);sheet.getRange(sheet.getLastRow()+1,1,1,MODEL_GOVERNANCE_HEADERS_.length).setValues([[modelId,d.revision,at,c.actor,JSON.stringify(d)]]);SpreadsheetApp.flush();return {success:true,decision:d,writable:true};
 }catch(e){return {success:false,error:String(e.message||e)};}finally{if(lock)lock.releaseLock();}}

