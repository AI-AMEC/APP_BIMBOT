// ==========================================
// BIMBOT v41.0 - 09_PublishReview
// ==========================================
// Human publish reviews are append-only and remain separate from Autodesk's
// publication facts. Review labels and evidence do not certify compliance.
const PUBLISH_REVIEW_LOG_SHEET_ = 'Publish_Review_Log';
const PUBLISH_REVIEW_HEADERS_ = ['Model ID','Revision','At','Actor','Cloud Version ID','Review JSON'];
const PUBLISH_REVIEW_CHECKS_ = ['discipline3d','allLinks3d','arSt3d','typical2d','sheets'];
const PUBLISH_REVIEW_STATUSES_ = ['UNREVIEWED','PASS','FAIL','NA'];

function publishReviewActor_() {
  const context = getAuthenticatedUserContext_();
  if (!context || !context.success || !context.email) throw new Error((context && context.error) || 'Cannot verify reviewer.');
  return String(context.email).toLowerCase().trim();
}
function publishReviewResolve_(folderId, projectNo, projectName, forWrite) {
  const actor = publishReviewActor_();
  const resolved = getLocalProjectDbAndBimFolder_(folderId, projectNo, projectName, !!forWrite);
  return {resolved:resolved,actor:actor};
}
function publishReviewSheet_(ss, create) {
  let sheet = ss ? ss.getSheetByName(PUBLISH_REVIEW_LOG_SHEET_) : null;
  if (!sheet && create) {
    sheet = ss.insertSheet(PUBLISH_REVIEW_LOG_SHEET_);
    sheet.getRange(1,1,1,PUBLISH_REVIEW_HEADERS_.length).setValues([PUBLISH_REVIEW_HEADERS_]);
    sheet.setFrozenRows(1);
    if (typeof safeProtect_ === 'function') safeProtect_(ss.getId());
  }
  if (sheet && sheet.getLastRow() > 0) {
    const headers = sheet.getRange(1,1,1,PUBLISH_REVIEW_HEADERS_.length).getDisplayValues()[0];
    if (headers.join('|') !== PUBLISH_REVIEW_HEADERS_.join('|')) throw new Error('Publish review log schema mismatch; recovery is required.');
  }
  return sheet;
}
function publishReviewValidate_(payload) {
  if (!payload || typeof payload !== 'object' || Array.isArray(payload) || !Number.isInteger(payload.expectedRevision) || payload.expectedRevision < 0) {
    throw new Error('Missing expected review revision. Reload before saving.');
  }
  if (!payload.checks || typeof payload.checks !== 'object' || Array.isArray(payload.checks) ||
      Object.keys(payload.checks).sort().join('|') !== PUBLISH_REVIEW_CHECKS_.slice().sort().join('|')) {
    throw new Error('All five publish review checks are required.');
  }
  const checks = {};
  PUBLISH_REVIEW_CHECKS_.forEach(function(id) {
    const item = payload.checks[id];
    if (!item || typeof item !== 'object' || Array.isArray(item) || PUBLISH_REVIEW_STATUSES_.indexOf(item.status) < 0) throw new Error('Invalid status for publish review check ' + id + '.');
    if (typeof item.note !== 'string' || item.note.length > 2000) throw new Error('Review note must contain at most 2000 characters.');
    if (!Array.isArray(item.viewIds) || item.viewIds.length > 100 || item.viewIds.some(function(viewId) { return typeof viewId !== 'string' || !viewId.trim() || viewId.length > 500; })) throw new Error('Review view references are invalid.');
    const viewIds = item.viewIds.map(function(viewId) { return viewId.trim(); });
    if (new Set(viewIds).size !== viewIds.length) throw new Error('Duplicate view references are not allowed.');
    if (item.status === 'PASS' && (!item.note.trim() || viewIds.length < 1)) throw new Error('PASS requires a note and at least one view reference.');
    if ((item.status === 'FAIL' || item.status === 'NA') && !item.note.trim()) throw new Error(item.status + ' requires a note.');
    checks[id] = {status:item.status,note:item.note.trim(),viewIds:viewIds};
  });
  return {expectedRevision:payload.expectedRevision,checks:checks};
}
function publishReviewRead_(ss, snapshot) {
  const latest = Object.create(null);
  const sheet = publishReviewSheet_(ss,false);
  if (!sheet || sheet.getLastRow() <= 1) return latest;
  const rows = sheet.getRange(2,1,sheet.getLastRow()-1,PUBLISH_REVIEW_HEADERS_.length).getValues();
  rows.forEach(function(row) {
    const modelId = String(row[0] || ''), revision = Number(row[1]), at = String(row[2] || ''), actor = String(row[3] || ''), cloudVersionId = String(row[4] || '');
    let saved;
    try { saved = JSON.parse(String(row[5] || '')); } catch(e) { throw new Error('Publish review log is malformed; recovery is required.'); }
    const expected = latest[modelId] ? latest[modelId].revision + 1 : 1;
    if (!modelId || !Number.isInteger(revision) || revision !== expected || !at || !actor || !cloudVersionId || !saved ||
        saved.modelId !== modelId || saved.revision !== revision || saved.updatedAt !== at || saved.updatedBy !== actor || saved.cloudVersionId !== cloudVersionId) {
      throw new Error('Publish review revision history is inconsistent; recovery is required.');
    }
    publishReviewValidate_({expectedRevision:0,checks:saved.checks});
    latest[modelId] = saved;
  });
  return latest;
}
function publishReviewSnapshot_(projectNo) {
  if (typeof getPublishSnapshot !== 'function') throw new Error('Publish snapshot service is unavailable.');
  const result = getPublishSnapshot(projectNo);
  if (!result || result.available !== true || !result.snapshot || !Array.isArray(result.snapshot.models)) throw new Error('A current Publish model snapshot is unavailable.');
  return result.snapshot;
}
function publishReviewModel_(snapshot, modelId, versionId) {
  const models = snapshot && Array.isArray(snapshot.models) ? snapshot.models : [];
  const matches = models.filter(function(model) { return model && String(model.id) === String(modelId); });
  if (matches.length !== 1 || String(matches[0].versionId) !== String(versionId)) throw new Error('Model or cloud version does not match the current Publish snapshot.');
  return matches[0];
}
function publishReviewValidateViews_(checks, model) {
  const knownViews = Object.create(null);
  (Array.isArray(model.views) ? model.views : []).forEach(function(view) { if (view && view.id != null) knownViews[String(view.id)] = String(view.kind || '').toLowerCase(); });
  const requiredKinds = {discipline3d:'3d',allLinks3d:'3d',arSt3d:'3d',typical2d:'2d',sheets:'sheet'};
  PUBLISH_REVIEW_CHECKS_.forEach(function(id) {
    checks[id].viewIds.forEach(function(viewId) {
      if (!Object.prototype.hasOwnProperty.call(knownViews,viewId)) throw new Error('A referenced view is not part of the current cloud model version.');
      if (checks[id].status === 'PASS' && knownViews[viewId] !== requiredKinds[id]) throw new Error('PASS evidence must reference a ' + requiredKinds[id] + ' view for ' + id + '.');
    });
    if (checks[id].status === 'PASS' && model.manifestComplete !== true) throw new Error('PASS requires a complete cloud model manifest.');
  });
}
function getPublishReviews(folderId, projectNo, projectName) {
  try {
    const auth = publishReviewActor_();
    const resolved = getLocalProjectDbAndBimFolder_(folderId,projectNo,projectName,false);
    const latest = publishReviewRead_(resolved.ss,publishReviewSnapshot_(projectNo));
    return {success:true,reviews:latest,writable:canWriteProject_(resolved.project)};
  } catch(error) { return {success:false,error:String(error.message || error)}; }
}
function savePublishReview(folderId, projectNo, projectName, modelId, versionId, payload) {
  let lock = null;
  try {
    const input = publishReviewValidate_(payload);
    if (typeof modelId !== 'string' || !modelId.trim() || modelId.length > 500 || typeof versionId !== 'string' || !versionId.trim() || versionId.length > 500) {
      throw new Error('Model and cloud version identifiers are required.');
    }
    const context = publishReviewResolve_(folderId,projectNo,projectName,true);
    const snapshot = publishReviewSnapshot_(projectNo);
    publishReviewModel_(snapshot,modelId,versionId);
    const live = getLivePublishModel_(projectNo,modelId,versionId);
    if (!live || !live.available || !live.model) throw new Error((live && live.error) || 'לא ניתן לאמת את מבטי גרסת הענן.');
    const model = live.model;
    publishReviewValidateViews_(input.checks,model);
    lock = LockService.getScriptLock();
    lock.waitLock(30000);
    if (!modelGovernanceIncluded_(modelGovernanceRead_(context.resolved.ss)[String(modelId)],'PUBLISH')) {
      throw new Error('המודל ממתין לסיווג או מוחרג מ־Publish. יש לעדכן את ההגדרה ב״מודלים ופירוק״ לפני שמירת בקרה.');
    }
    const latest = publishReviewRead_(context.resolved.ss,snapshot);
    const previous = latest[String(modelId)];
    const revision = previous ? previous.revision : 0;
    if (input.expectedRevision !== revision) throw new Error('Review changed in another window. Reload before saving.');
    const at = new Date().toISOString();
    const next = {revision:revision+1,modelId:String(modelId),cloudVersionId:String(versionId),updatedAt:at,updatedBy:context.actor,checks:input.checks};
    const sheet = publishReviewSheet_(context.resolved.ss,true);
    sheet.getRange(sheet.getLastRow()+1,1,1,PUBLISH_REVIEW_HEADERS_.length).setValues([[next.modelId,next.revision,at,context.actor,next.cloudVersionId,JSON.stringify(next)]]);
    SpreadsheetApp.flush();
    latest[next.modelId] = next;
    return {success:true,review:next,writable:true};
  } catch(error) { return {success:false,error:String(error.message || error)}; }
  finally { if (lock) lock.releaseLock(); }
}
