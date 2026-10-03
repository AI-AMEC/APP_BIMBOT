// ==========================================
// BIMBOT v41.0 - 98_Diagnostics
// Run from the Apps Script editor (Run -> diagnoseSetup / diagnoseProject) and read the Execution log.
// ==========================================

// Checks that the running account can open every central file in 00_Config.
function diagnoseSetup() {
  Logger.log('Running as: ' + Session.getEffectiveUser().getEmail() + ' | active user: ' + Session.getActiveUser().getEmail());
  Logger.log('APP_READ_ONLY = ' + APP_READ_ONLY);
  const ids = {
    MAIN_DB_ID: MAIN_DB_ID, INDEX_DB_ID: INDEX_DB_ID, PROCESS_DB_ID: PROCESS_DB_ID, TEMPLATE_DB_ID: TEMPLATE_DB_ID,
    DROPDOWN_DB_ID: DROPDOWN_DB_ID, USERS_DB_ID: USERS_DB_ID, MISSIONS_DB_ID: MISSIONS_DB_ID,
    KICKOFF_DOC_TEMPLATE_ID: KICKOFF_DOC_TEMPLATE_ID, PRESENTATION_TEMPLATE_ID: PRESENTATION_TEMPLATE_ID
  };
  Object.keys(ids).forEach(k => {
    try { Logger.log('OK      ' + k + ' -> ' + DriveApp.getFileById(ids[k]).getName()); }
    catch (e) { Logger.log('FAILED  ' + k + ' (' + ids[k] + '): ' + e.message); }
  });
  const user = getAuthenticatedUserContext_();
  Logger.log('User permission record: ' + JSON.stringify(user));
}

// Resolves one project exactly like the app does, without writing anything.
function diagnoseProject(projectNo) {
  projectNo = String(projectNo || '220109');
  const list = getCachedProjectData();
  if (!list.success) { Logger.log('Project list failed: ' + list.error); return; }
  const project = list.data.filter(p => p.no === projectNo)[0];
  if (!project) { Logger.log('Project ' + projectNo + ' is not in the authorized list for this user.'); return; }
  Logger.log('Project: ' + JSON.stringify(project));
  try {
    const db = getLocalProjectDbAndBimFolder_(project.folderId, project.no, project.name, false);
    Logger.log('BIM folder: ' + (db.bimFolder ? db.bimFolder.getName() + ' ' + db.bimFolder.getUrl() : 'none yet (created on first sync)'));
    Logger.log('BIM INFO sheet: ' + (db.ss ? db.ss.getName() + ' ' + db.ss.getUrl() : 'none yet (created on first sync)'));
    Logger.log('Writable: ' + canWriteProject_(db.project));
  } catch (e) {
    Logger.log('RESOLVE FAILED: ' + e.message);
  }
}

// Lists registered project spreadsheets (Script Properties) - useful after moving files.
function listRegisteredProjectDbs() {
  const props = PropertiesService.getScriptProperties().getProperties();
  Object.keys(props).filter(k => k.indexOf('BIM_PROJECT_DB:') === 0).sort().forEach(k => Logger.log(k + ' = ' + props[k]));
}
