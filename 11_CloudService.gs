// ==========================================
// BIMBOT v41.0 - 11_CloudService
// ACC views use the current user session. Captured pilot data is never a live fallback.
// ==========================================

function cloudProjectAccess_(projectNo) {
  const no = String(projectNo || '').trim();
  const access = getCachedProjectData();
  if (!access.success) throw new Error('לא ניתן לאמת הרשאת פרויקט.');
  const allowed = access.data.filter(p => String(p.no) === no && !p.ambiguousIdentity);
  if (allowed.length !== 1) throw new Error('הפרויקט אינו מורשה או שזהותו אינה חד־משמעית.');
  return no;
}

// Injects JSON into an HTML template. A replacer function avoids "$&"-style patterns in data.
function injectJson_(html, token, value) {
  const json = JSON.stringify(value).replace(/</g, '\\u003c');
  return html.replace(token, () => json);
}

// Client: google.script.run.getPilotCloudSnapshot (name kept for UI compatibility)
function getPilotCloudSnapshot(projectNo) {
  return getLiveCloudSnapshot_(cloudProjectAccess_(projectNo));
}

function getCloudPanel(projectNo, view) {
  const result = getPilotCloudSnapshot(projectNo);
  if (!result.available) return result;
  if (['overview', 'issues', 'publish'].indexOf(view) < 0) throw new Error('Unknown cloud view');
  if (view === 'issues') {
    const management = result.project.management || {};
    const issueData = {
      summary: management.summary || {}, activity: management.activity || {}, rows: management.rows || [],
      complete: management.complete, sourceReadAt: management.sourceReadAt,
      issuesComplete: management.dataQuality && management.dataQuality.issuesComplete
    };
    const issuePanel = HtmlService.createHtmlOutputFromFile('Panel_CloudIssues').getContent();
    return { available: true, html: injectJson_(issuePanel, '__BIM_ISSUES_DATA__', issueData).replace('__BIM_INITIAL_PAGE__', view) };
  }
  const panel = HtmlService.createHtmlOutputFromFile('Panel_CloudOverview').getContent();
  return { available: true, html: injectJson_(panel, '"__BIM_CLOUD_DATA__"', result.project).replace('__BIM_INITIAL_PAGE__', view) };
}

function getPublishSnapshot(projectNo) {
  return getLivePublishSnapshot_(cloudProjectAccess_(projectNo));
}
function getPublishModelThumbnails(projectNo, modelId, versionId) {
  return getLivePublishModelThumbnails_(cloudProjectAccess_(projectNo), modelId, versionId);
}
function getCloudIssueComments(projectNo, issueId) {
  return getLiveCloudIssueComments_(cloudProjectAccess_(projectNo), issueId);
}
