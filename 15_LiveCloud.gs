// Live, per-user ACC reads. Read-only: this module never creates/publishes ACC data.
// Requires apsUserToken_() from the 3-legged OAuth service.
const LIVE_CLOUD_BASE_ = 'https://developer.api.autodesk.com';
const LIVE_CLOUD_PAGE_CAP_ = 12; // bound execution time; incomplete is explicit
const LIVE_CLOUD_ITEM_CAP_ = 1200;

function liveCloudUser_() {
  const u = getAuthenticatedUserContext_();
  if (!u || !u.success || !u.email) throw new Error('לא ניתן לאמת משתמש BIMBOT.');
  return u;
}
function liveCloudProject_(projectNo) {
  const no = String(projectNo || '').trim();
  const list = getCachedProjectData();
  if (!list.success) throw new Error(list.error || 'לא ניתן לאמת הרשאת פרויקט.');
  const matches = list.data.filter(p => String(p.no) === no && !p.ambiguousIdentity);
  if (!no || matches.length !== 1) throw new Error('הפרויקט אינו מורשה או שזהותו אינה חד־משמעית.');
  return matches[0];
}
function liveCloudBindingKey_(no) { return 'LIVE_ACC_BINDING:' + String(no); }
function liveCloudBinding_() {
  const no = String(arguments[0] || '').trim(), user = liveCloudUser_();
  const raw = PropertiesService.getUserProperties().getProperty(liveCloudBindingKey_(no));
  if (!raw) return null;
  let b; try { b = JSON.parse(raw); } catch (e) { return null; }
  return b && b.user === user.email && b.hubId && b.projectId ? b : null;
}
function liveCloudFetch_(path, token) {
  const url = /^https:\/\/developer\.api\.autodesk\.com\//.test(path) ? path : LIVE_CLOUD_BASE_ + path;
  if (!url.startsWith(LIVE_CLOUD_BASE_ + '/')) throw new Error('Autodesk returned an unsafe pagination link.');
  const r = UrlFetchApp.fetch(url, {method:'get', headers:{Authorization:'Bearer ' + token, Accept:'application/json'}, muteHttpExceptions:true, followRedirects:false});
  let body = null; try { body = JSON.parse(r.getContentText()); } catch (e) {}
  let bytes=[];try{bytes=r.getBlob().getBytes();}catch(e){}
  return {code:r.getResponseCode(), body:body, text:r.getContentText(), bytes:bytes, headers:r.getHeaders?r.getHeaders():{}};
}
function liveCloudRequire200_(r) {
  if (r.code !== 200) throw new Error('קריאת Autodesk נכשלה (' + r.code + '). ' + String(r.text || '').slice(0,220));
  return r.body || {};
}
function liveCloudAssertAccess_(binding, token) {
  // Verify both the selected hub membership and selected project on this user's token.
  const hubs = liveCloudRequire200_(liveCloudFetch_('/project/v1/hubs',token));
  if (!(hubs.data || []).some(h => String(h.id) === String(binding.hubId))) throw new Error('ה-HUB שנבחר אינו נגיש בחשבון Autodesk המחובר.');
  const project = liveCloudRequire200_(liveCloudFetch_('/project/v1/hubs/' + encodeURIComponent(binding.hubId) + '/projects/' + encodeURIComponent(binding.projectId),token));
  if (!project.data || String(project.data.id) !== String(binding.projectId)) throw new Error('הפרויקט שנבחר אינו נגיש בחשבון Autodesk המחובר.');
  return project.data;
}
function liveCloudResolveBinding_(no) {
  const selected = liveCloudBinding_(no);
  if (selected) return selected;
  // Discover only an exact saved ACC link/project ID from the canonical project database.
  const p = liveCloudProject_(no);
  const resolved = getLocalProjectDbAndBimFolder_(p.folderId, no, p.name, false);
  const ss = resolved.ss;
  if (ss) {
    const sheet = typeof getSheetIgnoreCase === 'function' ? getSheetIgnoreCase(ss,'BIM_Info') : ss.getSheetByName('BIM_Info');
    if (sheet) {
      const rows = sheet.getDataRange().getDisplayValues();
      const val = key => { const row=rows.find(r=>String(r[0]||'').trim().toLowerCase()===key.toLowerCase()); return row?String(row[1]||'').trim():''; };
      const hubId=val('Autodesk Hub ID'), projectId=val('Autodesk Project ID');
      if (hubId && projectId) return {hubId:hubId,projectId:projectId,source:'BIM_Info'};
      const link=val('ACC Link');
      if (link && typeof apsResolveProjectLink==='function') {
        const r=apsResolveProjectLink(link);
        if (r && r.success) return {hubId:r.hub.id,projectId:r.project.id,source:'ACC Link'};
      }
    }
  }
  return null;
}
function setLiveAccProject(projectNo, hubId, projectId) {
  try {
    const user=liveCloudUser_(); liveCloudProject_(projectNo);
    const h=String(hubId||'').trim(), p=String(projectId||'').trim();
    if (!h || !p || h.length>300 || p.length>300 || /[\r\n]/.test(h+p)) throw new Error('יש לבחור HUB ופרויקט תקינים.');
    const token=apsUserToken_();
    const binding={hubId:h,projectId:p,user:user.email,selectedAt:new Date().toISOString()};
    liveCloudAssertAccess_(binding,token);
    PropertiesService.getUserProperties().setProperty(liveCloudBindingKey_(projectNo),JSON.stringify(binding));
    return {success:true,binding:{hubId:h,projectId:p}};
  } catch(e) { return {success:false,error:String(e.message||e)}; }
}
function liveCloudContext_(projectNo) {
  liveCloudProject_(projectNo);
  const b=liveCloudResolveBinding_(projectNo);
  if (!b) return {available:false,message:'לא הוגדר קישור ACC לפרויקט. בחר HUB ופרויקט מורשים.'};
  const token=apsUserToken_(), project=liveCloudAssertAccess_(b,token);
  return {available:true,token:token,binding:b,project:project};
}
function liveCloudPages_(firstUrl, token, cap, expectedPrefix) {
  const data=[], visited=Object.create(null); let url=firstUrl, pages=0, complete=false;
  while(url && pages < (cap || LIVE_CLOUD_PAGE_CAP_) && data.length < LIVE_CLOUD_ITEM_CAP_) {
    if (visited[url]) throw new Error('Autodesk returned a repeated pagination link.'); visited[url]=true;
    const body=liveCloudRequire200_(liveCloudFetch_(url,token));
    const page=Array.isArray(body.data)?body.data:[]; data.push.apply(data,page); pages++;
    let next=body.links && body.links.next && (body.links.next.href || body.links.next);
    if (next) {
      next=String(next);
      if (!next.startsWith(LIVE_CLOUD_BASE_+'/')) throw new Error('Autodesk returned a pagination link outside its API host.');
      if (expectedPrefix && !next.startsWith(expectedPrefix)) throw new Error('Autodesk returned a pagination link outside the selected project.');
      url=next;
    } else { complete=true; url=''; }
  }
  return {data:data,complete:complete,pages:pages,next:url||null};
}
function liveCloudIssuePath_(b, suffix) { return '/construction/issues/v1/projects/' + encodeURIComponent(String(b.projectId).replace(/^b\./,'')) + '/' + suffix; }
function liveCloudFolderPath_(b, suffix) { return '/data/v1/projects/' + encodeURIComponent(b.projectId) + '/' + suffix; }
function liveCloudTopFoldersPath_(b) { return '/project/v1/hubs/'+encodeURIComponent(b.hubId)+'/projects/'+encodeURIComponent(b.projectId)+'/topFolders'; }
function liveCloudIssues_(ctx) {
  const result=liveCloudIssueOffsetPages_(ctx,0);
  const rows=result.data.map(function(x){
    const a=x.attributes||x, status=String(a.status||'unknown').toLowerCase(), closed=['closed','resolved','void','cancelled'].indexOf(status)>=0;
    const comments=Array.isArray(a.comments)?a.comments:[];
    return {id:x.id,number:a.displayId||a.issueNumber||'',title:a.title||'',description:a.description||'',status:status,statusLabel:a.status||'לא ידוע',group:closed?'closed':'action',assignee:{id:(a.assignedTo&&a.assignedTo.id)||a.assignedTo||'',label:(a.assignedTo&&a.assignedTo.name)||a.assignedToName||'לא משויך'},dueDate:a.dueDate||'',updatedAt:a.updatedAt||'',discipline:(a.customAttributes||[]).filter(c=>/discipline/i.test(c.title||c.name||''))[0]?.value||'',priority:a.priority||'',commentCount:Number(a.commentCount||comments.length),commentsState:'not_loaded',comments:[],source:a};
  });
  const active=rows.filter(x=>x.group==='action').length, closed=rows.length-active;
  return {summary:{total:rows.length,active:active,closed:closed,void:0,overdue:0},activity:{weeks:[],coverageNoteHebrew:'היסטוריית פעילות אינה זמינה ב-API זה; מוצגים פרטי Issues שנקראו כעת.'},rows:rows,complete:result.complete,sourceReadAt:new Date().toISOString(),issuesComplete:result.complete,nextOffset:result.nextOffset,pagesRead:result.pages};
}
function liveCloudIssueOffsetPages_(ctx,offset) {
  const data=[];let current=Number(offset)||0,pages=0,total=null;
  while(pages<LIVE_CLOUD_PAGE_CAP_ && data.length<LIVE_CLOUD_ITEM_CAP_) {
    const path=liveCloudIssuePath_(ctx.binding,'issues?limit=100&offset='+current);
    const body=liveCloudRequire200_(liveCloudFetch_(path,ctx.token));
    const rows=Array.isArray(body.results)?body.results:(Array.isArray(body.data)?body.data:[]);
    const pg=body.pagination||{},rawTotal=pg.totalResults!=null?pg.totalResults:(body.totalResults!=null?body.totalResults:total); total=rawTotal==null?null:Number(rawTotal);
    data.push.apply(data,rows); pages++;
    const limit=Number(pg.limit)||100; if(!rows.length || rows.length<limit || (Number.isFinite(total)&&current+rows.length>=total)) {current=null;break;}
    current+=limit;
  }
  return {data:data,complete:current==null,nextOffset:current,pages:pages};
}
function getLiveCloudIssueComments_(projectNo, issueId) {
  try {
    const ctx=liveCloudContext_(projectNo); if(!ctx.available)return ctx;
    const id=String(issueId||'').trim(); if(!id||id.length>300)throw new Error('מזהה Issue אינו תקין.');
    const items=[];let offset=0,total=null,pages=0,ended=false;
    while(pages<LIVE_CLOUD_PAGE_CAP_) {
      const path=liveCloudIssuePath_(ctx.binding,'issues/'+encodeURIComponent(id)+'/comments?limit=100&offset='+offset);
      const body=liveCloudRequire200_(liveCloudFetch_(path,ctx.token)),rows=Array.isArray(body.results)?body.results:(Array.isArray(body.data)?body.data:[]),pg=body.pagination||{};
      const rawTotal=pg.totalResults!=null?pg.totalResults:(body.totalResults!=null?body.totalResults:total);total=rawTotal==null?null:Number(rawTotal);items.push.apply(items,rows);pages++;
      const limit=Number(pg.limit)||100;if(!rows.length||rows.length<limit||(Number.isFinite(total)&&offset+rows.length>=total)){ended=true;break;}offset+=limit;
    }
    const complete=ended || (Number.isFinite(total)&&items.length>=total);
    return {available:true,comments:items.map(function(x){const a=x.attributes||x;return {id:x.id,body:a.body||a.comment||a.message||'',createdAt:a.createdAt||'',author:{id:a.createdBy||'',label:a.createdByName||''}};}),complete:complete,nextOffset:complete?null:offset+100};
  } catch(e) {return {available:false,error:String(e.message||e)};}
}
function getLiveCloudIssuesNextPage(projectNo,nextOffset) {
  try {
    const ctx=liveCloudContext_(projectNo);if(!ctx.available)return ctx;
    const offset=Number(nextOffset);if(!Number.isInteger(offset)||offset<0)throw new Error('מיקום העמוד הבא אינו תקין.');
    const p=liveCloudIssueOffsetPages_(ctx,offset);
    return {available:true,items:p.data,complete:p.complete,nextOffset:p.nextOffset};
  }catch(e){return {available:false,error:String(e.message||e)};}
}
function getLiveCloudSnapshot_(projectNo) {
  try { const ctx=liveCloudContext_(projectNo); if(!ctx.available)return ctx;
    const issues=liveCloudIssues_(ctx);
    return {available:true,source:'Autodesk live user session',project:{acc:{projectName:(ctx.project.attributes||{}).name||'',hub:ctx.binding.hubId},management:issues,sourceReadAt:new Date().toISOString(),complete:issues.complete}};
  } catch(e){return {available:false,error:String(e.message||e)};}
}

// Publish model inventory. Each row is a current Autodesk file version with a lazy manifest.
function liveCloudFileInventory_(ctx) {
  const topPrefix=LIVE_CLOUD_BASE_+'/project/v1/hubs/'+encodeURIComponent(ctx.binding.hubId)+'/projects/'+encodeURIComponent(ctx.binding.projectId)+'/';
  const top=liveCloudPages_(liveCloudTopFoldersPath_(ctx.binding),ctx.token,1,topPrefix);
  const queue=(top.data||[]).map(x=>({id:x.id,path:(x.attributes&&x.attributes.displayName)||x.attributes&&x.attributes.name||''}));
  const files=[]; let complete=top.complete, steps=0;
  while(queue.length && steps<100 && files.length<LIVE_CLOUD_ITEM_CAP_) {
    const f=queue.shift(); steps++;
    const page=liveCloudPages_(liveCloudFolderPath_(ctx.binding,'folders/'+encodeURIComponent(f.id)+'/contents?limit=100'),ctx.token,3,LIVE_CLOUD_BASE_+'/data/v1/projects/'+encodeURIComponent(ctx.binding.projectId)+'/');
    complete=complete&&page.complete;
    (page.data||[]).forEach(x=>{if(x.type==='folders')queue.push({id:x.id,path:f.path+'/'+((x.attributes&&x.attributes.name)||'')});else if(x.type==='items')files.push({item:x,path:f.path});});
  }
  if(queue.length)complete=false;
  const models=[];let versionReads=0,versionReadsComplete=true;
  files.forEach(function(entry){if(versionReads>=100){versionReadsComplete=false;return;}const item=entry.item, rels=item.relationships||{}, versions=rels.tip&&rels.tip.data?[rels.tip.data]:[]; const v=versions[0]; if(!v)return;
    const a=item.attributes||{}, name=a.displayName||a.name||'', ext=(name.match(/\.([^.]+)$/)||[])[1]||'';
    if(!/^rvt$/i.test(ext))return;
    versionReads++;
    let versionData=null;
    try {const tip=liveCloudRequire200_(liveCloudFetch_(liveCloudFolderPath_(ctx.binding,'items/'+encodeURIComponent(item.id)+'/tip'),ctx.token));versionData=tip.data||null;}catch(e){versionReadsComplete=false;}
    const vId=String(versionData&&versionData.id||v.id||''), va=versionData&&versionData.attributes||v.attributes||{}, num=va.versionNumber||(vId.match(/[?&]version=(\d+)/i)||[])[1]||'';
    models.push({id:item.id,name:name,disciplineGroup:'unknown',discipline:'ללא דיסציפלינה מזוהה',disciplineSource:'filename',version:Number(num)||num||'',versionId:vId,webView:((item.links||{}).webView||{}).href||'',folderPath:entry.path,versionUpdatedAt:va.lastModifiedTime||a.lastModifiedTime||'',sourceCapturedAt:new Date().toISOString(),versionReadStatus:versionData?'READ':'UNAVAILABLE',viewsStatus:'NOT_READ',viewsMessage:'התצוגות ייטענו לפי דרישה.',views:[],manifestComplete:false,manifestStatus:'NOT_READ'});
  });
  return {models:models,complete:complete&&queue.length===0&&versionReadsComplete,checkedAt:new Date().toISOString(),pagesBounded:!complete||!versionReadsComplete};
}
function getLivePublishSnapshot_(projectNo) {
  try {const ctx=liveCloudContext_(projectNo);if(!ctx.available)return ctx;const inv=liveCloudFileInventory_(ctx);
    return {available:true,snapshot:{models:inv.models,complete:inv.complete,source:'Autodesk live user session',sourceReadAt:inv.checkedAt,inventoryComplete:inv.complete}};
  }catch(e){return {available:false,error:String(e.message||e)};}
}
function liveCloudEncodeUrn_(urn) { const s=String(urn);return /^urn:/i.test(s)?Utilities.base64EncodeWebSafe(s,'UTF-8').replace(/=+$/g,''):s; }
function liveCloudDerivativeUrn_(ctx,modelId,versionId) {
  const tip=liveCloudRequire200_(liveCloudFetch_('/data/v1/projects/'+encodeURIComponent(ctx.binding.projectId)+'/items/'+encodeURIComponent(modelId)+'/tip',ctx.token));
  if(!tip.data || String(tip.data.id)!==String(versionId))throw new Error('Autodesk item tip changed; reload the Publish inventory.');
  const rel=tip.data.relationships||{};let urn=rel.derivatives&&rel.derivatives.data&&rel.derivatives.data.id;
  if(!urn){const v=liveCloudRequire200_(liveCloudFetch_('/data/v1/projects/'+encodeURIComponent(ctx.binding.projectId)+'/versions/'+encodeURIComponent(versionId),ctx.token));const vr=v.data&&v.data.relationships||{};urn=vr.derivatives&&vr.derivatives.data&&vr.derivatives.data.id;}
  if(!urn) throw new Error('Autodesk did not return the published derivative URN from the version relationship.'); return String(urn);
}
function getLivePublishModel_(projectNo,modelId,versionId) {
  try {const ctx=liveCloudContext_(projectNo);if(!ctx.available)return ctx;
    const tip=liveCloudRequire200_(liveCloudFetch_('/data/v1/projects/'+encodeURIComponent(ctx.binding.projectId)+'/items/'+encodeURIComponent(modelId)+'/tip',ctx.token));
    if(!tip.data||String(tip.data.id)!==String(versionId))throw new Error('המודל או הגרסה אינם תואמים לגרסה הנוכחית ב-ACC.');
    const item=tip.data, a=item.attributes||{}, num=a.versionNumber||(String(versionId).match(/[?&]version=(\d+)/i)||[])[1]||'';
    const m={id:String(modelId),name:a.displayName||a.name||'',version:Number(num)||num||'',versionId:String(versionId),versionUpdatedAt:a.lastModifiedTime||'',views:[],viewsStatus:'READ_AVAILABLE',manifestComplete:false,manifestStatus:'NOT_READ'};
    const rawUrn=liveCloudDerivativeUrn_(ctx,modelId,versionId), urn=liveCloudEncodeUrn_(rawUrn);
    const res=liveCloudFetch_('/modelderivative/v2/designdata/'+encodeURIComponent(urn)+'/manifest',ctx.token);
    if(res.code!==200) return {available:true,model:m,manifestComplete:false,manifestStatus:'UNAVAILABLE',manifestHttpStatus:res.code,views:[]};
    const manifest=res.body&&res.body.data||res.body||{}, views=liveCloudNormalizeManifest_(manifest);
    m.urn=urn;m.derivativeUrn=urn;m.views=views;m.manifestComplete=manifest.status==='success'||manifest.status==='complete'||manifest.progress==='complete';m.manifestStatus=manifest.status||manifest.progress||'unknown';
    const failed=/^(failed|timeout)$/i.test(String(manifest.status||manifest.progress||''));
    const hasUnknown2d=views.some(v=>v.kind==='unknown'),has2d=views.some(v=>v.kind==='2d'||v.kind==='sheet'),has3d=views.some(v=>v.kind==='3d');
    m.viewsStatus=failed?'ERROR':(m.manifestComplete?(views.length?'LOADED':'EMPTY'):'READ_AVAILABLE');
    m.viewsMessage=failed?'Manifest derivation failed according to the Autodesk status.':!m.manifestComplete?'Manifest processing is incomplete; current view results may be partial.':!views.length?'Manifest completed but contained no identifiable views.':hasUnknown2d?'A 2D view was found, but the manifest did not provide enough evidence to classify it as a sheet or other 2D view.':has2d?'2D/3D types follow manifest roles; sheets require explicit thumbnail resource path evidence.':has3d?'3D views were identified by manifest role; no 2D views were identified.':'Views were found without a clear 2D/3D role.';
    views.forEach(v=>v.derivativeUrn=urn);m.derivativeUrnSource='version relationships.derivatives';
    return {available:true,model:m,manifestComplete:m.manifestComplete};
  }catch(e){return {available:false,error:String(e.message||e)};}
}
function liveCloudNormalizeManifest_(manifest) {
  const views=[];
  function gather(node,out){(node.children||[]).forEach(function(c){out.push(c);gather(c,out);});}
  function visit(node){const role=String(node.role||'').toLowerCase(),type=String(node.type||'').toLowerCase();
    if(type==='geometry'&&(role==='2d'||role==='3d')){
      const descendants=[];gather(node,descendants);const resources=descendants.filter(c=>c.type==='resource'&&c.urn),thumbs=resources.filter(c=>String(c.role||'').toLowerCase()==='thumbnail'),thumb=thumbs.filter(c=>Array.isArray(c.resolution)&&c.resolution[0]>=400)[0]||thumbs.filter(c=>Array.isArray(c.resolution)&&c.resolution[0]>=200)[0]||thumbs[0];
      const pdf=resources.filter(c=>String(c.role||'').toLowerCase()==='pdf-page'&&/\.pdf(?:$|\?)/i.test(c.urn))[0],meta=resources.filter(c=>String(c.role||'').toLowerCase()==='metadata')[0],id=String(node.guid||node.viewableID||'');
      if(id){let kind=role,evidence='manifest geometry role';if(role==='2d'){if(!thumb){kind='unknown';evidence='2d role without thumbnail resource evidence';}else if(/\/output\/Resource\/Sheet\//i.test(String(thumb.urn||''))){kind='sheet';evidence='manifest thumbnail resource path';}}
        const v={id:id,name:String(node.name||id),kind:kind,kindEvidence:evidence,isMasterView:node.isMasterView===true||node.isMaster===true||node.isInternalMaster===true,publishSet:null,sheetName:null,sheetNumber:null,derivative:{guid:String(node.guid||id),viewableID:node.viewableID?String(node.viewableID):null,role:role,type:node.type}};
        if(thumb&&thumb.urn)v.thumbnailUrn=String(thumb.urn);if(pdf&&pdf.urn)v.pdfUrn=String(pdf.urn);if(meta&&meta.urn)v.metadataUrn=String(meta.urn);views.push(v);}}
    (node.children||[]).forEach(visit);
  }
  (manifest.derivatives||[]).forEach(visit);return views;
}
// google.script.run RPC contract for the lazy Publish view loader.
function getLivePublishModel(projectNo,modelId,versionId) {
  const result=getLivePublishModel_(projectNo,modelId,versionId);
  return result.available ? {success:true,model:result.model,manifestComplete:result.manifestComplete===true} : {success:false,error:result.error||result.message||'Manifest is unavailable.'};
}
function getLivePublishModelThumbnails_(projectNo,modelId,versionId) {
  const detail=getLivePublishModel_(projectNo,modelId,versionId); if(!detail.available||!detail.model)return detail;
  const ctx=liveCloudContext_(projectNo), urn=detail.model.urn, images={}; let complete=true;
  detail.model.views.forEach(function(v){if(!v.id)return;const path='/modelderivative/v2/designdata/'+encodeURIComponent(urn)+'/thumbnail?guid='+encodeURIComponent(v.id)+'&width=200';const r=liveCloudFetch_(path,ctx.token);const type=String(r.headers['Content-Type']||r.headers['content-type']||'').split(';')[0].toLowerCase();if(r.code===200&&type==='image/png'&&r.bytes&&r.bytes.length)images[v.id]='data:image/png;base64,'+Utilities.base64Encode(r.bytes);else complete=false;});
  return {available:Object.keys(images).length>0,images:images,complete:complete};
}
function getLivePublishModelThumbnails(projectNo,modelId,versionId) {
  const result=getLivePublishModelThumbnails_(projectNo,modelId,versionId);
  return result.available ? {success:true,images:result.images,complete:result.complete} : {success:false,error:result.error||'No thumbnails are available.'};
}
