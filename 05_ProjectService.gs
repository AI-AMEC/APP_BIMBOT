// ==========================================
// BIMBOT v41.0 - 05_ProjectService
// Pull / save the project's BIM INFO spreadsheet + save guards
// ==========================================

function getSheetIgnoreCase(ss, name) {
  return ss.getSheets().find(s => s.getName().toLowerCase() === name.toLowerCase());
}

function localWipe(sheet) {
  try {
    if (sheet.getMaxRows() > 0 && sheet.getMaxColumns() > 0) {
      sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
    }
    sheet.clear();
  } catch(e) { throw e; }
}

function safeProtect_(id) {
    try {
        if (typeof protectProjectSpreadsheet_ !== 'undefined') {
            protectProjectSpreadsheet_(id);
        }
    } catch(e) {}
}

function pullProjectDataFromSheet(folderId, projectNo, projectName) {
  try {
    const db = getLocalProjectDbAndBimFolder_(folderId, projectNo, projectName, false);
    const ss = db.ss;
    let payload = {
      bimInfo: { code: "", weeklyDay: "", weeklyTime: "", revit: "", cdePlatform: "", docManagement: "", bimPlatform: "", driveFolder: "", acc: "", docLink: "", hub: "" },
      kickoff: { arcModel: "", grids: "", gridsX: "", gridsY: "", arLevel: "", urs: "", glLevel: "0.00", liveShared: "", licensing: "", ganttA: "", ganttB: "", ganttC: "", hasPhases: "" },
      lots: [], buildings: [], basements: [], genModels: [], disciplines: [], matrix: {}, contacts: [], models: [],
      siteGroups: [], designOptions: [], versionHistory: [], notebook: [],
      hasData: false, sheetUrl: ss ? ss.getUrl() : "" 
    };

    if (!ss) return { success: true, data: payload };
    let sInfo = getSheetIgnoreCase(ss, "BIM_Info");
    if(sInfo) {
      let data = sInfo.getDataRange().getDisplayValues();
      if(data.length > 1) {
        payload.hasData = true;
        let getVal = (key) => {
            let row = data.find(r => r[0] && r[0].toString().toLowerCase().includes(key.toLowerCase()));
            return row ? row[1] || "" : "";
        };

        payload.bimInfo.code = getVal("Project Code") || getVal("Project Identifier") || projectNo;
        payload.bimInfo.manager = getVal("Assigned BIM Manager"); 
        payload.bimInfo.coordinator = getVal("Assigned Coordinator"); 
        payload.bimInfo.revit = getVal("Revit Version"); 
        payload.bimInfo.cdePlatform = getVal("CDE Platform"); 
        payload.bimInfo.docManagement = getVal("Doc Management"); 
        payload.bimInfo.bimPlatform = getVal("BIM Platform"); 
        payload.bimInfo.driveFolder = getVal("Drive Folder Link") || getVal("Drive Folder"); 
        payload.bimInfo.acc = getVal("ACC Link"); 
        payload.bimInfo.docLink = getVal("Doc Management Link"); 
        payload.bimInfo.hub = getVal("HUB"); 
        payload.bimInfo.weeklyDay = getVal("Weekly Meeting");
        payload.bimInfo.weeklyTime = getVal("TIME");
        // Autodesk (BIM 360 / ACC) connection - exact key match ("HUB" above is a substring match).
        let getExact = (key) => {
            let row = data.find(r => String(r[0] || "").trim().toLowerCase() === key.toLowerCase());
            return row ? row[1] || "" : "";
        };
        payload.bimInfo.accHubName = getExact("Autodesk Hub");
        payload.bimInfo.accHubId = getExact("Autodesk Hub ID");
        payload.bimInfo.accProjectName = getExact("Autodesk Project");
        payload.bimInfo.accProjectId = getExact("Autodesk Project ID");
      }
    }

    let sKick = getSheetIgnoreCase(ss, "Kick_Off");
    if(sKick) {
      let data = sKick.getDataRange().getDisplayValues();
      if (data.length > 1) {
        payload.hasData = true;
        let getKickVal = (key) => {
            let row = data.find(r => r[0] && r[0].toString().toLowerCase().includes(key.toLowerCase()));
            return row ? row[1] || "" : "";
        };

        payload.kickoff.arcModel = mapYesNo(getKickVal("Arc Model")); 
        
        let rawGrids = getKickVal("Grids").trim();
        if (rawGrids.startsWith("No") && rawGrids.includes("|")) {
            payload.kickoff.grids = "No";
            let parts = rawGrids.split("|");
            payload.kickoff.gridsX = parts[1] ? parts[1].trim() : "";
            payload.kickoff.gridsY = parts[2] ? parts[2].trim() : "";
        } else {
            payload.kickoff.grids = mapYesNo(rawGrids);
        }

        payload.kickoff.arLevel = mapYesNo(getKickVal("AR Level")); 
        payload.kickoff.urs = mapYesNo(getKickVal("URS")); 
        payload.kickoff.glLevel = getKickVal("GL Level") || "0.00";
        
        let liveSharedStr = getKickVal("Live / Shared").toLowerCase();
        if (liveSharedStr.includes("live")) payload.kickoff.liveShared = "Live"; 
        else if (liveSharedStr.includes("shared")) payload.kickoff.liveShared = "Shared"; 
        else payload.kickoff.liveShared = getKickVal("Live / Shared");
        
        payload.kickoff.licensing = mapYesNo(getKickVal("Licensing"));
        
        payload.kickoff.ganttA = getKickVal("Gantt A") || "";
        payload.kickoff.ganttB = getKickVal("Gantt B") || "";
        payload.kickoff.ganttC = getKickVal("Gantt C") || "";
        payload.kickoff.hasPhases = getKickVal("Has Phases") || "";
      }
    }

    let sLots = getSheetIgnoreCase(ss, "Lots");
    if(sLots && sLots.getLastRow() > 1) {
      payload.hasData = true;
      let data = sLots.getDataRange().getDisplayValues();
      for(let i=1; i<data.length; i++) { if(data[i][0]) payload.lots.push({ number: data[i][0] }); }
    }

    let sBldgs = getSheetIgnoreCase(ss, "Buildings");
    if(sBldgs && sBldgs.getLastRow() > 1) {
      payload.hasData = true;
      let data = sBldgs.getDataRange().getDisplayValues();
      for(let i=1; i<data.length; i++) {
        let colA = data[i][0] ? String(data[i][0]).trim() : "";
        let colB = data[i][1] ? String(data[i][1]).trim() : "";
        let colC = data[i][2] ? String(data[i][2]).trim() : "";
        let colD = data[i][3] ? String(data[i][3]).trim() : "";
        let colE = data[i][4] ? String(data[i][4]).trim() : "";
        let colF = data[i][5] ? String(data[i][5]).trim() : "";
        
        if (!colA) continue;
        
        let isGlChecked = false;
        let customGlVal = "";
        if (colF !== "" && colF.toLowerCase() !== "false") {
            isGlChecked = true;
            if (colF.toUpperCase() !== "TRUE" && colF.toUpperCase() !== "V" && colF.toUpperCase() !== "YES" && colF !== "1") {
                customGlVal = colF;
            }
        }

        let isBasement = (colE !== "");
        if (isBasement) { 
            payload.basements.push({ code: colA, desc: colB, buildings: colE, floors: colC, lots: colD, gl: isGlChecked, glValue: customGlVal }); 
        } else { 
            payload.buildings.push({ name: colA, desc: colB, floors: colC, lot: colD, gl: isGlChecked, glValue: customGlVal }); 
        }
      }
    }

    let sGen = getSheetIgnoreCase(ss, "Revit_sites");
    let siteGroupRows = sGen && sGen.getLastRow() > 1 ? sGen.getDataRange().getDisplayValues() : [];
    if (siteGroupRows.length > 1) payload.hasData = true;
    payload.siteGroups = siteGroupRows.slice(1).map((row, offset) => {
      const firstBuilding = String(row[0] || '').trim();
      const modelBuildings = String(row[1] || '').trim();
      const customName = String(row[2] || '').trim();
      const buildings = modelBuildings ? modelBuildings.split('-').map(code => code.trim()).filter(Boolean) : [];
      if (!firstBuilding || !modelBuildings.includes('-')) return null;
      return {
        id: 'grp_' + (offset + 2) + '_' + buildings.join('-'),
        name: customName || modelBuildings,
        type: String(row[4] || '').trim() || (customName ? 'מאוחד' : 'טיפוסי'),
        color: String(row[3] || '').trim() || 'purple',
        buildings: buildings
      };
    }).filter(Boolean);

    let sDesOpt = getSheetIgnoreCase(ss, "Design_Option");
    let designOptions = [];
    if(sDesOpt && sDesOpt.getLastRow() > 1) {
      payload.hasData = true;
      let dataDes = sDesOpt.getDataRange().getDisplayValues();
      for (let i = 1; i < dataDes.length; i++) {
          let targetCode = dataDes[i][0] ? String(dataDes[i][0]).trim() : "";
          let optV = dataDes[i][1] ? String(dataDes[i][1]).trim().toUpperCase() : "";
          let title = dataDes[i][2] ? String(dataDes[i][2]).trim() : "";
          let desc = dataDes[i][3] ? String(dataDes[i][3]).trim() : "";

          if (targetCode && (optV === "V" || optV === "YES" || optV === "TRUE")) {
              let isGroup = payload.siteGroups.some(g => g.name === targetCode || g.buildings.join("-") === targetCode);
              designOptions.push({
                  id: 'do_' + i + '_' + Date.now(),
                  modelId: null, 
                  targetCode: targetCode,
                  isGroupTarget: isGroup,
                  title: title,
                  desc: desc
              });
          }
      }
    }
    payload.designOptions = designOptions;

    let sDisc = getSheetIgnoreCase(ss, "Disciplines"); 
    if(sDisc && sDisc.getLastRow() > 1) {
      payload.hasData = true;
      let data = sDisc.getDataRange().getDisplayValues();
      for(let i=1; i<data.length; i++) { 
        let discCode = data[i][0] ? data[i][0].toString().trim() : "";
        let discHe = data[i][1] ? data[i][1].toString().trim() : "";
        let orgCode = data[i][2] ? data[i][2].toString().trim() : "";
        let platform = data[i][3] ? data[i][3].toString().trim() : "Revit";
        let scope = data[i][4] ? data[i][4].toString().trim() : "בניין";
        
        if(discCode || discHe) { 
            payload.disciplines.push({ ac: discCode, he: discHe || discCode, orgCode: orgCode, platform: platform, scope: scope, checked: true }); 
        }
      }
    }

    let sCont = getSheetIgnoreCase(ss, "Contacts");
    if(sCont && sCont.getLastRow() > 1) {
      payload.hasData = true;
      let data = sCont.getDataRange().getDisplayValues();
      for(let i=1; i<data.length; i++) { if(data[i][0]) payload.contacts.push({name: data[i][0], email: data[i][1], discipline: data[i][2], role: data[i][3], company: data[i][4]}); }
    }
    
    let sNote = getSheetIgnoreCase(ss, "Notebook");
    if (sNote && sNote.getLastRow() > 1) {
      payload.hasData = true;
      let noteData = sNote.getDataRange().getDisplayValues();
      for(let i = 1; i < noteData.length; i++) {
        if(noteData[i][0]) {
          payload.notebook.push({
            date: noteData[i][0],
            user: noteData[i][1],
            content: noteData[i][2]
          });
        }
      }
    }

    let sHist = getSheetIgnoreCase(ss, "Version_History");
    if (sHist && sHist.getLastRow() > 1) {
      let histData = sHist.getDataRange().getDisplayValues();
      for(let i = 1; i < histData.length; i++) {
        if(histData[i][0]) {
          payload.versionHistory.push({
            timestamp: histData[i][0],
            user: histData[i][1],
            action: histData[i][2]
          });
        }
      }
      payload.versionHistory.reverse(); 
    }

    let sModel = getSheetIgnoreCase(ss, "Model_Table");
    payload.models = readModelTable_(sModel);
    if (payload.models.length > 0) payload.hasData = true;

    payload.modelSchema = readModelSchema_(ss);

    return { success: true, data: payload };
  } catch(e) { return { success: false, error: e.toString() }; }
}

function generateMasterProjectSheet(folderId, projectNo, projectName, payload) {
  let saveLock = null;
  let backupUrl = "";
  try {
    validateProjectPayload_(payload, projectNo);
    const db = getLocalProjectDbAndBimFolder_(folderId, projectNo, projectName);
    const ss = db.ss;
    saveLock = LockService.getScriptLock();
    saveLock.waitLock(30000);
    backupUrl = backupProjectBeforeSave_(db);
    
    function reorderSheet(sheetName, index) {
      let sheet = getSheetIgnoreCase(ss, sheetName);
      if (!sheet) { sheet = ss.insertSheet(sheetName); }
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(index + 1);
      return sheet;
    }
    
    let sInfo = reorderSheet("BIM_Info", 0);
    let sCont = reorderSheet("Contacts", 1);
    let sKick = reorderSheet("Kick_Off", 2);
    let sLots = reorderSheet("Lots", 3);
    let sBldgs = reorderSheet("Buildings", 4);
    let sDisc = reorderSheet("Disciplines", 5);
    let sGen = reorderSheet("Revit_sites", 6);
    let sDesOpt = reorderSheet("Design_Option", 7);
    let sModel = reorderSheet("Model_Table", 8);
    let sNote = reorderSheet("Notebook", 9);

    let sHist = getSheetIgnoreCase(ss, "Version_History");
    if (!sHist) {
      sHist = ss.insertSheet("Version_History");
      sHist.getRange(1, 1, 1, 3).setValues([["Timestamp", "User", "Action"]]).setFontWeight("bold");
    }
    let currentUser = "Unknown User";
    try { currentUser = Session.getActiveUser().getEmail() || "Unknown App User"; } catch(e){}
    let timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");


    // Legacy tabs are retained; migration must be explicit.
    projectNo = db.project.no;
    projectName = db.project.name;
    const info = payload.bimInfo || {};
    localWipe(sInfo); 
    let infoData = [
      ["Project Code", info.code || ""],
      ["Project Identifier", projectNo || ""],
      ["Project Name", projectName || ""],
      ["Assigned BIM Manager", info.manager || ""],
      ["Assigned Coordinator", info.coordinator || ""],
      ["Revit Version", info.revit || ""],
      ["CDE Platform", info.cdePlatform || ""],
      ["Doc Management", info.docManagement || ""],
      ["BIM Platform", info.bimPlatform || ""],
      ["Drive Folder Link", info.driveFolder || ""],
      ["ACC Link", info.acc || ""],
      ["Doc Management Link", info.docLink || ""],
      ["HUB", info.hub || ""],
      ["Weekly Meeting", info.weeklyDay || ""],
      ["TIME", info.weeklyTime || ""],
      // Keep these after "HUB": pull reads "HUB" by substring and takes the first match.
      ["Autodesk Hub", info.accHubName || ""],
      ["Autodesk Hub ID", info.accHubId || ""],
      ["Autodesk Project", info.accProjectName || ""],
      ["Autodesk Project ID", info.accProjectId || ""]
    ];
    sInfo.getRange(1, 1, infoData.length, 2).setValues(infoData).setHorizontalAlignment('center').setVerticalAlignment('middle');
    sInfo.getRange(1, 1, infoData.length, 1).setFontWeight("bold").setHorizontalAlignment('center');

    const contacts = payload.contacts || [];
    localWipe(sCont);
    const contHeaders = ["Name", "Email", "Discipline", "Role", "Company"];
    sCont.getRange(1, 1, 1, contHeaders.length).setValues([contHeaders]).setFontWeight("bold").setHorizontalAlignment('center');
    if(contacts.length > 0) {
       let rng = sCont.getRange(2, 1, contacts.length, 5);
       rng.setValues(contacts.map(c => [c.name || "", c.email || "", c.discipline || "", c.role || "", c.company || ""]));
       rng.setHorizontalAlignment('center').setVerticalAlignment('middle');
    }

    const kickoff = payload.kickoff || {};
    localWipe(sKick); 
    let fCell = kickoff.liveShared || "";
    if (kickoff.liveShared === "Live" && payload.liveLink) fCell = '=HYPERLINK("' + payload.liveLink + '", "Live")';
    else if (kickoff.liveShared === "Shared" && payload.sharedLink) fCell = '=HYPERLINK("' + payload.sharedLink + '", "Shared")';

    let saveGridsValue = kickoff.grids || "";
    if (saveGridsValue === "No" && (kickoff.gridsX || kickoff.gridsY)) {
        let xVal = kickoff.gridsX || "";
        let yVal = kickoff.gridsY || "";
        saveGridsValue = "No | " + xVal + " | " + yVal;
    }

    let kickData = [
        ["Arc Model", kickoff.arcModel || ""],
        ["Grids", saveGridsValue],
        ["AR Level", kickoff.arLevel || ""],
        ["URS", kickoff.urs || ""],
        ["GL Level", (kickoff.glLevel !== undefined && kickoff.glLevel !== "") ? kickoff.glLevel : "0.00"],
        ["Live / Shared", fCell],
        ["Licensing", kickoff.licensing || ""],
        ["Gantt A", kickoff.ganttA || ""],
        ["Gantt B", kickoff.ganttB || ""],
        ["Gantt C", kickoff.ganttC || ""],
        ["Has Phases", kickoff.hasPhases || ""]
    ];
    sKick.getRange(1, 1, kickData.length, 2).setValues(kickData).setHorizontalAlignment('center').setVerticalAlignment('middle');
    sKick.getRange(1, 1, kickData.length, 1).setFontWeight("bold").setHorizontalAlignment('center');

    const lots = payload.lots || [];
    localWipe(sLots);
    const lotsHeaders = ["Building Lot Number"];
    sLots.getRange(1, 1, 1, lotsHeaders.length).setValues([lotsHeaders]).setFontWeight("bold").setHorizontalAlignment('center');
    if(lots.length > 0) {
      let lotRows = lots.map(l => [l.number || ""]);
      let rng = sLots.getRange(2, 1, lotRows.length, 1);
      rng.setValues(lotRows);
      rng.setHorizontalAlignment('center').setVerticalAlignment('middle');
    }

    const buildings = payload.buildings || [];
    const basements = payload.basements || [];
    localWipe(sBldgs); 
    
    const bldgHeaders = ["Building_Code", "Building Description", "Floor Number", "Building Lot Number Association", "Building_Association", "GL"];
    sBldgs.getRange(1, 1, 1, bldgHeaders.length).setValues([bldgHeaders]).setFontWeight("bold").setHorizontalAlignment('center');
    
    let combinedBldgData = [];
    buildings.forEach(b => {
        let fVal = b.gl ? (b.glValue && b.glValue.trim() !== "" ? b.glValue.trim() : true) : false;
        combinedBldgData.push([b.name || "", b.desc || "", b.floors || "", b.lot || "", "", fVal]);
    });
    basements.forEach(b => {
        let fVal = b.gl ? (b.glValue && b.glValue.trim() !== "" ? b.glValue.trim() : true) : false;
        combinedBldgData.push([b.code || "", b.desc || "", b.floors || "", b.lots || "", b.buildings || "", fVal]);
    });
    
    if (combinedBldgData.length > 0) {
        sBldgs.getRange(2, 1, combinedBldgData.length, 6).setValues(combinedBldgData).setHorizontalAlignment('center').setVerticalAlignment('middle');
    }

    const siteGroups = projectSiteGroups_(payload);
    const designOptions = payload.designOptions || [];
    localWipe(sGen);    
    localWipe(sDesOpt); 
    
    const genHeaders = ["Building_Code", "Generic_Model", "Custom_Site_Name", "Group_Color", "Group_Type"];
    sGen.getRange(1, 1, 1, genHeaders.length).setValues([genHeaders]).setFontWeight("bold").setHorizontalAlignment('center');

    const desOptHeaders = ["Building Code", "Design Option", "Option Name", "Option Desc"];
    sDesOpt.getRange(1, 1, 1, desOptHeaders.length).setValues([desOptHeaders]).setFontWeight("bold").setHorizontalAlignment('center');
    
    let revitSitesRows = siteGroupRows_(siteGroups);
    let desOptRows = [];

    if (designOptions.length > 0) {
        designOptions.forEach(opt => {
            let targetCode = opt.targetCode || "";
            desOptRows.push([targetCode, "V", opt.title || "", opt.desc || ""]);
        });
    }

    if (revitSitesRows.length > 0) sGen.getRange(2, 1, revitSitesRows.length, 5).setValues(revitSitesRows).setHorizontalAlignment('center').setVerticalAlignment('middle');
    if (desOptRows.length > 0) sDesOpt.getRange(2, 1, desOptRows.length, 4).setValues(desOptRows).setHorizontalAlignment('center').setVerticalAlignment('middle');

    const disciplines = payload.disciplines || [];
    localWipe(sDisc);
    const discHeaders = ["Discipline Code", "Discipline Hebrew", "Organization Code", "Platform", "Scope"];
    sDisc.getRange(1, 1, 1, discHeaders.length).setValues([discHeaders]).setFontWeight("bold").setHorizontalAlignment('center');
    if(disciplines.length > 0) {
       let rng = sDisc.getRange(2, 1, disciplines.length, 5);
       rng.setValues(disciplines.map(d => [d.ac || "", d.he || "", d.orgCode || "", d.platform || "", d.scope || ""]));
       rng.setHorizontalAlignment('center').setVerticalAlignment('middle');
    }

    if (shouldReplaceModelTable_(payload)) {
        localWipe(sModel);
        sModel.getRange(1, 1, 1, MODEL_TABLE_HEADERS_.length).setValues([MODEL_TABLE_HEADERS_]).setFontWeight("bold").setHorizontalAlignment('center');
        if (payload.models.length > 0) {
            let modelTableRows = modelTableRows_(payload.models);
            sModel.getRange(2, 1, modelTableRows.length, MODEL_TABLE_HEADERS_.length).setValues(modelTableRows).setHorizontalAlignment('center').setVerticalAlignment('middle');
        }
    }

    const notebook = payload.notebook || [];
    localWipe(sNote);
    const noteHeaders = ["Date", "User", "Content"];
    sNote.getRange(1, 1, 1, noteHeaders.length).setValues([noteHeaders]).setFontWeight("bold").setHorizontalAlignment('center');
    if(notebook.length > 0) {
        let noteRows = notebook.map(n => [n.date || "", n.user || "", n.content || ""]);
        sNote.getRange(2, 1, noteRows.length, 3).setValues(noteRows).setHorizontalAlignment('left').setVerticalAlignment('top');
    }

    // Step-5 layout: only written when the client sends one (older clients leave it untouched).
    if (payload.modelSchema) writeModelSchema_(ss, payload.modelSchema);

    safeProtect_(ss.getId());
    SpreadsheetApp.flush();

    // Mirror the project code to the central Project_Code table (only when it changed).
    let warning = "";
    const newCode = String(info.code || "").trim();
    if (newCode) {
      try {
          const codeSheet = SpreadsheetApp.openById(MAIN_DB_ID).getSheetByName("Project_Code");
          if (codeSheet) {
              const codeData = codeSheet.getDataRange().getValues();
              const rowIndex = codeData.findIndex((r, i) => i > 0 && String(r[0]).trim() === String(projectNo).trim());
              if (rowIndex < 0) codeSheet.appendRow([projectNo, newCode]);
              else if (String(codeData[rowIndex][1]).trim() !== newCode) codeSheet.getRange(rowIndex + 1, 2).setValue(newCode);
          }
      } catch(e) {
          warning = "Project sheets saved, but Project_Code update failed: " + e;
          console.log(warning);
      }
    }

    sHist.appendRow([timestamp, currentUser, "Updated Project Data via App"]);
    return { success: true, url: ss.getUrl(), backupUrl: backupUrl, warning: warning };
  } catch(e) { return { success: false, error: e.toString(), backupUrl: backupUrl }; }
  finally { if (saveLock) saveLock.releaseLock(); }
}

// ---------- Save guards (payload validation, model table, backup) ----------

function validateProjectPayload_(payload, projectNo) {
  if (!payload || typeof payload !== 'object' || Array.isArray(payload)) throw new Error('Missing complete project payload.');
  ['bimInfo', 'kickoff'].forEach(key => {
    if (!payload[key] || typeof payload[key] !== 'object' || Array.isArray(payload[key])) throw new Error('Missing project section: ' + key);
  });
  ['contacts', 'lots', 'buildings', 'basements', 'designOptions', 'disciplines', 'models', 'notebook'].forEach(key => {
    if (!Array.isArray(payload[key])) throw new Error('Missing project table: ' + key);
    if (payload[key].some(row => !row || typeof row !== 'object' || Array.isArray(row))) throw new Error('Invalid row in ' + key);
  });
  if (Object.prototype.hasOwnProperty.call(payload, 'siteGroups')) {
    if (!Array.isArray(payload.siteGroups) || payload.siteGroups.some(group => !group || typeof group !== 'object' || Array.isArray(group) || !Array.isArray(group.buildings))) {
      throw new Error('Invalid project table: siteGroups');
    }
  } else if (payload.clientSchemaVersion === 1 && Array.isArray(payload.genModels)) {
    // Explicit v1 compatibility path for clients that predate siteGroups.
  } else {
    throw new Error('Missing project table: siteGroups');
  }
  if (Object.prototype.hasOwnProperty.call(payload, 'genModels') && !Array.isArray(payload.genModels)) {
    throw new Error('Invalid project table: genModels');
  }
  if (payload.modelSchema != null) normalizeModelSchema_(payload.modelSchema); // throws before any sheet changes
  if (Object.prototype.hasOwnProperty.call(payload, 'modelsDirty') && typeof payload.modelsDirty !== 'boolean') {
    throw new Error('Invalid modelsDirty intent flag.');
  }
  if (String(payload.bimInfo.projectNumber || '').trim() !== String(projectNo || '').trim()) throw new Error('Payload project identity does not match the requested project.');
}

function projectSiteGroups_(payload) {
  if (Object.prototype.hasOwnProperty.call(payload, 'siteGroups')) return payload.siteGroups;
  if (payload.clientSchemaVersion !== 1 || !Array.isArray(payload.genModels)) throw new Error('Canonical siteGroups are required.');
  return payload.genModels.filter(model => model && model.siteLoc && model.name).map(model => {
    const buildings = [model.name].concat(Array.isArray(model.siteBldgs) ? model.siteBldgs : []);
    return {
      id: model.id || '',
      name: model.customSiteName || buildings.join('-'),
      type: model.groupType || (model.customSiteName ? 'מאוחד' : 'טיפוסי'),
      color: model.groupColor || 'purple',
      buildings: buildings
    };
  });
}

function siteGroupRows_(groups) {
  return groups.map(group => {
    const buildings = group.buildings.map(code => String(code == null ? '' : code).trim()).filter(Boolean);
    return [buildings[0] || '', buildings.join('-'), group.name || '', group.color || 'purple', group.type || (group.name ? 'מאוחד' : 'טיפוסי')];
  }).filter(row => row[0] && row[1].includes('-'));
}

const MODEL_TABLE_HEADERS_ = ['Model Name', 'Project code', 'Discipline', 'Building/Basement', 'Lot', 'Description', 'Model comments'];
const MODEL_TABLE_FIELDS_ = ['code', 'projectCode', 'discCode', 'bldgCode', 'lotCode', 'desc', 'comment'];

function readModelTable_(sheet) {
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length <= 1) return [];
  const headerFields = new Map(MODEL_TABLE_HEADERS_.map((header, i) => [header.trim().toLowerCase(), MODEL_TABLE_FIELDS_[i]]));
  const fields = values[0].map(header => headerFields.get(String(header || '').trim().toLowerCase()) || '');
  return values.slice(1).map(row => {
    const model = {};
    fields.forEach((field, i) => { if (field) model[field] = row[i] == null ? '' : row[i]; });
    if (model.code) model.id = model.code;
    return model;
  }).filter(model => Object.keys(model).length > 0);
}

function modelTableRows_(models) {
  return models.map(model => MODEL_TABLE_FIELDS_.map(field => model[field] == null ? '' : model[field]));
}

function shouldReplaceModelTable_(payload) {
  return payload.modelsDirty === true;
}

function backupProjectBeforeSave_(db) {
  // Fail before any existing sheet is changed if backup cannot be created.
  // Throttled: autosave runs every few seconds, one copy per window is enough.
  const cache = CacheService.getScriptCache();
  const cacheKey = 'BIM_BACKUP:' + db.ss.getId();
  const recent = cache.get(cacheKey);
  if (recent) return recent;
  const folders = db.bimFolder.getFoldersByName(BACKUP_FOLDER_NAME);
  const folder = folders.hasNext() ? folders.next() : db.bimFolder.createFolder(BACKUP_FOLDER_NAME);
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH-mm');
  const url = DriveApp.getFileById(db.ss.getId()).makeCopy('Backup ' + db.project.no + ' ' + stamp, folder).getUrl();
  cache.put(cacheKey, url, Math.min(21600, BACKUP_MIN_INTERVAL_MINUTES * 60));
  return url;
}
