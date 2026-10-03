// ==========================================
// BIMBOT v41.0 - 03_DataService
// Project list (authorized per user), process table, dropdowns
// ==========================================

function getCachedProjectData() {
  try {
    const user = getAuthenticatedUserContext_();
    if (!user.success) return user;

    // Load source project rows by named columns. Projects_TL headers are the first row.
    let projectRows;
    try {
      const mainDb = SpreadsheetApp.openById(MAIN_DB_ID);
      const mainSheet = mainDb.getSheetByName("Projects_TL");
      if (!mainSheet) throw new Error("Projects_TL sheet was not found.");
      projectRows = mainSheet.getDataRange().getValues();
      if (!projectRows.length) throw new Error("Projects_TL is empty.");
    } catch (sourceError) {
      throw new Error("Projects For BimBot source read failed (MAIN_DB_ID): " + sourceError.toString());
    }
    const headers = projectRows[0].map(value => String(value || "").trim().toLowerCase());
    const column = name => headers.indexOf(name.toLowerCase());
    const requiredHeaders = {
      no: "project number",
      name: "project_name",
      teamLeader: "team_leader_name",
      status: "Project_status",
      bimLoad: "bim_load",
      mepLoad: "mep_load",
      department: "department",
      manager: "bim_manager_name",
      coordinator: "mep_coordinator_name"
    };
    const columns = {
      no: column(requiredHeaders.no),
      name: column(requiredHeaders.name),
      teamLeader: column(requiredHeaders.teamLeader),
      status: column(requiredHeaders.status),
      bimLoad: column(requiredHeaders.bimLoad),
      mepLoad: column(requiredHeaders.mepLoad),
      department: column(requiredHeaders.department),
      manager: column(requiredHeaders.manager),
      coordinator: column(requiredHeaders.coordinator)
    };
    const missing = Object.keys(columns).filter(key => columns[key] < 0);
    if (missing.length) throw new Error("Projects_TL is missing required columns: " + missing.join(", "));
    const duplicateHeaders = Object.keys(columns).filter(key => headers.filter(header => header === requiredHeaders[key].toLowerCase()).length > 1);
    if (duplicateHeaders.length) throw new Error("Projects_TL has duplicate required columns: " + duplicateHeaders.join(", "));

    // Check identity ambiguity against the entire source before any status or
    // permission filter can hide a conflicting row from the caller.
    const projectNumberCounts = {};
    for (let i = 1; i < projectRows.length; i++) {
      const projectNo = String(projectRows[i][columns.no] == null ? "" : projectRows[i][columns.no]).trim();
      if (projectNo) projectNumberCounts[projectNo] = (projectNumberCounts[projectNo] || 0) + 1;
    }

    const folderMap = {};
    const folderMapConflicts = {};
    let indexData;
    try {
      const indexSheets = SpreadsheetApp.openById(INDEX_DB_ID).getSheets();
      if (!indexSheets || !indexSheets.length) throw new Error("Folder index has no sheets.");
      indexData = indexSheets[0].getDataRange().getValues();
    } catch (sourceError) {
      throw new Error("Project folder index read failed (INDEX_DB_ID): " + sourceError.toString());
    }
    for (let i = 1; i < indexData.length; i++) {
      const prjNum = String(indexData[i][0] == null ? "" : indexData[i][0]).trim();
      const folderId = String(indexData[i][1] == null ? "" : indexData[i][1]).trim();
      if (!prjNum || !folderId) continue;
      if (Object.prototype.hasOwnProperty.call(folderMap, prjNum) && folderMap[prjNum] !== folderId) {
        folderMapConflicts[prjNum] = true;
        delete folderMap[prjNum];
      } else if (!folderMapConflicts[prjNum]) folderMap[prjNum] = folderId;
    }

    const codeMap = {};
    const codeMapConflicts = {};
    try {
      const codeSheet = SpreadsheetApp.openById(MAIN_DB_ID).getSheetByName("Project_Code");
      if (codeSheet) {
        const codeData = codeSheet.getDataRange().getValues();
        for (let i = 1; i < codeData.length; i++) {
          const prjNum = String(codeData[i][0] == null ? "" : codeData[i][0]).trim();
          const code = String(codeData[i][1] == null ? "" : codeData[i][1]).trim();
          if (!prjNum || !code) continue;
          if (Object.prototype.hasOwnProperty.call(codeMap, prjNum) && codeMap[prjNum] !== code) {
            codeMapConflicts[prjNum] = true;
            delete codeMap[prjNum];
          } else if (!codeMapConflicts[prjNum]) codeMap[prjNum] = code;
        }
      }
    } catch (e) { console.log("Failed to read Project_Code tab: " + e.toString()); }

    const cell = (row, key, fallback) => {
      const value = row[columns[key]];
      return value == null || String(value).trim() === "" ? fallback : String(value).trim();
    };
    const allActiveProjects = projectRows.slice(1).map(row => {
      const no = cell(row, "no", "");
      return {
        no: no,
        name: cell(row, "name", ""),
        projectCode: codeMap[no] || "—",
        teamLeader: cell(row, "teamLeader", "—"),
        status: cell(row, "status", "ACTIVE"),
        department: cell(row, "department", "—"),
        manager: cell(row, "manager", "—"),
        coordinator: cell(row, "coordinator", "—"),
        bimLoad: cell(row, "bimLoad", "—"),
        mepLoad: cell(row, "mepLoad", "—"),
        // Real Drive folder from the project index (INDEX_DB_ID).
        folderId: folderMap[no] || "",
        ambiguousIdentity: Boolean(projectNumberCounts[no] > 1 || folderMapConflicts[no] || codeMapConflicts[no])
      };
    }).filter(project => project.no !== "" && String(project.status).toUpperCase() !== "CANCELLED");

    allActiveProjects.sort((a, b) => (parseInt(a.no.replace(/\D/g, ""), 10) || 0) - (parseInt(b.no.replace(/\D/g, ""), 10) || 0));

    const missionAssignments = {};
    let missionCoverage = "complete";
    const sourceWarnings = [];
    try {
      // Drive permission preflight avoids a deferred Sheets authorization error.
      DriveApp.getFileById(MISSIONS_DB_ID).getName();
      const missionsSheet = SpreadsheetApp.openById(MISSIONS_DB_ID).getSheetByName("BIMBOT_MISSIONS");
      if (!missionsSheet) throw new Error("BIMBOT_MISSIONS sheet was not found.");
      const missionsData = missionsSheet.getDataRange().getValues();
      for (let i = 1; i < missionsData.length; i++) {
        const projectNo = String(missionsData[i][0] || "").trim();
        const employeeName = String(missionsData[i][6] || "").trim();
        if (projectNo && employeeName) {
          if (!missionAssignments[projectNo]) missionAssignments[projectNo] = [];
          if (!missionAssignments[projectNo].includes(employeeName)) missionAssignments[projectNo].push(employeeName);
        }
      }
    } catch (e) {
      missionCoverage = "unknown";
      sourceWarnings.push({source: "missions", status: "unavailable",
        message: "Missions assignments could not be read; mission-based project matches are unknown."});
      console.log("Error reading Missions DB: " + e.toString());
    }

    const cleanUserName = user.userName.toLowerCase().trim();
    allActiveProjects.forEach(project => {
      project.missionEmployees = missionCoverage === "complete" ? (missionAssignments[project.no] || []) : null;
      project.isUserMission = missionCoverage === "complete" ?
        cleanUserName !== "" && project.missionEmployees.some(name => name.toLowerCase().trim() === cleanUserName) : null;
    });

    const projectsToReturn = user.canViewAllProjects ? allActiveProjects : allActiveProjects.filter(project => {
      if (!cleanUserName) return false;
      return String(project.manager).toLowerCase().trim() === cleanUserName ||
        String(project.coordinator).toLowerCase().trim() === cleanUserName ||
        String(project.teamLeader).toLowerCase().trim() === cleanUserName ||
        project.isUserMission;
    });

    // Existing clients use allProjects for document/project pickers. It must remain
    // the same authorized set as data; returning the unfiltered list leaks records.
    return {
      success: true,
      data: projectsToReturn,
      allProjects: projectsToReturn,
      canViewAllProjects: user.canViewAllProjects,
      canViewDatabase: user.canViewDatabase,
      missionCoverage: missionCoverage,
      sourceWarnings: sourceWarnings
    };
  } catch (error) { return { success: false, error: error.toString() }; }
}

// The remaining readers are carried forward unchanged from the captured source.
function getProcessTableData() {
  try {
    const data = SpreadsheetApp.openById(PROCESS_DB_ID).getSheetByName("Process Table").getDataRange().getValues();
    data.shift();
    const processes = data.map(row => ({ stage: row[0] ? String(row[0]).trim() : "", process: row[1] ? String(row[1]).trim() : "", fileName: row[2] ? String(row[2]).trim() : "", type: row[3] ? String(row[3]).trim() : "טמפלייט", fileId: row[4] ? String(row[4]).trim() : "", discipline: row[6] ? String(row[6]).trim() : "All" })).filter(r => r.fileId !== "");
    return { success: true, data: processes };
  } catch(e) { return { success: false, error: e.toString() }; }
}

function getDropdownData() {
  try {
    const ss = SpreadsheetApp.openById(DROPDOWN_DB_ID);
    let result = { revitVersions: [], cdePlatforms: [], bimPlatforms: [], edmsPlatforms: [], allDisciplines: [], liveLink: "", sharedLink: "" };

    const infoSheet = ss.getSheetByName("BIM_Info_General");
    if (infoSheet) {
        const infoData = infoSheet.getDataRange().getDisplayValues();
        
        if (infoData.length > 0) {
            let rawRevit = infoData[0].slice(1).filter(v => v.trim() !== "");
            rawRevit.forEach(val => {
                if (val.includes("-")) {
                    let parts = val.split("-");
                    let start = parseInt(parts[0].trim());
                    let end = parseInt(parts[1].trim());
                    if (!isNaN(start) && !isNaN(end) && start <= end) {
                        for (let y = start; y <= end; y++) {
                            result.revitVersions.push(y.toString());
                        }
                    } else { result.revitVersions.push(val); }
                } else { result.revitVersions.push(val); }
            });
        }

        if (infoData.length > 1) result.cdePlatforms = infoData[1].slice(1).filter(v => v.trim() !== "");
        if (infoData.length > 2) result.bimPlatforms = infoData[2].slice(1).filter(v => v.trim() !== "");
    }

    const edmsSheet = ss.getSheetByName("EDMS_platforms");
    if (edmsSheet) {
        const edmsData = edmsSheet.getDataRange().getDisplayValues();
        for (let i = 1; i < edmsData.length; i++) {
            if (edmsData[i][0]) {
                result.edmsPlatforms.push({
                    name: edmsData[i][0].toString().trim(),
                    link: (edmsData[i][1] || "").toString().trim()
                });
            }
        }
    }

    const lsSheet = ss.getSheetByName("Live_Shared");
    if (lsSheet) {
        const lsData = lsSheet.getDataRange().getDisplayValues();
        for (let i = 1; i < lsData.length; i++) {
            let key = (lsData[i][0] || "").toString().toLowerCase().trim();
            if (key === "live") result.liveLink = (lsData[i][1] || "").toString().trim();
            if (key === "shared") result.sharedLink = (lsData[i][1] || "").toString().trim();
        }
    }

    const discSheet = ss.getSheetByName("Disipline");
    if (discSheet) {
        const discData = discSheet.getDataRange().getDisplayValues();
        for (let i = 1; i < discData.length; i++) {
            let heName = discData[i][0] ? discData[i][0].toString().trim() : ""; 
            // LOCKING TO EXACT COLUMNS REQUESTED:
            let acCode = discData[i][2] ? discData[i][2].toString().trim() : ""; // Column C (Index 2)
            let isDefault = discData[i][3] ? (discData[i][3].toString().trim().toUpperCase() === 'V') : false; 
            let platformVal = discData[i][4] ? discData[i][4].toString().trim() : "Revit"; 
            let scopeVal = discData[i][6] ? discData[i][6].toString().trim() : "Building"; // Column G (Index 6)
            
            if (heName) {
                result.allDisciplines.push({ 
                    he: heName, 
                    ac: acCode, 
                    isDefault: isDefault, 
                    platform: platformVal, 
                    scope: scopeVal 
                });
            }
        }
    }

    const cleanArray = (arr) => [...new Set(arr)].filter(val => val !== "");
    result.revitVersions = cleanArray(result.revitVersions);
    result.cdePlatforms = cleanArray(result.cdePlatforms);
    result.bimPlatforms = cleanArray(result.bimPlatforms);

    return { success: true, data: result };
  } catch(e) { return { success: false, error: e.toString() }; }
}
