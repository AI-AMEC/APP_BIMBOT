// ==========================================
// BIM Commander v40.0 - 05_DocumentService
// File and Document Generation
// ==========================================

function makeVersionedProjectOutputName(baseName) {
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmmss");
  const suffix = Utilities.getUuid().replace(/-/g, "").substring(0, 8);
  return baseName + " - " + stamp + "-" + suffix;
}

// Build GL rows from the same canonical AR model records used by the main model table.
// Physical building GL values are only shown when the model-to-building mapping is clear.
function buildKickoffGlEntries_(models, siteGroups, buildingRows) {
  const buildings = (buildingRows || []).slice(1).map(row => ({
    code: String(row[0] || "").trim(),
    gl: String(row[5] == null ? "" : row[5]).trim()
  })).filter(row => row.code);
  const buildingsByCode = new Map();
  buildings.forEach(row => {
    const key = row.code.toLowerCase();
    if (!buildingsByCode.has(key)) buildingsByCode.set(key, []);
    buildingsByCode.get(key).push(row);
  });

  return (models || []).filter(model => String(model.discCode || "").trim().toUpperCase() === "AR").map(model => {
    const modelBuilding = String(model.bldgCode || "").trim();
    const groups = (siteGroups || []).filter(item =>
      String(item.name || "").trim().toLowerCase() === modelBuilding.toLowerCase() ||
      (item.buildings || []).join("-").toLowerCase() === modelBuilding.toLowerCase()
    );
    let gl;
    if (groups.length > 1) {
      gl = "עמימות: כמה קבוצות למודל";
    } else {
      const directMatches = buildingsByCode.get(modelBuilding.toLowerCase()) || [];
      if (!groups.length && directMatches.length > 1) {
        return { code: String(model.code || ""), gl: "עמימות: קוד מבנה כפול " + modelBuilding };
      }
      const physicalCodes = groups.length ? groups[0].buildings : (directMatches.length === 1 ? [modelBuilding] : []);
      const mapped = physicalCodes.map(code => buildingsByCode.get(String(code).toLowerCase()) || []);
      if (!physicalCodes.length || mapped.some(matches => matches.length === 0)) {
        gl = "לא ידוע";
      } else if (mapped.some(matches => matches.length > 1)) {
        gl = "עמימות: קוד מבנה כפול";
      } else {
        const physicalRows = mapped.map(matches => matches[0]);
      const values = physicalRows.map(row => {
        const value = row.gl;
        if (/^(true|yes|v)$/i.test(value)) return "GL";
        if (/^(false|no|)$/i.test(value)) return "לא צוין";
        return value;
      });
      const distinct = [...new Set(values)];
        gl = distinct.length === 1 ? distinct[0] : "סתירה: " + physicalCodes.map((code, index) => code + "=" + values[index]).join(", ");
      }
    }
    return { code: String(model.code || ""), gl: gl };
  });
}

function generateKickoffDoc(folderId, projectNo, projectName, payload) {
  return generateKickoffDocWithResolvedDb_(folderId,projectNo,projectName,payload,null);
}
function generateKickoffDocWithResolvedDb_(folderId, projectNo, projectName, payload, resolvedDb) {
  let generatedFile = null;
  try {
    // Stage services may pass the already-resolved project DB to avoid nested ScriptLock acquisition.
    const db = resolvedDb || getLocalProjectDbAndBimFolder_(folderId, projectNo, projectName);
    projectNo = db.project.no;
    projectName = db.project.name;
    const bimFolder = db.bimFolder;
    const ss = db.ss; 

    let kickFolder;
    const subFolders = bimFolder.getFoldersByName("00. Kick-Off");
    if (subFolders.hasNext()) { kickFolder = subFolders.next(); } 
    else { kickFolder = bimFolder.createFolder("00. Kick-Off"); }

    const newFileName = makeVersionedProjectOutputName("Kick-Off Document - " + projectNo + " - " + projectName);

    const templateFile = DriveApp.getFileById(KICKOFF_DOC_TEMPLATE_ID);
    generatedFile = templateFile.makeCopy(newFileName, kickFolder);
    const newFile = generatedFile;
    
    const doc = DocumentApp.openById(newFile.getId());
    const body = doc.getBody();
    
    const d = new Date();
    const todayStr = Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy");
    const dPlus7 = new Date(d);
    dPlus7.setDate(dPlus7.getDate() + 7);
    const dPlus7Str = Utilities.formatDate(dPlus7, Session.getScriptTimeZone(), "dd/MM/yyyy");

    const header = doc.getHeader();
    if (header) {
      header.replaceText("<<PROJECT_NAME>>", projectName || "");
      header.replaceText("<<PROJECT_CODE>>", payload.bimInfo.code || projectNo);
      header.replaceText("<<DATE>>", todayStr);
      header.replaceText("\\[תאריך יציאת הקובץ\\]", todayStr); 
    }

    const footer = doc.getFooter();
    if (footer) {
      footer.replaceText("<<PROJECT_NAME>>", projectName || "");
      footer.replaceText("<<PROJECT_CODE>>", payload.bimInfo.code || projectNo);
      footer.replaceText("<<DATE>>", todayStr);
    }

    body.replaceText("<<PROJECT_NAME>>", projectName || "");
    body.replaceText("<<PROJECT_CODE>>", payload.bimInfo.code || projectNo);
    body.replaceText("<<DATE>>", todayStr);
    body.replaceText("<<DATE_PLUS_7>>", dPlus7Str);
    
    let glLevelValue = payload.kickoff.glLevel !== undefined && payload.kickoff.glLevel !== "" ? payload.kickoff.glLevel : "0.00";
    body.replaceText("<<GL_LEVEL_TEXT>>", "0.00=" + glLevelValue);

    function localRemoveBlock(docBody, startTag, endTag) {
        let numChildren = docBody.getNumChildren();
        let inBlock = false;
        let toDelete = [];
        for (let i = 0; i < numChildren; i++) {
            let child = docBody.getChild(i);
            let text = "";
            if (child.getType() === DocumentApp.ElementType.PARAGRAPH || child.getType() === DocumentApp.ElementType.LIST_ITEM) {
                text = child.getText().trim();
            } else if (child.getType() === DocumentApp.ElementType.TABLE) {
                text = child.editAsText().getText().trim();
            }
            
            if (text.indexOf(startTag) > -1) {
                inBlock = true;
                toDelete.push(child);
                continue;
            }
            if (text.indexOf(endTag) > -1) {
                inBlock = false;
                toDelete.push(child);
                continue;
            }
            if (inBlock) {
                toDelete.push(child);
            }
        }
        for (let i = toDelete.length - 1; i >= 0; i--) {
            try { docBody.removeChild(toDelete[i]); } catch(e) {}
        }
    }

    function localRemoveTags(docBody, startTag, endTag) {
        let numChildren = docBody.getNumChildren();
        let toDelete = [];
        for (let i = 0; i < numChildren; i++) {
            let child = docBody.getChild(i);
            if (child.getType() === DocumentApp.ElementType.PARAGRAPH || child.getType() === DocumentApp.ElementType.LIST_ITEM) {
                let text = child.getText().trim();
                if (text === startTag || text === endTag) {
                    toDelete.push(child);
                }
            }
        }
        for (let i = toDelete.length - 1; i >= 0; i--) {
            try { docBody.removeChild(toDelete[i]); } catch(e) {}
        }
        try { docBody.replaceText(startTag, ""); } catch(e){}
        try { docBody.replaceText(endTag, ""); } catch(e){}
    }

    if (payload.kickoff.grids === "Yes") {
       localRemoveBlock(body, "<<GRIDS_NO_START>>", "<<GRIDS_NO_END>>");
       localRemoveTags(body, "<<GRIDS_YES_START>>", "<<GRIDS_YES_END>>");
    } else if (payload.kickoff.grids === "No") {
       localRemoveBlock(body, "<<GRIDS_YES_START>>", "<<GRIDS_YES_END>>");
       localRemoveTags(body, "<<GRIDS_NO_START>>", "<<GRIDS_NO_END>>");
       let xVal = payload.kickoff.gridsX ? payload.kickoff.gridsX : "X";
       let yVal = payload.kickoff.gridsY ? payload.kickoff.gridsY : "Y";
       body.replaceText("<<GRID_X>>", xVal);
       body.replaceText("<<GRID_Y>>", yVal);
    } else {
       localRemoveBlock(body, "<<GRIDS_NO_START>>", "<<GRIDS_NO_END>>");
       localRemoveBlock(body, "<<GRIDS_YES_START>>", "<<GRIDS_YES_END>>");
    }

    if (payload.kickoff.arLevel === "Yes") {
       localRemoveBlock(body, "<<AR_LEVEL_NO_START>>", "<<AR_LEVEL_NO_END>>");
       localRemoveTags(body, "<<AR_LEVEL_YES_START>>", "<<AR_LEVEL_YES_END>>");
    } else if (payload.kickoff.arLevel === "No") {
       localRemoveBlock(body, "<<AR_LEVEL_YES_START>>", "<<AR_LEVEL_YES_END>>");
       localRemoveTags(body, "<<AR_LEVEL_NO_START>>", "<<AR_LEVEL_NO_END>>");
    } else {
       localRemoveBlock(body, "<<AR_LEVEL_NO_START>>", "<<AR_LEVEL_NO_END>>");
       localRemoveBlock(body, "<<AR_LEVEL_YES_START>>", "<<AR_LEVEL_YES_END>>");
    }

    if (payload.kickoff.urs === "Yes") {
       localRemoveBlock(body, "<<URS_NO_START>>", "<<URS_NO_END>>");
       localRemoveTags(body, "<<URS_YES_START>>", "<<URS_YES_END>>");
    } else if (payload.kickoff.urs === "No") {
       localRemoveBlock(body, "<<URS_YES_START>>", "<<URS_YES_END>>");
       localRemoveTags(body, "<<URS_NO_START>>", "<<URS_NO_END>>");
    } else {
       localRemoveBlock(body, "<<URS_NO_START>>", "<<URS_NO_END>>");
       localRemoveBlock(body, "<<URS_YES_START>>", "<<URS_YES_END>>");
    }

    let sheetContacts = [];
    let sContacts = ss.getSheets().find(s => s.getName().toLowerCase() === "contacts");
    
    if (sContacts) {
        let cData = sContacts.getDataRange().getDisplayValues();
        for (let k = 1; k < cData.length; k++) {
            let cName = String(cData[k][0] || "").trim();
            let cEmail = String(cData[k][1] || "").trim();
            let cDisc = String(cData[k][2] || "").trim();
            let cRole = String(cData[k][3] || "").trim();
            let cComp = String(cData[k][4] || "").trim();
            
            if (!cRole && cEmail && !cEmail.includes("@")) { cRole = cEmail; }
            if (cName || cComp || cRole) { sheetContacts.push({ name: cName, email: cEmail, disc: cDisc, role: cRole, company: cComp }); }
        }
    } else if (payload.contacts && payload.contacts.length > 0) {
        sheetContacts = payload.contacts;
    }

    const targetRoles = ["יזם", "אדריכל", "מנהל פרויקט"];
    let linesToInsert = [];

    targetRoles.forEach(targetRole => {
        let match = sheetContacts.find(c => {
            let r = String(c.role || "").trim();
            if (targetRole === "מנהל פרויקט") { return r.includes("מנהל פרויקט") || r.includes("מנהל פרוייקט") || r.includes("מנהל פרוייקטים") || r.includes("מנהל פרויקטים"); }
            return r.includes(targetRole);
        });

        if (match) {
            let lineParts = [targetRole];
            if (match.company && match.company.trim()) { lineParts.push(match.company.trim()); } 
            else if (match.name && match.name.trim()) { lineParts.push(match.name.trim()); }
            linesToInsert.push(lineParts.join(" - "));
        }
    });

    let contactsReplaced = false;

    const docTables = body.getTables();
    for (let t = docTables.length - 1; t >= 0; t--) {
        let tbl = docTables[t];
        let tblText = tbl.editAsText().getText();
        if (tblText.indexOf("CONTACT") > -1 || (tblText.indexOf("Company") > -1 && tblText.indexOf("Role") > -1)) {
            let parent = tbl.getParent();
            let childIdx = parent.getChildIndex(tbl);
            tbl.removeFromParent();

            for (let i = 0; i < linesToInsert.length; i++) {
                body.insertParagraph(childIdx + i, linesToInsert[i]);
            }
            contactsReplaced = true;
            break;
        }
    }

    if (!contactsReplaced) {
        let pars = body.getParagraphs();
        for (let i = 0; i < pars.length; i++) {
            let pText = pars[i].getText();
            if (pText.indexOf("<<CONTACT_") > -1 || pText.indexOf("<<CONTACTS_LIST>>") > -1) {
                let parent = pars[i].getParent();
                let childIdx = parent.getChildIndex(pars[i]);
                pars[i].removeFromParent();

                for (let j = 0; j < linesToInsert.length; j++) {
                    body.insertParagraph(childIdx + j, linesToInsert[j]);
                }
                break;
            }
        }
    }

    const tables = body.getTables();
    for (let t = 0; t < tables.length; t++) {
        let table = tables[t];

        // MAIN MODEL TABLE
        let templateRowIndexDesc = -1;
        for (let r = 0; r < table.getNumRows(); r++) {
            if (table.getRow(r).editAsText().getText().indexOf("<<MODEL_DESC>>") > -1) { 
                templateRowIndexDesc = r; break; 
            }
        }
        if (templateRowIndexDesc > -1) {
            let templateRow = table.getRow(templateRowIndexDesc);
            if (payload.models && payload.models.length > 0) {
                for (let i = payload.models.length - 1; i >= 0; i--) {
                    let model = payload.models[i];
                    let newRow = table.insertTableRow(templateRowIndexDesc + 1, templateRow.copy());
                    newRow.replaceText("<<MODEL_NAME>>", model.code || "");
                    newRow.replaceText("<<MODEL_DESC>>", model.desc || "");
                    newRow.replaceText("<<MODEL_DISC>>", model.discCode || "");
                    newRow.replaceText("<<MODEL_PROJ_CODE>>", model.projectCode || "");
                }
                table.removeRow(templateRowIndexDesc);
            } else {
                table.removeFromParent();
            }
            continue;
        }
        
        // GL LEVEL TABLE
        let templateRowIndexGl = -1;
        for (let r = 0; r < table.getNumRows(); r++) {
            if (table.getRow(r).editAsText().getText().indexOf("<<MODEL_GL>>") > -1) { 
                templateRowIndexGl = r; break; 
            }
        }
        
        if (templateRowIndexGl > -1) {
            let templateRow = table.getRow(templateRowIndexGl);
            let glEntries = [];
            let sBldgs = ss.getSheets().find(s => s.getName().toLowerCase() === "buildings");
            const buildingRows = sBldgs ? sBldgs.getDataRange().getDisplayValues() : [];
            glEntries = buildKickoffGlEntries_(payload.models, payload.siteGroups, buildingRows);

            if (glEntries.length > 0) {
                for (let i = glEntries.length - 1; i >= 0; i--) {
                    let entry = glEntries[i];
                    let newRow = table.insertTableRow(templateRowIndexGl + 1, templateRow.copy());
                    newRow.replaceText("<<MODEL_NAME>>", entry.code);
                    newRow.replaceText("<<MODEL_GL>>", entry.gl);
                }
                table.removeRow(templateRowIndexGl);
            } else {
                table.removeFromParent();
            }
        }
    }

    let sSites = ss.getSheets().find(s => s.getName().toLowerCase() === "revit_sites");
    let siteEntries = [];
    if (sSites) {
        let siteData = sSites.getDataRange().getDisplayValues();
        for (let k = 1; k < siteData.length; k++) {
            let bCode = String(siteData[k][0] || "").trim();
            let gModel = String(siteData[k][1] || "").trim();
            let cName = String(siteData[k][2] || "").trim();
            if (bCode || gModel || cName) { siteEntries.push({ bldg: bCode, generic: gModel, custom: cName }); }
        }
    }

    if (siteEntries.length === 0) {
        localRemoveBlock(body, "<<SITES_START>>", "<<SITES_END>>");
    } else {
        let pars = body.getParagraphs();
        let templatePar = null;
        let templateIndex = -1;
        
        for (let i = 0; i < pars.length; i++) {
            if (pars[i].getText().indexOf("<<SITE_BLDG>>") > -1) {
                templatePar = pars[i]; templateIndex = body.getChildIndex(templatePar); break;
            }
        }

        if (templatePar && templateIndex > -1) {
            for (let i = siteEntries.length - 1; i >= 0; i--) {
                let entry = siteEntries[i];
                let newPar = templatePar.copy();
                newPar.replaceText("<<SITE_BLDG>>", entry.bldg);
                newPar.replaceText("<<SITE_GENERIC>>", entry.generic);
                newPar.replaceText("<<SITE_CUSTOM>>", entry.custom);
                body.insertParagraph(templateIndex + 1, newPar);
            }
            templatePar.removeFromParent();
        }
        localRemoveTags(body, "<<SITES_START>>", "<<SITES_END>>");
    }

    let sOptions = ss.getSheets().find(s => s.getName().toLowerCase() === "design_option");
    let optionEntries = [];
    if (sOptions) {
        let optData = sOptions.getDataRange().getDisplayValues();
        for (let k = 1; k < optData.length; k++) {
            let bldg = String(optData[k][0] || "").trim(); 
            let name = String(optData[k][2] || "").trim(); 
            let desc = String(optData[k][3] || "").trim(); 
            if (bldg && name) { optionEntries.push({ bldg: bldg, name: name, desc: desc }); }
        }
    }

    if (optionEntries.length === 0) {
        localRemoveBlock(body, "<<DESOPT_START>>", "<<DESOPT_END>>");
    } else {
        let pars = body.getParagraphs();
        let templatePar = null;
        let templateIndex = -1;
        for (let i = 0; i < pars.length; i++) {
            if (pars[i].getText().indexOf("<<OPT_BLDG>>") > -1) {
                templatePar = pars[i]; templateIndex = body.getChildIndex(templatePar); break;
            }
        }

        if (templatePar && templateIndex > -1) {
            for (let i = optionEntries.length - 1; i >= 0; i--) {
                let entry = optionEntries[i];
                let newPar = templatePar.copy();
                newPar.replaceText("<<OPT_BLDG>>", entry.bldg);
                newPar.replaceText("<<OPT_NAME>>", entry.name);
                newPar.replaceText("<<OPT_DESC>>", entry.desc);
                body.insertParagraph(templateIndex + 1, newPar);
            }
            templatePar.removeFromParent();
        }
        localRemoveTags(body, "<<DESOPT_START>>", "<<DESOPT_END>>");
    }

    // THE SMART VACUUM: Removes excessive blank spaces without breaking pages
    function localCleanSpaces(docBody) {
        let pars = docBody.getParagraphs();
        let emptyCount = 0;
        for (let i = 0; i < pars.length; i++) {
            let p = pars[i];
            let text = p.getText().trim();
            
            let keep = false;
            let numChildren = p.getNumChildren();
            for(let j=0; j<numChildren; j++) {
                let type = p.getChild(j).getType();
                if(type === DocumentApp.ElementType.PAGE_BREAK || type === DocumentApp.ElementType.INLINE_IMAGE || type === DocumentApp.ElementType.HORIZONTAL_RULE) {
                    keep = true;
                    break;
                }
            }

            if (text === "" && !keep) {
                emptyCount++;
                if (emptyCount > 1) { 
                    try { p.removeFromParent(); } catch(e) {}
                }
            } else {
                emptyCount = 0;
            }
        }
    }
    localCleanSpaces(body);

    doc.saveAndClose();
    return { success: true, url: newFile.getUrl() };

  } catch(e) { return { success: false, error: e.toString(), orphanUrl: generatedFile ? generatedFile.getUrl() : "" }; }
}

function generatePresentationDoc(folderId, projectNo, projectName, payload) {
    try {
        const db = getLocalProjectDbAndBimFolder_(folderId, projectNo, projectName);
        projectNo = db.project.no;
        projectName = db.project.name;
        const bimFolder = db.bimFolder;
        const ss = db.ss; 

        let sSites = ss.getSheets().find(s => s.getName().toLowerCase() === "revit_sites");
        let siteEntries = [];
        if (sSites) {
            let siteData = sSites.getDataRange().getDisplayValues();
            for (let k = 1; k < siteData.length; k++) {
                let bCode = String(siteData[k][0] || "").trim();
                let gModel = String(siteData[k][1] || "").trim();
                let cName = String(siteData[k][2] || "").trim();
                if (bCode || gModel || cName) { siteEntries.push({ bldg: bCode, generic: gModel, custom: cName }); }
            }
        }

        let sOptions = ss.getSheets().find(s => s.getName().toLowerCase() === "design_option");
        let optionEntries = [];
        if (sOptions) {
            let optData = sOptions.getDataRange().getDisplayValues();
            for (let k = 1; k < optData.length; k++) {
                let bldg = String(optData[k][0] || "").trim(); 
                let name = String(optData[k][2] || "").trim(); 
                let desc = String(optData[k][3] || "").trim(); 
                if (bldg && name) { optionEntries.push({ bldg: bldg, name: name, desc: desc }); }
            }
        }

        let kickFolder;
        const subFolders = bimFolder.getFoldersByName("00. Kick-Off");
        if (subFolders.hasNext()) { kickFolder = subFolders.next(); } 
        else { kickFolder = bimFolder.createFolder("00. Kick-Off"); }

        const newFileName = makeVersionedProjectOutputName("BIM Presentation - " + projectNo + " - " + projectName);

        const templateFile = DriveApp.getFileById(PRESENTATION_TEMPLATE_ID);
        const newFile = templateFile.makeCopy(newFileName, kickFolder);
        
        const pres = SlidesApp.openById(newFile.getId());
        
        const d = new Date();
        const todayStr = Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy");
        
        const dPlus30 = new Date(d);
        dPlus30.setDate(dPlus30.getDate() + 30);
        const dPlus30Str = Utilities.formatDate(dPlus30, Session.getScriptTimeZone(), "dd/MM/yyyy");

        pres.replaceAllText("<<PROJECT_NAME>>", projectName || "");
        pres.replaceAllText("<<PROJECT_CODE>>", payload.bimInfo.code || projectNo);
        pres.replaceAllText("<<project_code>>", payload.bimInfo.code || projectNo); 
        
        pres.replaceAllText("<<DATE>>", todayStr);
        pres.replaceAllText("<<DATE_PLUS_30>>", dPlus30Str);
        
        pres.replaceAllText("<<Revit_Version>>", payload.bimInfo.revit || "");
        pres.replaceAllText("<<REVIT_VERSION>>", payload.bimInfo.revit || ""); 
        pres.replaceAllText("<<EDMS_name>>", payload.bimInfo.docManagement || "");
        pres.replaceAllText("<<DOC_MANAGEMENT>>", payload.bimInfo.docManagement || ""); 
        pres.replaceAllText("<<HUB>>", payload.bimInfo.hub || "");

        pres.replaceAllText("<<BIM_MANAGER>>", payload.bimInfo.manager || "");
        pres.replaceAllText("<<BIM_LEADER>>", payload.bimInfo.teamLeader || "");

        let slides = pres.getSlides();
        let slidesToDelete = [];

        for (let i = 0; i < slides.length; i++) {
            let slide = slides[i];
            let shapes = slide.getShapes();
            let deleteSlide = false;

            for (let j = 0; j < shapes.length; j++) {
                let shape = shapes[j];
                let textObj;
                try { textObj = shape.getText(); } catch(e) { continue; }
                let textStr = textObj.asString();
                
                let shapeReplaced = false;

                if (textStr.includes("<<DATE_PLUS_30>>")) {
                    textObj.replaceAllText("<<DATE_PLUS_30>>", dPlus30Str);
                }

                if (textStr.includes("<<project_management_logo>>")) {
                    if (payload.pmLogoBase64) {
                        try {
                            let parts = payload.pmLogoBase64.split(';base64,');
                            let blob = Utilities.newBlob(Utilities.base64Decode(parts[1]), parts[0].split(':')[1], "pm_logo");
                            shape.replaceWithImage(blob);
                            shapeReplaced = true;
                        } catch(e) {}
                    } else {
                        textObj.replaceAllText("<<project_management_logo>>", "");
                    }
                }

                if (!shapeReplaced && textStr.includes("<<developer_logo>>")) {
                    if (payload.devLogoBase64) {
                        try {
                            let parts = payload.devLogoBase64.split(';base64,');
                            let blob = Utilities.newBlob(Utilities.base64Decode(parts[1]), parts[0].split(':')[1], "dev_logo");
                            shape.replaceWithImage(blob);
                            shapeReplaced = true;
                        } catch(e) {}
                    } else {
                        textObj.replaceAllText("<<developer_logo>>", "");
                    }
                }
                
                if (shapeReplaced) continue;

                if (textStr.includes("<<SITE_BLDG>>") || textStr.includes("<<DELETE_IF_NO_SITES>>")) {
                    if (siteEntries.length === 0) {
                        deleteSlide = true;
                    } else {
                        if (textStr.includes("<<SITE_BLDG>>")) {
                            let templateLine = textStr.replace("<<DELETE_IF_NO_SITES>>", "").trim();
                            let finalBlock = siteEntries.map(e => {
                                return templateLine
                                    .replace("<<SITE_BLDG>>", e.bldg)
                                    .replace("<<SITE_GENERIC>>", e.generic)
                                    .replace("<<SITE_CUSTOM>>", e.custom);
                            }).join('\n\n');
                            textObj.setText(finalBlock);
                        } else {
                            textObj.replaceAllText("<<DELETE_IF_NO_SITES>>", "");
                        }
                    }
                }

                if (textStr.includes("<<OPT_BLDG>>") || textStr.includes("<<DELETE_IF_NO_OPTIONS>>")) {
                    if (optionEntries.length === 0) {
                        deleteSlide = true;
                    } else {
                        if (textStr.includes("<<OPT_BLDG>>")) {
                            let templateLine = textStr.replace("<<DELETE_IF_NO_OPTIONS>>", "").trim();
                            let finalBlock = optionEntries.map(e => {
                                return templateLine
                                    .replace("<<OPT_BLDG>>", e.bldg)
                                    .replace("<<OPT_NAME>>", e.name)
                                    .replace("<<OPT_DESC>>", e.desc);
                            }).join('\n\n');
                            textObj.setText(finalBlock);
                        } else {
                            textObj.replaceAllText("<<DELETE_IF_NO_OPTIONS>>", "");
                        }
                    }
                }

                if (textStr.includes("<<DELETE_IF_NO_PHASES>>")) {
                    if (payload.kickoff.hasPhases !== "Yes") deleteSlide = true;
                    else textObj.replaceAllText("<<DELETE_IF_NO_PHASES>>", "");
                }
            }
            if (deleteSlide) slidesToDelete.push(slide);
        }

        slidesToDelete.forEach(s => s.remove());
        slides = pres.getSlides(); 

        for (let i = 0; i < slides.length; i++) {
            let tables = slides[i].getTables();
            for (let t = 0; t < tables.length; t++) {
                let table = tables[t];
                let targetRow = -1;

                let numCols = table.getNumColumns();
                let numRows = table.getNumRows();
                for (let r = 0; r < numRows; r++) {
                    let rowText = "";
                    for(let c=0; c < numCols; c++){
                        try { rowText += table.getCell(r, c).getText().asString(); } catch(e) {}
                    }
                    if (rowText.includes("<<DISC_LICENSE>>") || rowText.includes("<<DISC_NAME>>")) {
                        targetRow = r; break;
                    }
                }

                if (targetRow > -1) {
                    let activeDiscs = (payload.disciplines || []).filter(d => d.checked !== false);
                    
                    if (activeDiscs.length > 0) {
                        for (let d = activeDiscs.length - 1; d >= 0; d--) {
                            let disc = activeDiscs[d];
                            let discName = disc.he || disc.ac;
                            let platform = disc.platform || "Revit";
                            let license = "";
                            
                            if (discName === "מנהל הפרויקט" || discName === "צפיין") {
                                discName = "צפיין";
                                license = "Forma Data Management";
                            } else if (platform.toLowerCase().includes("civil")) {
                                license = "Forma Design Collaboration+Civil 3D";
                            } else if (platform.toLowerCase().includes("revit")) {
                                license = "Forma Design Collaboration+Revit";
                            } else {
                                license = "Forma Design Collaboration+" + platform;
                            }

                            table.insertRow(targetRow + 1);
                            
                            let c0 = "", c1 = "", c2 = "";
                            try { c0 = table.getCell(targetRow, 0).getText().asString(); } catch(e){}
                            try { c1 = table.getCell(targetRow, 1).getText().asString(); } catch(e){}
                            try { c2 = table.getCell(targetRow, 2).getText().asString(); } catch(e){}
                            
                            try { table.getCell(targetRow + 1, 0).getText().setText(c0.includes("<<DISC_LICENSE>>") ? license : (c0.includes("<<DISC_PLATFORM>>") ? platform : discName)); } catch(e){}
                            try { table.getCell(targetRow + 1, 1).getText().setText(c1.includes("<<DISC_PLATFORM>>") ? platform : (c1.includes("<<DISC_NAME>>") ? discName : license)); } catch(e){}
                            try { table.getCell(targetRow + 1, 2).getText().setText(c2.includes("<<DISC_NAME>>") ? discName : (c2.includes("<<DISC_LICENSE>>") ? license : platform)); } catch(e){}
                        }
                        table.getRow(targetRow).remove();
                    } else {
                        table.remove();
                    }
                }
            }
        }

        const MAX_ROWS_PER_SLIDE = 10;
        for (let i = 0; i < slides.length; i++) {
            let tables = slides[i].getTables();
            for (let t = 0; t < tables.length; t++) {
                let table = tables[t];
                let targetRow = -1;

                let numCols = table.getNumColumns();
                let numRows = table.getNumRows();
                for (let r = 0; r < numRows; r++) {
                    let rowText = "";
                    for(let c=0; c < numCols; c++){
                        try { rowText += table.getCell(r, c).getText().asString(); } catch(e) {}
                    }
                    if (rowText.includes("<<MODEL_NAME>>")) {
                        targetRow = r; break;
                    }
                }

                if (targetRow > -1) {
                    let models = payload.models || [];
                    if (models.length === 0) {
                        table.remove();
                        continue;
                    }

                    let chunkedModels = [];
                    for (let m = 0; m < models.length; m += MAX_ROWS_PER_SLIDE) {
                        chunkedModels.push(models.slice(m, m + MAX_ROWS_PER_SLIDE));
                    }

                    let originalSlide = slides[i];
                    let slidesToPopulate = [originalSlide];

                    for (let c = 1; c < chunkedModels.length; c++) {
                        slidesToPopulate.push(originalSlide.duplicate());
                    }

                    for (let chunkIdx = 0; chunkIdx < chunkedModels.length; chunkIdx++) {
                        let chunk = chunkedModels[chunkIdx];
                        let workingSlide = slidesToPopulate[chunkIdx];
                        
                        let wTable;
                        let wTables = workingSlide.getTables();
                        for (let wt=0; wt<wTables.length; wt++) {
                            let textCheck = "";
                            let wCols = wTables[wt].getNumColumns();
                            for(let c=0; c < wCols; c++) {
                                try { textCheck += wTables[wt].getCell(targetRow, c).getText().asString(); } catch(e){}
                            }
                            if (textCheck.includes("<<MODEL_NAME>>")) {
                                wTable = wTables[wt]; break;
                            }
                        }

                        if (wTable) {
                            for (let d = chunk.length - 1; d >= 0; d--) {
                                let mod = chunk[d];
                                wTable.insertRow(targetRow + 1);
                                
                                let wCols = wTable.getNumColumns();
                                for (let c = 0; c < wCols; c++) {
                                    try {
                                        let templateText = wTable.getCell(targetRow, c).getText().asString();
                                        let val = "";
                                        
                                        if (templateText.includes("<<MODEL_NAME>>")) val = mod.code || "";
                                        else if (templateText.includes("<<MODEL_PROJ_CODE>>")) val = mod.projectCode || "";
                                        else if (templateText.includes("<<MODEL_DISC>>")) val = mod.discCode || "";
                                        else if (templateText.includes("<<MODEL_BLDG>>")) val = mod.bldgCode || "";
                                        else if (templateText.includes("<<MODEL_LOT>>")) val = mod.lotCode || "";
                                        else if (templateText.includes("<<MODEL_DESC>>")) val = mod.desc || "";
                                        else if (templateText.includes("<<MODEL_COM>>")) val = mod.comment || "";
                                        
                                        wTable.getCell(targetRow + 1, c).getText().setText(val);
                                    } catch(e) {}
                                }
                            }
                            wTable.getRow(targetRow).remove();
                        }
                    }
                }
            }
        }

        pres.saveAndClose();
        return { success: true, url: newFile.getUrl() };

    } catch(e) { 
        return { success: false, error: e.toString() }; 
    }
}

function generateProjectFiles(folderId, projectNo, projectName, filePrefix, processName, selectedFiles) {
  try {
    if (!processName || !Array.isArray(selectedFiles) || selectedFiles.length === 0) {
      throw new Error("A valid process and at least one selected file are required.");
    }

    const processResult = getProcessTableData();
    if (!processResult || processResult.success !== true || !Array.isArray(processResult.data)) {
      throw new Error("Cannot validate selected files against the process table.");
    }
    const processRows = processResult.data.filter(row => row && row.process === processName);
    if (processRows.length === 0) throw new Error("Unknown process.");

    const trustedFiles = [];
    const seenIds = new Set();
    selectedFiles.forEach(selected => {
      const fileId = selected && String(selected.fileId || "").trim();
      if (!fileId || seenIds.has(fileId)) throw new Error("Selected files contain a missing or duplicate file ID.");
      seenIds.add(fileId);
      const matches = processRows.filter(row => String(row.fileId || "").trim() === fileId);
      if (matches.length !== 1) throw new Error("A selected file is unknown or not part of the requested process.");
      trustedFiles.push(matches[0]);
    });

    const db = getLocalProjectDbAndBimFolder_(folderId, projectNo, projectName);
    const bimFolder = db.bimFolder;

    let mainTargetFolderUrl = "";
    const generatedFiles = [];

    trustedFiles.forEach(file => {
      const templateFile = DriveApp.getFileById(file.fileId);
      const newName = filePrefix + '_' + file.fileName;
      let targetSubFolder;
      const folderTargetName = (file.type === "מדריכים" || file.type === "Guides") ? "מדריכים" : processName;
      
      const subDirs = bimFolder.getFoldersByName(folderTargetName);
      if (subDirs.hasNext()) targetSubFolder = subDirs.next();
      else targetSubFolder = bimFolder.createFolder(folderTargetName);
      
      const copiedFile = templateFile.makeCopy(newName, targetSubFolder);
      
      mainTargetFolderUrl = targetSubFolder.getUrl(); 
      generatedFiles.push({ name: newName, url: copiedFile.getUrl() });
    });

    return { success: true, count: selectedFiles.length, folderUrl: mainTargetFolderUrl, files: generatedFiles };
  } catch(e) { return { success: false, error: e.toString() }; }
}

function forceAuth() {
  SlidesApp.create("Dummy Presentation");
}
