// ==========================================
// BIM Commander v40.0 - 01_Utils
// Helper Functions used across multiple services
// ==========================================

function clearDataPreserveHeaders(sheet) {
  try {
    if (sheet.getMaxRows() > 0 && sheet.getMaxColumns() > 0) {
      sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
    }
    if (sheet.getLastRow() > 1 && sheet.getMaxColumns() > 0) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getMaxColumns()).clearContent();
    }
  } catch(err) {
    console.warn("Failed to clear sheet: " + sheet.getName(), err);
  }
}

function clearDataPreserveHeadersVertical(sheet) {
  try {
    if (sheet.getMaxRows() > 0 && sheet.getMaxColumns() > 0) {
      sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
    }
    sheet.clearContent();
  } catch(err) {
    console.warn("Failed to clear vertical sheet: " + sheet.getName(), err);
  }
}

function wipeSheetClean(sheet) {
  try {
    if (sheet.getMaxRows() > 0 && sheet.getMaxColumns() > 0) {
      sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
    }
    sheet.clear(); 
  } catch(err) {
    console.warn("Failed to wipe sheet clean: " + sheet.getName(), err);
  }
}

function mapYesNo(v) {
  let s = String(v || "").trim().toLowerCase();
  let yesWords = ["כן", "קיים מודל אדריכלות", "יש grids", "קיים", "yes", "true", "v"];
  let noWords = ["לא", "לא קיים מודל אדריכלות", "אין grids", "לא קיים", "no", "false", "x"];
  if (yesWords.indexOf(s) > -1) return "Yes"; 
  if (noWords.indexOf(s) > -1) return "No"; 
  return String(v || "").trim();
}

function removeDocBlock(body, startTag, endTag) {
  try {
    let paragraphs = body.getParagraphs();
    let inBlock = false;
    let elementsToRemove = [];

    for (let i = 0; i < paragraphs.length; i++) {
      let text = paragraphs[i].editAsText().getText();
      
      if (text.indexOf(startTag) > -1) {
        inBlock = true;
        elementsToRemove.push(paragraphs[i]);
        continue;
      }
      
      if (text.indexOf(endTag) > -1) {
        inBlock = false;
        elementsToRemove.push(paragraphs[i]);
        continue;
      }
      
      if (inBlock) {
        elementsToRemove.push(paragraphs[i]);
      }
    }

    elementsToRemove.forEach(el => {
      try { el.removeFromParent(); } catch(e) {}
    });
  } catch(e) {
    console.log("Error removing block: " + e.toString());
  }
}

function removeParagraphsWithText(body, textToFind) {
  try {
    let paragraphs = body.getParagraphs();
    for (let i = paragraphs.length - 1; i >= 0; i--) {
      if (paragraphs[i].editAsText().getText().indexOf(textToFind) > -1) {
        paragraphs[i].removeFromParent();
      }
    }
  } catch(e) {}
}