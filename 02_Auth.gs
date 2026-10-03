// ==========================================
// BIMBOT v41.0 - 02_Auth
// Signed-in user identity, permissions (USERS_DB_ID), sheet protection
// ==========================================

function getAuthenticatedUserContext_() {
  let email = "";
  try {
    email = String(Session.getActiveUser().getEmail() || "").toLowerCase().trim();
  } catch (e) {
    return { success: false, error: "Unable to identify the signed-in user." };
  }

  // USER_DEPLOYING can make EffectiveUser the deployer. It is never a viewer identity.
  if (!email) return { success: false, error: "Unable to identify the signed-in user." };

  let userData;
  try {
    userData = SpreadsheetApp.openById(USERS_DB_ID).getSheets()[0].getDataRange().getValues();
  } catch (e) {
    console.log("Error reading Users DB: " + e.toString());
    return { success: false, error: "Unable to verify user permissions." };
  }

  let matchingAccounts = [];
  let userFound = false;
  for (let i = 1; i < userData.length; i++) {
    const rowEmail = String(userData[i][1] || "").toLowerCase().trim();
    if (rowEmail !== email) continue;
    userFound = true;
    matchingAccounts.push({
      userName: String(userData[i][0] || "").trim(),
      canViewAllProjects: String(userData[i][2] || "").trim().toUpperCase() === "V",
      canViewDatabase: String(userData[i][3] || "").trim().toUpperCase() === "V"
    });
  }

  if (!userFound) return { success: false, error: "User is not authorized." };
  const selected = matchingAccounts[0];
  const selectedName = selected.userName.toLowerCase().trim();
  const sameEmailConflict = matchingAccounts.some(account =>
    account.userName.toLowerCase().trim() !== selectedName ||
    account.canViewAllProjects !== selected.canViewAllProjects ||
    account.canViewDatabase !== selected.canViewDatabase
  );
  if (sameEmailConflict) return { success: false, error: "Conflicting permission records for user." };

  const canViewAllProjects = selected.canViewAllProjects;
  const canViewDatabase = selected.canViewDatabase;
  const userName = selected.userName;
  const distinctNameAccounts = new Set();
  if (selectedName) {
    for (let i = 1; i < userData.length; i++) {
      const rowName = String(userData[i][0] || "").trim().toLowerCase();
      const rowEmail = String(userData[i][1] || "").trim().toLowerCase();
      if (rowName === selectedName && rowEmail) distinctNameAccounts.add(rowEmail);
    }
  }
  if (!canViewAllProjects && (!selectedName || distinctNameAccounts.size > 1)) {
    return { success: false, error: "User name is blank or ambiguous; project access cannot be verified." };
  }

  return {
    success: true,
    email: email,
    userName: userName,
    canViewAllProjects: canViewAllProjects,
    canViewDatabase: canViewDatabase
  };
}

function authorizePermissionsHelper() {
  DriveApp.getRootFolder();
  DocumentApp.openById(KICKOFF_DOC_TEMPLATE_ID);
  Logger.log("Permissions requested successfully.");
}

function protectProjectSpreadsheet_(ssId) {
  try {
    const userSheet = SpreadsheetApp.openById(USERS_DB_ID).getSheets()[0];
    const userData = userSheet.getDataRange().getValues();
    const allowedEmails = [];
    for (let i = 1; i < userData.length; i++) {
      const email = String(userData[i][1] || "").toLowerCase().trim();
      const canViewAll = String(userData[i][2] || "").trim().toUpperCase() === "V";
      const canEditDB = String(userData[i][3] || "").trim().toUpperCase() === "V";
      if (email && (canEditDB || canViewAll)) allowedEmails.push(email);
    }

    let currentUser = "";
    try { currentUser = String(Session.getActiveUser().getEmail() || "").toLowerCase().trim(); } catch (e) {}
    if (currentUser && !allowedEmails.includes(currentUser)) allowedEmails.push(currentUser);

    const ss = SpreadsheetApp.openById(ssId);
    ss.getSheets().forEach(sheet => {
      let protection = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];
      if (!protection) protection = sheet.protect().setDescription("App Controlled - View Only");
      protection.removeEditors(protection.getEditors());
      if (protection.canDomainEdit()) protection.setDomainEdit(false);
      if (allowedEmails.length) protection.addEditors(allowedEmails);
    });
  } catch (e) { console.log("Protection error: " + e.toString()); }
}

function getActiveUserInfo() {
  const context = getAuthenticatedUserContext_();
  if (!context.success) return context;
  return {
    success: true,
    email: context.email,
    canViewAllProjects: context.canViewAllProjects,
    canViewDatabase: context.canViewDatabase
  };
}
