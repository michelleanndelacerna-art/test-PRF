// ============================================================
// MANPOWER REQUEST MANAGEMENT SYSTEM - GOOGLE APPS SCRIPT
// ============================================================

// Google Sheet Configuration
const SHEET_ID = "1qaC9H6Cbyeg2nhN-yotZFfUE8kzvbzW0ZktTFBAWvNI";
const ATTACHMENTS_FOLDER_ID = "1KhB0TavHpDj_iEmqIzrzOT-m3pPOeErM";
const SHEET_NAMES = {
  OPTIMUM_HC: "Optimum Manpower HC",
  ORG_CHART: "Org Chart",
  REQUESTS: "Requests",
  USERS: "Users"
};
const APPROVAL_HISTORY_KEY_PREFIX = "approvalHistory:";
const APPROVAL_REMARK_COLUMNS = [
  {
    label: "Plant Head",
    key: "remarks1",
    display: "Remarks 1",
    names: ["remarks 1", "remarks1", "remark 1", "plant head remarks", "plant head comment"]
  },
  {
    label: "BU Head",
    key: "remarks2",
    display: "Remarks 2",
    names: ["remarks 2", "remarks2", "remark 2", "bu head remarks", "bu head comment", "business unit head remarks"]
  },
  {
    label: "Managing Director",
    key: "remarks3",
    display: "Remarks 3",
    names: ["remarks 3", "remarks3", "remark 3", "managing director remarks", "managing director comment", "md remarks"]
  }
];

/**
 * Get sheet by name with fallback names for compatibility
 */
function getSheetByNameOrFallback(name, fallbackNames = []) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    for (let i = 0; i < fallbackNames.length; i++) {
      sheet = ss.getSheetByName(fallbackNames[i]);
      if (sheet) break;
    }
  }
  return sheet;
}

/**
 * Get the Requests sheet with fallback names for backward compatibility
 */
function getRequestsSheet() {
  const sheet = getSheetByNameOrFallback(SHEET_NAMES.REQUESTS, ["Request", "request"]);
  if (!sheet) {
    throw new Error('Requests sheet not found. Please ensure a sheet named "Requests" or "Request" exists.');
  }
  return sheet;
}

function getRequestHeaderMap() {
  const sheet = getRequestsSheet();
  const lastColumn = sheet.getLastColumn();
  const headerRow = sheet.getRange(1, 1, 1, lastColumn).getValues()[0] || [];
  const map = {};

  headerRow.forEach((header, idx) => {
    const key = (header || "").toString().trim().toLowerCase();
    if (key) {
      map[key] = idx;
    }
  });

  return {
    headerRow: headerRow,
    map: map
  };
}

function getRequestField(row, map, names, defaultValue) {
  const keys = Array.isArray(names) ? names : [names];
  for (const name of keys) {
    const idx = map[name.toLowerCase()];
    if (idx !== undefined && row[idx] !== undefined) {
      return row[idx];
    }
  }
  return defaultValue !== undefined ? defaultValue : "";
}

function getRequestFieldJson(row, map, names) {
  const value = getRequestField(row, map, names, "");
  try {
    return JSON.parse(value || "[]");
  } catch (e) {
    return [];
  }
}

function setRequestField(row, map, names, value) {
  const keys = Array.isArray(names) ? names : [names];
  for (const name of keys) {
    const idx = map[name.toLowerCase()];
    if (idx !== undefined) {
      row[idx] = value;
      return true;
    }
  }
  return false;
}

function normalizeRequestValue(value) {
  return (value || "").toString().trim().toLowerCase();
}

function getRequestColumnIndex(map, names) {
  const keys = Array.isArray(names) ? names : [names];
  for (const name of keys) {
    const idx = map[normalizeRequestValue(name)];
    if (idx !== undefined) {
      return idx;
    }
  }
  return -1;
}

function requireRequestColumnIndex(map, names, label) {
  const idx = getRequestColumnIndex(map, names);
  if (idx < 0) {
    throw new Error('Requests sheet missing required column: ' + label);
  }
  return idx;
}

function parseRequestJson(value, fallbackValue) {
  try {
    return JSON.parse(value || JSON.stringify(fallbackValue));
  } catch (e) {
    return fallbackValue;
  }
}

function getApprovalHistoryKey(requestID) {
  return APPROVAL_HISTORY_KEY_PREFIX + requestID;
}

function getApprovalHistory(requestID) {
  if (!requestID) {
    return [];
  }

  const raw = PropertiesService.getScriptProperties().getProperty(getApprovalHistoryKey(requestID));
  return parseRequestJson(raw, []);
}

function saveApprovalHistory(requestID, history) {
  if (!requestID) {
    return;
  }

  PropertiesService.getScriptProperties().setProperty(
    getApprovalHistoryKey(requestID),
    JSON.stringify(history || [])
  );
}

function appendApprovalHistoryEntry(requestID, entry) {
  const history = getApprovalHistory(requestID);
  history.push(entry);
  saveApprovalHistory(requestID, history);
  return history;
}

function getApprovalRemarkConfig(levelLabel) {
  const normalizedLabel = normalizeRequestValue(levelLabel);
  for (let i = 0; i < APPROVAL_REMARK_COLUMNS.length; i++) {
    const config = APPROVAL_REMARK_COLUMNS[i];
    if (normalizeRequestValue(config.label) === normalizedLabel) {
      return config;
    }
  }
  return null;
}

function getApprovalUserIdentifiers(user) {
  return [user && user.email, user && user.role, user && user.name]
    .map(normalizeRequestValue)
    .filter(Boolean);
}

function matchesApprovalIdentity(value, user) {
  const normalizedValue = normalizeRequestValue(value);
  if (!normalizedValue) {
    return false;
  }

  return getApprovalUserIdentifiers(user).indexOf(normalizedValue) >= 0;
}

function findApprovalStepIndex(steps, approverOrLabel) {
  const normalized = normalizeRequestValue(approverOrLabel);
  if (!normalized) {
    return -1;
  }

  for (let i = 0; i < steps.length; i++) {
    if (
      normalizeRequestValue(steps[i].approver) === normalized ||
      normalizeRequestValue(steps[i].label) === normalized
    ) {
      return i;
    }
  }

  return -1;
}

function findApprovalStepIndexByLabel(steps, label) {
  const normalized = normalizeRequestValue(label);
  if (!normalized) {
    return -1;
  }

  for (let i = 0; i < steps.length; i++) {
    if (normalizeRequestValue(steps[i].label) === normalized) {
      return i;
    }
  }

  return -1;
}

function getApprovalBucketLabel(bucketKey) {
  if (bucketKey === "approved") return "Approved";
  if (bucketKey === "disapproved") return "Disapproved";
  if (bucketKey === "on-hold") return "On Hold";
  return "Pending";
}

function getApprovalBucketFromAction(action) {
  if (action === "Approve") return "approved";
  if (action === "Disapprove") return "disapproved";
  if (action === "Hold") return "on-hold";
  return "pending";
}

function getRemarkValueForLevel(request, levelLabel, historyByLevel) {
  const config = getApprovalRemarkConfig(levelLabel);
  if (config && request[config.key]) {
    return request[config.key];
  }

  const historyEntry = historyByLevel[normalizeRequestValue(levelLabel)];
  return historyEntry && historyEntry.comment ? historyEntry.comment : "";
}

function resolveCurrentApprovalLevelLabel(request, approvalSteps, approvalHistory) {
  if (request.status === "Approved") {
    return approvalSteps.length ? approvalSteps[approvalSteps.length - 1].label : "Completed";
  }

  if (request.currentApprover && request.currentApprover !== "Completed" && request.currentApprover !== "Requestor") {
    const currentIndex = findApprovalStepIndex(approvalSteps, request.currentApprover);
    if (currentIndex >= 0) {
      return approvalSteps[currentIndex].label;
    }
    return request.currentApprover;
  }

  if (approvalHistory && approvalHistory.length > 0) {
    const lastEntry = approvalHistory[approvalHistory.length - 1];
    const levels = Array.isArray(lastEntry.levels) ? lastEntry.levels.filter(Boolean) : [];
    if (levels.length > 0) {
      return levels[levels.length - 1];
    }
    if (lastEntry.level) {
      return lastEntry.level;
    }
  }

  return request.currentApprover || "";
}

function buildApprovalStatusMessage(status, currentLevelLabel, nextApprovalLabel) {
  if (status === "Approved") {
    return "This request completed all required approvals.";
  }

  if (status === "Disapproved") {
    return currentLevelLabel
      ? "This request was disapproved at " + currentLevelLabel + ". Review the remarks below before submitting again."
      : "This request was disapproved. Review the remarks below before submitting again.";
  }

  if (status === "On Hold") {
    return currentLevelLabel
      ? "This request is currently on hold at " + currentLevelLabel + "."
      : "This request is currently on hold.";
  }

  if (nextApprovalLabel) {
    return "This request is waiting for " + nextApprovalLabel + " review.";
  }

  return "This request is waiting for the next approval action.";
}

function buildApprovalTrackerData(request, approvalSteps, approvalHistory) {
  const steps = (approvalSteps || []).map(step => ({
    label: step.label,
    approver: step.approver,
    state: "pending",
    action: "",
    comment: "",
    timestamp: "",
    actorEmail: "",
    actorName: ""
  }));

  if (!steps.length && request.currentApprover && request.currentApprover !== "Completed" && request.currentApprover !== "Requestor") {
    steps.push({
      label: request.currentApprover,
      approver: request.currentApprover,
      state: "pending",
      action: "",
      comment: "",
      timestamp: "",
      actorEmail: "",
      actorName: ""
    });
  }

  const historyByLevel = {};
  (approvalHistory || []).forEach(entry => {
    const levels = Array.isArray(entry.levels) && entry.levels.length > 0
      ? entry.levels.filter(Boolean)
      : entry.level ? [entry.level] : [];

    levels.forEach(level => {
      historyByLevel[normalizeRequestValue(level)] = {
        action: entry.action || "",
        comment: entry.comment || "",
        timestamp: entry.timestamp || "",
        actorEmail: entry.actorEmail || "",
        actorName: entry.actorName || ""
      };
    });
  });

  steps.forEach(step => {
    const historyEntry = historyByLevel[normalizeRequestValue(step.label)];
    step.comment = getRemarkValueForLevel(request, step.label, historyByLevel);
    if (historyEntry) {
      step.action = historyEntry.action;
      step.timestamp = historyEntry.timestamp;
      step.actorEmail = historyEntry.actorEmail;
      step.actorName = historyEntry.actorName;
      step.state = historyEntry.action === "Approve"
        ? "completed"
        : historyEntry.action === "Disapprove"
          ? "disapproved"
          : historyEntry.action === "Hold"
            ? "on_hold"
            : "pending";
    }
  });

  const currentLevelLabel = resolveCurrentApprovalLevelLabel(request, steps, approvalHistory);
  const currentIndex = findApprovalStepIndexByLabel(steps, currentLevelLabel);

  if (request.status === "Approved") {
    steps.forEach(step => {
      if (step.state === "pending" || step.state === "current") {
        step.state = "completed";
      }
    });
  } else if (request.status === "Pending") {
    if (currentIndex >= 0) {
      for (let i = 0; i < currentIndex; i++) {
        if (steps[i].state === "pending") {
          steps[i].state = "completed";
        }
      }
      if (steps[currentIndex].state === "pending") {
        steps[currentIndex].state = "current";
      }
    }
  } else if (request.status === "On Hold") {
    if (currentIndex >= 0) {
      for (let i = 0; i < currentIndex; i++) {
        if (steps[i].state === "pending") {
          steps[i].state = "completed";
        }
      }
      steps[currentIndex].state = "on_hold";
    }
  } else if (request.status === "Disapproved") {
    if (currentIndex >= 0) {
      for (let i = 0; i < currentIndex; i++) {
        if (steps[i].state === "pending") {
          steps[i].state = "completed";
        }
      }
      steps[currentIndex].state = "disapproved";
    }
  }

  const nextApprovalLabel = request.status === "Pending" || request.status === "On Hold"
    ? (currentLevelLabel || request.currentApprover || "Pending")
    : request.status === "Approved"
      ? "Completed"
      : "Returned to Requestor";

  return {
    steps: steps,
    currentLevelLabel: currentLevelLabel,
    nextApprovalLabel: nextApprovalLabel,
    statusMessage: buildApprovalStatusMessage(request.status, currentLevelLabel, nextApprovalLabel),
    completedCount: steps.filter(step => step.state === "completed").length,
    totalSteps: steps.length
  };
}

function getApprovalViewStateForUser(request, user) {
  if (request.status === "Pending" && matchesApprovalIdentity(request.currentApprover, user)) {
    return "pending";
  }

  if (request.status === "On Hold" && matchesApprovalIdentity(request.currentApprover, user)) {
    return "on-hold";
  }

  const history = request.approvalHistory || [];
  const userIdentifiers = getApprovalUserIdentifiers(user);
  const userHistory = history.filter(entry => {
    const entryIdentifiers = [entry.actorEmail, entry.actorRole, entry.actorName, entry.approver]
      .map(normalizeRequestValue)
      .filter(Boolean);

    return userIdentifiers.some(identifier => entryIdentifiers.indexOf(identifier) >= 0);
  });

  if (userHistory.length > 0) {
    return getApprovalBucketFromAction(userHistory[userHistory.length - 1].action);
  }

  const approvalSteps = request.approvalSteps || [];
  const userStepIndexes = approvalSteps
    .map((step, index) => matchesApprovalIdentity(step.approver, user) ? index : -1)
    .filter(index => index >= 0);

  if (!userStepIndexes.length) {
    return "";
  }

  const furthestUserIndex = userStepIndexes[userStepIndexes.length - 1];
  const currentIndex = findApprovalStepIndex(approvalSteps, request.currentApprover);

  if (request.status === "Approved") {
    return "approved";
  }

  if ((request.status === "Pending" || request.status === "On Hold") && currentIndex > furthestUserIndex) {
    return "approved";
  }

  return "";
}

function buildRequestRecord(row, headerMap, rowNumber) {
  const requestID = getRequestField(row, headerMap, ["request id", "id"], "");
  const manpowerData = parseRequestJson(
    getRequestField(row, headerMap, ["manpower data", "manpower_data"], ""),
    {}
  );

  const request = {
    rowNumber: rowNumber,
    requestID: requestID,
    createdDate: getRequestField(row, headerMap, ["timestamp", "date"], ""),
    email: getRequestField(row, headerMap, ["requestor", "requestor email", "requested by", "email"], ""),
    requestedBy: getRequestField(row, headerMap, ["requestor", "requested by", "requestor email"], ""),
    division: getRequestField(row, headerMap, ["division"], ""),
    group: getRequestField(row, headerMap, ["group"], ""),
    department: getRequestField(row, headerMap, ["department"], ""),
    section: getRequestField(row, headerMap, ["section"], ""),
    unit: getRequestField(row, headerMap, ["unit"], ""),
    line: getRequestField(row, headerMap, ["line"], ""),
    requestType: getRequestField(row, headerMap, ["type", "request type"], ""),
    details: getRequestField(row, headerMap, ["details", "detail"], ""),
    status: getRequestField(row, headerMap, ["status"], ""),
    approverLevel: getRequestField(row, headerMap, ["approver_level", "approver level", "approverlevel", "current approver", "current_approver"], ""),
    justification: getRequestField(row, headerMap, ["justification.", "justification"], ""),
    attachedFiles: getRequestField(row, headerMap, ["attached files", "attachments", "attachedfiles"], ""),
    category: getRequestField(row, headerMap, ["category"], ""),
    positions: getRequestFieldJson(row, headerMap, ["positions"]),
    approverNotes: getRequestField(row, headerMap, ["approver notes", "approver_notes", "notes"], ""),
    currentApprover: getRequestField(row, headerMap, ["next approver", "next_approver", "current approver", "current_approver", "approver_level", "approver level", "approverlevel"], ""),
    plantHeadApproval: getRequestField(row, headerMap, ["plant head approval", "plant_head_approval"], ""),
    buHeadApproval: getRequestField(row, headerMap, ["bu head approval", "bu_head_approval"], ""),
    typeSelectionApproval: getRequestField(row, headerMap, ["type selection approval", "type_selection_approval"], ""),
    reviewApproval: getRequestField(row, headerMap, ["review approval", "review_approval"], ""),
    legalApproval: getRequestField(row, headerMap, ["legal approval", "legal_approval"], ""),
    corporateHrodApproval: getRequestField(row, headerMap, ["corporate hrod approval", "corporate_hrod_approval"], ""),
    recruitmentNotification: getRequestField(row, headerMap, ["recruitment notification", "recruitment_notification"], ""),
    approvalChain: getRequestField(row, headerMap, ["approval chain", "approval_chain"], ""),
    manpowerData: manpowerData,
    optimumHC: getRequestField(row, headerMap, ["optimum manpower hc", "optimum hc", "optimum_hc"], manpowerData.optimumHC || ""),
    actualHC: getRequestField(row, headerMap, ["actual manpower hc", "actual hc", "actual_hc"], manpowerData.actualHC || ""),
    gap: getRequestField(row, headerMap, ["gap"], manpowerData.gap || ""),
    remarks1: getRequestField(row, headerMap, APPROVAL_REMARK_COLUMNS[0].names, ""),
    remarks2: getRequestField(row, headerMap, APPROVAL_REMARK_COLUMNS[1].names, ""),
    remarks3: getRequestField(row, headerMap, APPROVAL_REMARK_COLUMNS[2].names, "")
  };

  const approvalHistory = getApprovalHistory(requestID);
  const approvalSteps = getApprovalStepsForRequest(request);
  const approvalTracker = buildApprovalTrackerData(request, approvalSteps, approvalHistory);

  request.approvalHistory = approvalHistory;
  request.approvalSteps = approvalSteps;
  request.approvalTracker = approvalTracker;

  return request;
}

function updateApprovalRemarkColumns(sheet, rowNumber, headerMap, levels, comments) {
  const trimmedComments = (comments || "").toString().trim();
  if (!trimmedComments) {
    return;
  }

  const writtenColumns = {};
  (levels || []).forEach(level => {
    const config = getApprovalRemarkConfig(level);
    if (!config) {
      return;
    }

    const idx = getRequestColumnIndex(headerMap, config.names);
    if (idx >= 0 && !writtenColumns[idx]) {
      sheet.getRange(rowNumber, idx + 1).setValue(trimmedComments);
      writtenColumns[idx] = true;
    }
  });
}

function getSystemAppUrl() {
  try {
    return ScriptApp.getService().getUrl() || "";
  } catch (e) {
    return "";
  }
}

function escapeHtml(value) {
  return (value || "")
    .toString()
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function formatEmailMultilineText(value) {
  return escapeHtml(value).replace(/\n/g, "<br>");
}

function getAttachmentsParentFolder() {
  return DriveApp.getFolderById(ATTACHMENTS_FOLDER_ID);
}

// ============================================================
// 1. USER ROLE MANAGEMENT
// ============================================================

/**
 * Get current user's email
 */
function getCurrentUserEmail() {
  return Session.getActiveUser().getEmail();
}

/**
 * Get user role from Users sheet
 */
function getUserRole() {
  const email = getCurrentUserEmail();
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAMES.USERS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      return {
        email: email,
        role: data[i][1] || "Requestor",
        name: data[i][2] || "User",
        department: data[i][3] || "",
        division: data[i][4] || ""
      };
    }
  }
  
  // Default: new user is Requestor
  return {
    email: email,
    role: "Requestor",
    name: email.split("@")[0],
    department: "",
    division: ""
  };
}

/**
 * Get all data for initialization
 */
function getInitialData() {
  const userRole = getUserRole();
  const divisions = getCompanyStructure();
  
  return {
    user: userRole,
    divisions: divisions,
    timestamp: new Date().toISOString()
  };
}

// ============================================================
// 2. COMPANY STRUCTURE & MANPOWER DATA
// ============================================================

/**
 * Get company structure (Divisions, Groups, Departments, Sections, Units, Lines)
 */

function getCompanyStructure() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAMES.OPTIMUM_HC);
  const data = sheet.getDataRange().getValues();
  // Assume header row: Division, Group, Department, Section, Unit, Line, Optimum HC, Plant Head, BU Head, Managing Director, Category
  const header = data[0].map(h => (h || '').toString().trim().toLowerCase());
  const idx = {
    division: header.indexOf('division'),
    group: header.indexOf('group'),
    department: header.indexOf('department'),
    section: header.indexOf('section'),
    unit: header.indexOf('unit'),
    line: header.indexOf('line'),
    optimumHC: header.indexOf('optimum hc'),
    plantHead: header.indexOf('plant head'),
    buHead: header.indexOf('bu head'),
    managingDirector: header.indexOf('managing director'),
    category: header.indexOf('category')
  };

  const structure = {};

  for (let i = 1; i < data.length; i++) {
    const division = data[i][idx.division];
    const group = data[i][idx.group];
    const department = data[i][idx.department];
    const section = data[i][idx.section];
    const unit = data[i][idx.unit];
    const line = data[i][idx.line];
    let optimumHC = data[i][idx.optimumHC];
    if (optimumHC === undefined || optimumHC === null || optimumHC === '') {
      optimumHC = null;
    } else {
      optimumHC = Number(optimumHC);
      if (isNaN(optimumHC)) optimumHC = null;
    }
    const plantHead = (data[i][idx.plantHead] || '').toString().trim();
    const buHead = (data[i][idx.buHead] || '').toString().trim();
    const managingDirector = (data[i][idx.managingDirector] || '').toString().trim();
    const category = (data[i][idx.category] || '').toString().trim();

    if (division) {
      if (!structure[division]) structure[division] = { _categories: [], groups: {} };
      if (!structure[division].groups[group]) structure[division].groups[group] = { _categories: [], departments: {} };
      if (!structure[division].groups[group].departments[department]) structure[division].groups[group].departments[department] = { _categories: [], sections: {} };
      if (!structure[division].groups[group].departments[department].sections[section]) structure[division].groups[group].departments[department].sections[section] = { _categories: [], units: {} };
      if (!structure[division].groups[group].departments[department].sections[section].units[unit]) structure[division].groups[group].departments[department].sections[section].units[unit] = { _categories: [], lines: {} };

      structure[division].groups[group].departments[department].sections[section].units[unit].lines[line] = {
        optimumHC: optimumHC,
        plantHead: plantHead,
        buHead: buHead,
        managingDirector: managingDirector,
        category: category
      };

      // Add category to _categories array at each level if not already present
      function addCat(arr) { if (category && arr.indexOf(category) === -1) arr.push(category); }
      addCat(structure[division]._categories);
      addCat(structure[division].groups[group]._categories);
      addCat(structure[division].groups[group].departments[department]._categories);
      addCat(structure[division].groups[group].departments[department].sections[section]._categories);
      addCat(structure[division].groups[group].departments[department].sections[section].units[unit]._categories);
    }
  }
  return structure;
}

/**
 * Get approval email chain for a specific organization location.
 */
function getApprovalEmails(division, group, department, section, unit, line) {
  return getOrgApprovalSteps(division, group, department, section, unit, line)
    .map(step => step.approver);
}

function getOrgApprovalSteps(division, group, department, section, unit, line) {
  const structure = getCompanyStructure();

  function buildSteps(row) {
    return [
      { label: "Plant Head", approver: (row.plantHead || "").toString().trim() },
      { label: "BU Head", approver: (row.buHead || "").toString().trim() },
      { label: "Managing Director", approver: (row.managingDirector || "").toString().trim() }
    ].filter(step => step.approver);
  }

  if (unit && line && structure[division] && 
      structure[division].groups && structure[division].groups[group] && 
      structure[division].groups[group].departments && structure[division].groups[group].departments[department] &&
      structure[division].groups[group].departments[department].sections && structure[division].groups[group].departments[department].sections[section] && 
      structure[division].groups[group].departments[department].sections[section].units && structure[division].groups[group].departments[department].sections[section].units[unit] &&
      structure[division].groups[group].departments[department].sections[section].units[unit].lines && structure[division].groups[group].departments[department].sections[section].units[unit].lines[line]) {
    const exactSteps = buildSteps(structure[division].groups[group].departments[department].sections[section].units[unit].lines[line]);
    if (exactSteps.length) return exactSteps;
  }

  if (structure[division] && 
      structure[division].groups && structure[division].groups[group] && 
      structure[division].groups[group].departments && structure[division].groups[group].departments[department] &&
      structure[division].groups[group].departments[department].sections && structure[division].groups[group].departments[department].sections[section]) {
    const sectionData = structure[division].groups[group].departments[department].sections[section].units;
    for (const unitKey in sectionData) {
      const linesData = sectionData[unitKey].lines;
      if (linesData) {
        for (const lineKey in linesData) {
          const sectionSteps = buildSteps(linesData[lineKey]);
          if (sectionSteps.length) return sectionSteps;
        }
      }
    }
  }

  return [];
}

function getNormalizedRequestType(requestData) {
  if (requestData.requestType === "Additional" && requestData.positions && requestData.positions.length > 0) {
    return requestData.positions[0].type || "Regular";
  }
  return requestData.requestType;
}

function getApprovalStepsForData(requestData) {
  const orgSteps = getOrgApprovalSteps(
    requestData.division,
    requestData.group,
    requestData.department,
    requestData.section,
    requestData.unit,
    requestData.line
  );

  if (orgSteps.length) {
    return orgSteps;
  }

  const route = requestData.gap <= 0
    ? ["Plant Head", "BU Head", "Legal", "Corporate HROD"]
    : getApprovalRoute(getNormalizedRequestType(requestData), requestData.category);

  return route.map(label => ({ label: label, approver: label }));
}

function getApprovalStepsForRequest(request) {
  return getApprovalStepsForData({
    division: request.division,
    group: request.group,
    department: request.department,
    section: request.section,
    unit: request.unit,
    line: request.line,
    requestType: request.requestType,
    category: request.category,
    positions: request.positions,
    gap: request.gap
  });
}

function getRequestColumnIndex(map, names) {
  const keys = Array.isArray(names) ? names : [names];
  for (const name of keys) {
    const idx = map[normalizeRequestValue(name)];
    if (idx !== undefined) {
      return idx;
    }
  }
  return -1;
}

function getApprovalRemarkFieldNames(stepLabel, stepIndex) {
  if (stepLabel === "Plant Head" || stepIndex === 0) {
    return ["remarks 1", "remarks1", "remark 1", "remark1"];
  }
  if (stepLabel === "BU Head" || stepIndex === 1) {
    return ["remarks 2", "remarks2", "remark 2", "remark2"];
  }
  if (stepLabel === "Managing Director" || stepIndex === 2) {
    return ["remarks 3", "remarks3", "remark 3", "remark3"];
  }
  return ["approver notes", "approver_notes", "notes"];
}

function getApprovalColumnFieldNames(stepLabel, stepIndex) {
  // Maps approval role to approval column names
  const norm = (stepLabel || "").toLowerCase().trim();
  const config = { status: [], approver: [] };

  if (norm.indexOf("plant") !== -1) {
    config.status = ["plant head approval", "plant_head_approval"];
    config.approver = ["plant head approver", "plant_head_approver"];
  } else if (norm.indexOf("bu") !== -1 || norm.indexOf("business") !== -1) {
    config.status = ["bu head approval", "bu_head_approval"];
    config.approver = ["bu head approver", "bu_head_approver"];
  } else if (norm.indexOf("legal") !== -1) {
    config.status = ["legal approval", "legal_approval"];
    config.approver = ["legal approver", "legal_approver"];
  } else if (norm.indexOf("corporate") !== -1 && norm.indexOf("hrod") !== -1) {
    config.status = ["corporate hrod approval", "corporate_hrod_approval"];
    config.approver = ["corporate hrod approver", "corporate_hrod_approver"];
  } else if (norm.indexOf("type") !== -1 || norm === "type selection") {
    config.status = ["type selection approval", "type_selection_approval"];
  } else if (norm.indexOf("review") !== -1) {
    config.status = ["review approval", "review_approval"];
  } else if (norm.indexOf("recruitment") !== -1) {
    config.status = ["recruitment notification", "recruitment_notification"];
  }
  
  return config;
}

function getApprovalRemarkFromRow(row, map, stepLabel, stepIndex) {
  return getRequestField(row, map, getApprovalRemarkFieldNames(stepLabel, stepIndex), "");
}

function getApprovalRemarkFromRequest(request, stepLabel, stepIndex) {
  const names = getApprovalRemarkFieldNames(stepLabel, stepIndex);
  if (names.indexOf("remarks 1") !== -1) return request.remarks1 || "";
  if (names.indexOf("remarks 2") !== -1) return request.remarks2 || "";
  if (names.indexOf("remarks 3") !== -1) return request.remarks3 || "";
  return request.approverNotes || "";
}

function getApprovalActionFromRemark(remark) {
  const value = (remark || "").toString().trim().toLowerCase();
  if (value.indexOf("approved") === 0) return "Approved";
  if (value.indexOf("disapproved") === 0) return "Disapproved";
  if (value.indexOf("on hold") === 0) return "On Hold";
  return "";
}

function buildApprovalRemark(action, comments) {
  const label = action === "Approve"
    ? "Approved"
    : action === "Disapprove"
      ? "Disapproved"
      : "On Hold";
  const note = (comments || "").toString().trim();
  return note ? `${label}: ${note}` : label;
}

function buildRequestResponse(row, rowNumber, headerMap) {
  return buildRequestRecord(row, headerMap, rowNumber);
}

function getMatchingApprovalIndexes(request, userEmail, userRole) {
  const user = {
    email: userEmail,
    role: userRole,
    name: ""
  };

  return (request.approvalSteps || []).reduce((matches, step, index) => {
    if (matchesApprovalIdentity(step.approver, user)) {
      matches.push(index);
    }
    return matches;
  }, []);
}

function getApprovalReviewState(request, userEmail, userRole) {
  const matchingIndexes = getMatchingApprovalIndexes(request, userEmail, userRole);
  const user = {
    email: userEmail,
    role: userRole,
    name: ""
  };
  const isAssigned = matchesApprovalIdentity(request.currentApprover, user);

  if (!matchingIndexes.length && !isAssigned) {
    return null;
  }

  const reviewBucket = getApprovalViewStateForUser(request, user);
  const reviewState = reviewBucket ? getApprovalBucketLabel(reviewBucket) : "";

  if (!reviewState) {
    return null;
  }

  return {
    reviewState: reviewState,
    canAct: isAssigned && (request.status === "Pending" || request.status === "On Hold"),
    matchedLevels: matchingIndexes.map(index => request.approvalSteps[index] ? request.approvalSteps[index].label : "")
  };
}

/**
 * Build a details summary string for the request row.
 */
function buildRequestDetails(requestData) {
  const parts = [];
  if (requestData.category) parts.push("Category: " + requestData.category);
  if (requestData.gap !== undefined) parts.push("Gap: " + requestData.gap);
  if (requestData.positions && requestData.positions.length) {
    parts.push("Positions: " + requestData.positions.map(p => `${p.position} (${p.headcount})`).join(", "));
  }
  return parts.join(" | ");
}

/**
 * Get filtered structure based on user selections
 */
function getFilteredStructure(category, division, group, department, section) {
  const structure = getCompanyStructure();
  
  if (!division) {
    return Object.keys(structure);
  }
  
  if (!group && structure[division]) {
    return Object.keys(structure[division]);
  }
  
  if (!department && structure[division] && structure[division][group]) {
    return Object.keys(structure[division][group]);
  }
  
  if (!section && structure[division] && structure[division][group] && structure[division][group][department]) {
    return Object.keys(structure[division][group][department]);
  }
  
  if (!category && structure[division] && structure[division][group] && structure[division][group][department] && structure[division][group][department][section]) {
    return Object.keys(structure[division][group][department][section]);
  }
  
  return [];
}

/**
 * Calculate manpower gap
 */
function calculateManpowerGap(division, group, department, section, unit, line) {
  const structure = getCompanyStructure();
  const actualSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAMES.ORG_CHART);
  const actualData = actualSheet.getDataRange().getValues();
  
  // Get optimum HC - for Production (with unit & line) or Non-Production (section level)
  let optimumHC = 0;
  if (unit && line) {
    // Production: lookup by line level (correct nested structure)
    if (
      structure[division] &&
      structure[division].groups && structure[division].groups[group] &&
      structure[division].groups[group].departments && structure[division].groups[group].departments[department] &&
      structure[division].groups[group].departments[department].sections && structure[division].groups[group].departments[department].sections[section] &&
      structure[division].groups[group].departments[department].sections[section].units && structure[division].groups[group].departments[department].sections[section].units[unit] &&
      structure[division].groups[group].departments[department].sections[section].units[unit].lines && structure[division].groups[group].departments[department].sections[section].units[unit].lines[line]
    ) {
      optimumHC = structure[division]
        .groups[group]
        .departments[department]
        .sections[section]
        .units[unit]
        .lines[line].optimumHC || 0;
    }
  } else {
    // Non-Production: sum by section level (correct nested structure)
    if (
      structure[division] &&
      structure[division].groups && structure[division].groups[group] &&
      structure[division].groups[group].departments && structure[division].groups[group].departments[department] &&
      structure[division].groups[group].departments[department].sections && structure[division].groups[group].departments[department].sections[section]
    ) {
      const unitsObj = structure[division].groups[group].departments[department].sections[section].units;
      for (const unitKey in unitsObj) {
        const linesObj = unitsObj[unitKey].lines;
        for (const lineKey in linesObj) {
          const hc = linesObj[lineKey].optimumHC;
          if (typeof hc === 'number' && !isNaN(hc)) {
            optimumHC += hc;
          }
        }
      }
    }
  }
  
  // Get actual HC from Org Chart sheet by counting employee records
  // Org Chart columns: Position ID(0), Employee ID(1), Employee Name(2), ...
  // Division(7), Group(8), Department(9), Section(10), Unit(11), Line(12), ...
  // Competency(18), Employment Status(19), Position Status(20)
  let actualHC = 0;
  Logger.log("Searching Org Chart for: Division=" + division + ", Group=" + group + ", Dept=" + department + ", Section=" + section + ", Unit=" + unit + ", Line=" + line);
  
  for (let i = 1; i < actualData.length; i++) {
    // Skip empty rows
    if (!actualData[i][0] || actualData[i].length < 21) continue;
    
    // Get values with trim to handle whitespace
    const status = (actualData[i][20] || "").toString().trim(); // Position Status column
    const divRow = (actualData[i][7] || "").toString().trim();
    const grpRow = (actualData[i][8] || "").toString().trim();
    const deptRow = (actualData[i][9] || "").toString().trim();
    const sectRow = (actualData[i][10] || "").toString().trim();
    const unitRow = (actualData[i][11] || "").toString().trim();
    const lineRow = (actualData[i][12] || "").toString().trim();
    
    // Check if employee position is ACTIVE (from Position Status column)
    if (status === "ACTIVE") {
      // Match division, group, department, section
      if (divRow === division && grpRow === group && 
          deptRow === department && sectRow === section) {
        // For production with unit & line, match all fields
        if (unit && line) {
          if (unitRow === unit && lineRow === line) {
            actualHC += 1; // Count this active employee
            Logger.log("Matched employee at row " + (i+1) + ": " + actualData[i][2]);
          }
        } else {
          // For non-production, count all active units/lines in this section
          actualHC += 1; // Count this active employee
          Logger.log("Matched non-production employee at row " + (i+1) + ": " + actualData[i][2] + " (Unit: " + unitRow + ", Line: " + lineRow + ")");
        }
      }
    }
  }
  
  Logger.log("Total Actual HC returned: " + actualHC);
  
  return {
    optimumHC: optimumHC,
    actualHC: actualHC,
    gap: optimumHC - actualHC,
    gapPercentage: optimumHC > 0 ? ((optimumHC - actualHC) / optimumHC * 100).toFixed(2) : 0
  };
}

/**
 * Get revenue status (mock data for now)
 */
function getRevenueStatus() {
  // This can be replaced with actual calculation from a Revenue sheet
  return {
    current: 85000000,
    target: 100000000,
    percentage: 85,
    threshold: 80000000,
    status: "On Track"
  };
}

// ============================================================
// 3. REQUEST MANAGEMENT
// ============================================================

/**
 * Generate unique Request ID
 * Format: MR-YYYY-XXXX
 */
function generateRequestID() {
  const year = new Date().getFullYear();
  const sheet = getRequestsSheet();
  const data = sheet.getDataRange().getValues();
  const headerMap = getRequestHeaderMap().map;
  const requestIDCol = headerMap['request id'] !== undefined ? headerMap['request id'] : 0;
  
  let maxNum = 1000;
  for (let i = 1; i < data.length; i++) {
    if (data[i][requestIDCol] && data[i][requestIDCol].toString().startsWith("MR-" + year)) {
      const parts = data[i][requestIDCol].toString().split("-");
      const num = parseInt(parts[2]);
      if (!isNaN(num) && num > maxNum) maxNum = num;
    }
  }
  
  return "MR-" + year + "-" + (maxNum + 1);
}

/**
 * Create new manpower request
 */
function createRequest(requestData) {
  const sheet = getRequestsSheet();
  const requestID = generateRequestID();
  const timestamp = new Date().toISOString();
  const user = getUserRole();
  
  const approvalSteps = getApprovalStepsForData(requestData);
  const approvalChain = approvalSteps.map(step => step.approver).join(",");
  const firstApprover = approvalSteps.length > 0 ? approvalSteps[0].approver : "Plant Head";
  
  const headerInfo = getRequestHeaderMap();
  const headerMap = headerInfo.map;
  const rowData = Array(headerInfo.headerRow.length).fill("");

  setRequestField(rowData, headerMap, ["request id", "id"], requestID);
  setRequestField(rowData, headerMap, ["timestamp", "date"], timestamp);
  setRequestField(rowData, headerMap, ["requestor", "requestor email", "requested by", "email"], user.email);
  setRequestField(rowData, headerMap, ["division"], requestData.division);
  setRequestField(rowData, headerMap, ["group"], requestData.group);
  setRequestField(rowData, headerMap, ["department"], requestData.department);
  setRequestField(rowData, headerMap, ["section"], requestData.section);
  setRequestField(rowData, headerMap, ["unit"], requestData.unit);
  setRequestField(rowData, headerMap, ["line"], requestData.line);
  setRequestField(rowData, headerMap, ["type", "request type"], requestData.requestType);
  setRequestField(rowData, headerMap, ["details", "detail"], buildRequestDetails(requestData));
  setRequestField(rowData, headerMap, ["status"], "Pending");
  const nextApproverSet = setRequestField(rowData, headerMap, ["next approver", "next_approver", "current approver", "current_approver", "approver level", "approver_level", "approverlevel"], firstApprover);
  if (!nextApproverSet) {
    Logger.log("WARNING: Could not find 'Current/Next Approver' column in Request sheet. Available columns: " + Object.keys(headerMap.map).join(", "));
  }
  setRequestField(rowData, headerMap, ["justification", "justification."], requestData.justification || "");
  setRequestField(rowData, headerMap, ["attached files", "attachments", "attachedfiles"], (requestData.uploadedFiles || []).map(file => file.name).join('; '));
  setRequestField(rowData, headerMap, ["category"], requestData.category || "");
  setRequestField(rowData, headerMap, ["positions"], JSON.stringify(requestData.positions));
  setRequestField(rowData, headerMap, ["approver notes", "approver_notes", "notes"], "");
  setRequestField(rowData, headerMap, ["remarks 1", "remarks1", "remark 1", "remark1"], "");
  setRequestField(rowData, headerMap, ["remarks 2", "remarks2", "remark 2", "remark2"], "");
  setRequestField(rowData, headerMap, ["remarks 3", "remarks3", "remark 3", "remark3"], "");
  setRequestField(rowData, headerMap, ["approval chain", "approval_chain"], approvalChain);
  setRequestField(rowData, headerMap, ["manpower data", "manpower_data"], JSON.stringify({
    optimumHC: requestData.optimumHC,
    actualHC: requestData.actualHC,
    gap: requestData.gap,
    isExceptional: requestData.gap <= 0,
    exceptionalJustification: requestData.exceptionalJustification || ""
  }));

  setRequestField(rowData, headerMap, ["optimum manpower hc", "optimum hc", "optimum_hc"], requestData.optimumHC);
  setRequestField(rowData, headerMap, ["actual manpower hc", "actual hc", "actual_hc"], requestData.actualHC);
  setRequestField(rowData, headerMap, ["gap"], requestData.gap);

  if (headerInfo.headerRow.length === 0) {
    throw new Error('Requests sheet missing headers. Please restore the header row.');
  }

  sheet.appendRow(rowData);
  saveApprovalHistory(requestID, []);
  
  // Log the creation
  Logger.log("Request created: " + requestID + " (Gap: " + requestData.gap + ", Exceptional: " + (requestData.gap <= 0 ? "Yes" : "No") + ")");
  
  // Send notification to first approver (Plant Head)
  try {
    const newRequest = getRequestByID(requestID);
    if (newRequest && firstApprover) {
      const plantHeadEmail = getApproverEmailByRole(firstApprover, newRequest);
      if (plantHeadEmail) {
        sendNewRequestNotification(newRequest, firstApprover, plantHeadEmail);
        Logger.log("Notification sent to Plant Head: " + plantHeadEmail);
      } else {
        Logger.log("Could not find Plant Head email for notification. First Approver: " + firstApprover);
      }
    }
  } catch (notifError) {
    Logger.log("Error sending Plant Head notification: " + notifError);
    // Don't fail the request creation if notification fails
  }
  
  return {
    success: true,
    requestID: requestID,
    message: "Request submitted successfully"
  };
}

/**
 * Get all requests for current user
 */
function getMyRequests() {
  const sheet = getRequestsSheet();
  const data = sheet.getDataRange().getValues();
  const email = getCurrentUserEmail();
  const headerMap = getRequestHeaderMap().map;

  const requests = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowEmail = getRequestField(row, headerMap, ["requestor", "requestor email", "requested by", "email"], "").toString().trim();
    if (rowEmail === email) {
      requests.push(buildRequestResponse(row, i + 1, headerMap));
    }
  }

  return requests.sort((a, b) => new Date(b.createdDate) - new Date(a.createdDate));
}

/**
 * Get all requests pending approval
 */
function getPendingApprovals() {
  const sheet = getRequestsSheet();
  const data = sheet.getDataRange().getValues();
  const email = getCurrentUserEmail();
  const user = getUserRole();
  const headerMap = getRequestHeaderMap().map;

  const approvals = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const request = buildRequestResponse(row, i + 1, headerMap);
    const review = getApprovalReviewState(request, email, user.role);
    if (review) {
      request.reviewState = review.reviewState;
      request.canAct = review.canAct;
      request.matchedLevels = review.matchedLevels;
      approvals.push(request);
    }
  }
  
  return approvals.sort((a, b) => new Date(b.createdDate) - new Date(a.createdDate));
}

/**
 * Get request by ID
 */
function getRequestByID(requestID) {
  const sheet = getRequestsSheet();
  const data = sheet.getDataRange().getValues();
  const headerMap = getRequestHeaderMap().map;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowRequestID = getRequestField(row, headerMap, ["request id", "id"], "").toString().trim();
    if (rowRequestID === requestID) {
      return buildRequestResponse(row, i + 1, headerMap);
    }
  }
  
  return null;
}

// ============================================================
// 4. APPROVAL WORKFLOW
// ============================================================

/**
 * Approval routing configuration
 */
function getApprovalRoute(requestType, category) {
  if (requestType === "Replacement") {
    return ["Plant Head", "Legal", "Corporate HROD"];
  } else if (requestType === "Seasonal") {
    return category === "Production" 
      ? ["Plant Head", "Manufacturing Director", "BU Head", "Legal", "Corporate HROD"]
      : ["Plant Head", "BU Head", "Legal", "Corporate HROD"];
  } else if (requestType === "Regular") {
    return category === "Production"
      ? ["Plant Head", "Manufacturing Director", "BU Head", "Legal", "Corporate HROD"]
      : ["Plant Head", "BU Head", "Legal", "Corporate HROD"];
  }
  
  return ["Plant Head"];
}

/**
 * Approve or reject request
 */
function approveRequest(requestID, action, comments) {
  const request = getRequestByID(requestID);
  if (!request) return { success: false, message: "Request not found" };
  
  const sheet = getRequestsSheet();
  const headerMap = getRequestHeaderMap().map;
  const rowNumber = request.rowNumber;
  const user = getUserRole();
  const trimmedComments = (comments || "").toString().trim();
  
  if (action === "Disapprove" && !trimmedComments) {
    return { success: false, message: "Disapproval requires mandatory comments" };
  }
  
  // If approval chain contains "Legal" or "Corporate HROD", we're in Step 5 - prioritize stored chain
  const isInStep5 = request.approvalChain && (request.approvalChain.toLowerCase().indexOf("legal") !== -1);
  
  let approvalSteps = [];
  let approvalChain = [];
  
  if (isInStep5 && request.approvalChain) {
    // Step 5: Use the stored approval chain directly
    approvalChain = request.approvalChain.split(",").map(s => s.trim());
    approvalSteps = approvalChain.map(label => ({ label: label, approver: label }));
  } else {
    // Steps 1-2: Build approvalSteps from request data
    approvalSteps = getApprovalStepsForRequest(request);
    approvalChain = approvalSteps.length > 0
      ? approvalSteps.map(step => step.approver)
      : request.approvalChain ? request.approvalChain.split(",").map(s => s.trim()) : [];
  }
  
  const manpowerData = request.manpowerData || {};
  const isExceptional = manpowerData.isExceptional === true || manpowerData.isExceptional === "true";
  const currentIndex = findApprovalStepIndex(approvalSteps, request.currentApprover);
  const currentLevel = currentIndex >= 0 && approvalSteps[currentIndex]
    ? approvalSteps[currentIndex].label
    : request.currentApprover;
  const generalNotesIdx = getRequestColumnIndex(headerMap, ["approver notes", "approver_notes", "notes"]);
  const remarkText = buildApprovalRemark(action, trimmedComments);
  
  let newStatus = request.status;
  let nextApprover = "";
  let nextApproverLabel = "";
  let approvedLevels = [];
  let remarkIndexesToUpdate = [];
  let actedLevels = [];
  
  if (action === "Approve") {
    let nextIndex = currentIndex < 0 ? 0 : currentIndex;

    // Process current level and any identical subsequent levels in the chain
    if (nextIndex < approvalChain.length) {
      const currentApproverValue = approvalChain[nextIndex];
      const currentLabelValue = approvalSteps[nextIndex] ? approvalSteps[nextIndex].label : "";
      
      do {
        approvedLevels.push(approvalSteps[nextIndex] ? approvalSteps[nextIndex].label : approvalChain[nextIndex]);
        remarkIndexesToUpdate.push(nextIndex);
        nextIndex++;
      } while (nextIndex < approvalChain.length && 
               (approvalChain[nextIndex] === currentApproverValue || 
                (approvalSteps[nextIndex] && approvalSteps[nextIndex].label === currentLabelValue) ||
                approvalChain[nextIndex] === request.currentApprover ||
                (approvalSteps[nextIndex] && approvalSteps[nextIndex].label === request.currentApprover)));
    }

    actedLevels = approvedLevels.length > 0 ? approvedLevels.slice() : [currentLevel];

    if (nextIndex < approvalChain.length) {
      nextApprover = approvalChain[nextIndex];
      nextApproverLabel = approvalSteps[nextIndex] ? approvalSteps[nextIndex].label : approvalChain[nextIndex];
      newStatus = "Pending";
    } else {
      newStatus = "Approved";
      nextApprover = "Completed";
      nextApproverLabel = "Completed";
    }
  } else if (action === "Hold") {
    newStatus = "On Hold";
    nextApprover = request.currentApprover; // Stay with current approver
    nextApproverLabel = currentLevel;
    if (currentIndex >= 0) {
      remarkIndexesToUpdate.push(currentIndex);
    }
    actedLevels = [currentLevel];
  } else if (action === "Disapprove") {
    newStatus = "Disapproved";
    nextApprover = "Requestor";
    if (currentIndex >= 0) {
      remarkIndexesToUpdate.push(currentIndex);
    }
    actedLevels = [currentLevel];
  }
  
  const statusIdx = getRequestColumnIndex(headerMap, ["status"]);
  const nextApproverColumnNames = ["next approver", "next_approver", "current approver", "current_approver", "approver level", "approver_level", "approverlevel"];
  const approvalConfig = getApprovalColumnFieldNames(currentLevel, currentIndex);
  const approvalColumnIdx = getRequestColumnIndex(headerMap, approvalConfig.status);
  const approverEmailColumnIdx = getRequestColumnIndex(headerMap, approvalConfig.approver);

  if (statusIdx >= 0) {
    sheet.getRange(rowNumber, statusIdx + 1).setValue(newStatus);
  }
  if (generalNotesIdx >= 0) {
    sheet.getRange(rowNumber, generalNotesIdx + 1).setValue(remarkText);
  }
  
  // Update Next Approver - update multiple potential columns
  nextApproverColumnNames.forEach(colName => {
    const idx = headerMap[colName.toLowerCase()];
    if (idx !== undefined) {
      sheet.getRange(rowNumber, idx + 1).setValue(nextApprover);
    }
  });
  
  // Update the approval column for the current step with the decided action
  if (approvalColumnIdx >= 0) {
    const approvalDecision = action === "Approve" 
      ? "Approved" 
      : action === "Disapprove" 
        ? "Disapproved" 
        : action === "Hold" 
          ? "On Hold" 
          : action;
    sheet.getRange(rowNumber, approvalColumnIdx + 1).setValue(approvalDecision);
  }

  // Update the approver email column
  if (approverEmailColumnIdx >= 0 && action === "Approve") {
    sheet.getRange(rowNumber, approverEmailColumnIdx + 1).setValue(user.email);
  }

  remarkIndexesToUpdate.forEach(stepIndex => {
    const stepLabel = approvalSteps[stepIndex] ? approvalSteps[stepIndex].label : "";
    const remarkColumnIdx = getRequestColumnIndex(headerMap, getApprovalRemarkFieldNames(stepLabel, stepIndex));
    if (remarkColumnIdx >= 0) {
      sheet.getRange(rowNumber, remarkColumnIdx + 1).setValue(remarkText);
    }
  });

  const approvalHistory = appendApprovalHistoryEntry(requestID, {
    action: action,
    level: currentLevel,
    levels: actedLevels,
    comment: trimmedComments,
    remark: remarkText,
    timestamp: new Date().toISOString(),
    actorEmail: user.email,
    actorName: user.name,
    actorRole: user.role,
    approver: request.currentApprover,
    statusAfter: newStatus,
    nextApprover: nextApprover,
    nextApproverLabel: nextApproverLabel
  });
  
  // Send notifications
  if (action === "Approve") {
    // ALWAYS notify next approver if there's one waiting
    if (nextApprover && nextApprover !== "Completed") {
      const nextApproverEmail = getApproverEmailByRole(nextApproverLabel, request);
      if (nextApproverEmail) {
        sendApprovalNotificationToNextApprover(request, action, trimmedComments, nextApproverLabel, nextApproverEmail, isExceptional, {
          currentLevel: currentLevel,
          approvedLevels: approvedLevels,
          nextApproverLabel: nextApproverLabel,
          newStatus: newStatus,
          approvalHistory: approvalHistory
        });
        Logger.log("Sent notification to next approver: " + nextApproverEmail + " (" + nextApproverLabel + ")");
      } else {
        Logger.log("WARNING: No email found for next approver role: " + nextApproverLabel);
      }
    }
    
    // Notify requestor if moving to Step 3 (BU Head final approval before Step 3)
    if (newStatus === "Approved") {
      sendApprovalNotificationToRequestor(request, action, trimmedComments, newStatus, currentLevel, approvedLevels, isExceptional);
      Logger.log("Sent approval notification to requestor: " + request.email);
    }
  } else if (action === "Disapprove") {
    // Always notify requestor of disapproval
    sendApprovalNotificationToRequestor(request, action, trimmedComments, newStatus, currentLevel, approvedLevels, isExceptional);
    Logger.log("Sent disapproval notification to requestor: " + request.email);
  } else if (action === "Hold") {
    // Notify requestor when on hold
    sendApprovalNotificationToRequestor(request, action, trimmedComments, newStatus, currentLevel, approvedLevels, isExceptional);
    Logger.log("Sent hold notification to requestor: " + request.email);
  }
  
  // STEP 6: If this is Step 5 final approval by Corporate HROD, trigger recruitment notification and update status
  if (action === "Approve" && newStatus === "Approved" && 
      currentLevel && currentLevel.toLowerCase().indexOf("corporate") !== -1) {
    
    // Update status to "Sourcing In Progress"
    const sourcingStatusIdx = getRequestColumnIndex(headerMap, ["status"]);
    if (sourcingStatusIdx >= 0) {
      sheet.getRange(rowNumber, sourcingStatusIdx + 1).setValue("Sourcing In Progress");
      newStatus = "Sourcing In Progress";
    }
    
    // Notify Recruitment Department
    notifyRecruitmentDepartment(request);
    
    // Log this milestone in approval history
    appendApprovalHistoryEntry(requestID, {
      action: "Step 6: Recruitment Notified",
      level: "System",
      levels: ["System"],
      comment: "Final approval received. Recruitment department notified. Status: Sourcing In Progress",
      timestamp: new Date().toISOString(),
      statusAfter: "Sourcing In Progress"
    });
  }
  
  return {
    success: true,
    message: "Request " + action.toLowerCase() + " successfully",
    requestID: requestID,
    newStatus: newStatus,
    nextApprover: nextApprover,
    nextApproverLabel: nextApproverLabel,
    currentLevel: currentLevel,
    approvedLevels: approvedLevels
  };
}

/**
 * Proceed to Step 3: Save Step 3 data and move request to next workflow stage
 */
function proceedToStep3(requestID) {
  const request = getRequestByID(requestID);
  if (!request) return { success: false, message: "Request not found" };
  
  const status = (request.status || "").toString().trim();
  if (status !== "Approved") {
    return { success: false, message: "Only approved requests can proceed to Step 3. Current status: " + status };
  }
  
  const sheet = getRequestsSheet();
  const headerMap = getRequestHeaderMap().map;
  const rowNumber = request.rowNumber;
  
  // Update status to indicate Step 3 pending (requestor is filling in type selection)
  const statusIdx = getRequestColumnIndex(headerMap, ["status"]);
  if (statusIdx >= 0) {
    sheet.getRange(rowNumber, statusIdx + 1).setValue("Step 3: Type Selection");
  }
  
  // Store in approval history that requestor is proceeding to Step 3
  appendApprovalHistoryEntry(requestID, {
    action: "Proceed to Step 3",
    level: "Requestor",
    levels: ["Requestor"],
    comment: "Requestor proceeding to Step 3: Select Request Type",
    timestamp: new Date().toISOString(),
    actorEmail: getCurrentUserEmail(),
    actorName: getUserRole().name,
    actorRole: "Requestor",
    statusAfter: "Step 3: Type Selection"
  });
  
  return {
    success: true,
    message: "Proceeding to Step 3: Select Request Type",
    requestID: requestID,
    newStatus: "Step 3: Type Selection"
  };
}

/**
 * Save Step 3 data (request type selection and type-specific details)
 */
function saveStep3Data(requestID, step3Data) {
  const request = getRequestByID(requestID);
  if (!request) return { success: false, message: "Request not found" };
  
  const sheet = getRequestsSheet();
  const headerMap = getRequestHeaderMap().map;
  const rowNumber = request.rowNumber;
  
  // Update request type (Replacement, Seasonal, Regular)
  const typeIdx = getRequestColumnIndex(headerMap, ["type", "request type"]);
  if (typeIdx >= 0) {
    sheet.getRange(rowNumber, typeIdx + 1).setValue(step3Data.requestType || "");
  }
  
  // Update positions/details based on request type
  const positionsIdx = getRequestColumnIndex(headerMap, ["positions"]);
  if (positionsIdx >= 0 && step3Data.positions) {
    sheet.getRange(rowNumber, positionsIdx + 1).setValue(JSON.stringify(step3Data.positions));
  }
  
  // Store replacement-specific data if applicable
  if (step3Data.requestType === "Replacement" && step3Data.replacementData) {
    // Store replacement employee info in a custom field or manpower data
    const manpowerData = request.manpowerData || {};
    manpowerData.replacementEmployee = step3Data.replacementData;
    const mpIdx = getRequestColumnIndex(headerMap, ["manpower data", "manpower_data"]);
    if (mpIdx >= 0) {
      sheet.getRange(rowNumber, mpIdx + 1).setValue(JSON.stringify(manpowerData));
    }
  }
  
  // Update status to Step 4
  const statusIdx = getRequestColumnIndex(headerMap, ["status"]);
  if (statusIdx >= 0) {
    sheet.getRange(rowNumber, statusIdx + 1).setValue("Step 4: Review & Submit");
  }
  
  appendApprovalHistoryEntry(requestID, {
    action: "Step 3 Data Saved",
    level: "Requestor",
    levels: ["Requestor"],
    comment: "Request type: " + (step3Data.requestType || "Not specified"),
    timestamp: new Date().toISOString(),
    actorEmail: getCurrentUserEmail(),
    actorName: getUserRole().name,
    actorRole: "Requestor",
    statusAfter: "Step 4: Review & Submit"
  });
  
  return {
    success: true,
    message: "Step 3 data saved successfully",
    requestID: requestID,
    newStatus: "Step 4: Review & Submit"
  };
}

/**
 * Submit request for Step 5 approval (Legal Team → Corporate HROD)
 */
function submitForStep5Approval(requestID) {
  const request = getRequestByID(requestID);
  if (!request) return { success: false, message: "Request not found" };
  
  const sheet = getRequestsSheet();
  const headerMap = getRequestHeaderMap().map;
  const rowNumber = request.rowNumber;
  
  // Get Step 5 approval chain (Legal Team, Corporate HROD)
  const step5Chain = getStep5ApprovalChain();
  if (!step5Chain || step5Chain.length === 0) {
    return { success: false, message: "Step 5 approval chain not configured" };
  }
  
  // Set first approver in Step 5 chain
  const firstApprover = step5Chain[0];
  const currentApproverIdx = getRequestColumnIndex(headerMap, ["next approver", "next_approver", "current approver", "current_approver", "approver level", "approver_level", "approverlevel"]);
  if (currentApproverIdx >= 0) {
    sheet.getRange(rowNumber, currentApproverIdx + 1).setValue(firstApprover);
  } else {
    Logger.log("WARNING: Could not find 'Current/Next Approver' column in submitForStep5Approval. Available columns: " + Object.keys(headerMap).join(", "));
  }
  
  // Update status to Pending (now in Step 5 approval)
  const statusIdx = getRequestColumnIndex(headerMap, ["status"]);
  if (statusIdx >= 0) {
    sheet.getRange(rowNumber, statusIdx + 1).setValue("Pending");
  }
  
  // Update approval chain
  const chainIdx = getRequestColumnIndex(headerMap, ["approval chain", "approval_chain"]);
  if (chainIdx >= 0) {
    sheet.getRange(rowNumber, chainIdx + 1).setValue(step5Chain.join(","));
  }
  
  appendApprovalHistoryEntry(requestID, {
    action: "Submitted for Step 5 Approval",
    level: "Requestor",
    levels: ["Requestor"],
    comment: "Request submitted for final approval chain: " + step5Chain.join(" → "),
    timestamp: new Date().toISOString(),
    actorEmail: getCurrentUserEmail(),
    actorName: getUserRole().name,
    actorRole: "Requestor",
    statusAfter: "Pending",
    step: 5,
    approvalChain: step5Chain
  });
  
  // Send notification to first approver in Step 5 (Legal Team)
  const firstApproverEmail = getApproverEmailByRole(firstApprover, request);
  if (firstApproverEmail) {
    sendStep5ApprovalNotification(request, firstApprover, firstApproverEmail);
  }
  
  return {
    success: true,
    message: "Request submitted for Step 5 approval",
    requestID: requestID,
    newStatus: "Pending",
    approvalChain: step5Chain,
    nextApprover: firstApprover
  };
}

/**
 * Get Step 5 approval chain (only Legal + Corporate HROD)
 * Note: Uses naming consistent with getApprovalRoute
 */
function getStep5ApprovalChain() {
  return ["Legal", "Corporate HROD"];
}

/**
 * Send notification to Step 5 approver
 */
function sendStep5ApprovalNotification(request, approverLabel, approverEmail) {
  const subject = `ACTION REQUIRED: Manpower Request - Step 5 Final Approval - ${approverLabel} - MR ID: ${request.requestID}`;
  const systemUrl = getSystemUrl();
  
  const htmlBody = `
    <div style="font-family: Arial, sans-serif; color: #0f172a; line-height: 1.6; max-width: 640px; margin: 0 auto; border: 1px solid #e2e8f0; border-radius: 16px; overflow: hidden; background: #ffffff;">
      <div style="background: #0f172a; color: #ffffff; padding: 20px 24px;">
        <div style="font-size: 12px; letter-spacing: 0.08em; text-transform: uppercase; opacity: 0.75;">Manpower Request System - Step 5</div>
        <h2 style="margin: 8px 0 0; font-size: 24px;">Final Approval Required</h2>
      </div>
      <div style="padding: 24px;">
        <p style="margin: 0 0 12px;">Greetings <strong>${escapeHtml(approverLabel)}</strong>,</p>
        <p style="margin: 0 0 12px;">A manpower request requires your review and approval as part of the final approval chain (Step 5).</p>
        
        <div style="margin: 20px 0; padding: 16px; border-radius: 12px; background: #f0f9ff; border: 1px solid #0284c7;">
          <p style="margin: 0 0 10px;"><strong>Request Details:</strong></p>
          <ul style="margin: 0 0 10px 18px; padding: 0;">
            <li><strong>Request ID:</strong> ${escapeHtml(request.requestID)}</li>
            <li><strong>Department:</strong> ${escapeHtml(request.department)}</li>
            <li><strong>Request Type:</strong> ${escapeHtml(request.requestType)}</li>
            <li><strong>Category:</strong> ${escapeHtml(request.category)}</li>
            <li><strong>Submitted by:</strong> ${escapeHtml(request.requestedBy)}</li>
          </ul>
        </div>
        
        <p style="margin: 0 0 12px;"><strong>Approval Status:</strong></p>
        <p style="margin: 0 0 16px;">This request has already been approved by Plant Head and BU Head in the initial approval chain. It now requires your review for final approval.</p>
        
        ${systemUrl ? `<div style="margin: 24px 0 8px;"><a href="${escapeHtml(systemUrl)}" style="display: inline-block; background: #1d4ed8; color: #ffffff; text-decoration: none; padding: 12px 20px; border-radius: 10px; font-weight: 600;">Open Manpower Request System</a></div>` : ''}
        
        <hr style="border: 0; border-top: 1px solid #e2e8f0; margin: 24px 0 16px;">
        <p style="margin: 0; color: #475569;">Please log in to the system to review and approve/disapprove this request.</p>
      </div>
    </div>
  `;
  
  try {
    GmailApp.sendEmail(approverEmail, subject, "", {
      htmlBody: htmlBody,
      name: "Manpower Request System"
    });
    Logger.log("Step 5 notification sent to: " + approverEmail);
  } catch (e) {
    Logger.log("Error sending Step 5 notification: " + e);
  }
}

/**
 * Check if request is in Step 5 approval (determine by approval chain or status)
 */
function isStep5Approval(request) {
  const approvalChain = (request.approvalChain || "").toString().toLowerCase();
  return approvalChain.indexOf("legal") !== -1 || 
         approvalChain.indexOf("corporate hrod") !== -1;
}

/**
 * Notify Recruitment Department of approved request (Step 6 completion)
 */
function notifyRecruitmentDepartment(request) {
  const recruitmentEmail = getApproverEmailByRole("Recruitment", request);
  
  if (!recruitmentEmail) {
    Logger.log("Recruitment email not configured in Users sheet");
    return { success: false, message: "Recruitment email not configured" };
  }
  
  const subject = `✓ New Recruitment Request Ready - Request ID: ${request.requestID} - Status: Sourcing In Progress`;
  const systemUrl = getSystemUrl();
  
  let positionsHtml = "<ul>";
  if (request.positions && request.positions.length > 0) {
    request.positions.forEach(pos => {
      positionsHtml += `<li><strong>${escapeHtml(pos.position)}</strong> - Headcount: ${pos.headcount}${pos.type ? ' (' + pos.type + ')' : ''}${pos.startDate ? ' (From: ' + pos.startDate + ')' : ''}${pos.endDate ? ' To: ' + pos.endDate : ''}</li>`;
    });
  }
  positionsHtml += "</ul>";
  
  const htmlBody = `
    <div style="font-family: Arial, sans-serif; color: #0f172a; line-height: 1.6; max-width: 640px; margin: 0 auto; border: 1px solid #e2e8f0; border-radius: 16px; overflow: hidden; background: #ffffff;">
      <div style="background: linear-gradient(135deg, #10b981 0%, #059669 100%); color: #ffffff; padding: 20px 24px;">
        <div style="font-size: 12px; letter-spacing: 0.08em; text-transform: uppercase; opacity: 0.9;">Recruitment Department</div>
        <h2 style="margin: 8px 0 0; font-size: 24px;">✓ New Recruitment Request - Sourcing In Progress</h2>
      </div>
      <div style="padding: 24px;">
        <p style="margin: 0 0 12px;">Good news! A manpower request has been approved and is now ready for sourcing.</p>
        
        <div style="margin: 20px 0; padding: 16px; border-radius: 12px; background: #f0fdf4; border: 1px solid #86efac;">
          <p style="margin: 0 0 12px;"><strong>Request Summary:</strong></p>
          <ul style="margin: 0 0 10px 18px; padding: 0;">
            <li><strong>Request ID:</strong> ${escapeHtml(request.requestID)}</li>
            <li><strong>Department:</strong> ${escapeHtml(request.department)}</li>
            <li><strong>Request Type:</strong> ${escapeHtml(request.requestType)}</li>
            <li><strong>Category:</strong> ${escapeHtml(request.category)}</li>
            <li><strong>Requested By:</strong> ${escapeHtml(request.requestedBy)}</li>
          </ul>
        </div>
        
        <p style="margin: 0 0 12px;"><strong>Positions to Fill:</strong></p>
        ${positionsHtml}
        
        <p style="margin: 0 0 12px;"><strong>Request Details:</strong></p>
        <p style="margin: 0 0 16px;">${escapeHtml(request.details)}</p>
        
        ${systemUrl ? `<div style="margin: 24px 0 8px;"><a href="${escapeHtml(systemUrl)}" style="display: inline-block; background: #10b981; color: #ffffff; text-decoration: none; padding: 12px 20px; border-radius: 10px; font-weight: 600;">View Request Details</a></div>` : ''}
        
        <hr style="border: 0; border-top: 1px solid #e2e8f0; margin: 24px 0 16px;">
        <p style="margin: 0; color: #475569;">Please begin the sourcing process for the positions listed above. The request status has been updated to "Sourcing In Progress".</p>
      </div>
    </div>
  `;
  
  try {
    GmailApp.sendEmail(recruitmentEmail, subject, "", {
      htmlBody: htmlBody,
      name: "Manpower Request System"
    });
    Logger.log("Recruitment notification sent to: " + recruitmentEmail);
    return { success: true, message: "Recruitment notification sent" };
  } catch (e) {
    Logger.log("Error sending recruitment notification: " + e);
    return { success: false, message: "Error sending notification: " + e };
  }
}

/**
 * Combined function: Save Step 3 data and submit for Step 5 approval
 */
function saveStep3DataAndSubmit(requestID, step3Data) {
  // First save the Step 3 data
  const saveResult = saveStep3Data(requestID, step3Data);
  if (!saveResult.success) {
    return saveResult;
  }
  
  // Then submit for Step 5 approval
  const submitResult = submitForStep5Approval(requestID);
  if (!submitResult.success) {
    return submitResult;
  }
  
  return {
    success: true,
    message: "Request saved and submitted for final approval successfully",
    requestID: requestID,
    newStatus: submitResult.newStatus,
    approvalChain: submitResult.approvalChain,
    nextApprover: submitResult.nextApprover
  };
}

/**
 * Legacy: Mark an approved request as completed (deprecated - now uses Step 3-6 workflow)
 * Keeping for backward compatibility
 */
function completeApprovedRequest(requestID) {
  // Instead of completing, transition to Step 3
  return proceedToStep3(requestID);
}

/**
 * Send approval notification
 */
function getApproverEmailFromMasterList(request, approverRole) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAMES.ORG_CHART);
    if (!sheet) {
      Logger.log("Optimum Manpower HC sheet not found");
      return null;
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find column indices
    let divisionIdx = -1, groupIdx = -1, deptIdx = -1, sectionIdx = -1, lineIdx = -1;
    let plantHeadIdx = -1, buHeadIdx = -1, managingDirectorIdx = -1;
    
    for (let j = 0; j < headers.length; j++) {
      const header = (headers[j] || "").toString().toLowerCase().trim();
      if (header === "division") divisionIdx = j;
      else if (header === "group") groupIdx = j;
      else if (header === "department") deptIdx = j;
      else if (header === "section") sectionIdx = j;
      else if (header === "line") lineIdx = j;
      else if (header === "plant head") plantHeadIdx = j;
      else if (header === "bu head") buHeadIdx = j;
      else if (header === "managing director") managingDirectorIdx = j;
    }
    
    if (plantHeadIdx < 0 || buHeadIdx < 0) {
      Logger.log("Plant Head or BU Head columns not found in Optimum Manpower HC sheet");
      return null;
    }
    
    // Normalize approver role
    const normalizedRole = (approverRole || "").toString().toLowerCase().trim();
    
    // Search for matching row
    const reqDivision = (request.division || "").toString().toLowerCase().trim();
    const reqGroup = (request.group || "").toString().toLowerCase().trim();
    const reqDept = (request.department || "").toString().toLowerCase().trim();
    const reqLine = (request.line || "").toString().toLowerCase().trim();
    
    for (let i = 1; i < data.length; i++) {
      const division = (data[i][divisionIdx] || "").toString().toLowerCase().trim();
      const group = groupIdx >= 0 ? (data[i][groupIdx] || "").toString().toLowerCase().trim() : "";
      const dept = deptIdx >= 0 ? (data[i][deptIdx] || "").toString().toLowerCase().trim() : "";
      const line = lineIdx >= 0 ? (data[i][lineIdx] || "").toString().toLowerCase().trim() : "";
      
      // Match by division and department (best match)
      if (division === reqDivision && (dept === reqDept || dept === "")) {
        let approverEmail = "";
        
        // Map role names to master list columns
        if (normalizedRole.indexOf("plant") !== -1) {
          // Plant Head
          approverEmail = (data[i][plantHeadIdx] || "").toString().trim();
        } else if (normalizedRole.indexOf("bu") !== -1 || normalizedRole.indexOf("business") !== -1) {
          // BU Head
          approverEmail = (data[i][buHeadIdx] || "").toString().trim();
        } else if (normalizedRole.indexOf("managing") !== -1) {
          // Managing Director
          approverEmail = (data[i][managingDirectorIdx] || "").toString().trim();
        } else if (normalizedRole.indexOf("corporate") !== -1 || normalizedRole.indexOf("hrod") !== -1) {
          // Corporate HROD - map to Managing Director column
          approverEmail = (data[i][managingDirectorIdx] || "").toString().trim();
        } else if (normalizedRole.indexOf("legal") !== -1) {
          // Legal - for now, use Managing Director or fall back to Users sheet
          approverEmail = (data[i][managingDirectorIdx] || "").toString().trim();
        }
        
        if (approverEmail) {
          Logger.log("Found approver email from master list - Role: " + approverRole + ", Email: " + approverEmail + " (Div: " + division + ", Dept: " + dept + ")");
          return approverEmail;
        }
      }
    }
    
    Logger.log("No matching approver found in Optimum Manpower HC sheet for: Division=" + reqDivision + ", Department=" + reqDept + ", Role=" + approverRole + ". Will try Users sheet as fallback.");
    return null;
  } catch (e) {
    Logger.log("Error getting approver email from master list: " + e);
    return null;
  }
}

function getApproverEmailByRole(roleLabel, request = null) {
  const trimmedRole = (roleLabel || "").toString().trim();
  if (trimmedRole.indexOf("@") !== -1) {
    return trimmedRole;
  }
  const normalizedLabel = trimmedRole.toLowerCase();
  
  // Hardcoded fallback for system roles (can be changed in Users sheet later)
  const systemRoleEmails = {
    "legal": "corporate.training@uratex.com.ph",
    "corporate hrod": "michelleann.delacerna@uratex.com.ph",
    "recruitment": "corporate.training@uratex.com.ph" // can be updated
  };
  
  // First, try to get from the Optimum Manpower HC master list if request is provided
  if (request) {
    const masterListEmail = getApproverEmailFromMasterList(request, roleLabel);
    if (masterListEmail) {
      return masterListEmail;
    }
  }
  
  // Fallback to Users sheet
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAMES.USERS);
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      
      // Log all available roles for debugging
      const availableRoles = [];
      for (let i = 1; i < data.length; i++) {
        const role = (data[i][1] || "").toString().trim();
        if (role) {
          availableRoles.push(role);
        }
      }
      Logger.log("Looking for role '" + roleLabel + "' in Users sheet. Available roles: " + JSON.stringify(availableRoles));
      
      // First pass: exact match or partial match
      for (let i = 1; i < data.length; i++) {
        const email = (data[i][0] || "").toString().trim();
        const role = (data[i][1] || "").toString().trim().toLowerCase();
        
        if (!email) continue;
        
        // Exact match
        if (role === normalizedLabel) {
          Logger.log("Found exact match for role '" + roleLabel + "': " + email);
          return email;
        }
        
        // Partial match
        if (role.indexOf(normalizedLabel) !== -1 || normalizedLabel.indexOf(role) !== -1) {
          Logger.log("Found partial match for role '" + roleLabel + "': " + email + " (role: " + role + ")");
          return email;
        }
      }
    }
    
    // If not found in Users sheet, check system role fallbacks
    if (systemRoleEmails[normalizedLabel]) {
      Logger.log("Using fallback email for system role '" + roleLabel + "': " + systemRoleEmails[normalizedLabel]);
      return systemRoleEmails[normalizedLabel];
    }
    
    Logger.log("WARNING: No email found for role: '" + roleLabel + "'");
    Logger.log("To fix this, add a row to the Users sheet with email and role name matching: " + roleLabel);
    return null;
  } catch (e) {
    Logger.log("Error getting approver email: " + e);
    
    // Try system role fallback if Users sheet access fails
    if (systemRoleEmails[normalizedLabel]) {
      Logger.log("Using fallback email for role '" + roleLabel + "' due to error: " + systemRoleEmails[normalizedLabel]);
      return systemRoleEmails[normalizedLabel];
    }
    return null;
  }
}

function escapeHtml(value) {
  return (value || "").toString()
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function getSystemUrl() {
  try {
    return ScriptApp.getService().getUrl() || "";
  } catch (e) {
    Logger.log("Unable to resolve web app URL: " + e);
    return "";
  }
}

function buildApprovalEmailHtml(request, action, comments, approvedLevels, currentLevel, nextApproverLabel, nextApprover, isExceptional) {
  const subjectAction = action === "Approve" ? "Approved" : action === "Disapprove" ? "Disapproved" : "On Hold";
  const systemUrl = getSystemUrl();
  const buttonHtml = systemUrl
    ? `<div style="margin: 24px 0 8px;"><a href="${escapeHtml(systemUrl)}" style="display: inline-block; background: #1d4ed8; color: #ffffff; text-decoration: none; padding: 12px 20px; border-radius: 10px; font-weight: 600;">Open Manpower Request System</a></div>`
    : "";
  
  let htmlBody = `
    <div style="font-family: Arial, sans-serif; color: #0f172a; line-height: 1.6; max-width: 640px; margin: 0 auto; border: 1px solid #e2e8f0; border-radius: 16px; overflow: hidden; background: #ffffff;">
      <div style="background: #0f172a; color: #ffffff; padding: 20px 24px;">
        <div style="font-size: 12px; letter-spacing: 0.08em; text-transform: uppercase; opacity: 0.75;">Manpower Request System</div>
        <h2 style="margin: 8px 0 0; font-size: 24px;">Manpower Request ${escapeHtml(subjectAction)}</h2>
      </div>
      <div style="padding: 24px;">
        <p style="margin: 0 0 12px;"><strong>Request ID:</strong> ${escapeHtml(request.requestID)}</p>
        <p style="margin: 0 0 12px;"><strong>Status:</strong> ${escapeHtml(subjectAction)}</p>
        <p style="margin: 0 0 12px;"><strong>Department:</strong> ${escapeHtml(request.department)}</p>
        <p style="margin: 0 0 12px;"><strong>Request Type:</strong> ${escapeHtml(request.requestType)}</p>
        <p style="margin: 0 0 12px;"><strong>Category:</strong> ${escapeHtml(request.category)}</p>
  `;
  
  // Add exceptional case info if applicable
  if (isExceptional && request.manpowerData) {
    const mpData = request.manpowerData;
    htmlBody += `
        <div style="margin: 20px 0; padding: 16px; border-radius: 12px; background: #fef2f2; border: 1px solid #fecaca;">
          <h3 style="margin: 0 0 12px; color: #b91c1c;">Exceptional Case Notice</h3>
          <p style="margin: 0 0 10px;">This request was flagged as exceptional because:</p>
          <ul style="margin: 0 0 10px 18px; padding: 0;">
            <li><strong>Optimum HC:</strong> ${escapeHtml(mpData.optimumHC)}</li>
            <li><strong>Actual HC:</strong> ${escapeHtml(mpData.actualHC)}</li>
            <li><strong>Gap:</strong> ${mpData.gap <= 0 ? 'No gap or surplus' : 'Has gap'}</li>
          </ul>
          ${mpData.exceptionalJustification ? `<p style="margin: 0;"><strong>Your Justification:</strong> ${escapeHtml(mpData.exceptionalJustification)}</p>` : ''}
        </div>
    `;
  }
  
  if (comments) {
    htmlBody += `<p style="margin: 0 0 8px;"><strong>Comments from Approver:</strong></p><p style="margin: 0 0 16px;">${escapeHtml(comments)}</p>`;
  }
  
  if (action === "Approve") {
    const levelText = approvedLevels.length > 0 ? approvedLevels.join(" / ") : currentLevel;
    htmlBody += `<p style="margin: 0 0 12px;"><strong>Approved At:</strong> ${escapeHtml(levelText)}</p>`;
    if (nextApprover && nextApprover !== "Completed") {
      htmlBody += `<p style="margin: 0 0 12px;"><strong>Next Approval Level:</strong> ${escapeHtml(nextApproverLabel)}</p>`;
    }
  } else if (action === "Disapprove") {
    htmlBody += `<p style="margin: 0 0 12px;"><strong>Disapproved At:</strong> ${escapeHtml(currentLevel)}</p>`;
  } else if (action === "Hold") {
    htmlBody += `<p style="margin: 0 0 12px;"><strong>On Hold At:</strong> ${escapeHtml(currentLevel)}</p>`;
  }
  
  htmlBody += `
        ${buttonHtml}
        <hr style="border: 0; border-top: 1px solid #e2e8f0; margin: 24px 0 16px;">
        <p style="margin: 0; color: #475569;">Please log in to the system to review or take further action.</p>
      </div>
    </div>
  `;
  
  return htmlBody;
}

function sendApprovalNotificationToNextApprover(request, action, comments, nextApproverLabel, nextApproverEmail, isExceptional, context) {
  const approvedLevels = context && context.approvedLevels ? context.approvedLevels : [];
  const currentLevel = context && context.currentLevel ? context.currentLevel : request.currentApprover;
  const levelText = approvedLevels.length > 0 ? approvedLevels.join(" / ") : currentLevel;
  const subject = `ACTION REQUIRED: Manpower Request Pending Your Approval - ${nextApproverLabel} - MR ID: ${request.requestID}`;
  
  const htmlBody = buildApprovalEmailHtml(request, action, comments, approvedLevels, currentLevel, nextApproverLabel, null, isExceptional);
  
  try {
    GmailApp.sendEmail(nextApproverEmail, subject, "", {
      htmlBody: htmlBody,
      name: "Manpower Request System"
    });
    Logger.log("Notification sent to next approver: " + nextApproverEmail);
  } catch (e) {
    Logger.log("Error sending email to next approver: " + e);
  }
}

function sendNewRequestNotification(request, approverRole, approverEmail) {
  const subject = `ACTION REQUIRED: New Manpower Request Ready for Your Approval - ${approverRole} - MR ID: ${request.requestID}`;
  const isExceptional = request.manpowerData && request.manpowerData.isExceptional;
  
  const htmlBody = buildApprovalEmailHtml(request, "New", "", [], approverRole, approverRole, null, isExceptional);
  
  try {
    GmailApp.sendEmail(approverEmail, subject, "", {
      htmlBody: htmlBody,
      name: "Manpower Request System"
    });
    Logger.log("New request notification sent to " + approverRole + ": " + approverEmail);
  } catch (e) {
    Logger.log("Error sending new request email: " + e);
  }
}

function sendApprovalNotificationToRequestor(request, action, comments, newStatus, currentLevel, approvedLevels, isExceptional) {
  const recipientEmail = request.email;
  const subjectAction = action === "Approve" ? "Approved" : action === "Disapprove" ? "Disapproved" : "On Hold";
  const levelText = action === "Approve"
    ? (approvedLevels.length > 0 ? approvedLevels.join(" / ") : currentLevel)
    : currentLevel;
  const subject = `Manpower Request ${subjectAction} - ${levelText} - MR ID: ${request.requestID}`;
  
  const htmlBody = buildApprovalEmailHtml(request, action, comments, approvedLevels, currentLevel, null, null, isExceptional);
  
  try {
    GmailApp.sendEmail(recipientEmail, subject, "", {
      htmlBody: htmlBody,
      name: "Manpower Request System"
    });
    Logger.log("Status notification sent to requestor: " + recipientEmail);
  } catch (e) {
    Logger.log("Error sending email to requestor: " + e);
  }
}

// ============================================================
// 5. DASHBOARD ANALYTICS
// ============================================================

/**
 * Get dashboard statistics
 */
function getDashboardStats() {
  const sheet = getRequestsSheet();
  const data = sheet.getDataRange().getValues();
  const email = getCurrentUserEmail();
  const headerMap = getRequestHeaderMap().map;
  
  let total = 0;
  let pending = 0;
  let approved = 0;
  let disapproved = 0;
  let onHold = 0;
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowEmail = getRequestField(row, headerMap, ["requestor", "requestor email", "requested by", "email"], "").toString().trim();
    if (rowEmail === email) {
      total++;
      const status = getRequestField(row, headerMap, ["status"], "").toString().trim();
      if (status === "Pending") pending++;
      else if (status === "Approved") approved++;
      else if (status === "Disapproved") disapproved++;
      else if (status === "On Hold") onHold++;
    }
  }
  
  return {
    total: total,
    pending: pending,
    approved: approved,
    disapproved: disapproved,
    onHold: onHold,
    approvalRate: total > 0 ? ((approved / total) * 100).toFixed(1) : 0
  };
}

// ============================================================
// 6. EMPLOYEE DATA FOR REPLACEMENT
// ============================================================

/**
 * Get separated/resigned employees
 */
function getSeparatedEmployees(section, line) {
  // This would normally query your HR system
  // For now, returning mock data structure
  const employees = [
    { name: "John Doe", position: "Production Operator", status: "Resigned", dateLeft: "2024-01-15" },
    { name: "Jane Smith", position: "Technician", status: "Separated", dateLeft: "2024-02-20" },
    { name: "Bob Wilson", position: "Supervisor", status: "AWOL", dateLeft: "2024-03-10" }
  ];
  
  return employees;
}

// ============================================================
// 7. DEPLOYMENT & HELPER FUNCTIONS
// ============================================================

/**
 * Deploy as web app
 * Run this once to set up the web app
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile("index");
}

/**
 * Create request with file uploads (base64 encoded)
 */
function createRequestWithFiles(requestData) {
  try {
    const parentFolder = getAttachmentsParentFolder();

    // Create the request first
    const result = createRequest(requestData);
    
    // Handle files if provided
    if (requestData.filesData && requestData.filesData.length > 0) {
      const requestFolder = parentFolder.createFolder(result.requestID);
      const fileUrls = [];

      requestData.filesData.forEach(fileData => {
        try {
          // Parse the data URL
          const dataUrl = fileData.data;
          const base64Index = dataUrl.indexOf(',') + 1;
          const base64Data = dataUrl.substring(base64Index);
          const binaryData = Utilities.newBlob(Utilities.base64Decode(base64Data), fileData.type);
          
          const driveFile = requestFolder.createFile(binaryData.setName(fileData.name));
          fileUrls.push(fileData.name + ';' + driveFile.getUrl());
          
          Logger.log('File created: ' + fileData.name);
        } catch (fileError) {
          Logger.log('Error processing file ' + fileData.name + ': ' + fileError);
        }
      });

      if (fileUrls.length > 0) {
        updateRequestAttachments(result.requestID, fileUrls.join('| '));
      }
    }
    
    return {success: true, requestID: result.requestID, message: 'Request submitted successfully with files'};
  } catch (error) {
    Logger.log('createRequestWithFiles error: ' + error);
    if (String(error).indexOf('DriveApp.getFolderById') !== -1) {
      throw new Error('File upload failed: Drive authorization is required. Run authorizeDriveAccess() in the Apps Script editor, approve the permissions, then redeploy the web app.');
    }
    throw new Error('File upload failed: ' + error.toString());
  }
}

/**
 * Run this once from the Apps Script editor to grant Drive access.
 */
function authorizeDriveAccess() {
  const folder = getAttachmentsParentFolder();
  return {
    success: true,
    folderId: folder.getId(),
    folderName: folder.getName()
  };
}

/**
 * Initialize Users sheet if empty
 */
function initializeUsers() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAMES.USERS);
  const lastRow = sheet.getLastRow();
  
  if (lastRow === 0) {
    sheet.appendRow(["Email", "Role", "Name", "Department", "Division"]);
    sheet.appendRow(["user@company.com", "Requestor", "Sample User", "Production", "Plant A"]);
  }
}

/**
 * Get all requests (for admin view)
 */
function getAllRequests() {
  const sheet = getRequestsSheet();
  const data = sheet.getDataRange().getValues();
  const headerMap = getRequestHeaderMap().map;
  
  const requests = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    requests.push({
      requestID: getRequestField(row, headerMap, ["request id", "id"], ""),
      createdDate: getRequestField(row, headerMap, ["timestamp", "date"], ""),
      requestedBy: getRequestField(row, headerMap, ["requestor", "requested by", "requestor email"], ""),
      department: getRequestField(row, headerMap, ["department"], ""),
      requestType: getRequestField(row, headerMap, ["type", "request type"], ""),
      status: getRequestField(row, headerMap, ["status"], "")
    });
  }
  
  return requests;
}

/**
 * Update attached files for a request
 */
function updateRequestAttachments(requestID, attachments) {
  const sheet = getRequestsSheet();
  const data = sheet.getDataRange().getValues();
  const headerMap = getRequestHeaderMap().map;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (getRequestField(row, headerMap, ["request id", "id"], "") === requestID) {
      const idx = headerMap["attached files"] !== undefined ? headerMap["attached files"] :
                  headerMap["attachments"] !== undefined ? headerMap["attachments"] :
                  headerMap["attachedfiles"];
      if (idx !== undefined) {
        sheet.getRange(i + 1, idx + 1).setValue(attachments);
      }
      break;
    }
  }
}

/**
 * Remove exact duplicate requests from the Requests sheet.
 * This is useful for cleanup after accidental double submission.
 */
function removeDuplicateRequests() {
  const sheet = getRequestsSheet();
  const data = sheet.getDataRange().getValues();
  const seen = {};
  const rowsToDelete = [];

  for (let i = data.length - 1; i > 0; i--) {
    const row = data[i];
    const signature = [
      row[2] || '', // Requestor
      row[4] || '', // Division
      row[5] || '', // Group
      row[6] || '', // Department
      row[7] || '', // Section
      row[8] || '', // Unit
      row[9] || '', // Line
      row[11] || '', // Type
      row[12] || '', // Details
      row[13] || ''  // Justification
    ].join('|').toLowerCase();

    if (seen[signature]) {
      rowsToDelete.push(i + 1);
    } else {
      seen[signature] = true;
    }
  }

  rowsToDelete.forEach(rowIndex => sheet.deleteRow(rowIndex));
  return {
    removed: rowsToDelete.length,
    deletedRows: rowsToDelete
  };
}
