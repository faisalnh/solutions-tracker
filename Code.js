const SHEET_NAME = "Solution Requests";
const SPREADSHEET_NAME = "MAD Labs Solution Requests";

const HEADERS = [
  "Request ID",
  "Created At",
  "Request Title",
  "Department",
  "Requester Name",
  "Date of Request",
  "Problem Statement",
  "Impact / Why It Matters",
  "Current Data Sources",
  "Has Solution In Mind",
  "Proposed Solution",
  "Reference / Example",
  "Feature 1",
  "Feature 2",
  "Feature 3",
  "Users",
  "Success Criteria",
  "Priority Level",
  "Deadline",
  "Additional Notes",
  "Status",
  "Reviewed By",
  "Feasibility",
  "Impact Level",
  "Decision",
  "MAD Labs Notes",
  "Last Updated",
];

const DEFAULT_STATUS = "Submitted";
const VALID_PRIORITIES = ["Low", "Medium", "High"];
const VALID_STATUSES = [
  "Submitted",
  "Approved",
  "Rejected",
  "MVP",
  "Done",
  "Backlog",
];
const VALID_ASSESSMENTS = {
  feasibility: ["Easy", "Medium", "Complex"],
  impactLevel: ["Low", "Medium", "High"],
  decision: ["Approved", "Backlog", "Rejected"],
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("MAD Labs Solutions Tracker")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function submitRequest(formData) {
  const data = normalizeRequest_(formData || {});
  validateRequest_(data);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const sheet = getSheet_();
    const now = new Date();
    const requestId = createRequestId_(now);

    sheet.appendRow([
      requestId,
      now,
      data.requestTitle,
      data.department,
      data.requesterName,
      data.dateOfRequest,
      data.problemStatement,
      data.impact,
      data.dataSources,
      data.hasSolution,
      data.proposedSolution,
      data.reference,
      data.feature1,
      data.feature2,
      data.feature3,
      data.users,
      data.successCriteria,
      data.priority,
      data.deadline,
      data.additionalNotes,
      DEFAULT_STATUS,
      "",
      "",
      "",
      "",
      "",
      now,
    ]);

    applySheetFormatting_(sheet);
    return {
      ok: true,
      requestId: requestId,
      request: getRequestById_(requestId),
    };
  } finally {
    lock.releaseLock();
  }
}

function getRequests() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();

  if (values.length <= 1) {
    return {
      ok: true,
      requests: [],
      statuses: VALID_STATUSES,
      priorities: VALID_PRIORITIES,
      spreadsheetUrl: SpreadsheetApp.openById(getSpreadsheetId_()).getUrl(),
    };
  }

  const rows = values.slice(1).filter(function (row) {
    return row[0];
  });

  const requests = rows.map(rowToRequest_).sort(function (a, b) {
    return new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime();
  });

  return {
    ok: true,
    requests: requests,
    statuses: VALID_STATUSES,
    priorities: VALID_PRIORITIES,
    spreadsheetUrl: SpreadsheetApp.openById(getSpreadsheetId_()).getUrl(),
  };
}

function getRequestById_(requestId) {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();

  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === requestId) {
      return rowToRequest_(values[i]);
    }
  }

  return null;
}

function normalizeRequest_(formData) {
  return {
    requestTitle: cleanText_(formData.requestTitle),
    department: cleanText_(formData.department),
    requesterName: cleanText_(formData.requesterName),
    dateOfRequest: cleanText_(formData.dateOfRequest),
    problemStatement: cleanText_(formData.problemStatement),
    impact: cleanText_(formData.impact),
    dataSources: cleanText_(formData.dataSources),
    hasSolution: cleanText_(formData.hasSolution || "No"),
    proposedSolution: cleanText_(formData.proposedSolution),
    reference: cleanText_(formData.reference),
    feature1: cleanText_(formData.feature1),
    feature2: cleanText_(formData.feature2),
    feature3: cleanText_(formData.feature3),
    users: cleanText_(formData.users),
    successCriteria: cleanText_(formData.successCriteria),
    priority: cleanText_(formData.priority || "Medium"),
    deadline: cleanText_(formData.deadline),
    additionalNotes: cleanText_(formData.additionalNotes),
  };
}

function validateRequest_(data) {
  const requiredFields = [
    ["requestTitle", "Request title"],
    ["department", "Department"],
    ["requesterName", "Requester name"],
    ["dateOfRequest", "Date of request"],
    ["problemStatement", "Problem statement"],
    ["impact", "Why this matters"],
    ["users", "Who will use this"],
    ["successCriteria", "Success criteria"],
    ["priority", "Priority level"],
  ];

  requiredFields.forEach(function (field) {
    if (!data[field[0]]) {
      throw new Error(field[1] + " is required.");
    }
  });

  if (VALID_PRIORITIES.indexOf(data.priority) === -1) {
    throw new Error("Priority level must be Low, Medium, or High.");
  }

  if (["Yes", "No"].indexOf(data.hasSolution) === -1) {
    throw new Error("Solution in mind must be Yes or No.");
  }
}

function getSheet_() {
  const spreadsheet = getSpreadsheet_();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  ensureHeaders_(sheet);
  return sheet;
}

function getSpreadsheet_() {
  const boundSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (boundSpreadsheet) {
    ensureEditTrigger_(boundSpreadsheet.getId());
    return boundSpreadsheet;
  }

  const spreadsheetId = getSpreadsheetId_();
  return SpreadsheetApp.openById(spreadsheetId);
}

function getSpreadsheetId_() {
  const boundSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (boundSpreadsheet) {
    ensureEditTrigger_(boundSpreadsheet.getId());
    return boundSpreadsheet.getId();
  }

  const properties = PropertiesService.getScriptProperties();
  let spreadsheetId = properties.getProperty("REQUESTS_SPREADSHEET_ID");

  if (spreadsheetId) {
    ensureEditTrigger_(spreadsheetId);
    return spreadsheetId;
  }

  const spreadsheet = SpreadsheetApp.create(SPREADSHEET_NAME);
  spreadsheetId = spreadsheet.getId();
  properties.setProperty("REQUESTS_SPREADSHEET_ID", spreadsheetId);
  ensureEditTrigger_(spreadsheetId);

  const defaultSheet = spreadsheet.getSheets()[0];
  defaultSheet.setName(SHEET_NAME);
  ensureHeaders_(defaultSheet);
  applySheetFormatting_(defaultSheet);

  return spreadsheetId;
}

function handleRequestSheetEdit(e) {
  if (!e || !e.range) {
    return;
  }

  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_NAME || e.range.getRow() === 1) {
    return;
  }

  const editedColumn = e.range.getColumn();
  const firstAssessmentColumn = 21;
  const lastAssessmentColumn = 26;

  if (
    editedColumn < firstAssessmentColumn ||
    editedColumn > lastAssessmentColumn
  ) {
    return;
  }

  sheet.getRange(e.range.getRow(), 27).setValue(new Date());
}

function ensureEditTrigger_(spreadsheetId) {
  const triggerExists = ScriptApp.getProjectTriggers().some(function (trigger) {
    return trigger.getHandlerFunction() === "handleRequestSheetEdit";
  });

  if (!triggerExists) {
    ScriptApp.newTrigger("handleRequestSheetEdit")
      .forSpreadsheet(spreadsheetId)
      .onEdit()
      .create();
  }
}

function ensureHeaders_(sheet) {
  const currentHeaders = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const hasAnyHeader = currentHeaders.some(function (value) {
    return value !== "";
  });

  if (!hasAnyHeader) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    return;
  }

  HEADERS.forEach(function (header, index) {
    if (currentHeaders[index] !== header) {
      sheet.getRange(1, index + 1).setValue(header);
    }
  });
}

function applySheetFormatting_(sheet) {
  const maxRows = Math.max(sheet.getMaxRows() - 1, 1);

  sheet.setFrozenRows(1);
  sheet
    .getRange(1, 1, 1, HEADERS.length)
    .setBackground("#8BA888")
    .setFontColor("#ffffff")
    .setFontWeight("bold");
  sheet.autoResizeColumns(1, HEADERS.length);

  const statusRange = sheet.getRange(2, 21, maxRows, 1);
  statusRange.setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(VALID_STATUSES, true)
      .setAllowInvalid(false)
      .build(),
  );

  sheet
    .getRange(2, 18, maxRows, 1)
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(VALID_PRIORITIES, true)
        .setAllowInvalid(false)
        .build(),
    );

  sheet
    .getRange(2, 23, maxRows, 1)
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(VALID_ASSESSMENTS.feasibility, true)
        .setAllowInvalid(true)
        .build(),
    );

  sheet
    .getRange(2, 24, maxRows, 1)
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(VALID_ASSESSMENTS.impactLevel, true)
        .setAllowInvalid(true)
        .build(),
    );

  sheet
    .getRange(2, 25, maxRows, 1)
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(VALID_ASSESSMENTS.decision, true)
        .setAllowInvalid(true)
        .build(),
    );
}

function rowToRequest_(row) {
  return {
    id: String(row[0] || ""),
    createdAt: formatDateTime_(row[1]),
    requestTitle: String(row[2] || ""),
    department: String(row[3] || ""),
    requesterName: String(row[4] || ""),
    dateOfRequest: formatDate_(row[5]),
    problemStatement: String(row[6] || ""),
    impact: String(row[7] || ""),
    dataSources: String(row[8] || ""),
    hasSolution: String(row[9] || ""),
    proposedSolution: String(row[10] || ""),
    reference: String(row[11] || ""),
    features: [
      String(row[12] || ""),
      String(row[13] || ""),
      String(row[14] || ""),
    ].filter(Boolean),
    users: String(row[15] || ""),
    successCriteria: String(row[16] || ""),
    priority: String(row[17] || "Medium"),
    deadline: formatDate_(row[18]),
    additionalNotes: String(row[19] || ""),
    status: String(row[20] || DEFAULT_STATUS),
    reviewedBy: String(row[21] || ""),
    feasibility: String(row[22] || ""),
    impactLevel: String(row[23] || ""),
    decision: String(row[24] || ""),
    madLabsNotes: String(row[25] || ""),
    lastUpdated: formatDateTime_(row[26]),
  };
}

function createRequestId_(date) {
  const datePart = Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    "yyyyMMdd",
  );
  const randomPart = Utilities.getUuid().slice(0, 6).toUpperCase();
  return "MAD-" + datePart + "-" + randomPart;
}

function cleanText_(value) {
  return String(value || "").trim();
}

function formatDate_(value) {
  if (!value) {
    return "";
  }

  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd",
    );
  }

  return String(value);
}

function formatDateTime_(value) {
  if (!value) {
    return "";
  }

  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd HH:mm",
    );
  }

  return String(value);
}
