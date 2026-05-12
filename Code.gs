/*******************************************************
 * INTERN DAILY REPORT PLATFORM - Code.gs
 * Updated for:
 * 1. Correct Manager Dashboard count refresh
 * 2. Review Status bar update based on live data
 * 3. Approved report lock - no further edit/status change
 * 4. Clickable scorecards for status-wise filtering
 * 5. PDF export only for Approved reports
 * 6. Core Intern submission process remains unchanged
 *******************************************************/
 
const SHEET_NAME = "Daily_Report";
 
const REQUIRED_HEADERS = [
  "Timestamp",
  "Report ID",
  "Intern Name",
  "Department",
  "Reporting Manager",
  "Date",
  "Day",
  "Reporting Time",
  "Logout Time",
  "Total Working Hours",
  "Task Summary",
  "Work Execution Log",
  "Key Process Learned",
  "Tools Learned",
  "Data Observation",
  "MIS SOP Dashboard Status",
  "Challenges Faced",
  "Next Day Plan",
  "Self Assessment",
  "Manager Remarks",
  "Manager Rating",
  "Review Status",
  "Reviewed On"
];
 
function doGet() {
  setupSheet_();
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("Intern Daily Report Platform")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
 
function setupSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
 
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
 
  const firstRow = sheet.getRange(1, 1, 1, REQUIRED_HEADERS.length).getValues()[0];
  const isBlank = firstRow.every(v => String(v || "").trim() === "");
 
  if (isBlank) {
    sheet.getRange(1, 1, 1, REQUIRED_HEADERS.length).setValues([REQUIRED_HEADERS]);
    sheet.setFrozenRows(1);
  }
}
 
function getHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    map[String(h).trim()] = i + 1;
  });
  return map;
}
 
function normalizeStatus_(status) {
  const s = String(status || "Pending Review").trim().toLowerCase();
  if (s === "approved") return "Approved";
  if (s === "reviewed") return "Reviewed";
  if (s === "pending review" || s === "pending" || s === "") return "Pending Review";
  return "Pending Review";
}
 
function safeJson_(value) {
  if (Array.isArray(value)) return JSON.stringify(value);
  if (typeof value === "string") return value;
  return JSON.stringify(value || []);
}
 
function submitDailyReport(formData) {
  setupSheet_();
 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const reportId = "IDR-" + new Date().getTime();
 
  sheet.appendRow([
    new Date(),
    reportId,
    formData.internName,
    formData.department,
    formData.reportingManager,
    formData.reportDate,
    formData.day,
    formData.reportingTime,
    formData.logoutTime,
    formData.totalHours,
    safeJson_(formData.taskSummary),
    safeJson_(formData.workLog),
    formData.keyProcess,
    formData.toolsLearned,
    safeJson_(formData.dataObservation),
    safeJson_(formData.misStatus),
    safeJson_(formData.challenges),
    safeJson_(formData.nextDayPlan),
    safeJson_(formData.selfAssessment),
    "",
    "",
    "Pending Review",
    ""
  ]);
 
  SpreadsheetApp.flush();
  return "Daily Report submitted successfully. Report ID: " + reportId;
}
 
function getDashboardData() {
  setupSheet_();
 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
 
  if (!sheet) {
    throw new Error("Sheet not found: " + SHEET_NAME);
  }
 
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
 
  if (lastRow <= 1) {
    return {
      records: [],
      total: 0,
      pending: 0,
      reviewed: 0,
      approved: 0,
      timestamp: new Date().getTime()
    };
  }
 
  const range = sheet.getRange(1, 1, lastRow, lastCol);
  const values = range.getValues();
  const displayValues = range.getDisplayValues();
  const headers = values[0].map(h => String(h).trim());
  const records = [];
 
  for (let i = 1; i < values.length; i++) {
    let row = {};
 
    headers.forEach((h, j) => {
      let value = values[i][j];
      let displayValue = displayValues[i][j];
 
      if (value instanceof Date) {
        if (h === "Date") {
          value = Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else {
          value = Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        }
      } else {
        value = displayValue !== "" ? displayValue : value;
      }
 
      row[h] = value;
    });
 
    row["Review Status"] = normalizeStatus_(row["Review Status"]);
    row.rowNumber = i + 1;
    records.push(row);
  }
 
  const pending = records.filter(r => normalizeStatus_(r["Review Status"]) === "Pending Review").length;
  const reviewed = records.filter(r => normalizeStatus_(r["Review Status"]) === "Reviewed").length;
  const approved = records.filter(r => normalizeStatus_(r["Review Status"]) === "Approved").length;
 
  return {
    records: records.reverse(),
    total: records.length,
    pending: pending,
    reviewed: reviewed,
    approved: approved,
    timestamp: new Date().getTime()
  };
}
 
function updateReview(rowNumber, remarks, rating, status) {
  setupSheet_();
 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const map = getHeaderMap_(sheet);
 
  rowNumber = Number(rowNumber);
 
  if (!rowNumber || rowNumber < 2 || rowNumber > sheet.getLastRow()) {
    throw new Error("Invalid report row selected.");
  }
 
  if (!remarks || !rating || !status) {
    throw new Error("Manager remarks, rating and review status are mandatory.");
  }
 
  const finalStatus = normalizeStatus_(status);
  const currentStatus = normalizeStatus_(sheet.getRange(rowNumber, map["Review Status"]).getDisplayValue());
 
  if (currentStatus === "Approved") {
    throw new Error("This report is already Approved. Approved reports are locked and cannot be edited further.");
  }
 
  sheet.getRange(rowNumber, map["Manager Remarks"]).setValue(remarks);
  sheet.getRange(rowNumber, map["Manager Rating"]).setValue(rating);
  sheet.getRange(rowNumber, map["Review Status"]).setValue(finalStatus);
  sheet.getRange(rowNumber, map["Reviewed On"]).setValue(new Date());
 
  SpreadsheetApp.flush();
 
  return {
    message: "Manager review updated successfully.",
    status: finalStatus,
    reviewedOn: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")
  };
}
 
function generateReportPdf(rowNumber) {
  setupSheet_();
 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
 
  rowNumber = Number(rowNumber);
 
  if (!rowNumber || rowNumber < 2 || rowNumber > sheet.getLastRow()) {
    throw new Error("Invalid report selected for PDF export.");
  }
 
  const rowValues = sheet.getRange(rowNumber, 1, 1, lastCol).getDisplayValues()[0];
  const record = {};
  headers.forEach((h, i) => record[h] = rowValues[i]);
 
  const status = normalizeStatus_(record["Review Status"]);
 
  if (status !== "Approved") {
    throw new Error("PDF can be downloaded only after the report is Approved by the Reporting Manager.");
  }
 
  const reportId = record["Report ID"] || "Daily_Report";
  const internName = record["Intern Name"] || "Intern";
  const reportDate = record["Date"] || "";
 
  const doc = DocumentApp.create("Intern Daily Report - " + reportId);
  const body = doc.getBody();
 
  body.appendParagraph("INTERN DAILY REPORT")
    .setHeading(DocumentApp.ParagraphHeading.HEADING1)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
 
  body.appendParagraph("Corporate Office - Data Analysis & Process Orientation")
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
 
  body.appendParagraph(" ");
 
  appendKeyValueTable_(body, [
    ["Report ID", reportId],
    ["Intern Name", internName],
    ["Department", record["Department"] || ""],
    ["Reporting Manager", record["Reporting Manager"] || ""],
    ["Date", reportDate],
    ["Day", record["Day"] || ""],
    ["Reporting Time", record["Reporting Time"] || ""],
    ["Logout Time", record["Logout Time"] || ""],
    ["Total Working Hours", record["Total Working Hours"] || ""],
    ["Review Status", status],
    ["Manager Rating", record["Manager Rating"] || ""],
    ["Reviewed On", record["Reviewed On"] || ""]
  ]);
 
  appendJsonTable_(body, "1. Task Summary", record["Task Summary"]);
  appendJsonTable_(body, "2. Detailed Work Execution Log", record["Work Execution Log"]);
 
  body.appendParagraph("3. Process Understanding / Learning of the Day")
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);
  appendKeyValueTable_(body, [
    ["Key Process Learned", record["Key Process Learned"] || ""],
    ["Tools Learned", record["Tools Learned"] || ""]
  ]);
 
  appendJsonTable_(body, "4. Data Analysis / Observation Report", record["Data Observation"]);
  appendJsonTable_(body, "5. Dashboard / MIS / SOP Work Status", record["MIS SOP Dashboard Status"]);
  appendJsonTable_(body, "6. Challenges Faced", record["Challenges Faced"]);
  appendJsonTable_(body, "7. Plan for Next Working Day", record["Next Day Plan"]);
  appendJsonTable_(body, "8. Daily Self-Assessment", record["Self Assessment"]);
 
  body.appendParagraph("9. Manager Review")
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);
  appendKeyValueTable_(body, [
    ["Manager Remarks", record["Manager Remarks"] || ""],
    ["Manager Rating", record["Manager Rating"] || ""],
    ["Review Status", status],
    ["Reviewed On", record["Reviewed On"] || ""]
  ]);
 
  doc.saveAndClose();
 
  const docFile = DriveApp.getFileById(doc.getId());
  const pdfBlob = docFile.getAs(MimeType.PDF).setName(reportId + "_" + internName + "_" + reportDate + ".pdf");
  const pdfFile = DriveApp.createFile(pdfBlob);
 
  try {
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {
    // Domain-restricted environments may block public sharing. The file will still be created in the owner's Drive.
  }
 
  try {
    docFile.setTrashed(true);
  } catch (err2) {}
 
  return {
    message: "PDF generated successfully.",
    url: "https://drive.google.com/uc?export=download&id=" + pdfFile.getId(),
    previewUrl: pdfFile.getUrl(),
    fileName: pdfFile.getName()
  };
}
 
function appendKeyValueTable_(body, rows) {
  const table = body.appendTable();
  rows.forEach(pair => {
    const tr = table.appendTableRow();
    tr.appendTableCell(String(pair[0] || ""));
    tr.appendTableCell(String(pair[1] || ""));
  });
}
 
function appendJsonTable_(body, title, jsonText) {
  body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING2);
 
  let arr = [];
  try {
    arr = JSON.parse(jsonText || "[]");
  } catch (e) {
    arr = [];
  }
 
  if (!arr || arr.length === 0) {
    body.appendParagraph("No data available.");
    return;
  }
 
  const keys = Object.keys(arr[0]);
  const table = body.appendTable();
  const headerRow = table.appendTableRow();
  keys.forEach(k => headerRow.appendTableCell(k));
 
  arr.forEach(item => {
    const row = table.appendTableRow();
    keys.forEach(k => row.appendTableCell(String(item[k] || "")));
  });
}
 
