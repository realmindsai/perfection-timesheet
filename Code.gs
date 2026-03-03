/**
 * Perfection Services — Weekly Timesheet (API Backend)
 * Google Apps Script — receives data from external form hosted on GitHub Pages
 *
 * Setup:
 * 1. Create a Google Sheet
 * 2. Extensions → Apps Script → paste this into Code.gs
 * 3. Run initialSetup() once
 * 4. Add staff names to the "Staff List" tab
 * 5. Deploy → New deployment → Web app → Execute as "Me" → Who has access "Anyone" → Deploy
 * 6. Copy the deployment URL and paste it into the form's SCRIPT_URL
 */

const SHEET_NAME = "Timesheet Responses";
const SUMMARY_SHEET = "Pay Summary";
const STAFF_SHEET = "Staff List";

/**
 * Handles GET requests — returns staff names as JSON
 */
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || "getNames";

  if (action === "getNames") {
    const names = getStaffNames();
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, names: names }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ success: true, message: "Timesheet API running" }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Handles POST requests — processes timesheet submission
 */
function doPost(e) {
  try {
    // Handle both form-encoded (hidden iframe) and raw JSON (fetch) submissions
    let formData;
    if (e.parameter && e.parameter.data) {
      formData = JSON.parse(e.parameter.data);
    } else if (e.postData && e.postData.contents) {
      formData = JSON.parse(e.postData.contents);
    } else {
      throw new Error("No data received");
    }
    const result = processSubmission(formData);
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Returns staff names from the Staff List tab
 */
function getStaffNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(STAFF_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(STAFF_SHEET);
    sheet.getRange("A1").setValue("Staff Names").setFontWeight("bold");
    sheet.getRange("A2").setValue("Add staff names here");
    sheet.setColumnWidth(1, 200);
  }

  const data = sheet.getRange("A2:A").getValues();
  return data.filter(row => row[0] !== "").map(row => row[0]);
}

/**
 * Processes a timesheet submission and writes to sheet
 */
function processSubmission(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = createResponseSheet(ss);
  }

  const days = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"];
  const row = [
    new Date(),
    formData.staffName,
    formData.weekEnding,
  ];

  let totalRaw = 0;
  let totalBreaks = 0;
  let totalOrdinary = 0;
  let totalOvertime = 0;

  for (let i = 0; i < days.length; i++) {
    const start = formData[days[i] + "Start"] || "";
    const end = formData[days[i] + "End"] || "";
    const notes = formData[days[i] + "Notes"] || "";

    let rawHrs = "";
    let ordinary = "";
    let overtime = "";

    if (start && end) {
      rawHrs = calcRawHours(start, end);

      let worked = rawHrs;
      if (rawHrs > 5) {
        worked = rawHrs - 0.5;
        totalBreaks += 0.5;
      }

      if (worked > 8.5) {
        ordinary = 8.5;
        overtime = Math.round((worked - 8.5) * 100) / 100;
      } else {
        ordinary = Math.round(worked * 100) / 100;
        overtime = 0;
      }

      rawHrs = Math.round(rawHrs * 100) / 100;
      totalRaw += rawHrs;
      totalOrdinary += ordinary;
      totalOvertime += overtime;
    }

    row.push(start, end, rawHrs, ordinary, overtime, notes);
  }

  totalRaw = Math.round(totalRaw * 100) / 100;
  totalOrdinary = Math.round(totalOrdinary * 100) / 100;
  totalOvertime = Math.round(totalOvertime * 100) / 100;
  totalBreaks = Math.round(totalBreaks * 100) / 100;

  let status = "OK";
  if (totalRaw > 60) status = "CHECK: >60hrs";
  else if (totalRaw > 45) status = "WARN: >45hrs";

  row.push(totalRaw, totalBreaks, totalOrdinary, totalOvertime, status);
  row.push(formData.generalNotes || "");

  sheet.appendRow(row);
  updatePaySummary(ss);

  return {
    success: true,
    name: formData.staffName,
    weekEnding: formData.weekEnding,
    totalOrdinary: totalOrdinary,
    totalOvertime: totalOvertime
  };
}

function calcRawHours(start, end) {
  const [sh, sm] = start.split(":").map(Number);
  const [eh, em] = end.split(":").map(Number);
  let hours = (eh * 60 + em - sh * 60 - sm) / 60;
  if (hours < 0) hours += 24;
  return hours;
}

function createResponseSheet(ss) {
  const sheet = ss.insertSheet(SHEET_NAME, 0);
  const days = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];
  const headers = ["Timestamp", "Name", "Week Ending"];

  for (const day of days) {
    headers.push(day + " Start", day + " End", day + " Raw Hrs", day + " Ordinary", day + " Overtime", day + " Notes");
  }
  headers.push("Total Raw", "Total Breaks", "Total Ordinary", "Total Overtime", "Status", "General Notes");

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#6B3FA0")
    .setFontColor("#FFFFFF");
  sheet.setFrozenRows(1);

  return sheet;
}

function updatePaySummary(ss) {
  let summary = ss.getSheetByName(SUMMARY_SHEET);
  if (!summary) {
    summary = ss.insertSheet(SUMMARY_SHEET);
    const headers = ["Name", "Week Ending", "Total Raw Hrs", "Breaks Deducted", "Ordinary Hrs", "Overtime Hrs", "Status"];
    summary.getRange(1, 1, 1, headers.length).setValues([headers]);
    summary.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#6B3FA0")
      .setFontColor("#FFFFFF");
    summary.setFrozenRows(1);
  }

  const responses = ss.getSheetByName(SHEET_NAME);
  const data = responses.getDataRange().getValues();
  if (data.length < 2) return;

  if (summary.getLastRow() > 1) {
    summary.getRange(2, 1, summary.getLastRow() - 1, 7).clearContent();
  }

  const totalRawCol = data[0].indexOf("Total Raw");
  const totalBreaksCol = data[0].indexOf("Total Breaks");
  const totalOrdCol = data[0].indexOf("Total Ordinary");
  const totalOTCol = data[0].indexOf("Total Overtime");
  const statusCol = data[0].indexOf("Status");

  const summaryData = [];
  for (let i = 1; i < data.length; i++) {
    summaryData.push([
      data[i][1], data[i][2],
      data[i][totalRawCol], data[i][totalBreaksCol],
      data[i][totalOrdCol], data[i][totalOTCol],
      data[i][statusCol]
    ]);
  }

  if (summaryData.length > 0) {
    summary.getRange(2, 1, summaryData.length, 7).setValues(summaryData);
  }
}

function initialSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  createResponseSheet(ss);

  let staff = ss.getSheetByName(STAFF_SHEET);
  if (!staff) {
    staff = ss.insertSheet(STAFF_SHEET);
    staff.getRange("A1").setValue("Staff Names").setFontWeight("bold");
    staff.setColumnWidth(1, 200);
  }

  updatePaySummary(ss);
  SpreadsheetApp.getUi().alert("Setup complete! Add staff names to the Staff List tab, then deploy as web app.");
}
