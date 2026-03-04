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
const FORTNIGHT_SHEET = "Fortnight Summary";
const STAFF_SHEET = "Staff List";
const SETTINGS_SHEET = "Settings";

const DAYS = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"];
const DAYS_TITLE = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];

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

    let result;
    if (formData.version === 2) {
      result = processV2(formData);
    } else {
      result = processV1(formData);
    }

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

// ---------------------------------------------------------------------------
// v1 (legacy) processing — backward compatible with original flat format
// ---------------------------------------------------------------------------

/**
 * Processes a v1 (flat) timesheet submission and writes to sheet
 */
function processV1(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = createResponseSheet(ss);
  }

  const row = [
    new Date(),
    formData.staffName,
    formData.weekEnding,
  ];

  let totalRaw = 0;
  let totalBreaks = 0;
  let totalOrdinary = 0;
  let totalOvertime = 0;

  for (let i = 0; i < DAYS.length; i++) {
    const start = formData[DAYS[i] + "Start"] || "";
    const end = formData[DAYS[i] + "End"] || "";
    const notes = formData[DAYS[i] + "Notes"] || "";

    const calc = calcDayHours(start, end);
    totalRaw += calc.rawHrs || 0;
    totalBreaks += calc.breakHrs || 0;
    totalOrdinary += calc.ordinary || 0;
    totalOvertime += calc.overtime || 0;

    row.push(start, end, calc.rawHrs, calc.ordinary, calc.overtime, notes);
  }

  totalRaw = round2(totalRaw);
  totalBreaks = round2(totalBreaks);
  totalOrdinary = round2(totalOrdinary);
  totalOvertime = round2(totalOvertime);

  let status = "OK";
  if (totalRaw > 60) status = "CHECK: >60hrs";
  else if (totalRaw > 45) status = "WARN: >45hrs";

  row.push(totalRaw, totalBreaks, totalOrdinary, totalOvertime);
  row.push(0);                            // High Risk Hours
  row.push("N");                          // Is Estimate
  row.push("SUBMISSION");                 // Type
  row.push(status);                       // Status
  row.push("");                           // Unavailable Days
  row.push(formData.generalNotes || "");  // General Notes

  sheet.appendRow(row);
  updatePaySummary(ss);
  updateFortnightSummary(ss);

  return {
    success: true,
    name: formData.staffName,
    weekEnding: formData.weekEnding,
    totalOrdinary: totalOrdinary,
    totalOvertime: totalOvertime
  };
}

// ---------------------------------------------------------------------------
// v2 processing — multi-week, corrections, high-risk hours
// ---------------------------------------------------------------------------

/**
 * Processes a v2 (multi-week + corrections) submission
 */
function processV2(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = createResponseSheet(ss);
  }

  const results = [];

  // Process each week
  for (let i = 0; i < formData.weeks.length; i++) {
    const weekResult = processWeekSubmission(
      ss, sheet, formData.staffName, formData.weeks[i],
      formData.unavailableDays || "", formData.generalNotes || ""
    );
    results.push(weekResult);
  }

  // Process correction if present
  if (formData.correction) {
    processCorrection(ss, sheet, formData.staffName, formData.correction);
  }

  updatePaySummary(ss);
  updateFortnightSummary(ss);

  // Build summary from first week for response
  const primary = results[0] || {};
  return {
    success: true,
    name: formData.staffName,
    weeksProcessed: results.length,
    correctionProcessed: !!formData.correction,
    weekEnding: primary.weekEnding,
    totalOrdinary: primary.totalOrdinary,
    totalOvertime: primary.totalOvertime
  };
}

/**
 * Appends a single week's submission row to "Timesheet Responses"
 */
function processWeekSubmission(ss, sheet, staffName, weekData, unavailableDays, generalNotes) {
  const row = [
    new Date(),
    staffName,
    weekData.weekEnding,
  ];

  let totalRaw = 0;
  let totalBreaks = 0;
  let totalOrdinary = 0;
  let totalOvertime = 0;

  for (let i = 0; i < DAYS.length; i++) {
    const start = weekData[DAYS[i] + "Start"] || "";
    const end = weekData[DAYS[i] + "End"] || "";
    const notes = weekData[DAYS[i] + "Notes"] || "";

    const calc = calcDayHours(start, end);
    totalRaw += calc.rawHrs || 0;
    totalBreaks += calc.breakHrs || 0;
    totalOrdinary += calc.ordinary || 0;
    totalOvertime += calc.overtime || 0;

    row.push(start, end, calc.rawHrs, calc.ordinary, calc.overtime, notes);
  }

  totalRaw = round2(totalRaw);
  totalBreaks = round2(totalBreaks);
  totalOrdinary = round2(totalOrdinary);
  totalOvertime = round2(totalOvertime);

  let status = "OK";
  if (totalRaw > 60) status = "CHECK: >60hrs";
  else if (totalRaw > 45) status = "WARN: >45hrs";

  const highRiskHours = weekData.highRiskHours || 0;
  const isEstimate = weekData.isEstimate ? "Y" : "N";

  row.push(totalRaw, totalBreaks, totalOrdinary, totalOvertime);
  row.push(highRiskHours);
  row.push(isEstimate);
  row.push("SUBMISSION");
  row.push(status);
  row.push(unavailableDays);
  row.push(generalNotes);

  sheet.appendRow(row);

  return {
    weekEnding: weekData.weekEnding,
    totalOrdinary: totalOrdinary,
    totalOvertime: totalOvertime
  };
}

/**
 * Appends a CORRECTION row — only the corrected day has values
 */
function processCorrection(ss, sheet, staffName, correction) {
  const row = [
    new Date(),
    staffName,
    correction.weekEnding,
  ];

  let totalRaw = 0;
  let totalBreaks = 0;
  let totalOrdinary = 0;
  let totalOvertime = 0;

  for (let i = 0; i < DAYS.length; i++) {
    if (DAYS[i] === correction.day) {
      const start = correction.start || "";
      const end = correction.end || "";
      const calc = calcDayHours(start, end);

      totalRaw = calc.rawHrs || 0;
      totalBreaks = calc.breakHrs || 0;
      totalOrdinary = calc.ordinary || 0;
      totalOvertime = calc.overtime || 0;

      row.push(start, end, calc.rawHrs, calc.ordinary, calc.overtime, "");
    } else {
      row.push("", "", "", "", "", "");
    }
  }

  totalRaw = round2(totalRaw);
  totalBreaks = round2(totalBreaks);
  totalOrdinary = round2(totalOrdinary);
  totalOvertime = round2(totalOvertime);

  row.push(totalRaw, totalBreaks, totalOrdinary, totalOvertime);
  row.push(0);              // High Risk Hours
  row.push("N");            // Is Estimate
  row.push("CORRECTION");   // Type
  row.push("CORRECTION");   // Status
  row.push("");             // Unavailable Days
  row.push("");             // General Notes

  sheet.appendRow(row);
}

// ---------------------------------------------------------------------------
// Shared calculation helpers
// ---------------------------------------------------------------------------

/**
 * Calculates raw hours between two HH:MM time strings
 */
function calcRawHours(start, end) {
  const [sh, sm] = start.split(":").map(Number);
  const [eh, em] = end.split(":").map(Number);
  let hours = (eh * 60 + em - sh * 60 - sm) / 60;
  if (hours < 0) hours += 24;
  return hours;
}

/**
 * Calculates raw, break, ordinary, and overtime hours for a single day
 * Returns object with rawHrs, breakHrs, ordinary, overtime (all "" if no start/end)
 */
function calcDayHours(start, end) {
  if (!start || !end) {
    return { rawHrs: "", breakHrs: 0, ordinary: "", overtime: "" };
  }

  let rawHrs = calcRawHours(start, end);
  let breakHrs = 0;
  let worked = rawHrs;

  if (rawHrs > 5) {
    worked = rawHrs - 0.5;
    breakHrs = 0.5;
  }

  let ordinary, overtime;
  if (worked > 8.5) {
    ordinary = 8.5;
    overtime = round2(worked - 8.5);
  } else {
    ordinary = round2(worked);
    overtime = 0;
  }

  rawHrs = round2(rawHrs);

  return { rawHrs: rawHrs, breakHrs: breakHrs, ordinary: ordinary, overtime: overtime };
}

/**
 * Rounds a number to 2 decimal places
 */
function round2(n) {
  return Math.round(n * 100) / 100;
}

// ---------------------------------------------------------------------------
// Sheet creation and formatting
// ---------------------------------------------------------------------------

/**
 * Creates the "Timesheet Responses" sheet with headers
 */
function createResponseSheet(ss) {
  const sheet = ss.insertSheet(SHEET_NAME, 0);
  const headers = ["Timestamp", "Name", "Week Ending"];

  for (const day of DAYS_TITLE) {
    headers.push(day + " Start", day + " End", day + " Raw Hrs", day + " Ordinary", day + " Overtime", day + " Notes");
  }
  headers.push(
    "Total Raw", "Total Breaks", "Total Ordinary", "Total Overtime",
    "High Risk Hours", "Is Estimate", "Type", "Status",
    "Unavailable Days", "General Notes"
  );

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#6B3FA0")
    .setFontColor("#FFFFFF");
  sheet.setFrozenRows(1);

  return sheet;
}

/**
 * Creates the "Settings" sheet if it doesn't exist
 */
function ensureSettingsSheet(ss) {
  let sheet = ss.getSheetByName(SETTINGS_SHEET);
  if (sheet) return sheet;

  sheet = ss.insertSheet(SETTINGS_SHEET);
  sheet.getRange("A1").setValue("Fortnight Cycle Start").setFontWeight("bold");
  sheet.getRange("B1").setValue(new Date());
  sheet.getRange("A3").setValue("High Risk Premium ($/hr)").setFontWeight("bold");
  sheet.getRange("B3").setValue(10);
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 150);

  return sheet;
}

// ---------------------------------------------------------------------------
// Pay Summary
// ---------------------------------------------------------------------------

/**
 * Rebuilds the "Pay Summary" tab from Timesheet Responses data
 * Excludes CORRECTION rows (those are handled by Fortnight Summary)
 */
function updatePaySummary(ss) {
  let summary = ss.getSheetByName(SUMMARY_SHEET);
  const headers = [
    "Name", "Week Ending", "Total Raw Hrs", "Breaks Deducted",
    "Ordinary Hrs", "Overtime Hrs", "High Risk Hours",
    "Is Estimate", "Type", "Status"
  ];

  if (!summary) {
    summary = ss.insertSheet(SUMMARY_SHEET);
  }

  // Always rewrite headers (handles v1 → v2 upgrade)
  summary.getRange(1, 1, 1, headers.length).setValues([headers]);
  summary.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#6B3FA0")
    .setFontColor("#FFFFFF");
  summary.setFrozenRows(1);

  const responses = ss.getSheetByName(SHEET_NAME);
  if (!responses) return;

  const data = responses.getDataRange().getValues();
  if (data.length < 2) return;

  if (summary.getLastRow() > 1) {
    summary.getRange(2, 1, summary.getLastRow() - 1, headers.length).clearContent();
  }

  const hdr = data[0];
  const totalRawCol = hdr.indexOf("Total Raw");
  const totalBreaksCol = hdr.indexOf("Total Breaks");
  const totalOrdCol = hdr.indexOf("Total Ordinary");
  const totalOTCol = hdr.indexOf("Total Overtime");
  const highRiskCol = hdr.indexOf("High Risk Hours");
  const isEstimateCol = hdr.indexOf("Is Estimate");
  const typeCol = hdr.indexOf("Type");
  const statusCol = hdr.indexOf("Status");

  const summaryData = [];
  for (let i = 1; i < data.length; i++) {
    // Skip CORRECTION rows — they're merged in Fortnight Summary
    const rowType = typeCol >= 0 ? data[i][typeCol] : "";
    if (rowType === "CORRECTION") continue;

    summaryData.push([
      data[i][1],                                               // Name
      data[i][2],                                               // Week Ending
      totalRawCol >= 0 ? data[i][totalRawCol] : "",             // Total Raw Hrs
      totalBreaksCol >= 0 ? data[i][totalBreaksCol] : "",       // Breaks Deducted
      totalOrdCol >= 0 ? data[i][totalOrdCol] : "",             // Ordinary Hrs
      totalOTCol >= 0 ? data[i][totalOTCol] : "",               // Overtime Hrs
      highRiskCol >= 0 ? data[i][highRiskCol] : 0,              // High Risk Hours
      isEstimateCol >= 0 ? data[i][isEstimateCol] : "N",        // Is Estimate
      typeCol >= 0 ? data[i][typeCol] : "SUBMISSION",           // Type
      statusCol >= 0 ? data[i][statusCol] : ""                  // Status
    ]);
  }

  if (summaryData.length > 0) {
    summary.getRange(2, 1, summaryData.length, headers.length).setValues(summaryData);
  }
}

// ---------------------------------------------------------------------------
// Fortnight Summary
// ---------------------------------------------------------------------------

/**
 * Rebuilds the "Fortnight Summary" tab.
 *
 * Groups weeks into fortnights based on a cycle start date from Settings,
 * applies corrections, and calculates totals.
 */
function updateFortnightSummary(ss) {
  const settingsSheet = ensureSettingsSheet(ss);
  const cycleStartRaw = settingsSheet.getRange("B1").getValue();
  const cycleStart = new Date(cycleStartRaw);
  const hrPremium = settingsSheet.getRange("B3").getValue() || 10;

  const responses = ss.getSheetByName(SHEET_NAME);
  if (!responses) return;

  const data = responses.getDataRange().getValues();
  if (data.length < 2) return;

  const hdr = data[0];
  const colIdx = buildColumnIndex(hdr);

  // Separate submissions and corrections by person
  const submissions = {};  // { name: [ { weekEnding, ordinary, overtime, hrHrs, isEstimate, dayData } ] }
  const corrections = {};  // { name: [ { weekEnding, day, ... } ] }

  for (let i = 1; i < data.length; i++) {
    const name = data[i][1];
    const weekEnding = data[i][2];
    const rowType = colIdx.type >= 0 ? data[i][colIdx.type] : "SUBMISSION";

    if (rowType === "CORRECTION") {
      if (!corrections[name]) corrections[name] = [];
      corrections[name].push(extractCorrectionData(data[i], hdr, colIdx));
    } else {
      if (!submissions[name]) submissions[name] = [];
      submissions[name].push(extractSubmissionData(data[i], colIdx));
    }
  }

  // Apply corrections to submissions
  for (const name in corrections) {
    if (!submissions[name]) continue;
    for (const corr of corrections[name]) {
      applyCorrectionToSubmissions(submissions[name], corr);
    }
  }

  // Group into fortnights and build output rows
  const outputRows = [];
  const allNames = Object.keys(submissions).sort();

  for (const name of allNames) {
    const personWeeks = submissions[name].sort(function(a, b) {
      return parseWeekEnding(a.weekEnding) - parseWeekEnding(b.weekEnding);
    });

    // Assign each week to a fortnight bucket
    const fortnights = groupIntoFortnights(personWeeks, cycleStart);

    for (const fn of fortnights) {
      const wk1 = fn.week1;
      const wk2 = fn.week2;

      const wk1Ord = wk1 ? wk1.totalOrdinary : 0;
      const wk1OT = wk1 ? wk1.totalOvertime : 0;
      const wk1HR = wk1 ? wk1.highRiskHours : 0;
      const wk2Ord = wk2 ? wk2.totalOrdinary : 0;
      const wk2OT = wk2 ? wk2.totalOvertime : 0;
      const wk2HR = wk2 ? wk2.highRiskHours : 0;

      const totalOrd = round2(wk1Ord + wk2Ord);
      const totalOT = round2(wk1OT + wk2OT);
      const totalHR = round2(wk1HR + wk2HR);
      const hrPremiumVal = round2(totalHR * hrPremium);

      // Determine status
      let status;
      const hasWk1 = !!wk1;
      const hasWk2 = !!wk2;
      const hasEstimate = (wk1 && wk1.isEstimate === "Y") || (wk2 && wk2.isEstimate === "Y");

      if (hasWk1 && hasWk2 && !hasEstimate) {
        status = "Complete";
      } else if (hasWk1 && hasWk2 && hasEstimate) {
        status = "Complete (estimates)";
      } else if (hasWk1 || hasWk2) {
        status = "Partial";
      } else {
        continue; // "Missing" — don't show
      }

      outputRows.push([
        name,
        formatDate(fn.fortnightStart),
        formatDate(fn.fortnightEnd),
        wk1 ? wk1.weekEnding : "",
        wk1 ? wk1Ord : "",
        wk1 ? wk1OT : "",
        wk1 ? wk1HR : "",
        wk2 ? wk2.weekEnding : "",
        wk2 ? wk2Ord : "",
        wk2 ? wk2OT : "",
        wk2 ? wk2HR : "",
        totalOrd,
        totalOT,
        totalHR,
        hrPremiumVal,
        status
      ]);
    }
  }

  // Write to sheet
  const fnHeaders = [
    "Name", "Fortnight Start", "Fortnight End",
    "Wk1 Ending", "Wk1 Ordinary", "Wk1 Overtime", "Wk1 HR Hrs",
    "Wk2 Ending", "Wk2 Ordinary", "Wk2 Overtime", "Wk2 HR Hrs",
    "Total Ordinary", "Total Overtime", "Total High Risk",
    "HR Premium ($" + hrPremium + "/hr)", "Status"
  ];

  let fnSheet = ss.getSheetByName(FORTNIGHT_SHEET);
  if (!fnSheet) {
    fnSheet = ss.insertSheet(FORTNIGHT_SHEET);
  }

  // Clear everything and rewrite
  fnSheet.clearContents();
  fnSheet.clearFormats();

  fnSheet.getRange(1, 1, 1, fnHeaders.length).setValues([fnHeaders]);
  fnSheet.getRange(1, 1, 1, fnHeaders.length)
    .setFontWeight("bold")
    .setBackground("#6B3FA0")
    .setFontColor("#FFFFFF");
  fnSheet.setFrozenRows(1);

  if (outputRows.length > 0) {
    fnSheet.getRange(2, 1, outputRows.length, fnHeaders.length).setValues(outputRows);

    // Colour-code the Status column
    const statusColNum = fnHeaders.length; // last column
    for (let r = 0; r < outputRows.length; r++) {
      const statusCell = fnSheet.getRange(r + 2, statusColNum);
      const st = outputRows[r][fnHeaders.length - 1];
      if (st === "Complete") {
        statusCell.setBackground("#C6EFCE"); // green
      } else if (st === "Partial") {
        statusCell.setBackground("#FCD5B4"); // orange
      } else if (st === "Complete (estimates)") {
        statusCell.setBackground("#FFFFCC"); // yellow
      }
    }
  }
}

// ---------------------------------------------------------------------------
// Fortnight Summary helpers
// ---------------------------------------------------------------------------

/**
 * Builds a lookup of key column indices from the header row
 */
function buildColumnIndex(hdr) {
  return {
    totalRaw: hdr.indexOf("Total Raw"),
    totalBreaks: hdr.indexOf("Total Breaks"),
    totalOrdinary: hdr.indexOf("Total Ordinary"),
    totalOvertime: hdr.indexOf("Total Overtime"),
    highRisk: hdr.indexOf("High Risk Hours"),
    isEstimate: hdr.indexOf("Is Estimate"),
    type: hdr.indexOf("Type"),
    status: hdr.indexOf("Status")
  };
}

/**
 * Extracts submission data from a response row
 */
function extractSubmissionData(row, colIdx) {
  return {
    weekEnding: row[2],
    totalOrdinary: colIdx.totalOrdinary >= 0 ? (Number(row[colIdx.totalOrdinary]) || 0) : 0,
    totalOvertime: colIdx.totalOvertime >= 0 ? (Number(row[colIdx.totalOvertime]) || 0) : 0,
    highRiskHours: colIdx.highRisk >= 0 ? (Number(row[colIdx.highRisk]) || 0) : 0,
    isEstimate: colIdx.isEstimate >= 0 ? row[colIdx.isEstimate] : "N",
    // Store per-day data for correction merging
    dayData: extractDayData(row)
  };
}

/**
 * Extracts per-day start/end/hours from a response row
 * Each day occupies 6 columns: Start, End, Raw Hrs, Ordinary, Overtime, Notes
 * Starting at column index 3
 */
function extractDayData(row) {
  const result = {};
  for (let d = 0; d < DAYS.length; d++) {
    const base = 3 + d * 6;
    result[DAYS[d]] = {
      start: row[base] || "",
      end: row[base + 1] || "",
      rawHrs: row[base + 2] || 0,
      ordinary: row[base + 3] || 0,
      overtime: row[base + 4] || 0,
      notes: row[base + 5] || ""
    };
  }
  return result;
}

/**
 * Extracts correction data from a response row
 */
function extractCorrectionData(row, hdr, colIdx) {
  const weekEnding = row[2];
  // Find which day has data
  let corrDay = null;
  const dayData = extractDayData(row);

  for (const day of DAYS) {
    if (dayData[day].start && dayData[day].end) {
      corrDay = day;
      break;
    }
  }

  return {
    weekEnding: weekEnding,
    day: corrDay,
    dayData: corrDay ? dayData[corrDay] : null
  };
}

/**
 * Applies a correction to the matching submission (same week ending).
 * Overrides that day's hours and recalculates totals.
 */
function applyCorrectionToSubmissions(personSubmissions, corr) {
  if (!corr.day || !corr.dayData) return;

  for (let i = 0; i < personSubmissions.length; i++) {
    const sub = personSubmissions[i];
    if (normaliseWeekEnding(sub.weekEnding) === normaliseWeekEnding(corr.weekEnding)) {
      // Override the corrected day
      sub.dayData[corr.day] = corr.dayData;

      // Recalculate totals from day data
      let totalOrdinary = 0;
      let totalOvertime = 0;
      for (const day of DAYS) {
        totalOrdinary += Number(sub.dayData[day].ordinary) || 0;
        totalOvertime += Number(sub.dayData[day].overtime) || 0;
      }
      sub.totalOrdinary = round2(totalOrdinary);
      sub.totalOvertime = round2(totalOvertime);
      break;
    }
  }
}

/**
 * Normalises a week ending value to a comparable string (YYYY-MM-DD)
 */
function normaliseWeekEnding(we) {
  if (we instanceof Date) {
    return Utilities.formatDate(we, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(we);
}

/**
 * Parses a week ending value to a Date for sorting
 */
function parseWeekEnding(we) {
  if (we instanceof Date) return we;
  return new Date(we);
}

/**
 * Formats a Date as YYYY-MM-DD string
 */
function formatDate(d) {
  if (!(d instanceof Date)) return String(d);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

/**
 * Groups an array of week submissions into fortnight buckets
 * based on a cycle start date.
 *
 * Returns array of { fortnightStart, fortnightEnd, week1, week2 }
 */
function groupIntoFortnights(personWeeks, cycleStart) {
  // Map each week to its fortnight bucket number
  const buckets = {}; // bucketNum -> { week1, week2, fortnightStart, fortnightEnd }

  for (const week of personWeeks) {
    const weekEndDate = parseWeekEnding(week.weekEnding);
    // Week start is 6 days before week ending (Mon for a Sun week-ending)
    const weekStartDate = new Date(weekEndDate);
    weekStartDate.setDate(weekStartDate.getDate() - 6);

    // Calculate days since cycle start
    const msPerDay = 86400000;
    const daysSinceCycle = Math.floor((weekStartDate - cycleStart) / msPerDay);
    const bucketNum = Math.floor(daysSinceCycle / 14);

    if (!buckets[bucketNum]) {
      // Calculate fortnight start and end dates for this bucket
      const fnStart = new Date(cycleStart);
      fnStart.setDate(fnStart.getDate() + bucketNum * 14);
      const fnEnd = new Date(fnStart);
      fnEnd.setDate(fnEnd.getDate() + 13);

      buckets[bucketNum] = {
        fortnightStart: fnStart,
        fortnightEnd: fnEnd,
        week1: null,
        week2: null
      };
    }

    // Determine if this is week 1 or week 2 within the fortnight
    const fnStart = buckets[bucketNum].fortnightStart;
    const daysIntoFortnight = Math.floor((weekStartDate - fnStart) / msPerDay);

    if (daysIntoFortnight < 7) {
      buckets[bucketNum].week1 = week;
    } else {
      buckets[bucketNum].week2 = week;
    }
  }

  // Sort buckets by number and return
  const sortedKeys = Object.keys(buckets).map(Number).sort(function(a, b) { return a - b; });
  return sortedKeys.map(function(k) { return buckets[k]; });
}

// ---------------------------------------------------------------------------
// Setup
// ---------------------------------------------------------------------------

/**
 * Run once to create all required sheets
 */
function initialSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Create response sheet if needed
  if (!ss.getSheetByName(SHEET_NAME)) {
    createResponseSheet(ss);
  }

  // Create staff list if needed
  let staff = ss.getSheetByName(STAFF_SHEET);
  if (!staff) {
    staff = ss.insertSheet(STAFF_SHEET);
    staff.getRange("A1").setValue("Staff Names").setFontWeight("bold");
    staff.setColumnWidth(1, 200);
  }

  // Create settings sheet
  ensureSettingsSheet(ss);

  // Build summary sheets
  updatePaySummary(ss);
  updateFortnightSummary(ss);

  SpreadsheetApp.getUi().alert(
    "Setup complete! Add staff names to the Staff List tab, " +
    "set your fortnight cycle start in Settings, then deploy as web app."
  );
}
