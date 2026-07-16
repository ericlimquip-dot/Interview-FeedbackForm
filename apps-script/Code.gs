/**
 * MSBU Interview Feedback — Google Apps Script backend
 * =====================================================
 * What it does when the HTML form POSTs a submission:
 *   1. Appends the submission as a row in a Google Sheet (one row per submission).
 *   2. Creates a per-submission Google Spreadsheet copy in a Drive folder.
 *   3. Emails erika@ / hr@ / e@ (whatever the form sends) with the .xlsx attached.
 *
 * SETUP (once):
 *   1. Go to https://script.google.com  ->  New project. Paste this file.
 *   2. (Optional) set MASTER_SHEET_ID / FOLDER_ID below, or leave blank to auto-create.
 *   3. Deploy  ->  New deployment  ->  type "Web app".
 *        - Execute as: Me
 *        - Who has access: Anyone
 *   4. Copy the /exec URL, paste it into the form's Settings -> "Backend URL".
 *   5. First run will ask for authorization (Sheets, Drive, Gmail) — approve it.
 */

// ---- Optional fixed targets (leave "" to auto-create + remember) ----
var MASTER_SHEET_ID = "";   // master log spreadsheet
var FOLDER_ID       = "";   // Drive folder for per-submission files
var EMAIL_SUBJECT_PREFIX = "Interview Feedback";
// Recipients fallback if the form doesn't send any:
var DEFAULT_RECIPIENTS = "hr@msbu.co.id";

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var sheet = getMasterSheet();
    var flat = flatten(payload);

    // ---- ensure header row matches this submission's columns ----
    var headers = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
    if (sheet.getLastRow() === 0 || headers.join("") === "") {
      headers = Object.keys(flat);
      sheet.appendRow(headers);
    } else {
      // add any new columns that appeared (form questions can change over time)
      Object.keys(flat).forEach(function (k) {
        if (headers.indexOf(k) === -1) {
          headers.push(k);
          sheet.getRange(1, headers.length).setValue(k);
        }
      });
    }
    var row = headers.map(function (h) { return flat[h] !== undefined ? flat[h] : ""; });
    sheet.appendRow(row);
    formatMaster(sheet);

    // ---- per-submission spreadsheet (tidy, readable layout) ----
    var candidate = firstMetaValue(payload) || "candidate";
    var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm");
    var copyName = "Feedback - " + candidate + " - " + stamp;

    var perSS = SpreadsheetApp.create(copyName);
    writePerSubmission(perSS.getActiveSheet(), payload);
    SpreadsheetApp.flush();

    var file = DriveApp.getFileById(perSS.getId());
    var folder = getFolder();
    if (folder) { folder.addFile(file); DriveApp.getRootFolder().removeFile(file); }

    // ---- email (spreadsheet link only, no attachment) ----
    var recipients = (payload.recipients && payload.recipients.length)
      ? payload.recipients.join(",") : DEFAULT_RECIPIENTS;
    MailApp.sendEmail({
      to: recipients,
      subject: EMAIL_SUBJECT_PREFIX + " — " + candidate + " (" + (payload.recommendation || "n/a") + ")",
      htmlBody: buildEmailHtml(payload, perSS.getUrl())
    });

    return json({ ok: true, sheet: perSS.getUrl() });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

// Simple GET so you can verify the deployment in a browser
function doGet() {
  return json({ ok: true, service: "MSBU Interview Feedback backend is live." });
}

/**
 * DIAGNOSTIC — run this straight from the editor (Run ▸ runTest).
 * It calls doPost with a fake submission, so it exercises the whole path
 * (create Sheet -> create folder -> export xlsx -> send email) and shows
 * any REAL error in the execution log — unlike the form, which hides errors.
 * Sends the test email only to the address below; change it if you like.
 */
function runTest() {
  var fake = { postData: { contents: JSON.stringify({
    formTitle: "TEST — Recruiter (diagnostic)",
    role: "Recruiter",
    submittedAt: new Date().toISOString(),
    meta: {
      "Candidate name": "Test Candidate",
      "Position applied for": "QA Engineer",
      "Interviewer name": "Eric",
      "Interview date": "2026-07-16"
    },
    recommendation: "Hire",
    evalScore: 3,
    score: { total: 9, max: 12, percent: 75 },
    answers: [
      { section: "Core Competencies", question: "Technical expertise", type: "scorecard", value: "Hire" },
      { section: "", question: "General comments", type: "paragraph", value: "This is a diagnostic test submission." }
    ],
    recipients: ["e@msbu.co.id"]
  }) } };
  var out = doPost(fake);
  Logger.log(out.getContent());   // {"ok":true,...} on success
}

/* ---------------- helpers ---------------- */

function flatten(p) {
  var out = {};
  out["Submitted At"] = p.submittedAt || new Date().toISOString();
  out["Form"] = p.formTitle || "";
  Object.keys(p.meta || {}).forEach(function (k) { out[k] = p.meta[k]; });
  out["Overall Recommendation"] = p.recommendation || "";
  if (p.score) out["Score"] = p.score.total + "/" + p.score.max + " (" + p.score.percent + "%)";
  (p.answers || []).forEach(function (a) {
    out["[" + a.section + "] " + a.question] = a.value;
  });
  out["Summary Comments"] = p.overallComments || "";
  return out;
}

function firstMetaValue(p) {
  var m = p.meta || {};
  var keys = Object.keys(m);
  return keys.length ? m[keys[0]] : "";
}

function getMasterSheet() {
  var ss;
  if (MASTER_SHEET_ID) { ss = SpreadsheetApp.openById(MASTER_SHEET_ID); }
  else {
    var props = PropertiesService.getScriptProperties();
    var saved = props.getProperty("MASTER_SHEET_ID");
    if (saved) { ss = SpreadsheetApp.openById(saved); }
    else {
      ss = SpreadsheetApp.create("MSBU Interview Feedback — Master Log");
      props.setProperty("MASTER_SHEET_ID", ss.getId());
    }
  }
  return ss.getActiveSheet();
}

function getFolder() {
  if (FOLDER_ID) return DriveApp.getFolderById(FOLDER_ID);
  var props = PropertiesService.getScriptProperties();
  var saved = props.getProperty("FOLDER_ID");
  if (saved) return DriveApp.getFolderById(saved);
  var f = DriveApp.createFolder("MSBU Interview Feedback Submissions");
  props.setProperty("FOLDER_ID", f.getId());
  return f;
}

/* Tidy the master log: styled header, frozen top row, zebra striping, sane widths. */
function formatMaster(sheet) {
  var lastCol = sheet.getLastColumn(), lastRow = sheet.getLastRow();
  if (lastCol < 1 || lastRow < 1) return;
  var all = sheet.getRange(1, 1, lastRow, lastCol);
  all.setVerticalAlignment("top").setWrap(true).setFontFamily("Arial");
  try { if (sheet.getBandings().length === 0) all.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREEN, true, false); } catch (e) {}
  var header = sheet.getRange(1, 1, 1, lastCol);
  header.setFontWeight("bold").setFontColor("#ffffff").setBackground("#548235").setVerticalAlignment("middle");
  sheet.setRowHeight(1, 32);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, lastCol);
  for (var c = 1; c <= lastCol; c++) {
    var w = sheet.getColumnWidth(c);
    if (w > 340) sheet.setColumnWidth(c, 340);
    else if (w < 100) sheet.setColumnWidth(c, 100);
  }
}

/* Write one submission styled like a clean report template:
   a big title banner, a summary block with a mini score bar, and green
   section headers over zebra-striped rows. */
function writePerSubmission(sheet, p) {
  sheet.setName("Feedback");
  sheet.clear();

  var TITLE_BG = "#e2efda", TITLE_TX = "#375623";   // light-green banner / dark-green text
  var HEAD_BG  = "#548235", HEAD_TX  = "#ffffff";   // dark-green section header
  var BAND     = "#eaf1df", BORDER   = "#cfe0bd", BAR = "#70ad47";

  var rows = [], kind = [];
  function push(a, b, k) { rows.push([a, (b == null ? "" : b)]); kind.push(k); return rows.length; }
  var stripe = 0;
  function dataRow(a, b) { stripe++; push(a, b, stripe % 2 ? "a" : "b"); }
  function head(t) { push(t, "", "head"); stripe = 0; }

  push(p.formTitle || "Interview Feedback", "", "title");
  push("Overall recommendation", p.recommendation || "—", "sum");
  if (p.score) {
    var pct = p.score.percent, blocks = Math.max(1, Math.round(pct / 5));  // up to 20 blocks
    push("Scorecard", p.score.total + " / " + p.score.max + "   (" + pct + "%)   " + rep("█", blocks), "bar");
  }
  push("Submitted", fmtDate(p.submittedAt), "meta");

  head("CANDIDATE DETAILS");
  Object.keys(p.meta || {}).forEach(function (k) { dataRow(k, p.meta[k]); });
  var comps = (p.answers || []).filter(function (a) { return a.section; });
  if (comps.length) { head("CORE COMPETENCIES"); comps.forEach(function (a) { dataRow(a.question, a.value); }); }
  var qs = (p.answers || []).filter(function (a) { return !a.section; });
  if (qs.length) { head("FEEDBACK"); qs.forEach(function (a) { dataRow(a.question, a.value); }); }

  var n = rows.length;
  sheet.getRange(1, 1, n, 2).setValues(rows);
  sheet.setColumnWidth(1, 280); sheet.setColumnWidth(2, 540);
  var body = sheet.getRange(1, 1, n, 2);
  body.setVerticalAlignment("top").setWrap(true).setFontFamily("Arial").setFontSize(11);
  body.setBorder(true, true, true, true, true, true, BORDER, SpreadsheetApp.BorderStyle.SOLID);

  for (var i = 0; i < n; i++) {
    var r = i + 1, k = kind[i], rng = sheet.getRange(r, 1, 1, 2);
    if (k === "title") {
      rng.merge().setBackground(TITLE_BG).setFontColor(TITLE_TX).setFontSize(22).setFontWeight("bold").setVerticalAlignment("middle");
      sheet.setRowHeight(r, 56);
    } else if (k === "head") {
      rng.merge().setBackground(HEAD_BG).setFontColor(HEAD_TX).setFontWeight("bold").setFontSize(11).setVerticalAlignment("middle");
      sheet.setRowHeight(r, 28);
    } else if (k === "sum") {
      sheet.getRange(r, 1).setFontWeight("bold").setFontColor(TITLE_TX);
      sheet.getRange(r, 2).setFontWeight("bold").setFontSize(13).setFontColor(TITLE_TX);
      sheet.setRowHeight(r, 26);
    } else if (k === "bar") {
      sheet.getRange(r, 1).setFontWeight("bold").setFontColor(TITLE_TX);
      sheet.getRange(r, 2).setFontColor(BAR).setFontWeight("bold");
      sheet.setRowHeight(r, 24);
    } else if (k === "meta") {
      rng.setFontColor("#8a8f98").setFontSize(10);
    } else {                                   // data rows
      sheet.getRange(r, 1).setFontWeight("bold").setFontColor(TITLE_TX);
      if (k === "b") rng.setBackground(BAND);  // zebra stripe
    }
  }
  sheet.setFrozenRows(1);
}

function rep(ch, n) { var s = ""; for (var i = 0; i < n; i++) s += ch; return s; }

function fmtDate(iso) {
  try { return Utilities.formatDate(new Date(iso), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"); }
  catch (e) { return iso || ""; }
}

function buildEmailHtml(p, sheetUrl) {
  var rows = "";
  Object.keys(p.meta || {}).forEach(function (k) {
    rows += "<tr><td style='padding:4px 10px;color:#555'>" + k + "</td><td style='padding:4px 10px'><b>" + esc(p.meta[k]) + "</b></td></tr>";
  });
  var ans = "";
  (p.answers || []).forEach(function (a) {
    if (!a.value) return;
    ans += "<tr><td style='padding:4px 10px;color:#555'>" + esc(a.question) + "</td><td style='padding:4px 10px'>" + esc(a.value) + "</td></tr>";
  });
  var scoreLine = p.score ? "<p><b>Score:</b> " + p.score.total + "/" + p.score.max + " (" + p.score.percent + "%)</p>" : "";
  return "<div style='font-family:Arial,sans-serif;color:#1f2933'>"
    + "<h2 style='margin:0 0 6px'>" + esc(p.formTitle || "Interview Feedback") + "</h2>"
    + "<p style='margin:0 0 14px'><b>Recommendation:</b> " + esc(p.recommendation || "n/a") + "</p>"
    + scoreLine
    + "<table style='border-collapse:collapse;font-size:14px'>" + rows + ans + "</table>"
    + (p.overallComments ? "<p style='margin-top:14px'><b>Summary:</b><br>" + esc(p.overallComments) + "</p>" : "")
    + "<p style='margin-top:16px'><a href='" + sheetUrl + "'>Open the Google Sheet for this submission</a></p>"
    + "<p style='color:#999;font-size:12px'>Full log kept in the master sheet.</p></div>";
}

function esc(s) {
  return String(s == null ? "" : s).replace(/[&<>]/g, function (c) {
    return { "&": "&amp;", "<": "&lt;", ">": "&gt;" }[c];
  });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
