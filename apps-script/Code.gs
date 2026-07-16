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
var DEFAULT_RECIPIENTS = "erika@msbukonsultan.id, hr@msbu.co.id, e@msbu.co.id";

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

    // ---- per-submission spreadsheet + xlsx export ----
    var candidate = firstMetaValue(payload) || "candidate";
    var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm");
    var copyName = "Feedback - " + candidate + " - " + stamp;

    var perSheet = SpreadsheetApp.create(copyName);
    var ps = perSheet.getActiveSheet();
    ps.getRange(1, 1, headers.length, 2).setValues(headers.map(function (h) { return [h, flat[h] || ""]; }));
    ps.setColumnWidth(1, 240); ps.setColumnWidth(2, 420);
    SpreadsheetApp.flush();

    var file = DriveApp.getFileById(perSheet.getId());
    var folder = getFolder();
    if (folder) { folder.addFile(file); DriveApp.getRootFolder().removeFile(file); }

    var xlsx = exportXlsx(perSheet.getId(), copyName);

    // ---- email ----
    var recipients = (payload.recipients && payload.recipients.length)
      ? payload.recipients.join(",") : DEFAULT_RECIPIENTS;
    MailApp.sendEmail({
      to: recipients,
      subject: EMAIL_SUBJECT_PREFIX + " — " + candidate + " (" + (payload.recommendation || "n/a") + ")",
      htmlBody: buildEmailHtml(payload, perSheet.getUrl()),
      attachments: [xlsx]
    });

    return json({ ok: true, sheet: perSheet.getUrl() });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

// Simple GET so you can verify the deployment in a browser
function doGet() {
  return json({ ok: true, service: "MSBU Interview Feedback backend is live." });
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

function exportXlsx(spreadsheetId, name) {
  var url = "https://docs.google.com/feeds/download/spreadsheets/Export?key="
          + spreadsheetId + "&exportFormat=xlsx";
  var resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  return resp.getBlob().setName(name + ".xlsx");
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
    + "<p style='color:#999;font-size:12px'>The .xlsx copy is attached. Full log kept in the master sheet.</p></div>";
}

function esc(s) {
  return String(s == null ? "" : s).replace(/[&<>]/g, function (c) {
    return { "&": "&amp;", "<": "&lt;", ">": "&gt;" }[c];
  });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
