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
// Public URL to the MSBU logo (PNG) — shown in the email header (fetched + inlined).
// Falls back to a clean "M | S | B | U" text wordmark if this ever fails to load.
var LOGO_URL = "https://msbu.co.id/hs-fs/hubfs/W_Primary%20with%20no%20Slogan_No%20Bg2x.png";

// Scorecard rating scale (shown as a legend at the bottom of each submission sheet).
var SCALE = [
  { icon: "👎👎", label: "Critical Deficit / Does Not Meet Expectations",
    desc: "The candidate significantly lacks the required skill. They provided poor examples, demonstrated no understanding, or their approach would cause operational issues." },
  { icon: "👎", label: "Needs Development / Below Expectations",
    desc: "The candidate has a superficial understanding or limited experience. They might need heavy hand-holding, extensive training, or struggled to articulate a clear success story." },
  { icon: "👍", label: "Competent / Meets Expectations",
    desc: "The candidate is fully capable. They answered the situational questions well, possess the necessary experience to do the job independently day one, and meet the baseline requirements." },
  { icon: "👍👍", label: "Expert / Exceeds Expectations",
    desc: "The candidate went above and beyond. They demonstrated advanced mastery, shared impressive metrics/results, showed strategic thinking, and could easily mentor others in this skill." }
];

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

    // ---- email (branded HTML, link button, no attachment) ----
    var recipients = (payload.recipients && payload.recipients.length)
      ? payload.recipients.join(",") : DEFAULT_RECIPIENTS;
    var logo = getLogoBlob();
    var mail = {
      to: recipients,
      subject: EMAIL_SUBJECT_PREFIX + " — " + candidate + " (" + (payload.recommendation || "n/a") + ")",
      htmlBody: buildEmailHtml(payload, perSS.getUrl(), !!logo)
    };
    if (logo) mail.inlineImages = { msbulogo: logo };
    MailApp.sendEmail(mail);

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
    score: { total: 10, max: 12, percent: 83 },
    answers: [
      { section: "Core Competencies", question: "Cross-functional empathy", type: "scorecard", value: "👍 Competent / Meets Expectations", desc: "Coordinated well with product and ops during the migration project." },
      { section: "Core Competencies", question: "Technical expertise",      type: "scorecard", value: "👍👍 Expert / Exceeds Expectations", desc: "Deep Spring Boot knowledge; walked through a clean, scalable design." },
      { section: "Core Competencies", question: "Adaptability",             type: "scorecard", value: "👍 Competent / Meets Expectations" },
      { section: "", question: "Why is the candidate interested in this role?",     type: "paragraph", value: "Wants a bigger backend challenge and to grow into full-stack." },
      { section: "", question: "Summary of relevant experience",                    type: "paragraph", value: "4+ years front-end (React/Redux); moving toward full-stack with Spring Boot." },
      { section: "", question: "Salary expectations & notice period",               type: "paragraph", value: "IDR 18-20 million; one month notice." },
      { section: "", question: "Availability, location / work arrangement",         type: "paragraph", value: "Jakarta; hybrid; available in ~4 weeks." },
      { section: "", question: "Any concerns or red flags?",                        type: "paragraph", value: "Limited large-scale backend experience - probe in the technical round." },
      { section: "", question: "General comments",                                  type: "paragraph", value: "Strong communicator and culture fit. Recommend advancing to the technical interview." }
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
    out["[" + a.section + "] " + a.question] = a.value + (a.desc ? " — " + a.desc : "");
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
  if (comps.length) { head("CORE COMPETENCIES"); comps.forEach(function (a) { dataRow(a.question, a.value + (a.desc ? "\n" + a.desc : "")); }); }
  var qs = (p.answers || []).filter(function (a) { return !a.section; });
  if (qs.length) { head("FEEDBACK"); qs.forEach(function (a) { dataRow(a.question, a.value); }); }
  if (comps.length) { head("SCORECARD RATING SCALE"); SCALE.forEach(function (s) { dataRow(s.icon + "  " + s.label, s.desc); }); }

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

function getLogoBlob() {
  if (!LOGO_URL) return null;
  try { return UrlFetchApp.fetch(LOGO_URL, { muteHttpExceptions: true }).getBlob().setName("logo"); }
  catch (e) { return null; }
}

function metaVal(p, label) { return (p.meta && p.meta[label] != null) ? p.meta[label] : ""; }

/* Branded, job-portal-style notification email with a CTA button. */
function buildEmailHtml(p, sheetUrl, hasLogo) {
  var NAVY = "#123047", INK = "#3b4652", MUTE = "#8a929c", LINE = "#eceef0", BG = "#f4f5f7";

  var candidate = metaVal(p, "Candidate name") || firstMetaValue(p) || "Candidate";
  var position  = metaVal(p, "Position applied for");
  var role      = p.role || "";
  var initials  = (candidate.split(/\s+/).map(function (w) { return w.charAt(0); }).join("").substring(0, 2) || "?").toUpperCase();

  function sectionHead(t) {
    return "<p style='margin:22px 0 8px;color:" + NAVY + ";font-weight:bold;font-size:13px;text-transform:uppercase;letter-spacing:.04em'>" + t + "</p>";
  }
  function hr() { return "<hr style='border:0;border-top:1px solid " + LINE + ";margin:22px 0'>"; }

  var logo = hasLogo
    ? "<img src='cid:msbulogo' alt='MSBU' style='height:32px'>"
    : "<span style='font-size:22px;font-weight:800;letter-spacing:3px;color:" + NAVY + "'>M&nbsp;|&nbsp;S&nbsp;|&nbsp;B&nbsp;|&nbsp;U</span>";

  var scoreTxt = p.score ? (" &nbsp;·&nbsp; Scorecard " + p.score.total + "/" + p.score.max + " (" + p.score.percent + "%)") : "";

  // candidate details
  var det = "";
  ["Candidate name", "Position applied for", "Interviewer name", "Interview date"].forEach(function (k) {
    var v = metaVal(p, k); if (v === "") return;
    det += "<tr><td style='padding:6px 0;color:" + MUTE + ";width:42%;font-size:14px;vertical-align:top'>" + esc(k) + "</td>"
         + "<td style='padding:6px 0;color:" + INK + ";font-weight:bold;font-size:14px'>" + esc(v) + "</td></tr>";
  });

  // core competencies
  var comps = (p.answers || []).filter(function (a) { return a.section; });
  var compHtml = "";
  if (comps.length) {
    compHtml += sectionHead("Core competencies") + "<table width='100%' style='border-collapse:collapse'>";
    comps.forEach(function (a) {
      compHtml += "<tr><td style='padding:6px 0;color:" + INK + ";font-size:14px'>" + esc(a.question) + "</td>"
               + "<td style='padding:6px 0;text-align:right;color:" + NAVY + ";font-weight:bold;font-size:14px'>" + esc(a.value || "—") + "</td></tr>";
    });
    compHtml += "</table>";
  }

  // feedback answers (resume-entry style)
  var qs = (p.answers || []).filter(function (a) { return !a.section && a.value; });
  var qHtml = "";
  if (qs.length) {
    qHtml += sectionHead("Feedback");
    qs.forEach(function (a) {
      qHtml += "<p style='margin:14px 0 4px;color:" + NAVY + ";font-weight:bold;font-size:15px'>" + esc(a.question) + "</p>"
             + "<p style='margin:0;color:" + INK + ";font-size:14px;line-height:1.6'>" + esc(a.value).replace(/\n/g, "<br>") + "</p>";
    });
  }

  return ""
    + "<div style='background:" + BG + ";padding:28px 12px;font-family:Arial,Helvetica,sans-serif'>"
      + "<div style='text-align:center;padding:6px 0 20px'>" + logo + "</div>"
      + "<table align='center' width='600' cellpadding='0' cellspacing='0' style='max-width:600px;width:100%;background:#ffffff;border:1px solid #e6e8eb;border-radius:14px'>"
        + "<tr><td style='padding:34px 38px'>"
          + "<div style='text-align:center'>"
            + "<div style='width:66px;height:66px;border-radius:50%;background:" + NAVY + ";color:#ffffff;font-size:24px;font-weight:bold;line-height:66px'>" + esc(initials) + "</div>"
          + "</div>"
          + "<h1 style='text-align:center;color:#5f6b7a;font-size:21px;line-height:1.35;margin:18px 0 4px'>Interview feedback for<br>" + esc(candidate) + "</h1>"
          + "<p style='text-align:center;color:" + MUTE + ";margin:0;font-size:14px'>" + esc(position || role) + (position && role ? (" &nbsp;·&nbsp; " + esc(role)) : "") + "</p>"
          + hr()
          + "<p style='margin:0 0 6px;color:" + NAVY + ";font-weight:bold;font-size:13px;text-transform:uppercase;letter-spacing:.04em'>Recommendation</p>"
          + "<p style='margin:0;color:" + INK + ";font-size:18px;font-weight:bold'>" + esc(p.recommendation || "—")
            + "<span style='font-weight:normal;font-size:14px;color:" + MUTE + "'>" + scoreTxt + "</span></p>"
          + sectionHead("Candidate details")
          + "<table width='100%' style='border-collapse:collapse'>" + det + "</table>"
          + compHtml
          + qHtml
          + "<div style='text-align:center;padding:30px 0 4px'>"
            + "<a href='" + sheetUrl + "' style='background:" + NAVY + ";color:#ffffff;text-decoration:none;font-weight:bold;font-size:15px;padding:14px 30px;border-radius:8px;display:inline-block'>View Candidate Feedback Form</a>"
          + "</div>"
        + "</td></tr>"
      + "</table>"
      + "<p style='text-align:center;color:#aab1b8;font-size:12px;margin:18px 0 0'>MSBU · Interview Feedback &nbsp;·&nbsp; automated notification</p>"
    + "</div>";
}

function esc(s) {
  return String(s == null ? "" : s).replace(/[&<>]/g, function (c) {
    return { "&": "&amp;", "<": "&lt;", ">": "&gt;" }[c];
  });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
