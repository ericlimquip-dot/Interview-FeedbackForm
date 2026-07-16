# MSBU Interview Feedback Form

A flexible, self-contained interview feedback form for hiring managers. Questions,
sections, scorecard scales, and the recommendation options are **fully editable**
(add / update / delete) — no code changes needed. On submit, the answers are saved
to a **Google Sheet** and **emailed as an Excel (.xlsx)** to the hiring team.

## Files

| File | Purpose |
|------|---------|
| `Interview Feedback Form.html` | The form. Open in any browser. Contains the builder + submit logic. |
| `apps-script/Code.gs` | Google Apps Script backend — writes to Google Sheets + sends email. |

## How the form works

Open `Interview Feedback Form.html` in a browser. A hiring manager fills it in and
clicks **Submit feedback**.

- **Edit form** (top-right) — passcode-protected builder. Add/edit/delete sections
  and questions, reorder them, change rating scales & weights. Default passcode: `msbu`
  (change it in **Settings**).
- **Settings** — form title, recommendation options (e.g. Strong Yes / Yes / No /
  Strong No), email recipients, the backend URL, builder passcode. You can also
  **export/import** the whole form as a JSON config file to share it or version it.
- The form config is saved in the browser's `localStorage`, so each machine keeps
  its own setup. Use **Export config** + commit the JSON if you want one shared source of truth.

Supported question types: scorecard (1–N), rating, short text, long text, Yes/No,
dropdown, multiple choice, and checkboxes. Scorecard/rating questions roll up into
a weighted total score shown at the bottom.

## Enabling auto Google Sheet + email (one-time setup)

A browser page can't email or create spreadsheets by itself, so the form posts to a
tiny Google Apps Script that does both under your Google account.

1. Go to **https://script.google.com** → **New project**.
2. Delete the placeholder code, paste the contents of `apps-script/Code.gs`, and save.
3. Click **Deploy → New deployment → Web app**:
   - **Execute as:** Me
   - **Who has access:** Anyone
4. Click **Deploy**, approve the authorization prompts (Sheets, Drive, Gmail).
5. Copy the **Web app URL** (ends in `/exec`).
6. In the form → **Settings → Backend URL**, paste that URL and **Save settings**.

Now every submission:
- appends a row to a master log sheet (**MSBU Interview Feedback — Master Log**),
- creates a per-candidate spreadsheet in a Drive folder, and
- emails the recipients with the `.xlsx` attached.

> Recipients are sent from the form (`erika@msbukonsultan.id, hr@msbu.co.id, e@msbu.co.id`
> by default). Change them any time in **Settings** — no redeploy needed.

### Fallback (no backend configured)

If **Backend URL** is left blank, Submit will instead **download a CSV** (opens in
Google Sheets/Excel) and **open a pre-filled email draft** to the recipients — so the
form is still usable before the Apps Script is deployed.

## Notes

- The `mode:"no-cors"` fetch means the browser won't show the script's JSON response;
  a success toast is shown optimistically. Check the master sheet to confirm receipt.
- To share one identical form setup across the team: configure it once, **Export
  config**, and have others **Import config** (or host the HTML on a shared drive /
  intranet / GitHub Pages).
