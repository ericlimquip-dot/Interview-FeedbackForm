# MSBU Interview Feedback Forms

Two **fully independent**, self-contained feedback forms for the two stages of hiring:

- **Recruiter** — the initial screening call.
- **Hiring Manager** — the advance / technical interview.

On open you pick a **role**; you then see and submit **only that form** (a **Switch role**
button returns to the picker). The other form stays empty and sends nothing — so the
recruiter's submission and the hiring manager's submission are separate. Each form is
editable and fillable at once: edit question text inline, add/delete questions (the
Recruiter form can also drag-reorder), fill in answers, and submit. On submit the answers
are saved to a **Google Sheet** and **emailed as an Excel (.xlsx)** to the team.

Each form contains: locked instructions, required **candidate details** (with an
Initial/Advance **Interview stage** dropdown), a selectable + editable **evaluation**
(Strong Hire → Strong No Hire), a **Core Competencies** scorecard (thumb rating per row,
add/remove rows), and required paragraph questions.

## Files

| File | Purpose |
|------|---------|
| `Interview Feedback Form.html` | Both forms + role picker. Open in any browser. |
| `apps-script/Code.gs` | Google Apps Script backend — writes to Google Sheets + sends email. |

## How the forms work

Open `Interview Feedback Form.html` in a browser and choose **Recruiter** or
**Hiring Manager**. Fill the form and click **Submit feedback**.

- **Question text** is editable inline; **＋ Add another question** adds a row; the
  **⋯** menu on each question offers Duplicate / Delete. The Recruiter form supports
  drag-to-reorder via the ⠿ handle. All questions are required.
- **Evaluation** — click a row to select it; the labels are editable.
- **Core Competencies** — rate each row 👎👎 / 👎 / 👍 / 👍👍, edit names/descriptions,
  add or remove rows with **＋ Add another**.
- **Settings** — email recipients, the backend URL, brand label, and **export/import**
  the whole configuration (both forms) as JSON.
- Everything is saved in the browser's `localStorage`. Your typed *answers* are cleared
  after a successful submit; the *form setup* persists. Use **Export config** + commit
  the JSON if you want one shared source of truth across machines.

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
