# MSBU Interview Feedback Forms

Two **fully independent**, self-contained feedback forms for the two stages of hiring:

- **Recruiter** — the initial screening call.
- **Hiring Manager** — the advance / technical interview.

On open you pick a **role**; you then see and submit **only that form** (a **Switch role**
button returns to the picker). The other form stays empty and sends nothing — so the
recruiter's submission and the hiring manager's submission are separate. Each form is
editable and fillable at once: edit question text inline, add/delete questions (the
Recruiter form can also drag-reorder), fill in answers, and submit. On submit the answers
are saved to a **Google Sheet** and the team is **emailed a summary with a link to that
submission's own tidy Google Sheet** (no attachment).

Each form contains: locked instructions, required **candidate details** (candidate name,
position, interviewer, date), a selectable + editable **evaluation** (Strong Hire → Strong
No Hire), a **Core Competencies** scorecard (thumb rating per row, add/remove rows), and
required paragraph questions.

## Files

| File | Purpose |
|------|---------|
| `index.html` | Both forms + role picker. Open in any browser. |
| `apps-script/Code.gs` | Google Apps Script backend — writes to Google Sheets + sends email. |

## How the forms work

Open `index.html` in a browser and choose **Recruiter** or
**Hiring Manager**. Fill the form and click **Submit feedback**.

- **Question text** is editable inline; **＋ Add another question** adds a row; the
  **⋯** menu on each question offers Duplicate / Delete. The Recruiter form supports
  drag-to-reorder via the ⠿ handle. All questions are required.
- **Evaluation** — click a row to select it; the labels are editable.
- **Core Competencies** — rate each row 👎👎 / 👎 / 👍 / 👍👍, edit names/descriptions,
  add or remove rows with **＋ Add another**.
- **Settings** — email recipients (default `hr@msbu.co.id`) and brand label. The Apps
  Script backend URL is baked into the file, so there's nothing to configure to make
  submissions work.
- Everything is saved in the browser's `localStorage`. Your typed *answers* are cleared
  after a successful submit; the *form setup* persists per machine.

## Backend (Google Apps Script)

The form posts each submission to a Google Apps Script web app (deployed under the MSBU
Google account) which:
- appends a **formatted row** to the master log **MSBU Interview Feedback — Master Log**
  (styled header, frozen top row, zebra striping),
- creates a **tidy, grouped per-submission spreadsheet** (Candidate details / Evaluation /
  Core Competencies / Feedback) in the Drive folder **MSBU Interview Feedback Submissions**, and
- emails the recipients a summary **with a link to that submission's sheet** (no attachment).

The web app URL is already baked into the HTML and access is set to **Anyone**, so the form
works with no per-user setup. Recipients are sent from the form (default `hr@msbu.co.id`) and
are editable in **Settings** with no redeploy.

**If you edit `apps-script/Code.gs`**, publish a new version so the change goes live:
**Deploy ▸ Manage deployments ▸ ✏ Edit ▸ Version: New version ▸ Deploy** (the `/exec` URL
stays the same). Verify the backend any time by opening the `/exec` URL — it returns
`{"ok":true,...}`. The editor's **`runTest`** function exercises the whole path
(sheet + folder + email) and surfaces real errors.

## Notes

- The submit uses a `no-cors` request, so the browser can't read the script's reply; a
  success toast is shown optimistically. Check the master sheet to confirm receipt.
- To share the forms with the team, host the HTML at one URL (e.g. GitHub Pages) or send
  the file — the backend URL and defaults are already baked in.
