# TILN board dashboard

Circle leads fill in Google Sheets. A GitHub Action turns the sheets into this password-protected site, every night and whenever someone presses **Publish now**. Each board meeting is a separate sheet. The latest meeting is shown as current and earlier ones form the archive.

## How the pieces fit

```
TILN Board Meetings Index (Google Sheet)
  ├─ Meetings tab        one row per meeting: ID, date, sheet link, On dashboard Yes/No
  ├─ Circle leads tab    who gets edit access to the meeting sheets
  └─ menu "Board dashboard": Start a new board meeting · Publish now · Share a meeting with circle leads
        │
        ▼
TILN Board Meeting yyyy-mm (Google Sheet, one per meeting)
  Guide · Meeting · Overview · Policy School · Research · Media · Network · Finance
  every lead can edit every tab
        │  nightly, or Publish now
        ▼
.github/workflows/sync.yml → tools/publish.py   reads every sheet marked Yes, checks it, encrypts it
        ▼
index.html + data/manifest.json + data/<id>/meeting.json   served by GitHub Pages
```

## Who signs in where

| Person | Where | How they sign in | What they can do |
| --- | --- | --- | --- |
| Circle lead | Meeting sheet in Google Sheets | Their own Google account | Edit any tab (normally their own circle's) |
| Dashboard owner | Index sheet, meeting sheets, GitHub | Own Google and GitHub accounts | Start meetings, set permissions, publish, change the password |
| Board member | Dashboard | Shared login and password | View current and archived meetings |

Editing happens in Google Sheets because a GitHub Pages site cannot save anything: it has no server. Google accounts give each lead their own sign-in, and a version history showing who changed what, at no cost. The dashboard's Edit mode is a preview; its "Edit in the meeting sheet ↗" button takes a lead to the sheet.

## Decisions before setup

1. **Repository visibility.** On GitHub's free plan, Pages only publishes from a **public** repository. The data files and pictures are encrypted, but the code and the meeting dates are visible. A private repository needs GitHub Pro or Team. Either works; Team or Pro is cleaner.
2. **Viewer sign-in.** By default all board members share one login and password, which cannot be withdrawn from one person. For a sign-in per person (email one-time code, removable individually), put the site behind Cloudflare Access: deploy the same repository with Cloudflare Pages instead of GitHub Pages. Cloudflare's Zero Trust free plan covers up to 50 users.
3. **Ownership.** The Drive folder sits in one person's Drive. If that person leaves, the files go with their account. Moving the folder into a Google Shared Drive later keeps the same files and links working; the service account then needs to be a member of the Shared Drive.

## Where things live

| What | Where |
| --- | --- |
| Drive folder | [Board Meeting Dashboard 2026 onwards](https://drive.google.com/drive/folders/1IqBTSc2g3a9XObw3e2_bnmWXq64RWTZ8) (ID `1IqBTSc2g3a9XObw3e2_bnmWXq64RWTZ8`) |
| Pictures folder | `Pictures` inside it (ID `1G18tsanwOVSjNlQzVTIVyvEI82ubyp-r`) |
| Repository | `malathirenati/boardmeeting` |
| Dashboard address | `https://malathirenati.github.io/boardmeeting/` once Pages is on |

## One-time setup (dashboard owner, about an hour)

1. **Upload the sheets.** Open the Drive folder. Drag in `TILN Board Meeting 2026-06.xlsx`, `TILN Board Meeting 2026-02.xlsx` and `TILN Board Meetings Index.xlsx`. Open each one and choose File → Save as Google Sheets, then delete the three `.xlsx` copies. (Or turn on Drive → Settings → "Convert uploads" before dragging, and they convert as they upload.)
2. **Link the meetings in the index.** In the index's Meetings tab, replace the two file names in column C with the links of the two meeting sheets (open each sheet → Share → Copy link).
3. **Upload the pictures.** Unzip `TILN-sample-pictures.zip` and drag the pictures (not the folder) into the `Pictures` folder. The sheets already refer to them by file name, so no links need pasting.
4. **Circle leads.** On the index's Circle leads tab, enter each circle's Google accounts and the dashboard owners. Everyone listed can edit every tab.
5. **Service account.** In Google Cloud: create a project, enable the Google Drive API, create a service account and download its JSON key. Share the Drive folder with the service account's email as Viewer.
6. **Repository files.** Upload everything in this zip to `malathirenati/boardmeeting`: on GitHub, Add file → Upload files, drag the unzipped files and folders in, commit. On a Mac, Finder hides the `.github` folder: press Cmd + Shift + . in Finder to show it before dragging. If the workflow does not appear under the Actions tab afterwards, create it by hand: Add file → Create new file, name it `.github/workflows/sync.yml`, paste the contents of that file from the zip, commit.
7. **Secrets and variables.** In the repository: Settings → Secrets and variables → Actions.
   - Secrets: `DASHBOARD_PASSWORD` (the dashboard password), `GOOGLE_SERVICE_ACCOUNT_JSON` (the whole JSON key).
   - Variables: `INDEX_SHEET_ID` (the ID in the index sheet's address, the part after `/d/`), `PICTURES_FOLDER_ID` = `1G18tsanwOVSjNlQzVTIVyvEI82ubyp-r`.
8. **Pages.** Settings → Pages → Deploy from a branch → `main`, `/ (root)`. GitHub's free plan only publishes from a public repository; a private one needs GitHub Pro. Then Actions → Sync board meetings → Run workflow.
9. **Menu in the index sheet.** Extensions → Apps Script, paste `tools/Code.gs`, save. Project settings → Script properties: `GITHUB_REPO` = `malathirenati/boardmeeting`, `GITHUB_TOKEN` = a fine-grained token with Actions read and write on this repository only. Reload the sheet; the "Board dashboard" menu appears. Run **Share a meeting with circle leads…** for 2026-06 and 2026-02.

## Each board meeting

1. Anyone who can edit the index: **Board dashboard → Start a new board meeting…** and enter the meeting date and reporting window, of any length. The latest sheet is copied, its dates set, figures, pictures and notes cleared, shared with the circle leads, and the meeting listed with On dashboard = No.
2. Circle leads fill in the sheet, normally each their own tab; no tab is locked. The Guide tab and column K explain every row.
3. Set On dashboard = Yes, then **Board dashboard → Publish to the dashboard now**. The new meeting becomes current; the previous one moves to the archive.

Rows the dashboard cannot use (an unknown row type, a missing picture) are skipped and listed under "Sheet checks" in the dashboard's Edit mode, so leads can fix them.

## Things to know

- **Nightly sync in a public repository** stops after 60 days with no repository activity. Publish now always works, and any publish that changes data counts as activity. If the nightly run has stopped, re-enable it on the Actions tab.
- **Changing the password:** update the `DASHBOARD_PASSWORD` secret and run the workflow.
- **Pictures:** upload them to the Pictures folder and type the file name in the sheet's Picture row. A Drive link also works if the picture is shared with the service account.
- **Open editing:** any lead can change any tab. If a figure changes unexpectedly, File → Version history in the sheet shows who changed it and can restore the earlier version.
- **History:** Google Sheets keeps each sheet's version history (File → Version history). The repository keeps every published version.

## Local test

```
pip install openpyxl pillow cryptography
DASHBOARD_PASSWORD=… python tools/publish.py --index "sheets/TILN Board Meetings Index.xlsx" --out . --embed tools/template.html --light
python -m http.server   # open http://localhost:8000
```

Keep the sample `.xlsx` files out of the repository: they are not encrypted.
