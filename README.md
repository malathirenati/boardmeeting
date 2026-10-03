# Takshashila board dashboard

Circle leads fill in Google Sheets. Pressing **Sync** on the dashboard (or in the index sheet's menu) reads the sheets, moves any pictures inserted in them to the Pictures folder, and updates this password-protected site through a GitHub Action. It also runs every night. Each board meeting is a separate sheet; the latest is shown as current and earlier ones form the archive.

## How the pieces fit

```
Takshashila Board Meetings Index (Google Sheet)
  ├─ Meetings tab        one row per meeting: ID, date, sheet link, On dashboard Yes/No
  ├─ Circle leads tab    who gets edit access to the meeting sheets
  └─ menu "Board dashboard": Start a new board meeting · Sync the dashboard now · Share a meeting with circle leads · Turn on nightly sync
        │
        ▼
Takshashila Board Meeting yyyy-mm (Google Sheet, one per meeting)
  Guide · Meeting · Overview · Policy School · Research · Media · Network · Finance
  every lead can edit every tab
        │  Sync button on the dashboard, Sync in the sheet menu, or nightly
        ▼
.github/workflows/sync.yml → tools/publish.py   reads every sheet marked Yes, checks it, encrypts it
        ▼
index.html + data/manifest.json + data/<id>/meeting.json   served by GitHub Pages
```

## Who signs in where

| Person | Where | How they sign in | What they can do |
| --- | --- | --- | --- |
| Circle lead | Meeting sheet in Google Sheets | Their own Google account | Edit any tab (normally their own circle's) |
| Dashboard owner | Index sheet, meeting sheets, GitHub | Own Google and GitHub accounts | Start meetings, add leads, change the password |
| Board member | Dashboard | Shared login and password | View current and archived meetings |

Editing happens only in Google Sheets: a GitHub Pages site cannot save anything, because it has no server. Google accounts give each lead their own sign-in and a version history showing who changed what. The dashboard has two buttons for this: **Edit in Google Sheets ↗** opens the current meeting's sheet, and **Sync** brings the changes in.

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

1. **Upload the sheets.** Open the Drive folder. Drag in `Takshashila Board Meeting 2026-06.xlsx`, `Takshashila Board Meeting 2026-02.xlsx` and `Takshashila Board Meetings Index.xlsx`. Open each one and choose File → Save as Google Sheets, then delete the three `.xlsx` copies. (Or turn on Drive → Settings → "Convert uploads" before dragging, and they convert as they upload.)
2. **Link the meetings in the index.** In the index's Meetings tab, replace the two file names in column C with the links of the two meeting sheets (open each sheet → Share → Copy link).
3. **Upload the sample pictures.** Unzip `Takshashila-sample-pictures.zip` and drag the pictures (not the folder) into the `Pictures` folder. The sample sheets refer to them by file name. New pictures can simply be inserted into the sheet; Sync files them.
4. **Circle leads.** On the index's Circle leads tab, enter each circle's Google accounts and the dashboard owners. Everyone listed can edit every tab.
5. **Service account.** In Google Cloud: create a project, enable the Google Drive API, create a service account and download its JSON key. Share the Drive folder with the service account's email as Viewer.
6. **Repository files.** Upload everything in this zip to `malathirenati/boardmeeting`: on GitHub, Add file → Upload files, drag the unzipped files and folders in, commit. On a Mac, Finder hides the `.github` folder: press Cmd + Shift + . in Finder to show it before dragging. If the workflow does not appear under the Actions tab afterwards, create it by hand: Add file → Create new file, name it `.github/workflows/sync.yml`, paste the contents of that file from the zip, commit.
7. **Secrets and variables.** In the repository: Settings → Secrets and variables → Actions.
   - Secrets: `DASHBOARD_PASSWORD` (the dashboard password), `GOOGLE_SERVICE_ACCOUNT_JSON` (the whole JSON key).
   - Variables: `INDEX_SHEET_ID` (the ID in the index sheet's address, the part after `/d/`), `PICTURES_FOLDER_ID` = `1G18tsanwOVSjNlQzVTIVyvEI82ubyp-r`. `SYNC_URL` and the `SYNC_KEY` secret come in step 10.
8. **Pages.** Settings → Pages → Deploy from a branch → `main`, `/ (root)`. GitHub's free plan only publishes from a public repository; a private one needs GitHub Pro. Then Actions → Sync board meetings → Run workflow.
9. **Menu in the index sheet.** Extensions → Apps Script, paste `tools/Code.gs`, save. Project settings → Script properties: `GITHUB_REPO` = `malathirenati/boardmeeting`, `GITHUB_TOKEN` = a fine-grained token with Actions read and write on this repository only, `SYNC_KEY` = any long random text (keep a copy). Reload the sheet; the "Board dashboard" menu appears. Run **Share a meeting with circle leads…** for 2026-06 and 2026-02, then **Turn on nightly sync**.
10. **Sync button.** In Apps Script: Deploy → New deployment → Select type: Web app → Execute as: Me → Who has access: Anyone → Deploy, and approve the permissions. Copy the web app address. In the repository: add the variable `SYNC_URL` = that address and the secret `SYNC_KEY` = the same random text as in step 9. Run the workflow once more; the dashboard's Sync button then works. Anyone with the address and key could only start a sync; they cannot read or change data.

## Each board meeting

1. Anyone who can edit the index: **Board dashboard → Start a new board meeting…** and enter the meeting date and reporting window, of any length. The latest sheet is copied, its dates set, figures, pictures and notes cleared, shared with the circle leads, and the meeting listed with On dashboard = No.
2. Circle leads fill in the sheet, normally each their own tab; no tab is locked. The Guide tab and column K explain every row.
3. Set On dashboard = Yes, then press **Sync** on the dashboard. The new meeting becomes current; the previous one moves to the archive.

Rows the dashboard cannot use (an unknown row type, a missing picture) are skipped and listed in a notice at the top of the dashboard, so leads can fix them.

## Things to know

- **Nightly sync:** the index sheet's nightly trigger (step 9) collects pictures and starts the update. The repository's own nightly run is a fallback; in a public repository GitHub pauses it after 60 days without activity, which the sheet's trigger avoids.
- **Changing the password:** update the `DASHBOARD_PASSWORD` secret and run the workflow.
- **Pictures:** leads insert a picture into column D of a Picture row (Insert → Image → Image in cell). Sync saves it to the Pictures folder as `<meeting>-<panel>-row<n>.jpg` and writes that file name into the cell. Pictures floating over the sheet are not picked up. Typing the file name of a picture already in the Pictures folder also works.
- **Open editing:** any lead can change any tab. If a figure changes unexpectedly, File → Version history in the sheet shows who changed it and can restore the earlier version.
- **History:** Google Sheets keeps each sheet's version history (File → Version history). The repository keeps every synced version.

## Local test

```
pip install openpyxl pillow cryptography
DASHBOARD_PASSWORD=… python tools/publish.py --index "sheets/Takshashila Board Meetings Index.xlsx" --out . --embed tools/template.html --light
python -m http.server   # open http://localhost:8000
```

Keep the sample `.xlsx` files out of the repository: they are not encrypted.
