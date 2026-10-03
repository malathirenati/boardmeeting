/**
 * Paste into the "Takshashila Board Meetings Index" Google Sheet: Extensions → Apps Script.
 * Script properties (Project settings → Script properties):
 *   GITHUB_REPO        malathirenati/boardmeeting
 *   GITHUB_TOKEN       fine-grained token with "Actions: read and write" on that one repository
 *   SYNC_KEY           any long random text; the same value goes into the repository secret SYNC_KEY
 *   PICTURES_FOLDER_ID optional; defaults to the Pictures folder below
 * Then Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone) so the dashboard's Sync button can reach doGet.
 */
const CFG = PropertiesService.getScriptProperties();
const DEFAULT_PICTURES_FOLDER = '1G18tsanwOVSjNlQzVTIVyvEI82ubyp-r';
const CIRCLE_TABS = ['Overview', 'Policy School', 'Research', 'Media', 'Network', 'Finance'];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Board dashboard')
    .addItem('Start a new board meeting…', 'newMeeting')
    .addItem('Sync the dashboard now', 'publishNow')
    .addSeparator()
    .addItem('Share a meeting with circle leads…', 'applyPermissionsPrompt')
    .addItem('Turn on nightly sync', 'setupNightlySync')
    .addToUi();
}

/** Copies the latest meeting sheet, sets its dates, clears figures and pictures, and lists it (not yet on the dashboard). */
function newMeeting() {
  const ui = SpreadsheetApp.getUi();
  const ask = q => { const r = ui.prompt('New board meeting', q, ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK) throw new Error('cancelled'); return r.getResponseText().trim(); };
  let date, from, to;
  try { date = ask('Meeting date (yyyy-mm-dd)'); from = ask('Reporting window from (yyyy-mm-dd)'); to = ask('Reporting window to (yyyy-mm-dd)'); }
  catch (e) { return; }
  const re = /^\d{4}-\d{2}-\d{2}$/;
  if (![date, from, to].every(s => re.test(s)) || from > to) { ui.alert('Use yyyy-mm-dd dates, with the window starting before it ends.'); return; }
  const id = date.slice(0, 7);
  const sh = SpreadsheetApp.getActive().getSheetByName('Meetings');
  const rows = sh.getDataRange().getValues().slice(1).filter(r => r[0] && r[2] && /^\d{4}/.test(String(r[0])));
  if (rows.some(r => String(r[0]) === id)) { ui.alert('A meeting ' + id + ' is already listed.'); return; }
  if (!rows.length) { ui.alert('List at least one meeting sheet first.'); return; }
  rows.sort((a, b) => new Date(a[1]) - new Date(b[1]));
  const last = rows[rows.length - 1];
  const src = DriveApp.getFileById(idFrom(String(last[2])));
  const parents = src.getParents();
  const copy = src.makeCopy('Takshashila Board Meeting ' + id, parents.hasNext() ? parents.next() : DriveApp.getRootFolder());
  const ss = SpreadsheetApp.openById(copy.getId());
  const m = ss.getSheetByName('Meeting');
  m.getRange(2, 1, m.getLastRow() - 1, 2).getValues().forEach((r, i) => {
    const k = String(r[0]).toLowerCase(), cell = m.getRange(i + 2, 2);
    if (k === 'meeting id') cell.setValue(id);
    if (k === 'meeting date') cell.setValue(new Date(date));
    if (k === 'reporting window from') cell.setValue(new Date(from));
    if (k === 'reporting window to') cell.setValue(new Date(to));
  });
  ['Overview', 'Policy School', 'Research', 'Media', 'Network', 'Finance'].forEach(n => {
    const t = ss.getSheetByName(n); if (!t || t.getLastRow() < 2) return;
    t.getRange(2, 1, t.getLastRow() - 1, 3).getValues().forEach((r, i) => {
      const type = String(r[1]).toLowerCase(), row = i + 2;
      if (type === 'figure') t.getRange(row, 4, 1, 2).clearContent();          // value and context
      if (type === 'picture') t.getRange(row, 4).clearContent();               // image link
      if (type === 'note' || type === 'needs checking') t.getRange(row, 3).clearContent();
      // Charts and text are kept: leads add the new period's rows and rewrite points.
    });
  });
  sh.appendRow([id, new Date(date), copy.getUrl(), 'No', 'Copied from ' + last[0] + ' on ' + new Date().toDateString()]);
  applyPermissions(ss);
  ui.alert('Created "Takshashila Board Meeting ' + id + '". Everyone on the Circle leads tab can now edit it. Set On dashboard to Yes when every tab is filled in, then Sync.');
}

/** Menu: moves inserted pictures to the Pictures folder, then starts the GitHub update. */
function publishNow() {
  const r = syncAll();
  SpreadsheetApp.getUi().alert(r.ok
    ? 'Syncing. ' + (r.pictures ? r.pictures + ' picture(s) moved to the Pictures folder. ' : '') + 'The dashboard updates in about two minutes.'
    : 'Sync failed: ' + r.error);
}

/** Called by the dashboard's Sync button (web app). */
function doGet(e) {
  const key = CFG.getProperty('SYNC_KEY');
  if (!key || !e || !e.parameter || e.parameter.key !== key) return json_({ ok: false, error: 'not allowed' });
  return json_(syncAll());
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/** Collects pictures from every listed meeting sheet, then triggers the GitHub Action. Also run nightly. */
function syncAll() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { ok: true, pictures: 0, note: 'A sync is already running.' };
  try {
    let pictures = 0;
    const sh = SpreadsheetApp.openById(indexId_()).getSheetByName('Meetings');
    sh.getDataRange().getValues().slice(1).filter(r => r[0] && r[2] && /^\d{4}/.test(String(r[0]))).forEach(r => {
      try { pictures += collectPictures(SpreadsheetApp.openById(idFrom(String(r[2]))), String(r[0])); }
      catch (err) { console.warn('Pictures for ' + r[0] + ': ' + err); }
    });
    const code = triggerWorkflow_();
    return code === 204 ? { ok: true, pictures } : { ok: false, pictures, error: 'GitHub answered ' + code + '. Check GITHUB_REPO and GITHUB_TOKEN.' };
  } finally { lock.releaseLock(); }
}

/** The index sheet's own ID, also when run from a trigger or the web app. */
function indexId_() {
  let id = CFG.getProperty('INDEX_SHEET_ID');
  if (!id) { id = SpreadsheetApp.getActive().getId(); CFG.setProperty('INDEX_SHEET_ID', id); }
  return id;
}

function triggerWorkflow_() {
  const repo = CFG.getProperty('GITHUB_REPO'), token = CFG.getProperty('GITHUB_TOKEN');
  if (!repo || !token) return 0;
  const r = UrlFetchApp.fetch('https://api.github.com/repos/' + repo + '/actions/workflows/sync.yml/dispatches', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + token, Accept: 'application/vnd.github+json' },
    payload: JSON.stringify({ ref: 'main' })
  });
  return r.getResponseCode();
}

/**
 * Pictures inserted into a cell (Insert → Image → Image in cell) in column D of a Picture row are saved
 * to the Pictures folder as <meeting>-<panel>-row<n>.<ext>, and the cell is replaced by that file name.
 */
function collectPictures(ss, meetingId) {
  const folder = DriveApp.getFolderById(CFG.getProperty('PICTURES_FOLDER_ID') || DEFAULT_PICTURES_FOLDER);
  let moved = 0;
  CIRCLE_TABS.forEach(name => {
    const t = ss.getSheetByName(name); if (!t || t.getLastRow() < 5) return;
    const vals = t.getRange(1, 1, t.getLastRow(), 4).getValues();
    vals.forEach((r, i) => {
      if (String(r[1]).trim().toLowerCase() !== 'picture') return;
      const v = r[3];
      if (!v || typeof v !== 'object' || v.valueType !== SpreadsheetApp.ValueType.IMAGE) return;
      const url = v.getContentUrl(); if (!url) return;
      const blob = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob();
      const ext = String(blob.getContentType() || 'image/jpeg').split('/')[1].replace('jpeg', 'jpg');
      const file = meetingId + '-' + String(r[0]).replace(/[^A-Za-z0-9.]/g, '') + '-row' + (i + 1) + '.' + ext;
      folder.createFile(blob.setName(file));
      t.getRange(i + 1, 4).setValue(file);
      moved++;
    });
  });
  return moved;
}

/** Menu: runs syncAll every night at about 01:00 (the sheet's time zone). */
function setupNightlySync() {
  indexId_();
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'syncAll').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('syncAll').timeBased().everyDays(1).atHour(1).create();
  SpreadsheetApp.getUi().alert('Nightly sync is on. Pictures are collected and the dashboard updated every night around 1 am.');
}

/** Reads the Circle leads tab: { 'Policy School': ['a@x.org', …], … } */
function circleLeads() {
  const t = SpreadsheetApp.getActive().getSheetByName('Circle leads'); const out = {};
  if (!t) return out;
  t.getRange(2, 1, Math.max(1, t.getLastRow() - 1), 2).getValues().forEach(r => {
    if (!r[0]) return;
    out[String(r[0]).trim()] = String(r[1] || '').split(/[,;\s]+/).map(e => e.trim().toLowerCase()).filter(e => e.indexOf('@') > 0);
  });
  return out;
}

/**
 * Gives everyone on the Circle leads tab edit access to the whole meeting sheet.
 * No tab is locked: any lead can edit any tab. Removes tab locks left by earlier versions of this script.
 */
function applyPermissions(ss) {
  const leads = circleLeads();
  const everyone = [...new Set(Object.values(leads).flat())];
  const file = DriveApp.getFileById(ss.getId());
  const existing = file.getEditors().map(u => u.getEmail().toLowerCase());
  const toAdd = everyone.filter(e => existing.indexOf(e) < 0);
  if (toAdd.length) file.addEditors(toAdd);
  ss.getSheets().forEach(sheet => sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove()));
}

function applyPermissionsPrompt() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Share with circle leads', 'Meeting ID (yyyy-mm) as listed on the Meetings tab:', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const id = r.getResponseText().trim();
  const row = SpreadsheetApp.getActive().getSheetByName('Meetings').getDataRange().getValues().find(x => String(x[0]) === id);
  if (!row) { ui.alert('No meeting ' + id + ' on the Meetings tab.'); return; }
  applyPermissions(SpreadsheetApp.openById(idFrom(String(row[2]))));
  ui.alert('Done. Everyone on the Circle leads tab can now edit every tab of ' + id + '.');
}

function idFrom(s) { const m = s.match(/\/d\/([A-Za-z0-9_-]{20,})/); return m ? m[1] : s; }
