/**
 * Paste into the "TILN Board Meetings Index" Google Sheet: Extensions → Apps Script.
 * Script properties (Project settings): GITHUB_REPO = owner/name, GITHUB_TOKEN = fine-grained token
 * with "Actions: read and write" on that one repository.
 */
const CFG = PropertiesService.getScriptProperties();

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Board dashboard')
    .addItem('Start a new board meeting…', 'newMeeting')
    .addItem('Publish to the dashboard now', 'publishNow')
    .addSeparator()
    .addItem('Share a meeting with circle leads…', 'applyPermissionsPrompt')
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
  const copy = src.makeCopy('TILN Board Meeting ' + id, parents.hasNext() ? parents.next() : DriveApp.getRootFolder());
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
  ui.alert('Created "TILN Board Meeting ' + id + '". Everyone on the Circle leads tab can now edit it. Set On dashboard to Yes when every tab is filled in, then publish.');
}

/** Starts the GitHub Action that reads the sheets and updates the dashboard. */
function publishNow() {
  const repo = CFG.getProperty('GITHUB_REPO'), token = CFG.getProperty('GITHUB_TOKEN');
  if (!repo || !token) { SpreadsheetApp.getUi().alert('Set GITHUB_REPO and GITHUB_TOKEN under Extensions → Apps Script → Project settings → Script properties.'); return; }
  const r = UrlFetchApp.fetch('https://api.github.com/repos/' + repo + '/actions/workflows/sync.yml/dispatches', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + token, Accept: 'application/vnd.github+json' },
    payload: JSON.stringify({ ref: 'main' })
  });
  SpreadsheetApp.getUi().alert(r.getResponseCode() === 204
    ? 'Publishing. The dashboard updates in about two minutes.'
    : 'Publishing failed (' + r.getResponseCode() + '): ' + r.getContentText().slice(0, 200));
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
