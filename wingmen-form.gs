/**
 * Your Body's Wingmen: creates the Google Form, the response Sheet and a "Scores" tab.
 *
 * FIRST TIME: choose setupWingmenForm in the menu at the top, press Run, allow the permissions.
 * ALREADY MADE THE FORM? Choose wingmenLinks and press Run. It finds your form, finishes any
 *   setup that is missing, and prints the links again. It never makes a second form.
 *
 * The links print in the Execution log at the bottom of the screen. They are also saved
 * in a "Setup links" tab in the response spreadsheet.
 */
const FORM_TITLE = "Your Body's Wingmen: student results";
const FIELDS = [            // key must match the page. Do not rename the keys.
  ['name',    'Name', true],
  ['cls',     'Class', true],
  ['points',  'Points (out of 100)', true],
  ['parts',   'Parts finished (of 8)', false],
  ['acts',    'Activities done (of 11)', false],
  ['checks',  'Quick checks done (of 36)', false],
  ['first',   'Right first time (of 36)', false],
  ['explain', 'Explain-it boxes done (of 6)', false],
  ['answers', 'Explain-it answers', false],
  ['ver',     'Page version', false]
];
const SAMPLE = { name: 'Test Student', cls: '4A', points: '73.5', parts: '5', acts: '9', checks: '30',
                 first: '20', explain: '4', answers: 'Part 2: sample answer', ver: 'test' };

function setupWingmenForm() {
  const form = FormApp.create(FORM_TITLE);
  form.setDescription('Filled in automatically by the Wingmen page. Students do not need to open this.');
  form.setCollectEmail(false);
  form.setConfirmationMessage('Got it!');
  FIELDS.forEach(([key, title, req]) => {
    const it = (key === 'answers') ? form.addParagraphTextItem() : form.addTextItem();
    it.setTitle(title).setRequired(req);
  });
  PropertiesService.getScriptProperties().setProperty('FORM_ID', form.getId());
  wingmenLinks();
}

function wingmenLinks() {
  const props = PropertiesService.getScriptProperties();

  // 1. Find the form
  let form = null;
  const savedId = props.getProperty('FORM_ID');
  if (savedId) { try { form = FormApp.openById(savedId); } catch (e) {} }
  if (!form) {
    const files = DriveApp.searchFiles("title contains 'Wingmen: student results' and mimeType = 'application/vnd.google-apps.form' and trashed = false");
    if (!files.hasNext()) throw new Error('No Wingmen form found. Run setupWingmenForm first.');
    form = FormApp.openById(files.next().getId());
    props.setProperty('FORM_ID', form.getId());
  }
  // School accounts often limit forms to people signed in to the school. Students are not, so open it up.
  try { form.setRequireLogin(false); } catch (e) {}

  // 2. Make sure it has a response spreadsheet
  let ssId = props.getProperty('SHEET_ID');
  if (!ssId) { try { ssId = form.getDestinationId(); } catch (e) { ssId = null; } }
  if (!ssId) {
    const newSs = SpreadsheetApp.create("Your Body's Wingmen: responses");
    form.setDestination(FormApp.DestinationType.SPREADSHEET, newSs.getId());
    ssId = newSs.getId();
  }
  props.setProperty('SHEET_ID', ssId);
  const ss = SpreadsheetApp.openById(ssId);

  // 3. Make sure the Scores tab updates on every submission
  if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'refreshScores')) {
    ScriptApp.newTrigger('refreshScores').forForm(form).onFormSubmit().create();
  }

  // 4. Build the test link and read the entry numbers from it
  const byTitle = {};
  FIELDS.forEach(([key, title]) => { byTitle[title] = key; });
  const resp = form.createResponse();
  const order = [];
  form.getItems().forEach(item => {
    const key = byTitle[item.getTitle()];
    if (!key) return;
    const t = item.getType();
    const typed = (t === FormApp.ItemType.PARAGRAPH_TEXT) ? item.asParagraphTextItem()
                : (t === FormApp.ItemType.TEXT) ? item.asTextItem() : null;
    if (!typed) return;
    resp.withItemResponse(typed.createResponse(SAMPLE[key]));
    order.push(key);
  });
  const testUrl = resp.toPrefilledUrl();
  const found = [...testUrl.matchAll(/[?&](entry\.\d+)=/g)].map(m => m[1]);
  const entries = {};
  order.forEach((key, i) => { entries[key] = found[i]; });
  const action = form.getPublishedUrl().replace(/\/viewform.*$/, '/formResponse');
  const formLine = 'var FORM=' + JSON.stringify({ action: action, entries: entries }) + ';';

  // 5. Print them, and save them in the spreadsheet as a backup
  const lines = [
    ['1. Send this line to Claude (or paste it into the page)', formLine],
    ['2. One-click test: opens the form filled in. Press Submit.', testUrl],
    ['3. Response spreadsheet', ss.getUrl()],
    ['4. Form editor', form.getEditUrl()],
    ['5. After testing', 'Delete the "Test Student" row in the responses tab.']
  ];
  lines.forEach(([label, value]) => { Logger.log(label); Logger.log(value); });
  const tab = ss.getSheetByName('Setup links') || ss.insertSheet('Setup links');
  tab.clear();
  tab.getRange(1, 1, lines.length, 2).setValues(lines);
  tab.setColumnWidth(1, 380); tab.setColumnWidth(2, 700);
}

/** Builds/updates the "Scores" tab: each student's BEST points per name + class. Runs on every submission. */
function refreshScores() {
  const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'));
  const raw = ss.getSheets().find(s => s.getName().indexOf('Form Responses') === 0);
  if (!raw) return;
  const rows = raw.getDataRange().getValues().slice(1);        // A=time, B=name, C=class, D=points, E=parts ...
  const best = {};
  rows.forEach(r => {
    if (!r[1]) return;
    const key = String(r[1]).trim().toLowerCase() + '|' + String(r[2]).trim();
    const pts = parseFloat(r[3]) || 0;
    const n = ((best[key] || {}).n || 0) + 1;
    if (!best[key] || pts >= best[key].pts) {
      best[key] = { name: String(r[1]).trim(), cls: r[2], pts: pts, parts: r[4], acts: r[5], checks: r[6], first: r[7], explain: r[8], when: r[0], n: n };
    } else { best[key].n = n; }
  });
  const list = Object.values(best).sort((a, b) => String(a.cls).localeCompare(String(b.cls)) || a.name.localeCompare(b.name));
  const sc = ss.getSheetByName('Scores') || ss.insertSheet('Scores');
  sc.clear();
  const head = ['Class', 'Name', 'Best points', 'Parts finished (of 8)', 'Activities (of 11)', 'Quick checks (of 36)', 'Right first time', 'Explain-its (of 6)', 'Last sent', 'Times sent'];
  sc.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#EDE7F6');
  if (list.length) sc.getRange(2, 1, list.length, head.length).setValues(list.map(x => [x.cls, x.name, x.pts, x.parts, x.acts, x.checks, x.first, x.explain, x.when, x.n]));
  sc.setFrozenRows(1);
}
