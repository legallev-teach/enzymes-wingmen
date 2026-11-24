/**
 * Your Body's Wingmen: creates the Google Form, the response Sheet and a "Scores" tab.
 * HOW TO USE: script.google.com > New project > paste this > choose setupWingmenForm > Run.
 * Allow the permissions. Then open View > Logs (Execution log) and copy what it prints.
 */
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

function setupWingmenForm() {
  const form = FormApp.create("Your Body's Wingmen: student results");
  form.setDescription('Filled in automatically by the Wingmen page. Students do not need to open this.');
  form.setCollectEmail(false);
  try { form.setRequireLogin(false); } catch (e) {}
  form.setConfirmationMessage('Got it!');

  const items = {};
  FIELDS.forEach(([key, title, req]) => {
    const it = (key === 'answers') ? form.addParagraphTextItem() : form.addTextItem();
    it.setTitle(title).setRequired(req);
    items[key] = it;
  });

  // Sample response -> prefilled link -> entry IDs
  const sample = { name: 'Test Student', cls: '4A', points: '73.5', parts: '5', acts: '9', checks: '30',
                   first: '20', explain: '4', answers: 'Part 2: sample answer', ver: 'test' };
  const resp = form.createResponse();
  FIELDS.forEach(([key]) => resp.withItemResponse(items[key].createResponse(sample[key])));
  const testUrl = resp.toPrefilledUrl();
  const entries = {};
  const found = [...testUrl.matchAll(/[?&](entry\.\d+)=/g)].map(m => m[1]);
  // Prefilled URL lists the entries in item order, so zip them with FIELDS
  FIELDS.forEach(([key], i) => { entries[key] = found[i]; });

  const action = form.getPublishedUrl().replace('/viewform', '/formResponse');

  const ss = SpreadsheetApp.create("Your Body's Wingmen: responses");
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  const props = PropertiesService.getScriptProperties();
  props.setProperty('FORM_ID', form.getId());
  props.setProperty('SHEET_ID', ss.getId());
  ScriptApp.newTrigger('refreshScores').forForm(form).onFormSubmit().create();

  Logger.log('=== 1. PASTE THIS IN config.js (replace the line starting var FORM=) ===');
  Logger.log('var FORM=' + JSON.stringify({ action: action, entries: entries }) + ';');
  Logger.log('=== 2. ONE-CLICK TEST LINK (opens the form already filled in; press Submit) ===');
  Logger.log(testUrl);
  Logger.log('=== 3. YOUR SPREADSHEET ===');
  Logger.log(ss.getUrl());
  Logger.log('Form editor: ' + form.getEditUrl());
  Logger.log('Remember: delete the "Test Student" row after you test.');
}

/** Builds/updates the "Scores" tab: each student's BEST points per name + class. Runs on every submission. */
function refreshScores() {
  const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'));
  const raw = ss.getSheets().find(s => s.getName().indexOf('Form Responses') === 0);
  const rows = raw.getDataRange().getValues().slice(1);        // A=time, B=name, C=class, D=points, E=parts ...
  const best = {};
  rows.forEach(r => {
    if (!r[1]) return;
    const key = String(r[1]).trim().toLowerCase() + '|' + String(r[2]).trim();
    const pts = parseFloat(r[3]) || 0;
    if (!best[key] || pts >= best[key].pts) {
      best[key] = { name: String(r[1]).trim(), cls: r[2], pts: pts, parts: r[4], acts: r[5], checks: r[6], first: r[7], explain: r[8], when: r[0], n: ((best[key] || {}).n || 0) + 1 };
    } else { best[key].n++; }
  });
  const list = Object.values(best).sort((a, b) => String(a.cls).localeCompare(String(b.cls)) || a.name.localeCompare(b.name));
  let sc = ss.getSheetByName('Scores') || ss.insertSheet('Scores');
  sc.clear();
  const head = ['Class', 'Name', 'Best points', 'Parts finished (of 8)', 'Activities (of 11)', 'Quick checks (of 36)', 'Right first time', 'Explain-its (of 6)', 'Last sent', 'Times sent'];
  sc.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#EDE7F6');
  if (list.length) sc.getRange(2, 1, list.length, head.length).setValues(list.map(x => [x.cls, x.name, x.pts, x.parts, x.acts, x.checks, x.first, x.explain, x.when, x.n]));
  sc.setFrozenRows(1);
}
