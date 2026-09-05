/**
 * Runs build/CreateSpellExchangeForm.gs against a mock FormApp and asserts that
 * the questions come out in the exact order the workbook reads response columns.
 *
 *   node scripts/verify_form_script.js
 *
 * The workbook addresses the response tab by POSITION, so this ordering is the
 * single thing most likely to break everything quietly. Page breaks and section
 * headers create no response column and are ignored here, matching Forms.
 */
'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.dirname(__dirname);
const GS = path.join(ROOT, 'build', 'CreateSpellExchangeForm.gs');
const CSV = path.join(ROOT, 'data', 'wizard-spells.csv');

// ---- mock Apps Script ------------------------------------------------------
const items = [];     // response-producing questions, in order
const settings = {};
const logs = [];

function mkItem(type) {
  const it = { type, title: null, help: null, choices: null, required: false };
  items.push(it);
  const api = {
    setTitle(t) { it.title = t; return api; },
    setHelpText(t) { it.help = t; return api; },
    setChoiceValues(v) { it.choices = v; return api; },
    setRequired(b) { it.required = b; return api; },
    showOtherOption() { return api; },
  };
  return api;
}

// Page breaks / section headers produce no column -- record separately.
const layout = [];
function mkLayout(type) {
  const api = {
    setTitle(t) { layout.push({ type, title: t }); return api; },
    setHelpText() { return api; },
    setGoToPage() { return api; },
  };
  return api;
}

const form = {
  setDescription(d) { settings.description = d; return form; },
  setProgressBar(v) { settings.progressBar = v; return form; },
  setCollectEmail(v) { settings.collectEmail = v; return form; },
  setAllowResponseEdits(v) { settings.allowEdits = v; return form; },
  setLimitOneResponsePerUser(v) { settings.limitOne = v; return form; },
  setShuffleQuestions(v) { settings.shuffle = v; return form; },
  setConfirmationMessage(v) { settings.confirmation = v; return form; },
  setDestination(t, id) { settings.destination = id; return form; },
  setAcceptingResponses() { return form; },
  addTextItem: () => mkItem('text'),
  addListItem: () => mkItem('list'),
  addCheckboxItem: () => mkItem('checkbox'),
  addParagraphTextItem: () => mkItem('paragraph'),
  addMultipleChoiceItem: () => mkItem('mc'),
  addPageBreakItem: () => mkLayout('page'),
  addSectionHeaderItem: () => mkLayout('section'),
  getId: () => 'MOCK_FORM_ID',
  getPublishedUrl: () => 'https://forms.gle/MOCK',
  getEditUrl: () => 'https://docs.google.com/forms/d/MOCK/edit',
};

const sandbox = {
  FormApp: {
    create: () => form,
    openById: () => form,
    DestinationType: { SPREADSHEET: 'SPREADSHEET' },
  },
  Logger: { log: (m) => logs.push(String(m)) },
  PropertiesService: {
    getScriptProperties: () => ({ setProperty() {}, getProperty: () => null }),
  },
  String,
};

vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(GS, 'utf8'), sandbox);
sandbox.createSpellExchangeForm();

// ---- expectations ----------------------------------------------------------
// Column A is Timestamp and B is Email (both added by Google), so the first
// scripted question lands in column C.
const EXPECTED = [
  ['C', 'Player name', 'text', true],
  ['D', 'Character name', 'text', true],
  ['E', 'Wizard level', 'list', true],
  ['F', 'Table number', 'text', false],
  ['G', 'Wizard subclass', 'list', false],
  ['H', 'Contact after the event', 'text', false],
  ['I', 'Level 1 spells in your spellbook', 'checkbox', false],
  ['J', 'Level 2 spells in your spellbook', 'checkbox', false],
  ['K', 'Level 3 spells in your spellbook', 'checkbox', false],
  ['L', 'Level 4 spells in your spellbook', 'checkbox', false],
  ['M', 'Level 5 spells in your spellbook', 'checkbox', false],
  ['N', 'Level 6 spells in your spellbook', 'checkbox', false],
  ['O', 'Level 7 spells in your spellbook', 'checkbox', false],
  ['P', 'Level 8 spells in your spellbook', 'checkbox', false],
  ['Q', 'Level 9 spells in your spellbook', 'checkbox', false],
  ['R', 'Other spells not in the lists above', 'paragraph', false],
];

// counts straight from the CSV, so a stale .gs is caught
const wanted = {};
for (const line of fs.readFileSync(CSV, 'utf8').split(/\r?\n/).slice(1)) {
  if (!line.trim()) continue;
  const lv = Number(line.split(',')[0]);
  wanted[lv] = (wanted[lv] || 0) + 1;
}

let failures = 0;
function check(label, got, want) {
  const ok = String(got) === String(want);
  console.log(`  ${label.padEnd(52)} ${String(got).padEnd(12)} ${ok ? 'ok' : 'EXPECTED ' + want}`);
  if (!ok) failures++;
}

console.log('\nResponse column order (workbook reads these by position)');
check('question count', items.length, EXPECTED.length);
EXPECTED.forEach(([col, title, type, required], i) => {
  const it = items[i] || {};
  check(`col ${col}: ${title}`, it.title, title);
  if (it.type !== type) check(`col ${col} type`, it.type, type);
  if (it.required !== required) check(`col ${col} required`, it.required, required);
});

console.log('\nSpell choices per level (must match data/wizard-spells.csv)');
let total = 0;
for (let lv = 1; lv <= 9; lv++) {
  const it = items[5 + lv];
  const n = it && it.choices ? it.choices.length : 0;
  total += n;
  check(`level ${lv} choices`, n, wanted[lv]);
}
check('total spells offered', total, Object.values(wanted).reduce((a, b) => a + b, 0));

console.log('\nSettings');
check('collects email (gives column B)', settings.collectEmail, true);
check('response editing allowed', settings.allowEdits, true);
check('not limited to one response', settings.limitOne, false);
check('questions not shuffled', settings.shuffle, false);

console.log('\nLayout');
check('page breaks (9 spell levels + final)', layout.length, 10);

console.log('\nSanity');
const dupes = [];
for (let lv = 1; lv <= 9; lv++) {
  const c = items[5 + lv].choices || [];
  if (new Set(c).size !== c.length) dupes.push(lv);
}
check('no duplicate choices within a level', dupes.length, 0);
const blank = items.filter((i) => i.choices && i.choices.some((c) => !c || !c.trim()));
check('no blank choice values', blank.length, 0);
check('receipts reminder present in log',
  logs.some((l) => /copy of their response/i.test(l)), true);

console.log(failures ? `\nFAILED: ${failures} check(s)` : '\nALL CHECKS PASSED');
process.exit(failures ? 1 : 0);
