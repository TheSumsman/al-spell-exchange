/**
 * Runs build/PolishSpellExchangeSheet.gs against a mock SpreadsheetApp and
 * asserts it touches exactly the cells the workbook puts things in.
 *
 *   node scripts/verify_polish_script.js
 *
 * The polish script addresses Copy Planner by position -- the Want column, the
 * Status column, the budget panel. Nothing else notices if the workbook's
 * layout moves and the script keeps colouring the old cells: the sheet just
 * quietly loses its checkboxes or colours. So the expectations here come from
 * the BUILT WORKBOOK (via Python), not from the script's own constants.
 *
 * It also runs the script twice, because SETUP.md promises re-running is safe:
 * the second run must leave the same rules, not a second copy of each.
 */
'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { execFileSync } = require('child_process');

const ROOT = path.dirname(__dirname);
const GS = path.join(ROOT, 'build', 'PolishSpellExchangeSheet.gs');

// ---- what the workbook actually looks like ---------------------------------
const PY = `
import json, sys
sys.path.insert(0, 'scripts')
import build_workbook as bw
from openpyxl import load_workbook
cp = load_workbook('build/SpellExchange.xlsx')['Copy Planner']
hdr = bw.PLAN_TOP - 1
print(json.dumps({
    'planTop': bw.PLAN_TOP,
    'nWiz': bw.N_WIZ,
    'lastRow': cp.max_row,
    'statusHeader': cp.cell(row=hdr, column=5).value,
    'wantHeader': cp.cell(row=hdr, column=9).value,
    'lastHeader': cp.cell(row=hdr, column=10).value,
    'statusFormula': cp.cell(row=bw.PLAN_TOP, column=5).value,
    'budgetRange': bw.CP_BUDGET_FLAGS,
    'gpLeft': bw.CP_GP_LEFT,
    'dtLeft': bw.CP_DT_LEFT,
    'gpLeftFormula': cp[bw.CP_GP_LEFT].value,
    'dtLeftFormula': cp[bw.CP_DT_LEFT].value,
}))
`;
const wb = JSON.parse(execFileSync(process.env.PYTHON || 'python', ['-c', PY],
                                   { cwd: ROOT, encoding: 'utf8' }));

// ---- mock Apps Script ------------------------------------------------------
function colLetter(n) {
  let s = '';
  for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}
function colNumber(s) {
  return s.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
}

const events = { validations: [], checkboxes: [], filters: [] };
const logs = [];

function mkRange(sheet, row, col, nRows, nCols) {
  const a1 = colLetter(col) + row +
    (nRows > 1 || nCols > 1 ? ':' + colLetter(col + nCols - 1) + (row + nRows - 1) : '');
  return {
    sheet, a1,
    getA1Notation: () => a1,
    getColumn: () => col,
    getRow: () => row,
    setDataValidation(v) { events.validations.push({ at: a1, v }); return this; },
    insertCheckboxes() { events.checkboxes.push(a1); return this; },
    createFilter() { sheet.filter = { at: a1, remove() { sheet.filter = null; } };
                     events.filters.push(a1); return sheet.filter; },
  };
}

function mkSheet(name, lastRow) {
  const sheet = {
    name, filter: null, rules: [],
    getLastRow: () => lastRow,
    getRange(a, b, c, d) {
      if (typeof a === 'string') {
        const m = a.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
        const c1 = colNumber(m[1]), r1 = Number(m[2]);
        const c2 = m[3] ? colNumber(m[3]) : c1, r2 = m[4] ? Number(m[4]) : r1;
        return mkRange(sheet, r1, c1, r2 - r1 + 1, c2 - c1 + 1);
      }
      return mkRange(sheet, a, b, c || 1, d || 1);
    },
    getFilter: () => sheet.filter,
    getConditionalFormatRules: () => sheet.rules.slice(),
    setConditionalFormatRules(r) { sheet.rules = r.slice(); },
  };
  return sheet;
}

function ruleBuilder() {
  const r = { ranges: [] };
  const b = {
    whenTextEqualTo(t) { r.text = t; return b; },
    whenFormulaSatisfied(f) { r.formula = f; return b; },
    setBackground(c) { r.bg = c; return b; },
    setFontColor(c) { r.fg = c; return b; },
    setBold(v) { r.bold = v; return b; },
    setRanges(rs) { r.ranges = rs; return b; },
    build() { r.getRanges = () => r.ranges; return r; },
  };
  return b;
}

function validationBuilder() {
  const v = {};
  const b = {
    requireValueInRange(range, dropdown) { v.source = range.sheet.name + '!' + range.a1;
                                           v.dropdown = dropdown; return b; },
    setAllowInvalid(x) { v.allowInvalid = x; return b; },
    setHelpText(t) { v.help = t; return b; },
    build: () => v,
  };
  return b;
}

const planner = mkSheet('Copy Planner', wb.lastRow);
const wizards = mkSheet('Wizards', wb.nWiz + 1);
// An organizer's own rule on another column, which the script must leave alone.
const theirs = ruleBuilder().whenTextEqualTo('mine')
  .setRanges([planner.getRange('B11:B20')]).build();
planner.rules.push(theirs);

const sandbox = {
  SpreadsheetApp: {
    getActive: () => ({
      getSheetByName: (n) => ({ 'Copy Planner': planner, Wizards: wizards })[n] || null,
    }),
    newDataValidation: validationBuilder,
    newConditionalFormatRule: ruleBuilder,
  },
  Logger: { log: (m) => logs.push(String(m)) },
};
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(GS, 'utf8'), sandbox);

// ---- checks ----------------------------------------------------------------
let failures = 0;
function check(label, got, want) {
  const ok = String(got) === String(want);
  console.log(`  ${label.padEnd(52)} ${String(got).padEnd(14)} ${ok ? 'ok' : 'EXPECTED ' + want}`);
  if (!ok) failures++;
}

const top = wb.planTop;
const spellRows = `${top}:${wb.lastRow}`;
const [want, status] = [`I${top}:I${wb.lastRow}`, `E${top}:E${wb.lastRow}`];

console.log('\nThe columns the script assumes are where the workbook put them');
check('column E is Status', wb.statusHeader, 'Status');
check('column I is Want', wb.wantHeader, 'Want');
check('nothing after Want (filter stops at I)', wb.lastHeader, 'null');

sandbox.polishSheet();
const firstRules = planner.rules.length;

console.log('\nFirst run');
check('all steps applied', logs.some((l) => /^5 of 5 steps applied/.test(l)), true);
check('no step failed', logs.filter((l) => /^FAILED/.test(l)).length, 0);

const dd = events.validations.find((e) => e.at === 'B1');
check('character dropdown on B1', !!dd, true);
check('dropdown lists Wizards!A2:A' + (wb.nWiz + 1), dd && dd.v.source, `Wizards!A2:A${wb.nWiz + 1}`);
check('dropdown rejects other values', dd && dd.v.allowInvalid, false);

check('checkboxes on the Want column only', events.checkboxes.join(' '), want);
check('filter over the spell table', events.filters[0], `A${top - 1}:I${wb.lastRow}`);

const onStatus = planner.rules.filter((r) => r.ranges.some((g) => g.a1 === status));
check('Status colour rules', onStatus.length, 5);
// Each colour fires on exact text, so every one must be a status the
// workbook's formula can actually produce -- a renamed status would otherwise
// just silently lose its colour.
for (const r of onStatus) {
  check(`  "${r.text}" is produced by the Status formula`,
    wb.statusFormula.includes(`"${r.text}"`), true);
}

const onBudget = planner.rules.filter((r) => r.ranges.some((g) => g.a1 === wb.budgetRange));
check('budget colour rules on ' + wb.budgetRange, onBudget.length, 2);
// The rule tests the "left after copying" column, relative to the range's top
// row, so it must name that column and the range must start on the GP row.
const leftCol = wb.gpLeft.replace(/\d+/, '');
const leftRef = `$${leftCol}${wb.gpLeft.replace(/^[A-Z]+/, '')}`;
check('budget range starts on the gold row', wb.budgetRange.split(':')[0], wb.gpLeft);
check('downtime "left" cell is in the same column', wb.dtLeft.replace(/\d+/, ''), leftCol);
check('red rule: left < 0', onBudget.some((r) => r.formula === `=AND(ISNUMBER(${leftRef}),${leftRef}<0)` && r.bg === '#fde2e2'), true);
check('green rule: left >= 0', onBudget.some((r) => r.formula === `=AND(ISNUMBER(${leftRef}),${leftRef}>=0)` && r.bg === '#d6f2d6'), true);
check('gold "left" cell holds a subtraction', /-/.test(wb.gpLeftFormula), true);
check('downtime "left" cell holds a subtraction', /-/.test(wb.dtLeftFormula), true);
check("organizer's own rule kept", planner.rules.includes(theirs), true);

console.log('\nSecond run (SETUP.md says re-running is safe)');
logs.length = 0;
sandbox.polishSheet();
check('all steps applied again', logs.some((l) => /^5 of 5 steps applied/.test(l)), true);
check('same number of rules, not doubled', planner.rules.length, firstRules);
check("organizer's own rule still kept", planner.rules.includes(theirs), true);
check('one filter, not two', planner.filter && planner.filter.at, `A${top - 1}:I${wb.lastRow}`);

console.log(failures ? `\nFAILED: ${failures} check(s)` : '\nALL CHECKS PASSED');
process.exit(failures ? 1 : 0);
