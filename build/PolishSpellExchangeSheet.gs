/**
 * Wizard Spell Exchange -- Google Sheets polish (Apps Script)
 * ==========================================================
 *
 * WHAT THIS IS
 *   Applies the things an .xlsx import drops on the floor: the character
 *   dropdown on Copy Planner, real checkboxes in the Want column, and colour
 *   coding on the Status column.
 *
 * HOW TO RUN (about a minute)
 *   1. Open the converted Google Sheet.
 *   2. Extensions > Apps Script. Delete the placeholder, paste this file in.
 *   3. Run > polishSheet. Approve the authorization prompt on first run.
 *   4. Read the Execution log (Ctrl+Enter) for what it changed.
 *
 * Safe to re-run: it clears its own rules first rather than stacking them up.
 * Run it AFTER linking the form and repointing the formulas, not before --
 * the character dropdown reads the Wizards tab, which is empty until then.
 *
 * Each step is isolated, so if one fails the others still apply and the log
 * tells you which one to chase.
 */

var N_WIZ = 24;
var PLAN_TOP = 11;   // first spell row on Copy Planner


// Run one step in isolation: a failure is reported and the rest still run,
// rather than aborting the script half-applied.
function step_(name, fn) {
  try {
    fn();
    return true;
  } catch (e) {
    Logger.log('FAILED - ' + name + ': ' + (e && e.message ? e.message : e));
    return false;
  }
}


function polishSheet() {
  var ss = SpreadsheetApp.getActive();
  var cp = ss.getSheetByName('Copy Planner');
  if (!cp) { throw new Error('No "Copy Planner" tab -- wrong spreadsheet?'); }
  var wz = ss.getSheetByName('Wizards');
  if (!wz) { throw new Error('No "Wizards" tab -- wrong spreadsheet?'); }

  var lastRow = cp.getLastRow();
  var nSpells = lastRow - PLAN_TOP + 1;
  if (nSpells < 1) { throw new Error('Copy Planner looks empty (last row ' + lastRow + ').'); }
  Logger.log('Copy Planner: ' + nSpells + ' spell rows (' + PLAN_TOP + '..' + lastRow + ').');

  var ok = 0;
  var total = 4;

  // ---- 1. character picker in B1 ------------------------------------------
  // .xlsx validation pointing at another sheet's range does not survive import.
  if (step_('character dropdown', function () {
    var chars = wz.getRange(2, 1, N_WIZ, 1);
    cp.getRange('B1').setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInRange(chars, true)
        .setAllowInvalid(false)
        .setHelpText('Pick your character.')
        .build());
    Logger.log('B1: dropdown bound to Wizards!A2:A' + (N_WIZ + 1));
  })) { ok++; }

  // ---- 2. real checkboxes in the Want column ------------------------------
  // .xlsx has no checkbox cell type, so this can only be done here. The
  // formulas count TRUE, which is exactly what a ticked checkbox stores.
  if (step_('Want checkboxes', function () {
    cp.getRange(PLAN_TOP, 9, nSpells, 1).insertCheckboxes();
    Logger.log('Want column: ' + nSpells + ' checkboxes inserted.');
  })) { ok++; }

  // ---- 3. colour the Status column ----------------------------------------
  // getConditionalFormatRules() and setConditionalFormatRules() are methods on
  // SHEET, not on Spreadsheet. Calling them on the spreadsheet object throws
  // "ss.getConditionalFormatRules is not a function".
  if (step_('Status colours', function () {
    var status = cp.getRange(PLAN_TOP, 5, nSpells, 1);
    var rules = cp.getConditionalFormatRules().filter(function (r) {
      // drop only our own rules on the Status column, leave anything else alone
      return r.getRanges().every(function (rg) { return rg.getColumn() !== 5; });
    });
    function rule(text, bg, fg, bold) {
      var b = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(text).setBackground(bg).setFontColor(fg);
      if (bold) { b = b.setBold(true); }
      return b.setRanges([status]).build();
    }
    rules.push(rule('CAN COPY',      '#d6f2d6', '#14532d', true));  // green, act on these
    rules.push(rule('OWNED',         '#eeeeee', '#777777', false)); // grey, already have it
    rules.push(rule('TOO HIGH',      '#fde2e2', '#7f1d1d', false)); // red, out of reach
    rules.push(rule('NOBODY HAS IT', '#f6f2e8', '#8a7040', false)); // sand, no source here
    rules.push(rule('RESTRICTED',    '#efe2f7', '#5b2d82', false)); // purple, subclass gated
    cp.setConditionalFormatRules(rules);
    Logger.log('Status column: 5 colour rules applied (green = CAN COPY).');
  })) { ok++; }

  // ---- 4. a filter, so players can narrow to CAN COPY ---------------------
  if (step_('spell-table filter', function () {
    var existing = cp.getFilter();
    if (existing) { existing.remove(); }
    cp.getRange(PLAN_TOP - 1, 1, nSpells + 1, 9).createFilter();
    Logger.log('Filter created over the spell table.');
  })) { ok++; }

  Logger.log('');
  Logger.log(ok + ' of ' + total + ' steps applied.');
  if (ok === total) {
    Logger.log('Done. Tick a checkbox in Want and the totals at the top should move.');
  } else {
    Logger.log('Some steps failed above. Fix and re-run -- re-running is safe.');
  }
}
