"""Generate build/CreateSpellExchangeForm.gs from data/wizard-spells.csv.

The Google Form is built by an Apps Script rather than by hand because the
workbook addresses response columns BY POSITION -- one mis-ordered or inserted
question silently breaks every formula. Generating it also means a change to the
spell list is one re-run, not nine re-pastes.

Generating the form from the same CSV that builds the workbook is what keeps
the form's question order and the workbook's column positions in agreement.
"""
import argparse
import csv
import json
import os
import re
import sys
from collections import OrderedDict

sys.stdout.reconfigure(encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CSV_PATH = os.path.join(ROOT, "data", "wizard-spells.csv")
OUT = os.path.join(ROOT, "build", "CreateSpellExchangeForm.gs")

SUBCLASSES = [
    "Abjurer", "Diviner", "Evoker", "Illusionist",
    "Conjurer", "Enchanter", "Necromancer", "Transmuter",
    "Bladesinger", "Order of Scribes", "Chronurgist", "Graviturgist",
    "Other / not listed",
]

HEAD = '''/**
 * Wizard Spell Exchange -- Google Form builder (Apps Script)
 * =========================================================
 *
 * WHAT THIS IS
 *   Creates the spellbook-registration Form in YOUR Google account, with every
 *   question in the exact order the SpellExchange workbook expects. Nothing to
 *   type by hand.
 *
 * WHY SCRIPTED
 *   The workbook reads the response tab BY COLUMN POSITION. One question added,
 *   removed or reordered shifts every column and quietly breaks every formula.
 *   This script is the guarantee that the order is right.
 *
 * HOW TO RUN (about 2 minutes)
 *   1. Go to https://script.google.com  ->  "New project".
 *   2. Delete the placeholder code, paste this whole file in.
 *   3. Press Run and select the function `createSpellExchangeForm`.
 *   4. First run asks you to authorize -- approve it. Run again if it stops at
 *      the consent screen.
 *   5. Open the Execution log (Ctrl+Enter). It prints the LIVE and EDIT links.
 *
 * LINKING TO THE WORKBOOK
 *   Upload build/SpellExchange.xlsx to Drive, open it as a Google Sheet, then
 *   put its id in SPREADSHEET_ID below and re-run. The id is the long string in
 *   the sheet URL between /d/ and /edit. Leave it empty to link by hand later
 *   via Responses > Link to Sheets.
 *
 *   Google ALWAYS makes its own new tab ("Form Responses 1") -- it will never
 *   write into the workbook's own "Form Responses" tab. Repoint it afterwards,
 *   and mind the order: in Sheets, DELETING a referenced sheet turns every
 *   formula into #REF! for good, while RENAMING one rewrites the formulas to
 *   follow it. So:
 *     1. rename the workbook's "Form Responses" tab to "OldResponses"
 *     2. rename Google's new tab to exactly "Form Responses"
 *     3. Ctrl+H, find `OldResponses!` replace `'Form Responses'!`,
 *        Search: all sheets, tick "Also search within formulas"
 *     4. delete "OldResponses"
 *   SETUP.md walks through this.
 *
 * ONE SETTING THIS SCRIPT CANNOT SET
 *   Apps Script has no API for response receipts. After running, open the form
 *   and set Settings > Responses > "Send responders a copy of their response"
 *   to ALWAYS. That receipt email carries the edit link players need to fix
 *   their entry after they get home -- don't skip it.
 *
 * GENERATED FILE -- do not hand-edit.
 *   Regenerate with:  python scripts/make_form_script.py
 *   Spell list: %(count)d leveled wizard spells, AL-legal Forgotten Realms sources.
 */

// The workbook's Google Sheets id -- responses are linked here automatically.
// Set it with:  python scripts/make_form_script.py --spreadsheet-id <id>
// The id is the part of the sheet URL between /d/ and /edit (NOT the gid).
var SPREADSHEET_ID = %(sheet_id)s;

var SUBCLASSES = %(subclasses)s;

// Spell choices by level, generated from data/wizard-spells.csv.
var SPELLS = {
%(spells)s};


function createSpellExchangeForm() {
  var form = FormApp.create('Wizard Spell Exchange -- AL Epic (Forgotten Realms)');

  form.setDescription(
    'Register what is in your wizard\\'s spellbook so other wizards at the Epic ' +
    'can see what they could copy -- and so you can see what you could copy from them.\\n\\n' +
    'Tick only the spells you ALREADY have. Cantrips are not included: they are ' +
    'not kept in a spellbook.\\n\\n' +
    'Copying costs 50 GP per spell level, plus 1 Downtime Day per spell for ' +
    'spell levels 1-4 and 2 DT per spell for levels 5-9. You may only copy a ' +
    'spell of a level you can already prepare.\\n\\n' +
    'AL rule (ALPG p.3): you may copy from another character\\'s spellbook ' +
    'immediately after a session in which you both played -- so register today. ' +
    'For this Epic the whole event counts as one session, so any wizard here may ' +
    'copy from any other wizard here.\\n\\n' +
    'One entry per wizard character. You will be emailed a link to edit your ' +
    'answers later.');

  form.setProgressBar(true)
      .setCollectEmail(true)
      .setAllowResponseEdits(true)
      .setLimitOneResponsePerUser(false)
      .setShuffleQuestions(false)
      .setConfirmationMessage(
        'Registered. Check your email for a link to edit this entry, and open ' +
        'the results sheet to see what you can copy and from whom.');

  // ---- Page 1: who you are -------------------------------------------------
  // ORDER IS LOAD-BEARING. The workbook expects response columns:
  //   A Timestamp | B Email | C Player | D Character | E Wizard level
  //   F Table | G Subclass | H Contact | I..Q Level 1..9 spells | R Other
  // Page breaks and section headers create no columns, so they are free.

  form.addTextItem()
      .setTitle('Player name')
      .setRequired(true);

  form.addTextItem()
      .setTitle('Character name')
      .setHelpText('One entry per wizard. Submit again for a second character.')
      .setRequired(true);

  var levels = [];
  for (var i = 1; i <= 20; i++) { levels.push(String(i)); }
  form.addListItem()
      .setTitle('Wizard level')
      .setHelpText('Your levels in the Wizard class. This sets which spell ' +
                   'levels you are allowed to copy.')
      .setChoiceValues(levels)
      .setRequired(true);

  form.addTextItem()
      .setTitle('Table number')
      .setHelpText('Which table you are playing at today.');

  form.addListItem()
      .setTitle('Wizard subclass')
      .setHelpText('Order of Scribes matters: it lets you copy ten level 1-4 ' +
                   'spells, or five level 5-9 spells, for 1 Downtime Day.')
      .setChoiceValues(SUBCLASSES);

  form.addTextItem()
      .setTitle('Contact after the event')
      .setHelpText('Discord handle or similar, so people can reach you to ' +
                   'arrange copying. Optional.');

  // ---- One page per spell level -------------------------------------------
  // A page each, rather than one enormous page: a Tier 1 wizard would otherwise
  // have to scroll past 300 irrelevant checkboxes to reach the end.
  for (var lv = 1; lv <= 9; lv++) {
    form.addPageBreakItem()
        .setTitle('Level ' + lv + ' spells')
        .setHelpText(lv === 1
          ? 'Tick every level 1 spell already in your spellbook.'
          : 'Tick every level ' + lv + ' spell already in your spellbook. ' +
            'Nothing at this level? Just press Next.');

    form.addCheckboxItem()
        .setTitle('Level ' + lv + ' spells in your spellbook')
        .setHelpText(SPELLS[lv].length + ' spells available at this level.')
        .setChoiceValues(SPELLS[lv]);
  }

  // ---- Final page ----------------------------------------------------------
  form.addPageBreakItem()
      .setTitle('Anything missing?')
      .setHelpText('Almost done.');

  form.addParagraphTextItem()
      .setTitle('Other spells not in the lists above')
      .setHelpText('One per line, or comma separated. Use this for anything ' +
                   'the lists missed -- spell names are matched by name, so ' +
                   'spell them as printed.');

  if (SPREADSHEET_ID) {
    form.setDestination(FormApp.DestinationType.SPREADSHEET, SPREADSHEET_ID);
    Logger.log('Responses linked to spreadsheet ' + SPREADSHEET_ID);
    Logger.log('CHECK THE TAB NAME: it must be exactly "Form Responses".');
  } else {
    Logger.log('No SPREADSHEET_ID set -- link by hand: Responses > Link to Sheets.');
  }

  PropertiesService.getScriptProperties()
      .setProperty('SPELLEXCHANGE_FORM_ID', form.getId());

  Logger.log('LIVE (share this):    ' + form.getPublishedUrl());
  Logger.log('EDIT (open to tweak): ' + form.getEditUrl());
  Logger.log('');
  Logger.log('STILL TO DO BY HAND: Settings > Responses > "Send responders a ' +
             'copy of their response" = ALWAYS. Apps Script cannot set it, and ' +
             'that email is how players edit their entry from home.');
}


// Stop the form accepting responses (run by hand after the event if you like).
function closeForm() {
  var id = PropertiesService.getScriptProperties()
      .getProperty('SPELLEXCHANGE_FORM_ID');
  if (!id) { Logger.log('No stored form id -- run createSpellExchangeForm first.'); return; }
  FormApp.openById(id).setAcceptingResponses(false);
  Logger.log('Form closed.');
}


function reopenForm() {
  var id = PropertiesService.getScriptProperties()
      .getProperty('SPELLEXCHANGE_FORM_ID');
  if (!id) { Logger.log('No stored form id -- run createSpellExchangeForm first.'); return; }
  FormApp.openById(id).setAcceptingResponses(true);
  Logger.log('Form reopened.');
}
'''


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--spreadsheet-id", default="",
                    help="workbook's Google Sheets id, or its full URL")
    args = ap.parse_args()

    sheet_id = args.spreadsheet_id.strip()
    if "/d/" in sheet_id:  # a full URL was pasted -- pull the id out of it
        m = re.search(r"/d/([a-zA-Z0-9_-]+)", sheet_id)
        if not m:
            sys.exit("Could not find a spreadsheet id in that URL")
        sheet_id = m.group(1)
    if sheet_id and not re.fullmatch(r"[a-zA-Z0-9_-]{20,}", sheet_id):
        sys.exit("That does not look like a spreadsheet id: %r\n"
                 "It is the part of the sheet URL between /d/ and /edit."
                 % sheet_id)

    by_level = OrderedDict((lv, []) for lv in range(1, 10))
    with open(CSV_PATH, encoding="utf-8") as fh:
        for r in csv.DictReader(fh):
            by_level[int(r["Level"])].append(r["Name"])
    for lv in by_level:
        by_level[lv].sort(key=str.lower)

    # ensure_ascii=True so curly apostrophes travel as ’ escapes and cannot
    # be mangled by copy-paste into the Apps Script editor.
    chunks = []
    for lv, names in by_level.items():
        items = ",\n    ".join(json.dumps(n, ensure_ascii=True) for n in names)
        chunks.append("  %d: [\n    %s\n  ]" % (lv, items))

    total = sum(len(v) for v in by_level.values())
    body = HEAD % {
        "count": total,
        "sheet_id": json.dumps(sheet_id),
        "subclasses": json.dumps(SUBCLASSES, ensure_ascii=True),
        "spells": ",\n".join(chunks) + "\n",
    }

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(body)

    print("Wrote %s" % OUT)
    print("  %d spells across %d questions" % (total, len(by_level)))
    if sheet_id:
        print("  responses will link to spreadsheet %s" % sheet_id)
    else:
        print("  no --spreadsheet-id given: link by hand via Responses > Link to Sheets")
    for lv, names in by_level.items():
        print("    level %d: %d" % (lv, len(names)))


if __name__ == "__main__":
    main()
