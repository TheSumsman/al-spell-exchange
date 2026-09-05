# Wizard Spell Exchange

Lets wizards at an Adventurers League **Epic** register their spellbooks on a
phone, then works out afterwards which spells they can copy from each other,
from whom, and what it costs in gold and downtime.

A Google Form captures the spellbooks; a Google Sheet does the analysis. There
is nothing to host and nothing to install.

Scope: the **Forgotten Realms** campaign. Other AL campaigns have different
legal content and are deliberately out of scope.

---

## Getting started

You do not need to run anything. The three files you need are in `build/`, ready
to use:

```
build/SpellExchange.xlsx             upload to Drive, open as Google Sheets
build/CreateSpellExchangeForm.gs     paste into script.google.com to build the Form
build/PolishSpellExchangeSheet.gs    paste into the Sheet to add dropdowns and colours
```

1. Follow **[SETUP.md](SETUP.md)** — about 10 minutes.
2. Work through **[TESTPLAN.md](TESTPLAN.md)** once before the event — about 20
   minutes, with exact expected numbers at every step.
3. Print `build/table-tent.html` at A5 for the registration table.

## What's in the box

```
build/SpellExchange.xlsx          the workbook: registers wizards, matches spells, prices them
build/CreateSpellExchangeForm.gs  builds the registration Form, all 16 questions
build/PolishSpellExchangeSheet.gs adds the dropdown, checkboxes and status colours
build/table-tent.html             printable A5 sign, with a QR code to your Form
data/wizard-spells.csv            350 leveled wizard spells: level, name, school, source
data/form-options/                the same lists as plain text, if you build the Form by hand
scripts/                          the generators and their tests
```

The workbook has eight tabs. The ones you will use are **Wizards** (who
registered), **Matrix** (who can copy what from whom), **Copy Planner** (pick a
wizard, tick spells, get a price) and **Copy Log** (what to write on a logsheet).

## The rules it applies

Worth reading before the event, because players will ask.

**Cost.** 50 GP and 2 hours per spell level (PHB 2024). ALPG v2026.4 p.3
converts the time into downtime: **1 DT per spell for levels 1–4, 2 DT for
levels 5–9**. Order of Scribes wizards instead copy ten level 1–4 spells, or
five level 5–9 spells, per 1 DT (ALPG p.2) — their gold cost is unchanged.

**Downtime is the binding constraint, not gold.** You earn 10 DT per session and
levelling costs 10 DT, so a Tier 3 wizard with 40,000 GP can still only copy
about ten spells. The Copy Planner is built around a DT budget for that reason.

**Eligibility.** You may only copy a spell of a level you can already prepare:
`MaxSpellLevel = MIN(9, roundup(WizardLevel / 2))`.

**Timing.** ALPG p.3: *"You may copy spells from a character's spellbook
immediately after a session in which you both played."* Registration therefore
has to happen at the event, not afterwards.

### Two organizer rulings are baked in

Both are judgement calls, not rules text. Announce them at the start, and change
them if you disagree — see [Rebuilding](#rebuilding).

- **The whole Epic counts as one session**, so any wizard present may copy from
  any other, regardless of table. ALPG never addresses multi-table Epics, so
  this fills a genuine gap.
- **50 GP per spell level.** The PHB's adjacent *"Copying the Book"* clause
  prices a wizard copying a spell they already know into another book at 10 GP
  and 1 hour per level — one fifth the cost. That reading is defensible when the
  owner does the scribing, and a player will raise it. This tool charges 50 GP,
  and the workbook's Read Me tab records the ruling.

---

## Rebuilding

Only needed if you want to change something — the roster size, the costs, the
rulings above, or the layout. `build/` is committed, so a fresh clone can
rebuild everything without any extra setup.

```bash
python scripts/build_workbook.py      # -> build/SpellExchange.xlsx
python scripts/make_form_script.py    # -> build/CreateSpellExchangeForm.gs
python scripts/make_polish_script.py  # -> build/PolishSpellExchangeSheet.gs
python scripts/make_table_tent.py --url https://forms.gle/xxxx
```

Common changes:

| To change | Edit |
|---|---|
| Roster size (default 24) | `N_WIZ` in `scripts/build_workbook.py` |
| Gold or downtime costs | `scripts/build_workbook.py` |
| The spell list | `data/wizard-spells.csv` |

Rebuild the workbook **before** linking a Form to it — rebuilding means
re-uploading, and re-uploading means re-linking.

### Testing a change

```bash
python scripts/verify_workbook.py     # evaluates every workbook formula (~5 min)
node   scripts/verify_form_script.js  # runs the .gs against a mock FormApp
```

The first evaluates the real formulas against three known wizards and checks
every number. The second checks the Form's questions still line up with the
columns the workbook reads. Run both after any change; they are what stops a
broken workbook reaching an event.

Requirements: Python 3 with `openpyxl`, plus `formulas` for `verify_workbook.py`
and `qrcode` for the table tent's QR code. `verify_form_script.js` needs Node.

### Rebuilding the spell list

`data/wizard-spells.csv` is committed and everything above reads it, so you do
not need this. Two scripts regenerate it, and **neither runs from a clone**:

```bash
python scripts/extract_spells.py     # saved listing pages -> data/wizard-spells.csv
python scripts/crosscheck_books.py   # independent check of that scrape
```

They need two things that are not in this repository and cannot be:

| | what it is | why it's absent |
|---|---|---|
| `TheMasterSpellbook/` | ~140 MB of saved D&D Beyond listing pages | copyrighted WotC content |
| `$SPELLEXCHANGE_BOOKS` | a directory of D&D Beyond book exports | personal copies of purchased books |

Both scripts check for them and exit with an explanation rather than running on
nothing.

**Without them** you can still change the costs, the rulings, the roster size
and the layout, rebuild the workbook and the Form, and run the full test suite.
You just can't regenerate the spell list — and you don't need to, because it
ships with the repo. To edit it, edit the CSV directly.

**With your own D&D Beyond exports**, point `SPELLEXCHANGE_BOOKS` at them and
recreate `TheMasterSpellbook/` as described in
[DESIGN-NOTES.md](DESIGN-NOTES.md#re-scraping-the-listing).

---

## Going deeper

**[DESIGN-NOTES.md](DESIGN-NOTES.md)** explains why the code is shaped the way
it is — why the Form is generated rather than hand-built, the constraints the
workbook's formulas have to respect, and how the spell list is parsed. Read it
before modifying the scripts.

## Licence and content

The code is **MIT** licensed — see [LICENSE](LICENSE).

`data/wizard-spells.csv` lists spell **names, levels, schools and source books**
and nothing else: no descriptions, no rules text, no mechanics. The saved D&D
Beyond pages and book exports it was derived from are not published here.
[NOTICE.md](NOTICE.md) sets that out in full, along with the Wizards of the
Coast Fan Content Policy this is published under.

Unofficial Fan Content permitted under the Fan Content Policy. Not approved or
endorsed by Wizards. Portions of the materials used are property of Wizards of
the Coast. ©Wizards of the Coast LLC.
