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

**Gold and downtime both matter**, so the Copy Planner totals both. Downtime is
the one people forget: you earn 10 DT per session, levelling costs 10 DT and a
Bastion turn 7 DT, so a wealthy wizard can still run out of days before gold.

The planner's **Downtime budget** cell starts at **10** — the minimum a character
is sure to have after the Epic. Downtime can be banked in a character's log, and
the tool has no way to know how much, so tell players to change that cell to
their real total before planning.

**Eligibility.** You may only copy a spell of a level you can already prepare:
`MaxSpellLevel = MIN(9, roundup(WizardLevel / 2))`.

**Timing.** ALPG p.3: *"You may copy spells from a character's spellbook
immediately after a session in which you both played."* Registration therefore
has to happen at the event, not afterwards.

### One organizer ruling is baked in

**The whole Epic counts as one session**, so any wizard present may copy from any
other, regardless of table. ALPG never addresses multi-table Epics, so this fills
a genuine gap. It is a judgement call — announce it at the start, and change it
if you disagree.

Everything else above is the rules as written, not a judgement call.

> **If a player cites the PHB's *"Copying the Book"* clause** at 10 GP per level:
> that clause covers duplicating your *own* spellbook into a replacement, not
> learning a new spell from another wizard. It doesn't apply here. The workbook's
> Read Me tab says so, so you only have the conversation once.

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

> **Can't you just drag extra rows down in the Sheet?** No — and it fails
> quietly, which is worse. Each wizard's row pulls its data with a literal row
> number, `INDEX('Form Responses'!$D$1:$D$200, 7)`. Filling down copies that `7`
> unchanged while the other references do move, so the new rows look right and
> read blank. Adding a wizard also needs a new *column* on Matrix, Calc and Copy
> Log, which fill-right gets wrong the same way. Raise `N_WIZ` and rebuild.

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

## Something wrong? Ideas?

Please say so — bug reports and rules corrections are both welcome, and you do
not need to be a programmer to file one.

**[Open an issue here.](https://github.com/TheSumsman/al-spell-exchange/issues)**
It's a web form. You need a free GitHub account, and that's the only barrier;
there is nothing to install and no code involved. Click **New issue**, describe
what happened, and submit.

Useful things to include, if you have them:

- What you expected and what you got instead — a screenshot of the tab is ideal.
- Which step of [SETUP.md](SETUP.md) you were on.
- Whether you're using the committed `build/SpellExchange.xlsx` or one you
  rebuilt yourself.

Particularly worth reporting: **a spell that is missing, or listed at the wrong
level or school.** That list is derived automatically, and a mistake in it is
invisible until it costs a player the wrong amount of gold. Name the spell and
the book it's from.

Rules disagreements are welcome too. The one organizer ruling above is a
judgement call, and if you run Epics and think it's wrong, the reasoning is
worth hearing.

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
