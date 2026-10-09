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

**Copying spends both gold and downtime**, so the Copy Planner budgets both.
Players enter their character's gold and downtime. The planner shows what the
ticked spells cost, what's left of each, and flags **over budget** in red
when either goes below zero. Which one a character runs short of first depends
entirely on the character, so the tool doesn't assume either.

**Gold** starts blank, because the tool can't guess it. **Downtime** starts at
**10**, one session's award (ALPG p.6). Both can be banked in a character's log
and spent on other things, so tell players to enter their real totals before
planning.

**Eligibility.** You may only copy a spell of a level you can already prepare:
`MaxSpellLevel = MIN(9, roundup(WizardLevel / 2))`.

**Levelling up.** A character who levels up at the end of the Epic copies at
their new level. Players register the level they are *now* — the Form records
the character as they sit at the table — and enter the new level in the Copy
Planner's **New wizard level** cell when they plan. The Wizards tab keeps the
registered level; only that player's planner moves.

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
python scripts/make_form_script.py    # -> build/CreateSpellExchangeForm.gs + data/form-options/
python scripts/make_polish_script.py  # -> build/PolishSpellExchangeSheet.gs
python scripts/make_table_tent.py --url https://forms.gle/xxxx
```

Common changes:

| To change | Edit |
|---|---|
| Roster size (default 24) | `N_WIZ` in `scripts/build_workbook.py` |
| Gold or downtime costs | `scripts/build_workbook.py` |
| The spell list | `data/wizard-spells.csv`, then re-run `build_workbook.py` and `make_form_script.py` |

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
pip install -r requirements.txt

python scripts/check_artifacts.py       # build/ is current and safe to publish
python scripts/verify_workbook.py       # evaluates the workbook's formulas (~20 s)
node   scripts/verify_form_script.js    # the Form's questions, in column order
node   scripts/verify_polish_script.js  # the polish script hits the right cells
python scripts/verify_table_tent.py     # the sign prints on one A5 page
```

| Check | What it stops |
|---|---|
| `check_artifacts.py` | A stale `build/`; a lock file, PDF, live form link or used workbook reaching the public repo |
| `verify_workbook.py` | Wrong numbers — evaluates the real formulas against three known wizards |
| `verify_form_script.js` | The Form's questions drifting out of the column order the workbook reads |
| `verify_polish_script.js` | Dropdown, checkboxes or colours landing on the wrong cells after a layout change |
| `verify_table_tent.py` | The sign spilling onto a second page. Needs Chrome, Chromium or Edge |

**All five run on GitHub for every push** (`.github/workflows/checks.yml`).

**Turn on the pre-commit hook once per clone.** It runs `check_artifacts.py`
and the two mocks before every commit, in a couple of seconds:

```bash
git config core.hooksPath scripts/hooks
```

CI only sees a commit once it is public, so the hook is what actually stops a
leak. It checks exactly what is staged, so stage `build/` together with the
script change that produced it.

None of these touch a real Google Sheet. That part is still
[TESTPLAN.md](TESTPLAN.md).

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
ships with the repo. To edit it, edit the CSV directly, then re-run
`build_workbook.py` and `make_form_script.py`. The second also rewrites the
plain-text lists in `data/form-options/`; the pre-commit hook refuses a commit
where they disagree with the CSV.

**With your own D&D Beyond exports**, point `SPELLEXCHANGE_BOOKS` at them and
recreate `TheMasterSpellbook/` as described in
[DESIGN-NOTES.md](DESIGN-NOTES.md#re-scraping-the-listing).

---

## Something wrong? Ideas?

Please say so. You do not need to be a programmer, and there is nothing to
install — it's a web form, and a free GitHub account is the only requirement.

**[Report it here.](https://github.com/TheSumsman/al-spell-exchange/issues/new/choose)**
Pick whichever fits; each one asks a few short questions so you don't have to
guess what's useful.

| | Use it for |
|---|---|
| **Wrong or missing spell** | A spell absent from the list, or at the wrong level, school or source book |
| **Something didn't work** | Setup, the form, or the workbook not behaving as documented |
| **Rules or ruling disagreement** | You think an AL or PHB rule is applied incorrectly |
| **Idea or suggestion** | Something that would make this more useful at an event |

None of them fit? File a blank issue — questions are welcome, and you can't get
the format wrong.

**Spell errors are the most valuable reports.** That list is generated
automatically, so a wrong level or a missing spell is invisible until it
misprices someone at the table. Name the spell and, if you have it, the book and
page.

**Rules disagreements are welcome too.** One thing in this tool is a judgement
call — the whole Epic counting as one session. Everything else is meant to be
the rules as written, so if something else looks wrong, that's a mistake worth
telling me about rather than a position I took deliberately.

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
