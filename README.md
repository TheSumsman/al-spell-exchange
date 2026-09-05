# Wizard Spell Exchange

Lets wizards at an Adventurers League **Epic** register their spellbooks on a
phone, then work out afterwards which spells they can copy from each other, from
whom, and what it costs in gold and downtime.

A Google Form captures spellbooks; a Google Sheet does the analysis. Nothing is
self-hosted. **[SETUP.md](SETUP.md)** is the runbook.

Scope: the **Forgotten Realms** campaign. Other AL campaigns have different legal
content and are deliberately out of scope.

---

## Using it

You do not need to run anything. The three files an organizer actually uses are
committed and ready:

```
build/SpellExchange.xlsx             upload to Drive, open as Google Sheets
build/CreateSpellExchangeForm.gs     paste into script.google.com to build the Form
build/PolishSpellExchangeSheet.gs    paste into the Sheet to add dropdowns and colours
```

Follow **[SETUP.md](SETUP.md)** — about 10 minutes — then walk
**[TESTPLAN.md](TESTPLAN.md)** once before the event.

## What's here

```
data/wizard-spells.csv            350 leveled wizard spells: level, name, school, source
data/form-options/                the same lists as plain text (fallback only)
build/SpellExchange.xlsx          the workbook to upload to Drive
build/CreateSpellExchangeForm.gs  paste into script.google.com to build the Form
build/PolishSpellExchangeSheet.gs paste into the Sheet after linking the Form
build/table-tent.html             printable A5 sign for the event table
scripts/                          the generators and their tests
```

## Rebuilding

Everything below runs from a fresh clone. `build/` is committed, so this is only
needed if you change the roster size, the costs, or the spell list.

```bash
python scripts/build_workbook.py      # CSV -> build/SpellExchange.xlsx
python scripts/verify_workbook.py     # evaluates every workbook formula for real (~5 min)
python scripts/make_form_script.py    # CSV -> build/CreateSpellExchangeForm.gs
node   scripts/verify_form_script.js  # runs the .gs against a mock FormApp
python scripts/make_polish_script.py  # -> build/PolishSpellExchangeSheet.gs
python scripts/make_table_tent.py --url https://forms.gle/xxxx
```

`verify_workbook.py` needs `formulas`; `build_workbook.py` needs `openpyxl`; the
table tent's QR needs `qrcode`; the form verifier needs Node.

> Re-run `make_table_tent.py` with **no** `--url` before committing, so your
> event's form link isn't left sitting in the repo. Rendered PDFs are gitignored
> for the same reason — the link survives into the PDF.

### Rebuilding the spell list

`data/wizard-spells.csv` is committed, and everything above reads it. Two scripts
regenerate it, and **neither one runs from a clone**:

```bash
python scripts/extract_spells.py     # saved listing pages -> data/wizard-spells.csv
python scripts/crosscheck_books.py   # independent check of that scrape
```

They need two things that are not in this repository and cannot be:

| | what it is | why it's absent |
|---|---|---|
| `TheMasterSpellbook/` | ~140 MB of saved D&D Beyond listing pages | copyrighted WotC content |
| `$SPELLEXCHANGE_BOOKS` | a directory of D&D Beyond book exports | personal copies of purchased books |

Both scripts check for them first and **exit with an explicit message** rather
than running on nothing — `crosscheck_books.py` in particular would otherwise
compare zero spells and print a meaningless `PASS`.

If you have your own D&D Beyond exports, point `SPELLEXCHANGE_BOOKS` at them and
recreate `TheMasterSpellbook/` as described under
[Re-scraping](#re-scraping). If you don't: you can still change costs, rules,
roster size and layout, rebuild the workbook and the form, and run the full test
suite. You just can't regenerate the spell list itself — and you don't need to,
because it's committed.

## Why the Form is scripted

The workbook addresses the response tab **by column position**, so one question
added, removed or reordered shifts every column and quietly breaks every
formula. Generating the form from the same CSV that builds the workbook is the
guarantee that they agree — and a change to the spell list becomes one re-run
rather than nine re-pastes.

`verify_form_script.js` executes the generated `.gs` against a mock `FormApp`
and asserts the resulting question order matches the columns the workbook reads
(C=Player … R=Other), that the per-level choice counts match the CSV, and that
the email/editing settings are right.

### The response-tab traps

Two of them, and the second one shipped broken.

**Linking always makes a new tab.** A Form linked to an existing spreadsheet
always creates its own tab (`Form Responses 1`); it will never write into one you
made. So the shipped `Form Responses` placeholder must be repointed afterwards —
one find-and-replace, *then* delete the placeholder. Never delete first:
deleting a sheet that formulas reference turns every reference into `#REF!`
permanently, and recreating a sheet by that name does not heal them.

**Form responses INSERT rows, and plain references drift.** Google Forms inserts
a row per response instead of filling the next blank one, and the insert lands
exactly where a plain reference points. `'Form Responses'!D2` becomes `D3` after
one response, `D4` after two. Three test submissions left every reference in the
live sheet off by exactly three rows, and the workbook read blank throughout.

Every reference into the response tab is therefore a **bounded `INDEX`** —
`INDEX('Form Responses'!$D$1:$D$200, 2)`. Row insertion extends the range end but
leaves the literal row index alone. `verify_workbook.py` scans the built workbook
and fails if a plain row reference reappears.

(Whole-column `$D:$D` is equally safe but makes the offline evaluator crawl
through a million rows per reference, hence the bound.)

---

## Where the spell list comes from

**Primary source: `TheMasterSpellbook/`** — saved pages from D&D Beyond's
`/spells` listing, filtered to **Class = Wizard** and the 15 FR-legal sources.
21 pages, 20 rows each. *(Not in this repository — see
[Rebuilding the spell list](#rebuilding-the-spell-list).)* The wizard filter is the whole trick: the listing rows
carry no class, so filtering makes attribution implicit and nothing needs to be
expanded or guessed.

Two things the saved rows do **not** contain:

- **Class** — solved by the filter, as above.
- **Source book** — the badge you see bottom-right of each card is rendered
  client-side, so `Ctrl+S` never captures it. Source is instead joined by spell
  name from a local directory of D&D Beyond book exports — personal copies of
  purchased books, pointed at by `SPELLEXCHANGE_BOOKS` and not part of this
  repository. All 350 spells currently resolve.

Cantrips are dropped throughout: they are not kept in a spellbook.

### Re-scraping

Apply Class = Wizard in D&D Beyond's own filter UI (don't hand-edit the query
string — the page exposes two class-id schemes, `2190886` and `8`), then save
each page into `TheMasterSpellbook/`. Filenames don't matter; the extractor
globs `*.htm` and dedupes on `data-slug`, so overlapping saves are harmless. It
reads the advertised page count from the pagination and **warns loudly if pages
are missing**, so a truncated scrape can't silently shrink the list.

---

## Traps in the book exports

Worth knowing before touching the parsers — each of these cost real time.

**Find class lists by anchor id, never by heading text.** Xanathar's wizard list
is marked up as `id="WizardSpells"` — no space. Searching for the visible string
`"Wizard Spells"` finds nothing and makes the book look like it has no spells at
all. Sweep for `id="\w*Spell\w*"` to enumerate these sections; that also surfaces
`ArtificerSpellList`, `DunamancySpellList` and warlock `ExpandedSpellList`s,
which must **not** be read as wizard spells.

**Always whitelist the eight schools.** A pattern like "heading followed by
`Nth-level`" also matches class features — "Breath of the Dragon — *3rd-level
feature*" — and invents hundreds of phantom spells. Requiring
`3rd-level evocation` is what separates a spell from a feature.

**Three markup dialects, not two.** Metadata appears as
`<em>Level 3 Evocation (Sorcerer, Wizard)</em>` (2024 books),
`<em>3rd-level evocation</em>`, and `<p class="Core-Styles_Core-Metadata">8th-level
necromancy</p>` (Xanathar's). Anchor on the `spell-tooltip` heading and accept
any of them.

**An illustration can sit between a heading and its metadata**, so search as far
as the next spell heading rather than a fixed character window.

**Normalise case and apostrophes.** Xanathar's lists are sentence case
(`Absorb elements`), PHB 2024 is title case, and curly apostrophes appear in
`Tasha’s`. Dedupe on a normalised key; take display names from the listing.

**Exports can be partial.** Several book folders contain only some chapters.
Kwalish's three wizard spells live in its Appendix E, which was missing at first.
`extract_spells.py` supports an optional `data/source-overrides.csv`
(`Name,Source,Why`) for spells whose defining chapter isn't available locally.
None are needed right now.

---

## Verification

`crosscheck_books.py` derives wizard spells from the local books *independently*
of the listing and diffs the two. It currently confirms **297 spells across four
books** (PHB 2024, Xanathar's, Heroes of Faerûn, Tasha's) with **0 missing and 0
level disagreements** — that's the check that catches a truncated scrape.

`verify_workbook.py` builds a small workbook using the same formula-generating
code, seeds it with three known wizards, evaluates every formula, and asserts the
results: max spell levels, presence detection, `OWNED` / `CAN COPY` / `TOO HIGH`
statuses, per-spell and total costs, the Copy Log rows and log entry, and the Order of
Scribes rate (four level 1–4 spells cost 4 DT normally, 1 DT for a Scribe, with
gold unchanged).

**One thing it cannot settle:** whether `TEXTJOIN` drops formula-produced blanks.
The offline evaluator joins them anyway, producing `, Aria, Bexley`. Excel and
Google Sheets should both ignore them, but confirm the *Owners* and *Which
spells* columns by eye once in Sheets — SETUP.md step 3 covers it.

---

## The rules this implements

**Cost.** PHB 2024 sets 50 GP and 2 hours per spell level. ALPG v2026.4 p.3
converts the time into downtime: **1 DT per spell for levels 1–4, 2 DT for levels
5–9**. Order of Scribes wizards instead copy ten level 1–4 spells, or five level
5–9 spells, per 1 DT (ALPG p.2) — gold is unaffected.

**Downtime is the binding constraint, not gold.** You earn 10 DT per session,
and levelling costs 10 DT. A Tier 3 wizard with 40,000 GP can still only copy
about ten spells. The Copy Planner is built around a DT budget for that reason.

**Eligibility.** You may only copy a spell of a level you can already prepare:
`MaxSpellLevel = MIN(9, roundup(WizardLevel / 2))`.

**Timing.** ALPG p.3: *"You may copy spells from a character's spellbook
immediately after a session in which you both played."* Registration therefore
has to happen at the event.

### Two organizer rulings baked in

- **The whole Epic counts as one session**, so any wizard present may copy from
  any other. ALPG never addresses multi-table Epics — this fills a genuine gap
  and should be announced.
- **50 GP per spell level.** The PHB's adjacent *"Copying the Book"* clause
  prices a wizard copying a spell they already know into another book at 10 GP
  and 1 hour per level. That reading is defensible when the owner scribes, and
  players will raise it; this event uses 50 GP.


---

## Licence and content

The code is **MIT** licensed — see [LICENSE](LICENSE).

`data/wizard-spells.csv` lists spell **names, levels, schools and source books**
and nothing else: no descriptions, no rules text, no mechanics. The saved D&D
Beyond pages and book exports it was derived from are not published here.
[NOTICE.md](NOTICE.md) sets that out in full, along with the Wizards of the Coast
Fan Content Policy this is published under.

Unofficial Fan Content permitted under the Fan Content Policy. Not approved or
endorsed by Wizards. Portions of the materials used are property of Wizards of
the Coast. ©Wizards of the Coast LLC.
