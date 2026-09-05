# Organizer runbook

How to stand up the Wizard Spell Exchange for one event. See
**[README.md](README.md)** for what it is and why it works this way.

About 10 minutes. You need a Google account and nothing else — the workbook and
both scripts are committed under `build/`, so there is nothing to build and
nothing to type by hand.

---

## 1. Upload the workbook

1. Upload `build/SpellExchange.xlsx` to Google Drive.
2. Right-click → **Open with → Google Sheets**, then **File → Save as Google
   Sheets**. Work with that native copy from here on, not the .xlsx.
3. Copy its **spreadsheet id** from the URL — the long string between `/d/` and
   `/edit`.

## 2. Create the form

1. Go to [script.google.com](https://script.google.com) → **New project**.
2. Delete the placeholder code and paste in the whole of
   `build/CreateSpellExchangeForm.gs`.
3. Put the spreadsheet id from step 1 into `SPREADSHEET_ID` at the top. (Leave
   it blank to link by hand later via *Responses → Link to Sheets*.)
4. **Run** → select `createSpellExchangeForm`. Approve the authorization prompt
   on first run; run again if it stops at the consent screen.
5. Open the **Execution log** (Ctrl+Enter) for the LIVE and EDIT links.

The script creates all 16 questions in the exact order the workbook expects,
with all 350 spells as checkbox options, one page per spell level, and every
setting configured — except one.

### The one setting the script can't set

Apps Script has no API for response receipts. Open the form and set
**Settings → Responses → "Send responders a copy of their response" → Always**.

Don't skip it: that receipt email is how a player edits their entry after they
get home, which is half the point of the exercise.

### Repoint the response tab (you will always need this)

Google **always** creates its own new tab — typically `Form Responses 1` — when
you link a form to a spreadsheet. It will never write into the `Form Responses`
tab that ships with the workbook. That is expected, not a fault.

**Rename nothing.** Just repoint the formulas at Google's tab, then delete the
placeholder:

1. Note the exact name Google gave its new tab — usually `Form Responses 1`.
2. **Ctrl+H** (Find and replace):
   - Find `'Form Responses'!`
   - Replace with `'Form Responses 1'!` — or whatever Google actually called it
   - Search: **All sheets**
   - Tick **Also search within formulas** — nothing happens without it
   - Replace all
3. Delete the now-empty `Form Responses` tab. This is safe *because step 2 left
   nothing pointing at it*.

> Order matters. Deleting a sheet that formulas still reference turns every one
> of them into `#REF!` permanently, and recreating a sheet by the same name does
> not heal them. Repoint first, delete second — never the other way round.

Check the **Wizards** tab picks up your submissions and the **Matrix** fills in.

### Then run the polish script

Three things cannot survive an .xlsx import and must be applied in Sheets:
the character dropdown on Copy Planner (cross-sheet data validation is dropped),
the **checkboxes** in the Want column (.xlsx has no checkbox cell type), and the
**colour coding** on the Status column.

1. In the Sheet: **Extensions → Apps Script**.
2. Paste in the whole of `build/PolishSpellExchangeSheet.gs`.
3. **Run → polishSheet**, approve the prompt, read the Execution log.

Run it **after** linking the form and repointing the formulas — the dropdown
reads the Wizards tab, which is empty until then. Safe to re-run; it clears its
own formatting rules rather than stacking them.

### Why the formulas can survive this at all

Google Forms **inserts** a row for each response rather than filling the next
blank one, and that insert lands exactly where a plain reference points. A
formula reading `'Form Responses'!D2` silently becomes `D3` after one response,
`D4` after two — the whole workbook drifts off the data, one row per submission.
It reads blank rather than wrong, which is a slow way to notice.

So every reference into the response tab is written as a **bounded `INDEX`**:

```
INDEX('Form Responses'!$D$1:$D$200, 2)
```

Inserting a row extends the range's end but leaves `INDEX`'s literal row number
alone, so it cannot drift. `verify_workbook.py` fails the build if a plain row
reference ever reappears.

The 200-row bound is headroom for a 24-wizard roster; past ~199 responses the
workbook needs rebuilding anyway.

> `data/form-options/level-*.txt` are the same lists as plain text, one spell
> per line. They're only a fallback if you'd rather build the form by hand —
> Google Forms turns a newline-separated paste into one option per line.

---

## 3. Verify before the event

Work through **[TESTPLAN.md](TESTPLAN.md)** — three test wizards with exact
expected numbers at every step. Twenty minutes, done once.

It exists because the offline suite already proves the formulas and the form's
column order; what it *cannot* prove is that the xlsx-to-Sheets conversion kept
the dropdowns, filters and `TEXTJOIN` behaviour intact, or that the response tab
is wired to the right columns. That is what the test plan checks.

## 4. Share it

- Share the **Sheet** as *Anyone with the link → **Viewer***. Players never need
  edit access; everything is calculated.
- Share the **Form** link for registration. Check it in a logged-out browser to
  confirm it doesn't demand a Google account.
- Build the table tent with your form link embedded as a QR code:
  ```
  python scripts/make_table_tent.py --url https://forms.gle/your-form-link
  ```
  Then either open `build/table-tent.html` and print to PDF at **A5**, or render
  it headlessly:
  ```
  chrome --headless=new --no-pdf-header-footer \
         --print-to-pdf=build/table-tent.pdf \
         file:///absolute/path/to/build/table-tent.html
  ```
  It should come out as **one** A5 page. If you get two, the print dialog is
  adding its own margins on top of the page's — set margins to None/Default.

  Re-run `make_table_tent.py` with no `--url` afterwards to reset the file to a
  blank placeholder, so your event's form link isn't left sitting in the repo.
  The rendered PDF carries the link too — in a link annotation, where a search
  of the file's text won't obviously show it — so `*.pdf` is gitignored.

---

## 5. Announce at the start of the Epic

Four things, or players will ask all day:

1. **The whole Epic counts as one session** for spell copying — any wizard here
   may copy from any other wizard here, regardless of table. *(This is an
   organizer ruling. ALPG says "a session in which you both played" and never
   addresses multi-table Epics, so say it out loud.)*
2. **Cost is 50 GP per spell level**, and 1 DT per spell for levels 1–4, 2 DT
   for levels 5–9.
3. **You must register today.** The AL rule is "immediately after a session in
   which you both played" — the arithmetic can wait, the record cannot.
4. **You can only copy spells of a level you can already prepare.**

Expect one player to point out the PHB's *"Copying the Book"* clause, which
prices a wizard copying a spell they already know into another book at 10 GP and
1 hour per level — one fifth the cost. It is a genuinely defensible reading when
the owner does the scribing. The ruling for this event is 50 GP; the Read Me tab
records that so you only have to have the argument once.

---

## Roster size

The workbook is pre-sized for **24 wizards**. If more turn up, edit `N_WIZ` at
the top of `scripts/build_workbook.py` and rebuild — but do it *before* linking
the form, since rebuilding means re-uploading.
