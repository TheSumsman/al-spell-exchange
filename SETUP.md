# Organizer runbook

How to stand up the Wizard Spell Exchange for one event. About 10 minutes.

You need a Google account and nothing else — the workbook and both scripts are
in `build/`, ready to use. See [README.md](README.md) for what the tool does and
[DESIGN-NOTES.md](DESIGN-NOTES.md) for why it is built this way.

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

That creates all 16 questions in the order the workbook expects, with all 350
spells as checkbox options, one page per spell level, and every setting
configured except one.

> Prefer to build the form by hand? `data/form-options/level-*.txt` are the same
> lists as plain text, one spell per line — Google Forms turns a
> newline-separated paste into one option per line. The question order then
> matters; check it against `verify_form_script.js`.

### Turn on response receipts

Open the form and set **Settings → Responses → "Send responders a copy of their
response" → Always**. The script cannot set this one; Apps Script has no API
for it.

Don't skip it. That receipt email is how a player edits their entry after they
get home, which is half the point of the exercise.

### Repoint the response tab

You will always need this. Google creates its own new tab — typically
`Form Responses 1` — when you link a form to a spreadsheet, and never writes
into the `Form Responses` tab that ships with the workbook. That is expected,
not a fault.

**Rename nothing.** Repoint the formulas at Google's tab, then delete the
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

> **Order matters.** Deleting a sheet that formulas still reference turns every
> one of them into `#REF!` permanently, and recreating a sheet by the same name
> does not heal them. Repoint first, delete second — never the other way round.

Check the **Wizards** tab picks up your submissions and the **Matrix** fills in.

### Run the polish script

Three things don't survive the .xlsx import and have to be added in Sheets: the
character dropdown on Copy Planner, the checkboxes in the Want column, and the
colour coding on the Status column.

1. In the Sheet: **Extensions → Apps Script**.
2. Paste in the whole of `build/PolishSpellExchangeSheet.gs`.
3. **Run → polishSheet**, approve the prompt, read the Execution log.

Run it **after** linking the form and repointing the formulas — the dropdown
reads the Wizards tab, which is empty until then. Safe to re-run.

---

## 3. Verify before the event

Work through **[TESTPLAN.md](TESTPLAN.md)** — three test wizards with exact
expected numbers at every step. Twenty minutes, done once.

It checks the things only a real Google Sheet can show: that the conversion kept
the formulas, dropdowns and filters intact, and that the response tab is wired
to the right columns.

## 4. Share it

- Share the **Sheet** as *Anyone with the link → **Viewer***. Players never need
  edit access; everything is calculated.
- Share the **Form** link for registration. Check it in a logged-out browser to
  confirm it doesn't demand a Google account.
- Build the table tent with your form link as a QR code:
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

---

## 5. Announce at the start of the Epic

Four things, or players will ask all day:

1. **The whole Epic counts as one session** for spell copying — any wizard here
   may copy from any other, regardless of table. This is the one organizer
   ruling, so say it out loud.
2. **Cost is 50 GP per spell level**, and 1 DT per spell for levels 1–4, 2 DT
   for levels 5–9.
3. **You must register today.** The AL rule is "immediately after a session in
   which you both played" — the arithmetic can wait, the record cannot.
4. **You can only copy spells of a level you can already prepare.**
5. **Set your own downtime budget.** The Copy Planner starts at 10 DT, the
   minimum everyone will have after the Epic. Anyone with downtime banked in
   their log should change that cell to their real total.

If a player raises the PHB's *"Copying the Book"* clause at 10 GP per level:
that covers duplicating your own spellbook into a replacement, not learning a
spell from another wizard, so it doesn't apply. The workbook's Read Me tab says
so too.

---

## Roster size

The workbook is pre-sized for **24 wizards** and reads the first ~199 form
responses. If more turn up, edit `N_WIZ` in `scripts/build_workbook.py` and
rebuild — do it *before* linking the form, since rebuilding means re-uploading.

Don't try to drag extra rows down in the Sheet instead: the new rows will look
correct and read blank. [README.md](README.md#rebuilding) explains why.
