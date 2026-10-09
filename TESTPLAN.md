# Pre-event test plan

Twenty minutes, done once, before you share anything. Every expected number
below is exact — if you see a different one, something is wrong.

The offline checks (README, *Testing a change*) already evaluate the workbook
formulas, the Form's question order and where the polish script puts things, so
this is **not** re-checking the arithmetic. It checks the
things only a real Google Sheet can tell us: that the xlsx→Sheets conversion kept
the formulas, dropdowns and filters intact, and that the response tab is wired to
the right columns.

Use a phone for at least one submission — that's how players will actually do it.

---

## 0. Redeploy first

This test assumes a **freshly built** sheet. The workbook changed since the last
round — response references are now insertion-proof, the Want column uses real
checkboxes, and the Trade Sheet became a per-spell Copy Log — so an older
sheet will not behave as described here.

- [ ] Rebuild locally:
      ```
      python scripts/build_workbook.py
      python scripts/make_form_script.py
      python scripts/make_polish_script.py
      ```
- [ ] Delete (or set aside) the old spreadsheet. Its formulas have the old row
      drift baked in and cannot be repaired by hand.
- [ ] Upload the new `build/SpellExchange.xlsx`, open as Google Sheets, note its
      spreadsheet id.
- [ ] Link a form to it — either re-run `CreateSpellExchangeForm.gs` with that id
      in `SPREADSHEET_ID`, or link your existing form via
      *Responses → Link to Sheets*.
- [ ] **Repoint the formulas**: Ctrl+H, find `'Form Responses'!`, replace with
      `'Form Responses 1'!` (or whatever Google actually named its tab),
      **Search: all sheets**, **tick "Also search within formulas"**.
- [ ] Delete the now-unreferenced `Form Responses` placeholder tab — *after* the
      replace, never before.
- [ ] Run `build/PolishSpellExchangeSheet.gs` via **Extensions → Apps Script** to
      add the dropdown, checkboxes and Status colours.
- [ ] Set the form's **Settings → Responses → "Send responders a copy" → Always**
      if you're re-creating the form.

SETUP.md has the same sequence with the reasoning attached.

> Use **one real email address** on at least one submission this time. Step 5
> tests editing from home via the receipt link, which is the one path still
> unverified.

---

## The test data

Three wizards, chosen to sit on every boundary that matters.

| | **Aria** | **Bexley** | **Cirilla** |
|---|---|---|---|
| Player name | Test A | Test B | Test C |
| Wizard level | **3** | **9** | **17** |
| Table number | 1 | 2 | 3 |
| Subclass | Evoker | **Order of Scribes** | Diviner |
| Contact | testA | testB | testC |
| Level 1 | Magic Missile, Shield | Shield | Magic Missile |
| Level 2 | Misty Step | Misty Step, Web | — |
| Level 3 | — | Fireball | Fireball, Counterspell |
| Level 4 | — | Polymorph | — |
| Level 5 | — | Wall of Force | — |
| Level 9 | — | — | Wish |
| Other spells | *(leave blank)* | *(leave blank)* | `Alarm` |

Why these three: Aria is capped at spell level 2 so most of the list must be
unavailable to her; Bexley is Order of Scribes so his downtime is charged at a
different rate; Cirilla can reach level 9 and holds a spell nobody else has.
Cirilla's `Alarm` in the free-text box tests the catch-all field.

---

## 1. Submit the three responses

Submit **Aria from a phone**, the other two from anywhere.

- [ ] The form asks for your email at the top (if not, email collection is off —
      fix it in Settings before going further).
- [ ] After the basics you get **one page per spell level**, nine in all, with a
      progress bar. Levels you skip just need **Next**.
- [ ] Each submission produces a **confirmation email** containing an
      *Edit your response* link. **Keep Aria's** — step 5 needs it.

> No email arrives? Settings → Responses → *Send responders a copy of their
> response* → **Always**. Apps Script can't set this, so it's easy to miss.

---

## 2. Wizards tab

- [ ] All three characters appear, one row each, no blanks between them.
- [ ] **Max spell level** reads exactly:

| Character | Wizard level | Max spell level |
|---|---|---|
| Aria | 3 | **2** |
| Bexley | 9 | **5** |
| Cirilla | 17 | **9** |

- [ ] **Order of Scribes?** is `Yes` for Bexley only, `No` for the other two.
- [ ] **Spells in book**: Aria **3**, Bexley **6**, Cirilla **5**.
      Cirilla's five includes `Alarm`, which she entered in the free-text box —
      if she shows **4**, the "Other spells" column isn't being read.
- [ ] **Rare spells held** — spells nobody else has: Aria **0**, Bexley **3**
      (Web, Polymorph, Wall of Force), Cirilla **3** (Counterspell, Wish, Alarm).

If this tab is empty or full of `#REF!`, the formulas aren't pointing at
Google's response tab — see *Repoint the response tab* in SETUP.md.

If it is empty but the response tab clearly has your rows, check what
`Wizards!A2` actually references. It should read
`INDEX('Form Responses 1'!$D$1:$D$200,2)`. If it instead names a plain cell like
`!D5`, the formulas have drifted — the workbook is an old build and needs
replacing.

---

## 3. Matrix tab

Use Ctrl+F to jump to each spell.

- [ ] **Shield** — *# Owners* = **2**, *Owners* = `Aria, Bexley`
- [ ] **Fireball** — *# Owners* = **2**, *Owners* = `Bexley, Cirilla`
- [ ] **Wish** — *# Owners* = **1**, *Owners* = `Cirilla`
- [ ] **Feather Fall** — *# Owners* = **0**, *Owners* blank (nobody took it)

> The *Owners* text must read exactly `Aria, Bexley` — not `, Aria, Bexley`, not
> `Aria, , Bexley`. The offline suite checks this exactly now, so a stray comma
> here means Sheets evaluates the formula differently. Tell me if you see one.

---

## 4. Copy Planner — the tab players actually use

### 4a. Picking a character

- [ ] Click **B1**. You should get a **dropdown** listing the three characters.
- [ ] The **Want** column shows real **checkboxes**, not empty cells.
- [ ] The **Status** column is colour-coded — green `CAN COPY`, grey `OWNED`,
      red `TOO HIGH`, sand `NOBODY HAS IT`.
- [ ] The budget panel at the top: **B3** (Gold) is empty, **B4** (Downtime)
      reads **10**, and **E4** reads `WITHIN BUDGET` on a **green** background.
      E3 reads `<- enter your total to compare`, with no colour.

> All of these come from `build/PolishSpellExchangeSheet.gs`. An .xlsx import drops
> cross-sheet data validation and conditional formatting, and .xlsx has no
> checkbox cell type at all, so they can only be applied once the workbook is in
> Sheets. If any are missing, you haven't run that script yet — see SETUP.md
> step 2. Run it *after* linking the form, since the dropdown reads the Wizards
> tab.

### 4b. As Aria (level 3, max spell level 2)

Set B1 to `Aria`, then check the **Status** column:

| Spell | Level | Expected status | Why |
|---|---|---|---|
| Magic Missile | 1 | **OWNED** | already in her book |
| Web | 2 | **CAN COPY** | Bexley has it, within her reach |
| Fireball | 3 | **TOO HIGH** | above her max spell level |
| Wish | 9 | **TOO HIGH** | far above |
| Feather Fall | 1 | **NOBODY HAS IT** | in reach, but no owner |

- [ ] `Feather Fall` shows **NOBODY HAS IT**, not `CAN COPY` — this is the check
      that the *# Owners* wiring is right.
- [ ] Apply a filter on **Status = CAN COPY**. Aria should see **exactly two
      spells: `Alarm` and `Web`.** Nothing else in the whole 350-row list is both
      within her reach and held by somebody. If she can see level 5 spells, the
      max-spell-level gate is broken.

> `Alarm` appearing here is a second, stronger test of the free-text field:
> Cirilla typed it into "Other spells", and it has flowed all the way through to
> another player's shopping list.

### 4c. Costs at the normal rate

> **Clear the Want column whenever you switch character.** The totals add up
> every ticked row, whatever its status — they don't skip rows marked OWNED or
> TOO HIGH. Leftover ticks from another wizard give nonsense totals. Row 7 of
> the planner says so too.

The panel reads across: **Your total** (B), **Cost of N ticked** (C), **Left
after copying** (D), then the verdict (E). Gold is row 3, downtime row 4.

Still as Aria, tick **Web** only:

- [ ] C2 reads `Cost of 1 ticked`. Gold cost **100**, downtime cost **1**
- [ ] Gold left is **blank** — no gold entered, so nothing to compare
- [ ] Downtime left **9**, `WITHIN BUDGET`, green

Clear that tick. Set B1 to `Cirilla` and tick **four** spells she can actually
copy — Shield (1), Misty Step (2), Web (2), Polymorph (4):

- [ ] Gold cost **450** (50 + 100 + 100 + 200)
- [ ] Downtime cost **4** — four spells of level 1–4, one downtime day each
- [ ] Downtime left **6**

### 4d. Over budget

Keep Cirilla's four ticks. Type **400** into **B3** (Gold) and **3** into
**B4** (Downtime):

- [ ] Gold left **-50**, `OVER BUDGET`, **red**
- [ ] Downtime left **-1**, `OVER BUDGET`, **red**
- [ ] Type **-5** into B3. It's **rejected** — budgets can't be negative. Put
      400 back.

> No colour, but the words are right? The polish script's budget step didn't
> run — check its Execution log for `budget colours`.

### 4e. Order of Scribes — the headline feature

Clear those ticks. Set B1 to `Bexley` and tick the three spells **he** can copy —
Magic Missile (1), Alarm (1), Counterspell (3):

- [ ] Gold cost **250** (50 + 50 + 150) — gold is per spell level as usual
- [ ] Downtime cost **1**, *not 3* — Order of Scribes copies **ten** level 1–4
      spells per downtime day (ALPG p.2), so three of them still round up to
      just one
- [ ] Against the same 400 GP / 3 DT: gold left **150**, downtime left **2**,
      both `WITHIN BUDGET`, both **green**

That's the contrast: Cirilla paid **4 DT for four** level 1–4 spells, Bexley pays
**1 DT for three**. If Bexley shows 3 DT, the Scribes branch isn't firing — check
his subclass reads exactly `Order of Scribes` on the Wizards tab.

Now **clear B3** and set **B4 back to 10**.

### 4f. Levelling up at the end of the Epic

Clear those ticks. Set B1 to `Aria` and type **5** into **B5** (*New wizard
level*) — she registered at 3 and is levelling up twice for the test's sake:

- [ ] The note beside B5 reads *Registered at level 3, copying up to spell
      level 3*.
- [ ] The **Status = CAN COPY** filter now shows **exactly four spells**:
      `Alarm`, `Web`, `Fireball`, `Counterspell`. Level 3 opened up; nothing at
      level 4 or above did.
- [ ] The **Wizards** tab still shows Aria at wizard level **3**, max spell
      level **2**. The new level is her planning input, not her registration.

Now type **2** into B5 — lower than she registered:

- [ ] The filter is back to `Alarm` and `Web` only. A lower level is ignored,
      not trusted.

**Clear B5** before moving on; step 5 relies on Aria being back at level 3.

---

## 5. Editing from home

This is what makes the whole thing work after the event, so don't skip it.

- [ ] Open **Aria's** edit link from her confirmation email.
- [ ] Add **Counterspell** (level 3) and resubmit.
- [ ] Back in the Sheet: Aria's *Spells in book* goes **3 → 4**, and
      Counterspell's *# Owners* on Matrix goes **1 → 2** (only Cirilla had it).
- [ ] Crucially, this **updates her existing row** rather than adding a fourth
      one. A new row means response editing is off.

Counterspell is level 3, so it still shows **TOO HIGH** for Aria — she can hold a
spell she can't yet prepare, which is correct.

---

## 6. Copy Log — what you copied, what it cost, from whom

Set B1 back to `Aria` and tick **Web** only, then open **Copy Log**.

Gold and downtime are **expended by the copying wizard** — nobody is paid. So
this tab is a record of what Aria bought, not an invoice addressed to Bexley.
There is deliberately **no per-lender table**.

- [ ] Row 8 carries a dark banner: **PASTE THIS ONE CELL INTO YOUR CHARACTER LOG**.
- [ ] The summary row reads **Character** = Aria, **Spells** 1,
      **Total GP** 100, **Total DT** 1.
- [ ] The table starting at row 11 has **exactly one row** of data, and the rows
      beneath it are blank:

| # | Spell | Level | GP | Copied from (character) | Player |
|---|---|---|---|---|---|
| 1 | Web | 2 | 100 | Test B Character | Test B player |

- [ ] **Click A9** — the yellow cell under the banner. It is the one cell you
      copy, and it reads, near enough:
      `Copied 1 wizard spell(s) into spellbook: Web (L2, 100 GP) from Test B Character. Total: 100 GP and 1 DT.`
- [ ] Copy A9 and paste it somewhere plain. You should get that whole sentence
      as text — spell, level, per-spell GP, who from, and both totals.
- [ ] **No stray `; ;` or `(from )` fragments.** The previous build had a guard
      testing a numeric column, which for unused rows produced the non-empty
      string `" (from )"` — 21 of them. Fixed by guarding on the row index.

Now tick **three** spells Aria can reach and confirm the log grows to three
entries separated by `; `, with GP and DT totals moving accordingly.

> Downtime is spent as one batch, so **Total DT** already applies the Order of
> Scribes rate where it is due. Do not add up per spell — that is why there is
> no per-spell DT column.

## 7. Sharing and permissions

- [ ] Open the **Sheet link in a private window**. It should open **read-only** —
      no edit access, no Google sign-in demanded.
- [ ] Open the **Form link in a private window**. It should accept a submission
      **without** a Google account.

> A viewer opening the Sheet read-only can still *look* at Copy Planner, but
> can't set B1 to their own character. Tell players to use **File → Make a copy**,
> or hand out an edit link if you trust the room. Worth deciding before the day.

- [ ] If anyone will plan in **Excel** instead: **File → Download → Microsoft
      Excel (.xlsx)** and open it. Matrix's *Owners* column and Copy Log's
      log entry show names, not `#NAME?`. The workbook avoids every function
      newer than Excel 2007, so this should hold in any Excel version. While
      you're there, tick a Want box and check the budget panel's costs move.

---

## 8. Clean up

- [ ] Delete the three test rows from the response tab.
- [ ] Untick every checkbox in the Copy Planner **Want** column. **B1**, **B3**
      (Gold) and **B5** (New wizard level) are empty; **B4** (Downtime) is
      **10**.
- [ ] Confirm the Wizards tab is empty again and shows no `#REF!`.

---

## If something fails

Tell me **which step and what you actually saw**. The most likely failures, in
order:

1. **`#REF!` everywhere** — the placeholder tab was deleted *before* the
   formulas were repointed. Recoverable: re-upload the workbook and re-link.
2. **Wizards tab empty** — the find-and-replace didn't reach every formula
   (usually "Also search within formulas" was left unticked), or it named the
   wrong tab.
3. **Everything shifted one column** — a question was added or reordered in the
   form. Re-run `verify_form_script.js` and rebuild the form from the script.
4. **`#NAME?` in a downloaded Excel copy** — a formula is using a function
   that version of Excel lacks. `verify_workbook.py` is meant to catch this, so
   tell me which cell.
5. **No dropdown, no checkboxes, no Status or budget colours** — `PolishSpellExchangeSheet.gs`
   hasn't been run, or was run before the Wizards tab had data. Re-run it; it's
   safe to run repeatedly.
6. **Want ticks do nothing to the totals** — the Want cells are plain text rather
   than checkboxes. The totals count `TRUE`, so the polish script must run before
   ticking anything.
7. **Stray commas or semicolons** in *Owners* or the log entry — Sheets is
   evaluating the joins differently from the offline suite. Cosmetic, but tell
   me which cell.
