"""Build build/SpellExchange.xlsx from data/wizard-spells.csv.

The workbook is authored as .xlsx but LIVES IN GOOGLE SHEETS, and organisers
also download it back out to Excel. So every formula is restricted to what
Sheets and every Excel since 2007 share: COUNTIF(S), SUMIF, INDEX, MATCH, IF,
IFERROR, SEARCH, SMALL, MID, CEILING, & -- filled down explicitly.

Deliberately avoided:
  * QUERY / ARRAYFORMULA / FILTER      -- Google-only
  * spilling dynamic arrays / XLOOKUP  -- Excel-365-only, converts badly
  * TEXTJOIN / CONCAT                  -- Excel 2019+ only, AND Google's .xlsx
                                          export drops the _xlfn. prefix Excel
                                          needs, so even 365 reads #NAME?.
                                          See joined().
  * checkbox data validation           -- does not survive xlsx -> Sheets
"""
import csv
import os
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

sys.stdout.reconfigure(encoding="utf-8")

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
CSV_PATH = os.path.join(ROOT, "data", "wizard-spells.csv")
OUT = os.path.join(ROOT, "build", "SpellExchange.xlsx")

N_WIZ = 24          # pre-sized roster
FIRST_RESP = 2      # first data row on the response tab
PLAN_TOP = 11       # first spell row on Copy Planner

# The placeholder tab this workbook ships. Google Forms always creates its OWN
# tab when a form is linked (typically "Form Responses 1"), so after linking you
# repoint the formulas at Google's tab with ONE find-and-replace and then delete
# this one. Do NOT name it "Form Responses 1" -- Google would collide with it and
# fall back to "Form Responses 2", which is just confusing.
RESP_TAB = "Form Responses"

# Bound the response ranges rather than using whole columns ($D:$D). Whole
# columns are equally insertion-safe but make the offline formula evaluator
# unusably slow, and 200 rows is far more headroom than a 24-wizard roster
# needs. Past ~199 responses the workbook would need rebuilding anyway.
RESP_LAST = 200


def resp(col, row):
    """A reference to the response tab that survives new form submissions.

    Google Forms INSERTS a row for each response instead of filling a blank
    one, and the insert lands exactly where a direct reference points -- so
    'Form Responses'!D2 silently becomes D3, D4, D5... one row per submission,
    and the whole workbook drifts off the data.

    INDEX over the WHOLE column is immune: inserting rows never rewrites a
    full-column range, and the row index is a plain number, not a reference.
    """
    return "INDEX('%s'!$%s$1:$%s$%d,%d)" % (RESP_TAB, col, col, RESP_LAST, row)

# Copy Planner cells that formulas elsewhere, the polish script and the
# verifier all address. Named once here so the layout can move without anyone
# hunting for "$B$4" inside string literals.
CP_CHAR = "B1"           # character picker
CP_GP_BUDGET = "B3"      # player's gold
CP_DT_BUDGET = "B4"      # player's downtime
CP_NEW_LEVEL = "B5"      # optional wizard level after levelling up
CP_GP_TOTAL = "C3"       # cost of the ticked spells
CP_DT_TOTAL = "C4"
CP_GP_LEFT = "D3"        # budget minus cost; blank until a budget is entered
CP_DT_LEFT = "D4"
CP_GP_STATUS = "E3"      # OVER BUDGET / WITHIN BUDGET
CP_DT_STATUS = "E4"
CP_BUDGET_FLAGS = "D3:E4"  # what PolishSpellExchangeSheet.gs colours
CP_LEVEL_NOTE = "C5"
CP_IDX = "L1"            # hidden helpers: labels in column K, values in L
CP_MAXLVL = "L2"
CP_SCRIBES = "L3"
CP_LVL_USED = "L4"
CP_COUNT = "L5"          # number of spells ticked


def ab(cell):
    """'B4' -> '$B$4'."""
    col = cell.rstrip("0123456789")
    return "$%s$%s" % (col, cell[len(col):])


# Excel's limit on the text in one cell. MID needs a length; this one never
# truncates.
MAX_TEXT = 32767


def joined(cells, sep):
    """A delimited list without TEXTJOIN, which Excel 2016 lacks entirely.

    Each cell in `cells` must hold either "" or `sep` + its value -- e.g.
    ", Bexley". Concatenating them all gives ", Aria, Bexley"; MID then drops
    the leading separator. Empty cells contribute nothing, so there is no
    stray-delimiter case to handle. Formula length grows ~6 characters per
    cell, nowhere near Excel's 8,192 limit at any roster this workbook can read.
    """
    return "MID(%s,%d,%d)" % ("&".join(cells), len(sep) + 1, MAX_TEXT)


def planner(cell):
    """An absolute reference to a Copy Planner cell, from another tab."""
    return "'Copy Planner'!%s" % ab(cell)


HDR_FILL = PatternFill("solid", fgColor="2F3E46")
HDR_FONT = Font(color="FFFFFF", bold=True)
NOTE_FONT = Font(italic=True, color="555555")
TITLE_FONT = Font(bold=True, size=14)
BOX_FILL = PatternFill("solid", fgColor="EDF2F4")
THIN = Side(style="thin", color="BBBBBB")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def header(ws, row, labels, widths=None):
    for i, lab in enumerate(labels, start=1):
        c = ws.cell(row=row, column=i, value=lab)
        c.fill, c.font = HDR_FILL, HDR_FONT
        c.alignment = Alignment(vertical="center", wrap_text=True)
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w


def load_spells():
    with open(CSV_PATH, encoding="utf-8") as fh:
        rows = list(csv.DictReader(fh))
    rows.sort(key=lambda r: (int(r["Level"]), r["Name"].lower()))
    return rows


def build(spells=None, n_wiz=None, out=None):
    """Build the workbook. Parameterised so the verifier can build a small one."""
    global N_WIZ
    spells = load_spells() if spells is None else spells
    if n_wiz is not None:
        N_WIZ = n_wiz
    out = out or OUT
    n = len(spells)
    wb = Workbook()

    # ------------------------------------------------------------ Read Me
    ws = wb.active
    ws.title = "Read Me"
    ws.column_dimensions["A"].width = 108
    lines = [
        ("Wizard Spell Exchange - AL Epic (Forgotten Realms)", TITLE_FONT),
        ("", None),
        ("Players register their spellbook via the Google Form. Everything else on this "
         "workbook is calculated - do not type into any tab except the yellow cells on "
         "Copy Planner.", None),
        ("", None),
        ("HOW TO USE", TITLE_FONT),
        ("1. Fill in the Google Form at the event (one submission per wizard character).", None),
        ("2. Open the 'Copy Planner' tab and fill in the yellow cells at the top: your "
         "character, their", None),
        ("   gold and downtime, and - if levelling up at the end of this Epic - their new "
         "wizard level.", None),
        ("3. Filter the Status column to 'CAN COPY' to see what is available to you.", None),
        ("4. Tick the 'Want' checkbox. The panel at the top shows what the ticked spells "
         "cost, what you", None),
        ("   have left, and turns red if either gold or downtime is over budget.", None),
        ("5. Open 'Copy Log' for the list of what you copied, what it cost, and from", None),
        ("   whom - plus a line of text ready to paste into your character log.", None),
        ("", None),
        ("THE COST", TITLE_FONT),
        ("50 GP per spell level, and 1 Downtime Day per spell for spell levels 1-4, "
         "2 DT per spell for levels 5-9.", None),
        ("", None),
        ("PHB 2024, Wizard - 'Expanding and Replacing a Spellbook':", Font(bold=True)),
        ('  "When you find a level 1+ Wizard spell, you can copy it into your spellbook if '
         "it's of a level you can", NOTE_FONT),
        ('   prepare and if you have time to copy it. For each level of the spell, the '
         'transcription takes 2 hours', NOTE_FONT),
        ('   and costs 50 GP."', NOTE_FONT),
        ("", None),
        ("ALPG v2026.4 p.3, Downtime - 'Copying Spells':", Font(bold=True)),
        ('  "Use \'Expanding and Replacing a Spellbook\' (PH) to copy spells found in '
         "adventures at 1 DT per spell", NOTE_FONT),
        ("   up to level 4 and 2 DT per spell at levels 5-9. You may copy spells from a "
         "character's spellbook", NOTE_FONT),
        ('   immediately after a session in which you both played."', NOTE_FONT),
        ("", None),
        ("ALPG p.2 - Order of Scribes wizards copy ten level 1-4 spells, or five level 5-9 "
         "spells, for 1 DT.", None),
        ("This workbook applies that rate automatically when the subclass is recorded.", NOTE_FONT),
        ("", None),
        ("GOLD AND DOWNTIME", TITLE_FONT),
        ("Copying spends both, so Copy Planner compares both against the totals you "
         "enter from your", None),
        ("character's log. Which one runs out first depends on the character.", None),
        ("", None),
        ("Gold starts blank - the workbook cannot guess it. Downtime starts at 10, one "
         "session's award", NOTE_FONT),
        ("(ALPG p.6). Change both to your real totals, after anything else you are "
         "spending them on.", NOTE_FONT),
        ("", None),
        ("LEVELLING UP", TITLE_FONT),
        ("A character who levels up at the end of this Epic can copy spells of the level "
         "they can prepare", None),
        ("once levelled. Enter the new wizard level on Copy Planner; the form keeps the "
         "level you registered at.", None),
        ("", None),
        ("RULING FOR THIS EVENT", TITLE_FONT),
        ("* The whole Epic counts as one session: any wizard present may copy from any "
         "other wizard present.", None),
        ("  ALPG does not address multi-table Epics, so this fills a genuine gap.", NOTE_FONT),
        ("", None),
        ("Everything else here is the rules as written, quoted above - 50 GP per spell "
         "level, and you may", None),
        ("only copy a spell of a level you can already prepare. Neither is a judgement "
         "call.", None),
        ("", None),
        ("If a player cites the PHB's 'Copying the Book' clause at 10 GP per level: that "
         "covers duplicating", NOTE_FONT),
        ("your own spellbook into a replacement, not learning a spell from another "
         "wizard. It does not apply.", NOTE_FONT),
        ("", None),
        ("Spell list: %d leveled wizard spells from the AL-legal Forgotten Realms sources. "
         "Cantrips are excluded" % n, NOTE_FONT),
        ("- they are never kept in a spellbook.", NOTE_FONT),
    ]
    for i, (text, font) in enumerate(lines, start=1):
        c = ws.cell(row=i, column=1, value=text)
        if font:
            c.font = font
        c.alignment = Alignment(wrap_text=False)

    # ------------------------------------------------------------- Spells
    sp = wb.create_sheet("Spells")
    header(sp, 1, ["Level", "Spell", "School", "Source", "Restriction", "GP", "DT"],
           [7, 34, 15, 22, 22, 8, 6])
    for i, r in enumerate(spells, start=2):
        sp.cell(row=i, column=1, value=int(r["Level"]))
        sp.cell(row=i, column=2, value=r["Name"])
        sp.cell(row=i, column=3, value=r["School"])
        sp.cell(row=i, column=4, value=r["Source"])
        sp.cell(row=i, column=5, value=r["Restriction"])
        sp.cell(row=i, column=6, value="=A%d*50" % i)
        sp.cell(row=i, column=7, value="=IF(A%d<=4,1,2)" % i)
    sp.freeze_panes = "A2"
    sp.auto_filter.ref = "A1:G%d" % (n + 1)

    # ----------------------------------------------------- Form Responses
    fr = wb.create_sheet(RESP_TAB)
    fr_cols = ["Timestamp", "Email Address", "Player name", "Character name",
               "Wizard level", "Table number", "Wizard subclass",
               "Contact after the event"]
    fr_cols += ["Level %d spells" % i for i in range(1, 10)]
    fr_cols += ["Other spells not in the lists above"]
    header(fr, 1, fr_cols, [18, 24, 18, 20, 12, 12, 22, 20] + [30] * 10)
    fr.cell(row=N_WIZ + 3, column=1,
            value="Google Forms writes into this tab. Do not edit or reorder columns - "
                  "the whole workbook references them by position.").font = NOTE_FONT
    fr.freeze_panes = "C2"

    # ------------------------------------------------------------ Wizards
    wz = wb.create_sheet("Wizards")
    header(wz, 1, ["Character", "Player", "Wizard level", "Max spell level", "Table",
                   "Subclass", "Order of Scribes?", "Contact", "Spells in book",
                   "Rare spells held"],
           [22, 18, 12, 13, 8, 22, 15, 20, 13, 13])
    for k in range(1, N_WIZ + 1):
        r = k + 1
        src = FIRST_RESP + k - 1
        wz.cell(row=r, column=1,
                value="=IF(%s=\"\",\"\",%s)" % (resp("D", src), resp("D", src)))
        wz.cell(row=r, column=2, value="=IF($A%d=\"\",\"\",%s)" % (r, resp("C", src)))
        wz.cell(row=r, column=3, value="=IF($A%d=\"\",\"\",%s)" % (r, resp("E", src)))
        # Wizard slot levels arrive at levels 1,3,5,...,17 -> MIN(9, roundup(lvl/2))
        wz.cell(row=r, column=4, value="=IF($A%d=\"\",\"\",MIN(9,CEILING($C%d/2,1)))" % (r, r))
        wz.cell(row=r, column=5, value="=IF($A%d=\"\",\"\",%s)" % (r, resp("F", src)))
        wz.cell(row=r, column=6, value="=IF($A%d=\"\",\"\",%s)" % (r, resp("G", src)))
        wz.cell(row=r, column=7,
                value="=IF($A%d=\"\",\"\",IF(ISNUMBER(SEARCH(\"Scribes\",%s)),\"Yes\",\"No\"))"
                      % (r, resp("G", src)))
        wz.cell(row=r, column=8, value="=IF($A%d=\"\",\"\",%s)" % (r, resp("H", src)))
        col = get_column_letter(3 + k)          # this wizard's column on Matrix
        wz.cell(row=r, column=9,
                value="=IF($A%d=\"\",\"\",COUNT(Matrix!$%s$2:$%s$%d))" % (r, col, col, n + 1))
        # spells this wizard holds that nobody else has
        wz.cell(row=r, column=10,
                value="=IF($A%d=\"\",\"\",COUNTIFS(Matrix!$%s$2:$%s$%d,1,Matrix!$%s$2:$%s$%d,1))"
                      % (r, col, col, n + 1,
                         get_column_letter(3 + N_WIZ + N_WIZ + 2),
                         get_column_letter(3 + N_WIZ + N_WIZ + 2), n + 1))
    wz.freeze_panes = "A2"

    # ------------------------------------------------------------- Matrix
    mx = wb.create_sheet("Matrix")
    name_first = 4                                  # column D
    name_last = 3 + N_WIZ                           # column AA
    help_first = name_last + 1                      # helper name block
    help_last = help_first + N_WIZ - 1
    own_col = help_last + 1                         # Owners
    cnt_col = own_col + 1                           # # Owners

    header(mx, 1, ["Level", "Spell", "School"], [7, 34, 15])
    for k in range(1, N_WIZ + 1):
        c = mx.cell(row=1, column=name_first + k - 1,
                    value="=IF(Wizards!A%d=\"\",\"\",Wizards!A%d)" % (k + 1, k + 1))
        c.fill, c.font = HDR_FILL, HDR_FONT
        mx.column_dimensions[get_column_letter(name_first + k - 1)].width = 14
        h = mx.cell(row=1, column=help_first + k - 1, value="helper %d" % k)
        h.fill, h.font = HDR_FILL, HDR_FONT
        mx.column_dimensions[get_column_letter(help_first + k - 1)].width = 14
        mx.column_dimensions[get_column_letter(help_first + k - 1)].hidden = True
    for col, lab, w in ((own_col, "Owners", 46), (cnt_col, "# Owners", 10)):
        c = mx.cell(row=1, column=col, value=lab)
        c.fill, c.font = HDR_FILL, HDR_FONT
        mx.column_dimensions[get_column_letter(col)].width = w

    for i, r in enumerate(spells, start=2):
        mx.cell(row=i, column=1, value=int(r["Level"]))
        mx.cell(row=i, column=2, value=r["Name"])
        mx.cell(row=i, column=3, value=r["School"])
        for k in range(1, N_WIZ + 1):
            src = FIRST_RESP + k - 1
            # The response cell for this spell's level: columns I..Q = levels 1..9.
            # Whole-column INDEX for the same insertion-safety reason as resp().
            lvl_cell = ("INDEX('%s'!$I$1:$Q$%d,%d,$A%d)"
                        % (RESP_TAB, RESP_LAST, src, i))
            other = resp("R", src)
            # Delimiter-wrapped so "Fire Bolt" never matches inside "Wall of Fire".
            test = ('OR(ISNUMBER(SEARCH(", "&$B{r}&", ", ", "&{lv}&", ")),'
                    'ISNUMBER(SEARCH(", "&$B{r}&", ", ", "&{ot}&", ")))'
                    ).format(r=i, lv=lvl_cell, ot=other)
            mx.cell(row=i, column=name_first + k - 1,
                    value="=IF(%s=\"\",\"\",IF(%s,1,\"\"))"
                          % ("Wizards!$A$%d" % (k + 1), test))
            nm = get_column_letter(name_first + k - 1)
            # ", Name" or "" -- the shape joined() expects.
            mx.cell(row=i, column=help_first + k - 1,
                    value="=IF(%s%d=1,\", \"&%s$1,\"\")" % (nm, i, nm))
        mx.cell(row=i, column=own_col,
                value="=" + joined(["$%s%d" % (get_column_letter(c), i)
                                    for c in range(help_first, help_last + 1)],
                                   ", "))
        mx.cell(row=i, column=cnt_col,
                value="=COUNT($%s%d:$%s%d)"
                      % (get_column_letter(name_first), i,
                         get_column_letter(name_last), i))
    mx.freeze_panes = "D2"

    # -------------------------------------------------------- Copy Planner
    cp = wb.create_sheet("Copy Planner")
    # A is wide enough for the input labels: they sit beside filled cells, so
    # they cannot overflow into B the way a label beside an empty cell would.
    cp.column_dimensions["A"].width = 19
    cp.column_dimensions["B"].width = 34
    cp.column_dimensions["C"].width = 15
    cp.column_dimensions["D"].width = 22
    cp.column_dimensions["E"].width = 16
    cp.column_dimensions["F"].width = 44
    cp.column_dimensions["G"].width = 8
    cp.column_dimensions["H"].width = 6
    cp.column_dimensions["I"].width = 8

    yellow = PatternFill("solid", fgColor="FFF3B0")

    def bold(cell, text):
        cp[cell] = text
        cp[cell].font = Font(bold=True)

    def note(cell, text):
        cp[cell] = text
        cp[cell].font = NOTE_FONT

    def entry(cell, value=None):
        c = cp[cell]
        c.value = value
        c.fill = yellow
        c.border = BORDER
        return c

    bold("A1", "Your character:")
    pick = entry(CP_CHAR)
    note("C1", "<- pick from the list, then filter Status to CAN COPY")

    # Budget panel: one row per resource, read left to right -- what you have,
    # what the ticked spells cost, what is left, and whether that is over.
    # Neither resource is assumed to be the one that runs out: which one a
    # character is short of is entirely down to the character.
    last = PLAN_TOP + n - 1
    want, lvl_col, gp_col = ("$I$%d:$I$%d" % (PLAN_TOP, last),
                             "$A$%d:$A$%d" % (PLAN_TOP, last),
                             "$G$%d:$G$%d" % (PLAN_TOP, last))
    for cell, text in (("B2", "Your total"), ("D2", "Left after copying")):
        cp[cell] = text
    cp["C2"] = '="Cost of "&%s&" ticked"' % ab(CP_COUNT)
    for cell in ("B2", "C2", "D2"):
        cp[cell].font = Font(bold=True, size=9)
        cp[cell].alignment = Alignment(horizontal="center")
        cp[cell].border = Border(bottom=THIN)

    # Order of Scribes (ALPG p.2) changes only the DT rate: ten level 1-4 spells
    # or five level 5-9 spells per 1 DT. GP is unaffected.
    # TRUE, not "x": the Want column is a Google Sheets checkbox, which stores
    # a boolean. COUNTIF/SUMIF match TRUE identically in Excel and Sheets.
    lo = 'COUNTIFS(%s,TRUE,%s,"<=4")' % (want, lvl_col)
    hi = 'COUNTIFS(%s,TRUE,%s,">=5")' % (want, lvl_col)
    dt_total = ('=IF({sc}="Yes",CEILING({lo}/10,1)+CEILING({hi}/5,1),{lo}+2*{hi})'
                .format(sc=ab(CP_SCRIBES), lo=lo, hi=hi))
    gp_total = "=SUMIF(%s,TRUE,%s)" % (want, gp_col)

    # Gold starts blank: the tool cannot guess it, and a made-up default would
    # read as a real comparison. Downtime starts at one session's award.
    for row_label, budget, start, total, left, status in (
            ("Gold (GP):", CP_GP_BUDGET, None, gp_total, CP_GP_LEFT, CP_GP_STATUS),
            ("Downtime (DT):", CP_DT_BUDGET, 10, dt_total, CP_DT_LEFT, CP_DT_STATUS)):
        row = budget[1:]
        bold("A" + row, row_label)
        entry(budget, start)
        cost = cp["C" + row]
        cost.value = total
        cp[left] = '=IF({b}="","",{b}-{c})'.format(b=ab(budget), c=ab("C" + row))
        cp[status] = ('=IF({b}="","<- enter your total to compare",'
                      'IF({l}<0,"OVER BUDGET","WITHIN BUDGET"))'
                      .format(b=ab(budget), l=ab(left)))
        for c in (cost, cp[left]):
            c.font = Font(bold=True, size=13)
            c.alignment = Alignment(horizontal="center")
            c.fill = BOX_FILL
            c.border = BORDER
        cp[status].font = Font(bold=True)
        # Over-budget colouring is a Sheets conditional format, applied by
        # PolishSpellExchangeSheet.gs -- it does not survive the .xlsx import.
        # The status text carries the same message without it.
    note("F4", "starts at 10, one session's award - change it to your real total")

    dv_gp = DataValidation(type="decimal", operator="greaterThanOrEqual",
                           formula1="0", allow_blank=True)
    dv_gp.error, dv_gp.errorTitle = "Your character's gold, 0 or more.", "Gold"
    dv_dt = DataValidation(type="whole", operator="greaterThanOrEqual",
                           formula1="0", allow_blank=True)
    dv_dt.error, dv_dt.errorTitle = "Your character's downtime, 0 or more.", "Downtime"
    for dv, cell in ((dv_gp, CP_GP_BUDGET), (dv_dt, CP_DT_BUDGET)):
        cp.add_data_validation(dv)
        dv.add(cp[cell])

    # Levelling up at the end of the Epic raises the spell level a character can
    # copy. It lives here, not on the Form: the Form records what the character
    # IS at registration, and whether to level up is often decided afterwards.
    # Like the budgets, it is the player's own planning input.
    reg_level = "INDEX(Wizards!$C$2:$C$%d,%s)" % (N_WIZ + 1, ab(CP_IDX))
    bold("A" + CP_NEW_LEVEL[1:], "New wizard level:")
    lvl_in = entry(CP_NEW_LEVEL)
    cp[CP_LEVEL_NOTE] = (
        '=IF({ix}="","<- levelling up at the end of this Epic? Enter your new wizard level",'
        '"<- levelling up at the end of this Epic? Enter your new wizard level. '
        'Registered at level "&{reg}&", copying up to spell level "&{mx})'
    ).format(ix=ab(CP_IDX), reg=reg_level, mx=ab(CP_MAXLVL))
    cp[CP_LEVEL_NOTE].font = NOTE_FONT
    dv_lvl = DataValidation(type="whole", operator="between",
                            formula1="1", formula2="20", allow_blank=True)
    dv_lvl.error = "A wizard level from 1 to 20."
    dv_lvl.errorTitle = "New wizard level"
    cp.add_data_validation(dv_lvl)
    dv_lvl.add(lvl_in)

    note("A7", "Switching character? Untick the Want column first - the costs "
               "add up every tick, whatever its Status.")

    # Hidden lookups, in hidden columns K:L rather than hidden rows, so the
    # visible panel above can grow without colliding with them.
    helpers = [
        (CP_IDX, "wizard index",
         '=IFERROR(MATCH({ch},Wizards!$A$2:$A${w},0),"")'),
        (CP_MAXLVL, "max spell level",
         '=IF({ix}="","",MIN(9,CEILING({lv}/2,1)))'),
        (CP_SCRIBES, "order of scribes",
         '=IF({ix}="","",INDEX(Wizards!$G$2:$G${w},{ix}))'),
        # The registered level, unless the new-level cell holds a higher one.
        # Levels are never lost, so a lower or non-numeric entry is ignored
        # rather than trusted.
        (CP_LVL_USED, "wizard level used",
         '=IF({ix}="","",IF(AND(ISNUMBER({nl}),{nl}>{reg}),MIN(20,{nl}),{reg}))'),
        (CP_COUNT, "spells ticked", "=COUNTIF(%s,TRUE)" % want),
    ]
    for cell, text, formula in helpers:
        cp.cell(row=cp[cell].row, column=cp[cell].column - 1, value=text)
        cp[cell] = formula.format(ch=ab(CP_CHAR), w=N_WIZ + 1, ix=ab(CP_IDX),
                                  lv=ab(CP_LVL_USED), nl=ab(CP_NEW_LEVEL),
                                  reg=reg_level)
    for col in ("K", "L"):
        cp.column_dimensions[col].hidden = True

    header(cp, PLAN_TOP - 1,
           ["Level", "Spell", "School", "Source", "Status", "Available from",
            "GP", "DT", "Want"])
    for i, r in enumerate(spells, start=PLAN_TOP):
        mrow = i - PLAN_TOP + 2                     # matching Matrix row
        cp.cell(row=i, column=1, value=int(r["Level"]))
        cp.cell(row=i, column=2, value=r["Name"])
        cp.cell(row=i, column=3, value=r["School"])
        cp.cell(row=i, column=4, value=r["Source"])
        mine = "INDEX(Matrix!$%s%d:$%s%d,1,%s)" % (
            get_column_letter(name_first), mrow,
            get_column_letter(name_last), mrow, ab(CP_IDX))
        # OWNED is tested first, so by the time we reach the "# Owners" test the
        # selected wizard is not among the owners -- a count of 0 really does
        # mean nobody else has it.
        cp.cell(row=i, column=5, value=(
            '=IF({ix}="","",'
            'IF({mine}=1,"OWNED",'
            'IF(Spells!$E${srow}<>"","RESTRICTED",'
            'IF($A{row}>{mx},"TOO HIGH",'
            'IF(Matrix!${cc}{mrow}=0,"NOBODY HAS IT","CAN COPY")))))'
        ).format(ix=ab(CP_IDX), mx=ab(CP_MAXLVL), mine=mine, srow=mrow, row=i,
                 mrow=mrow, cc=get_column_letter(cnt_col)))
        cp.cell(row=i, column=6, value="=Matrix!$%s%d" % (get_column_letter(own_col), mrow))
        cp.cell(row=i, column=7, value="=$A%d*50" % i)
        cp.cell(row=i, column=8, value="=IF($A%d<=4,1,2)" % i)
        w = cp.cell(row=i, column=9)
        w.fill = PatternFill("solid", fgColor="FFF3B0")
        w.border = BORDER
        w.alignment = Alignment(horizontal="center")
    cp.freeze_panes = "A%d" % PLAN_TOP
    cp.auto_filter.ref = "A%d:I%d" % (PLAN_TOP - 1, last)

    dv_char = DataValidation(type="list",
                             formula1="=Wizards!$A$2:$A$%d" % (N_WIZ + 1),
                             allow_blank=True, showDropDown=False)
    cp.add_data_validation(dv_char)
    dv_char.add(pick)
    # No .xlsx data validation for Want: real checkboxes are applied by
    # build/PolishSpellExchangeSheet.gs once the workbook is in Google Sheets.

    # ---------------------------------------------------------------- Calc
    # One column: for each spell row, its 1-based index if the player ticked
    # Want, otherwise blank. SMALL() over this compacts the ticked spells into
    # a gap-free list on the Copy Log without needing array formulas.
    cal = wb.create_sheet("Calc")
    cal["A1"] = "selected spell index (hidden helper for Copy Log)"
    for i in range(n):
        cal.cell(row=2 + i, column=1,
                 value="=IF('Copy Planner'!$I%d=TRUE,%d,\"\")" % (PLAN_TOP + i, i + 1))
    cal.sheet_state = "hidden"

    # ------------------------------------------------------------ Copy Log
    # Per SPELL, not per lender. Gold and downtime are expended by the copier,
    # so an invoice-shaped table addressed to each lender was simply the wrong
    # model: the wizard doing the copying pays, and what they need is a record
    # of what they bought, what it cost, and who they got it from.
    LOG_LABEL = 8                     # "paste this" banner
    LOG_TEXT = 9                      # the one cell players copy
    LOG_TOP = 12                      # first spell row
    LOG_ROWS = 40                     # a 10 DT budget cannot buy more than this
    ts = wb.create_sheet("Copy Log")
    ts.column_dimensions["A"].width = 5
    ts.column_dimensions["B"].width = 34
    ts.column_dimensions["C"].width = 7
    ts.column_dimensions["D"].width = 9
    ts.column_dimensions["E"].width = 34
    ts.column_dimensions["F"].width = 30
    ts.column_dimensions["G"].width = 3

    ts["A1"] = "COPY LOG"
    ts["A1"].font = TITLE_FONT
    ts["A2"] = ("Everything below is what YOUR selected character spends. "
                "Gold and downtime are expended, not paid to the lender.")
    ts["A2"].font = NOTE_FONT

    for col, label, ref in (("A", "Character", planner(CP_CHAR)),
                            ("C", "Spells", planner(CP_COUNT)),
                            ("D", "Total GP", planner(CP_GP_TOTAL)),
                            ("E", "Total DT", planner(CP_DT_TOTAL))):
        h = ts["%s4" % col]
        h.value = label
        h.font = Font(bold=True, size=9)
        v = ts["%s5" % col]
        v.value = "=%s" % ref
        v.font = Font(bold=True, size=13)
        v.fill = BOX_FILL
        v.border = BORDER

    ts["A6"] = ("Downtime is spent as one batch, so Total DT already applies the "
                "Order of Scribes rate where it is due - which is why there is no "
                "per-spell DT column. Per-spell GP is shown, because gold is per "
                "spell level regardless of subclass.")
    ts["A6"].font = NOTE_FONT

    frag_first = 8                                  # hidden per-spell log fragment
    ply_first = frag_first + 1                      # hidden per-wizard player names
    ply_last = ply_first + N_WIZ - 1
    log_cells = ["$%s$%d" % (get_column_letter(frag_first), LOG_TOP + k)
                 for k in range(LOG_ROWS)]

    # The single cell players actually copy. It previously sat unlabelled between
    # two notes and the table header, so it read as more commentary and nobody
    # could tell it was the thing to take -- hence the banner directly above.
    banner = ts.cell(row=LOG_LABEL, column=1,
                     value=("PASTE THIS ONE CELL INTO YOUR CHARACTER LOG   "
                            "(click A%d, then copy)" % LOG_TEXT))
    banner.font = HDR_FONT
    banner.fill = HDR_FILL
    for c in range(2, 7):
        ts.cell(row=LOG_LABEL, column=c).fill = HDR_FILL

    txt = ts.cell(row=LOG_TEXT, column=1, value=(
        '=IF({ix}="",'
        '"Pick your character on Copy Planner, then tick the spells you want.",'
        'IF({cnt}=0,"Nothing ticked yet.",'
        '"Copied "&{cnt}&" wizard spell(s) into spellbook: "'
        '&{frags}&". Total: "&{gp}'
        '&" GP and "&{dt}&" DT."))').format(
            ix=planner(CP_IDX), cnt=planner(CP_COUNT), gp=planner(CP_GP_TOTAL),
            dt=planner(CP_DT_TOTAL), frags=joined(log_cells, "; ")))
    txt.fill = PatternFill("solid", fgColor="FFF3B0")
    txt.border = BORDER
    txt.alignment = Alignment(wrap_text=True, vertical="top")
    ts.row_dimensions[LOG_TEXT].height = 58

    # Merge both rows across the table's width. The banner then reads as one
    # heading rather than text bleeding over five columns, and the log entry
    # becomes a single obvious click-target -- which is the whole point of it.
    # Merges survive the .xlsx -> Google Sheets conversion, and copying a merged
    # cell still yields its text.
    ts.merge_cells(start_row=LOG_LABEL, start_column=1, end_row=LOG_LABEL, end_column=6)
    ts.merge_cells(start_row=LOG_TEXT, start_column=1, end_row=LOG_TEXT, end_column=6)

    header(ts, LOG_TOP - 1,
           ["#", "Spell", "Level", "GP", "Copied from (character)",
            "Player", "", "log fragment"])
    for c in range(frag_first, ply_last + 1):
        ts.column_dimensions[get_column_letter(c)].hidden = True
    ts.column_dimensions[get_column_letter(frag_first)].width = 60

    own_letter = get_column_letter(own_col)
    for k in range(LOG_ROWS):
        r = LOG_TOP + k
        # k-th ticked spell, compacted out of the Calc column by SMALL().
        ts.cell(row=r, column=1,
                value="=IFERROR(SMALL(Calc!$A$2:$A$%d,%d),\"\")" % (n + 1, k + 1))
        ts.cell(row=r, column=2,
                value="=IF($A%d=\"\",\"\",INDEX(Spells!$B$2:$B$%d,$A%d))" % (r, n + 1, r))
        ts.cell(row=r, column=3,
                value="=IF($A%d=\"\",\"\",INDEX(Spells!$A$2:$A$%d,$A%d))" % (r, n + 1, r))
        ts.cell(row=r, column=4,
                value="=IF($A%d=\"\",\"\",$C%d*50)" % (r, r))
        ts.cell(row=r, column=5,
                value="=IF($A%d=\"\",\"\",INDEX(Matrix!$%s$2:$%s$%d,$A%d))"
                      % (r, own_letter, own_letter, n + 1, r))
        ts.cell(row=r, column=6,
                value="=IF($A%d=\"\",\"\",%s)"
                      % (r, joined(["$%s%d" % (get_column_letter(c), r)
                                    for c in range(ply_first, ply_last + 1)],
                                   ", ")))
        # Guard on $A (the index), never on a numeric column: an unused row
        # returns "" not 0, and testing =0 built stray " (from )" fragments.
        ts.cell(row=r, column=frag_first,
                value=("=IF($A%d=\"\",\"\",\"; \"&$B%d&\" (L\"&$C%d&\", \"&$D%d"
                       "&\" GP) from \"&$E%d)") % (r, r, r, r, r))
        for w in range(1, N_WIZ + 1):
            mcol = get_column_letter(name_first + w - 1)
            ts.cell(row=r, column=ply_first + w - 1,
                    value=("=IF($A%d=\"\",\"\",IF(INDEX(Matrix!$%s$2:$%s$%d,$A%d)=1,"
                           "\", \"&Wizards!$B$%d,\"\"))")
                          % (r, mcol, mcol, n + 1, r, w + 1))

    ts.cell(row=LOG_TOP + LOG_ROWS + 1, column=1,
            value=("=IF(%s>%d,\"More than %d spells ticked - "
                   "the list above is truncated.\",\"\")"
                   % (planner(CP_COUNT), LOG_ROWS, LOG_ROWS))
            ).font = NOTE_FONT
    ts.cell(row=LOG_TOP + LOG_ROWS + 3, column=1,
            value=("AL rule: you may copy from a character's spellbook immediately "
                   "after a session in which you both played - settle it at the event.")
            ).font = NOTE_FONT
    ts.freeze_panes = "A%d" % LOG_TOP

    os.makedirs(os.path.dirname(out), exist_ok=True)
    wb.save(out)
    print("Wrote %s" % out)
    print("  %d spells, %d wizard slots" % (n, N_WIZ))
    print("  tabs: %s" % ", ".join(w.title for w in wb.worksheets))


if __name__ == "__main__":
    build()
