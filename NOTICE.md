# Content notice

The **MIT licence in [LICENSE](LICENSE) covers the code in this repository** —
the Python scripts, the generated Apps Scripts, the workbook layout and the
documentation. It does not, and cannot, license anything owned by Wizards of
the Coast.

## What is included from D&D

`data/wizard-spells.csv` lists 350 wizard spells by **name, level, school and
source book**. That is all it contains. There are no spell descriptions, no
rules text, no statistics and no mechanics — it is an index, of the kind a book
prints at the back.

Everything else in this repository — the cost arithmetic, the eligibility rule,
the copy planner — implements rules from the *Player's Handbook* and the
*Adventurers League Player's Guide*. Those rules are cited where they are used;
their text is not reproduced.

## What is deliberately excluded

- **`TheMasterSpellbook/`** — saved D&D Beyond listing pages, the input to
  `extract_spells.py`. Copyrighted WotC content; never published here.
- **The local book exports** (`SPELLEXCHANGE_BOOKS`) — personal exports of
  purchased books, used only to attribute each spell to a source. Never
  published here.

Both are gitignored, and the scripts that need them fail with an explicit
message rather than degrading quietly.

## Fan Content Policy

This is unofficial Fan Content permitted under the [Wizards of the Coast Fan
Content Policy](https://company.wizards.com/en/legal/fancontentpolicy). It is
not approved or endorsed by Wizards. Portions of the materials used are
property of Wizards of the Coast. ©Wizards of the Coast LLC.

Adventurers League is a Wizards of the Coast organized play programme. This
project is a volunteer organizer's tool and has no affiliation with it.

## If you own the rights and object

Open an issue and the material will be removed.
