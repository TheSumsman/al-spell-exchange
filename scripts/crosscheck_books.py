"""Independently derive wizard spells from the local book exports and diff them
against data/wizard-spells.csv (which comes from the D&D Beyond listing).

This is a correctness check on the scrape, NOT a data source. The listing should
be a strict superset: anything found here but missing there means a listing page
was never saved, or a filter was wrong.

Wizard attribution lives in three different places depending on a book's vintage:
  * PHB 2024 / Heroes of Faerun -- inline metadata "Level 3 Evocation (.., Wizard)"
  * Xanathar's Guide            -- class list at  id="WizardSpells"
  * Tasha's Cauldron            -- class list at  id="AdditionalWizardSpells"

Find these sections by ANCHOR ID, never by visible heading text: the markup is
id="WizardSpells" with no space, so searching for "Wizard Spells" finds nothing.
"""
import csv
import glob
import html
import os
import re
import sys
import unicodedata
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8")

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
CSV_PATH = os.path.join(ROOT, "data", "wizard-spells.csv")
# The D&D Beyond book exports are personal copies of purchased books and are
# deliberately not part of this repository. Set SPELLEXCHANGE_BOOKS to your
# own export directory; the fallback is a sibling of the repo root.
BOOKS = os.environ.get(
    "SPELLEXCHANGE_BOOKS",
    os.path.join(os.path.dirname(ROOT), "ALLog-sources", "Books"))

SCHOOLS = ("abjuration conjuration divination enchantment evocation "
           "illusion necromancy transmutation").split()
# Abbreviated forms appear in Tasha's ("evoc.", "ench.") and full names in XGE
# ("conjuration"). The abbreviations are prefixes of the full names, so one
# prefix alternation matches both dialects.
SCH_RE = "|".join(s[:5] for s in SCHOOLS)

# Sections that look similar but are NOT the wizard list -- never treat as wizard.
FORBIDDEN_ANCHORS = ("ArtificerSpellList", "DunamancySpellList", "ExpandedSpellList")


def norm(name):
    s = unicodedata.normalize("NFKC", html.unescape(name))
    s = s.replace("\u2019", "'").replace("\u2018", "'")
    return re.sub(r"\s+", " ", s).strip().lower()


def from_inline(folder_frag, label):
    """2024 markup: '<em>Level 3 Evocation (Sorcerer, Wizard)</em>'."""
    found = {}
    name_re = re.compile(r'spell-tooltip"[^>]*>([^<]+)</a></h3>')
    for folder in glob.glob(os.path.join(BOOKS, "*")):
        if folder_frag.lower() not in os.path.basename(folder).lower():
            continue
        for path in glob.glob(os.path.join(folder, "*.htm")):
            raw = open(path, encoding="utf-8", errors="replace").read()
            blocks = re.split(
                r'<h3 class="compendium-hr h4-override with-metadata heading-anchor"',
                raw)[1:]
            for blk in blocks:
                nm = name_re.search(blk)
                meta = re.search(r"<em>([^<]+)</em>", blk)
                if not (nm and meta):
                    continue
                m = re.match(r"Level (\d+) \w+\s*\(([^)]*)\)",
                             html.unescape(meta.group(1)).strip())
                if m and "Wizard" in m.group(2):
                    found[norm(nm.group(1))] = (html.unescape(nm.group(1)).strip(),
                                                int(m.group(1)), label)
    return found


def from_class_list(folder_frag, anchor, label):
    """2014 markup: a class spell list under id="<anchor>".

    Entries look like:  <strong>3rd Level</strong> ...
                        <em><a href="#IceKnife">Ice knife</a></em> (conjuration)
    """
    found = {}
    for folder in glob.glob(os.path.join(BOOKS, "*")):
        if folder_frag.lower() not in os.path.basename(folder).lower():
            continue
        for path in glob.glob(os.path.join(folder, "*.htm")):
            raw = open(path, encoding="utf-8", errors="replace").read()
            m = re.search(r'id="%s"' % anchor, raw)
            if not m:
                continue
            seg = raw[m.start():]
            # stop at the next class list so we never absorb another class
            nxt = re.search(r'<h\d[^>]*id="(?!%s)\w*Spell' % anchor, seg)
            if nxt:
                seg = seg[:nxt.start()]
            for bad in FORBIDDEN_ANCHORS:
                if bad != anchor and ('id="%s"' % bad) in seg:
                    seg = seg[:seg.find('id="%s"' % bad)]
            lvl = None
            for tok in re.finditer(
                    r"<strong>(Cantrip|\d)(?:st|nd|rd|th)?[^<]*</strong>"
                    r"|>([^<]+)</(?:a|em)>[^(]{0,14}\((?:%s)" % SCH_RE, seg, re.I):
                if tok.group(1):
                    lvl = 0 if tok.group(1).lower().startswith("c") else int(tok.group(1))
                elif lvl and tok.group(2):
                    nm = html.unescape(tok.group(2)).strip().rstrip("*").strip()
                    if len(nm) >= 3 and re.search(r"[A-Za-z]", nm):
                        found.setdefault(norm(nm), (nm, lvl, label))
    return found


def require_books():
    """Stop rather than report a meaningless PASS.

    This check passes when nothing found in the local books is missing from the
    listing. With no book exports nothing is found locally, so nothing can be
    missing, and the run prints PASS having compared zero spells. Refuse to
    start instead.
    """
    if not os.path.isdir(BOOKS):
        sys.exit("""Book exports not found: """ + BOOKS + """

  These are personal D&D Beyond exports of purchased books. They are
  deliberately not part of this repository, so a fresh clone cannot run
  this check. See "Rebuilding the spell list" in README.md.

  If you do have them, set SPELLEXCHANGE_BOOKS to that directory, or edit
  BOOKS at the top of this file.

  This checks the scrape, not the build. Skipping it does not affect the
  workbook, which is built from the committed data/wizard-spells.csv.
""")

def main():
    require_books()
    if not os.path.exists(CSV_PATH):
        sys.exit("Run extract_spells.py first -- %s not found" % CSV_PATH)

    listing = {}
    for r in csv.DictReader(open(CSV_PATH, encoding="utf-8")):
        listing[norm(r["Name"])] = (r["Name"], int(r["Level"]))

    local = {}
    for src in (from_inline("Player's Handbook", "PHB 2024"),
                from_inline("Forgotten Realms Heroes of Faer", "Heroes of Faerun"),
                from_class_list("Xanathar's Guide to Everything",
                                "WizardSpells", "Xanathar's Guide"),
                from_class_list("Tasha", "AdditionalWizardSpells",
                                "Tasha's Cauldron")):
        for k, v in src.items():
            local.setdefault(k, v)

    print("Listing (data/wizard-spells.csv) : %d leveled wizard spells" % len(listing))
    print("Local books (independent)        : %d" % len(local))
    for lbl, n in Counter(v[2] for v in local.values()).most_common():
        print("    %-20s %d" % (lbl, n))

    missing = {k: v for k, v in local.items() if k not in listing}
    print("\nIn local books but MISSING from the listing: %d" % len(missing))
    for k in sorted(missing, key=lambda k: (local[k][1], k)):
        nm, lv, lbl = local[k]
        print("    ! L%d %-34s %s" % (lv, nm, lbl))

    mismatch = [(local[k][0], local[k][1], listing[k][1])
                for k in local if k in listing and local[k][1] != listing[k][1]]
    print("\nSpell-level disagreements: %d" % len(mismatch))
    for nm, a, b in mismatch:
        print("    ! %-34s local L%d vs listing L%d" % (nm, a, b))

    # Canaries called out in the plan.
    print("\nSpot checks:")
    for name, lv in [("Fireball", 3), ("Wish", 9),
                     ("Abi-Dalzim's Horrid Wilting", 8), ("Ice Knife", 1),
                     ("Wardaway", None), ("Tasha's Mind Whip", 2),
                     ("Absorb Elements", 1)]:
        k = norm(name)
        if k not in listing:
            print("    ! %-30s ABSENT from listing" % name)
        elif lv is not None and listing[k][1] != lv:
            print("    ! %-30s L%d, expected L%d" % (name, listing[k][1], lv))
        else:
            print("      %-30s ok (L%d)" % (name, listing[k][1]))

    ok = not missing and not mismatch
    print("\n%s" % ("PASS - listing is a superset of the local books."
                    if ok else "REVIEW NEEDED - see items marked ! above."))
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
