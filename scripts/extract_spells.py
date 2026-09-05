"""Build the AL-legal (Forgotten Realms) wizard spell list.

Primary source : saved D&D Beyond /spells listing pages, filtered to Class=Wizard
                 and the FR-legal source set  ->  name, level, school
Source labels  : joined by name from a local directory of D&D Beyond book exports
                 (personal copies of purchased books; see SPELLEXCHANGE_BOOKS below)

Cantrips are dropped: they are never kept in a spellbook.
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
LISTING = os.path.join(ROOT, "TheMasterSpellbook")
# The D&D Beyond book exports are personal copies of purchased books and are
# deliberately not part of this repository. Set SPELLEXCHANGE_BOOKS to your
# own export directory; the fallback is a sibling of the repo root.
BOOKS = os.environ.get(
    "SPELLEXCHANGE_BOOKS",
    os.path.join(os.path.dirname(ROOT), "ALLog-sources", "Books"))
OUT_CSV = os.path.join(ROOT, "data", "wizard-spells.csv")
OUT_OPTS = os.path.join(ROOT, "data", "form-options")

WIZARD_CLASS_ID = "2190886"

SCHOOLS = ("abjuration conjuration divination enchantment evocation "
           "illusion necromancy transmutation").split()

LEVEL_WORD = {"Cantrip": 0, "1st": 1, "2nd": 2, "3rd": 3, "4th": 4,
              "5th": 5, "6th": 6, "7th": 7, "8th": 8, "9th": 9}

# Most-recent-first. ALPG p.1: "use the most recent version of all D&D content",
# so an earlier entry here wins when a spell is reprinted in several books.
SOURCE_PRIORITY = [
    ("Arcana Unleashed", "Arcana Unleashed"),
    ("Player's Handbook", "PHB 2024"),
    ("Forgotten Realms Heroes of Faer", "Heroes of Faerun"),
    ("Eberron Forge of the Artificer", "Forge of the Artificer"),
    ("Ravenloft The Horrors Within", "Ravenloft: THW"),
    ("The Book of Many Things", "Book of Many Things"),
    ("Bigby Presents Glory of the Giants", "Bigby Presents"),
    ("Planescape Adventures in the Multiverse", "Planescape"),
    ("Spelljammer Adventures in Space", "Spelljammer"),
    ("Fizban's Treasury of Dragons", "Fizban's"),
    ("Tasha", "Tasha's Cauldron"),
    ("Icewind Dale Rime of the Frostmaiden", "Icewind Dale"),
    ("Xanathar's Guide to Everything", "Xanathar's Guide"),
    ("Sword Coast Adventurer", "SCAG"),
    ("Lost Laboratory of Kwalish", "Kwalish"),
]


def norm(name):
    """Case- and apostrophe-insensitive key.

    Xanathar's lists are sentence case ('Absorb elements'), PHB 2024 is title
    case ('Absorb Elements'), and curly apostrophes appear in 'Tasha's'.
    """
    s = unicodedata.normalize("NFKC", html.unescape(name))
    s = s.replace("\u2019", "'").replace("\u2018", "'")
    return re.sub(r"\s+", " ", s).strip().lower()


def parse_listing():
    """Read every saved listing page. Returns (spells, n_files, pages_advertised)."""
    files = sorted(glob.glob(os.path.join(LISTING, "*.htm")))
    if not files:
        sys.exit("""No saved listing pages found in: """ + LISTING + """

  TheMasterSpellbook/ holds ~140 MB of saved D&D Beyond listing pages. It is
  copyrighted WotC content and is deliberately not part of this repository.

  To recreate it, see "Rebuilding the spell list" in README.md.

  data/wizard-spells.csv is committed, so nothing downstream needs this
  script: build_workbook.py and make_form_script.py both read the CSV.
""")

    spells = {}
    advertised = 0
    for path in files:
        raw = open(path, encoding="utf-8", errors="replace").read()

        seen_class = set(re.findall(r"filter-class=(\d+)", raw))
        if seen_class and seen_class != {WIZARD_CLASS_ID}:
            print("  ! %s: unexpected class filter %s"
                  % (os.path.basename(path), sorted(seen_class)))

        linked = [int(p) for p in re.findall(r"&amp;page=(\d+)\"", raw)]
        advertised = max([advertised] + linked)

        body = raw[raw.find('<div class="listing-body">'):]
        for row in re.split(r'(?=<div class="info" data-isopen)', body)[1:]:
            slug = re.search(r'data-slug="([^"]+)"', row)
            school = re.search(r'<div class="school (\w+)"', row)
            level = re.search(r'spell-level">\s*<span>([^<]+)</span>', row)
            name = re.search(r'spell-name">.*?class="link">([^<]+)</a>', row, re.S)
            if not (slug and school and level and name):
                continue

            lv = LEVEL_WORD.get(level.group(1).strip())
            if lv is None or lv == 0:  # cantrips are not kept in a spellbook
                continue

            sch = school.group(1).lower()
            if sch not in SCHOOLS:
                print("  ! unknown school %r for %s" % (sch, name.group(1)))
                continue

            key = norm(name.group(1))
            entry = spells.setdefault(key, {
                "Name": html.unescape(name.group(1)).strip(),
                "Level": lv,
                "School": sch.title(),
                "slugs": set(),
            })
            entry["slugs"].add(slug.group(1))
            if entry["Level"] != lv or entry["School"] != sch.title():
                print("  ! conflicting data for %s: L%d/%s vs L%d/%s"
                      % (entry["Name"], entry["Level"], entry["School"],
                         lv, sch.title()))

    return spells, len(files), advertised


# A spell heading always carries a spell-tooltip anchor holding the display name.
_SPELL_HEADING = re.compile(
    r'<a class="tooltip-hover spell-tooltip"[^>]*>([^<]+)</a>\s*</h\d>')

# Immediately after the heading comes the level/school line, in one of two
# dialects.  The school whitelist is what keeps CLASS FEATURES out: a naive
# "heading followed by Nth-level" pattern also matches "Breath of the Dragon --
# 3rd-level feature" and inflates the map with hundreds of phantom entries.
_META = re.compile(
    r"(?:Level \d+ (?:%(s)s)\b"                 # 2024: "Level 8 Necromancy (Wizard)"
    r"|\d(?:st|nd|rd|th)-level\s+(?:%(s)s)\b"   # 2014: "8th-level necromancy"
    r"|(?:%(s)s)\s+cantrip"                     # 2014 cantrip
    r"|(?:%(s)s)\s+Cantrip)" % {"s": "|".join(SCHOOLS)}, re.I)

# 2014 books that predate the tooltip markup put the name in a bare heading.
_BARE_HEADING = re.compile(
    r"<h\d[^>]*>\s*(?:<[^>]+>\s*)*([A-Z][^<]{2,45}?)\s*(?:<[^>]+>\s*)*</h\d>"
    r"\s*<p[^>]*>\s*<em>\s*\d(?:st|nd|rd|th)-level\s+(?:%s)" % "|".join(SCHOOLS),
    re.I)


def _names_in(raw):
    """Every spell name defined in one book page, both markup dialects."""
    names = list(_BARE_HEADING.findall(raw))

    # The level/school line follows the heading, but an illustration block can
    # sit between the two -- so search as far as the NEXT spell heading rather
    # than a fixed window, which silently dropped e.g. XGE's Shadow Blade.
    heads = list(_SPELL_HEADING.finditer(raw))
    for i, m in enumerate(heads):
        end = heads[i + 1].start() if i + 1 < len(heads) else len(raw)
        if _META.search(raw, m.end(), min(end, m.end() + 3000)):
            names.append(m.group(1))
    return names


def source_map():
    """norm(name) -> display source, built from the local book exports."""
    out = {}
    for folder_frag, display in SOURCE_PRIORITY:  # priority order: first wins
        for folder in sorted(glob.glob(os.path.join(BOOKS, "*"))):
            if not os.path.isdir(folder):
                continue
            if folder_frag.lower() not in os.path.basename(folder).lower():
                continue
            for path in glob.glob(os.path.join(folder, "*.htm")):
                raw = open(path, encoding="utf-8", errors="replace").read()
                for n in _names_in(raw):
                    out.setdefault(norm(n), display)
    return out


def require_books():
    """Stop rather than quietly emit a plausible-looking wrong answer.

    With no book exports the source join finds nothing, every spell falls back
    to the generic 'Other AL-legal (FR)' label, and the run still exits 0. That
    is a silent downgrade of the data, so refuse to start instead.
    """
    if not os.path.isdir(BOOKS):
        sys.exit("""Book exports not found: """ + BOOKS + """

  These are personal D&D Beyond exports of purchased books. They are
  deliberately not part of this repository, so a fresh clone cannot run
  this script. See "Rebuilding the spell list" in README.md.

  If you do have them, set SPELLEXCHANGE_BOOKS to that directory, or edit
  BOOKS at the top of this file.

  data/wizard-spells.csv is committed, so nothing downstream needs this
  script: build_workbook.py and make_form_script.py both read the CSV.
""")

def main():
    require_books()
    spells, n_files, advertised = parse_listing()
    print("Parsed %d listing page(s); listing advertises %d pages."
          % (n_files, advertised))
    complete = n_files >= advertised
    if not complete:
        print("  ! INCOMPLETE: %d page(s) still missing -- spell list will be short."
              % (advertised - n_files))

    collapsed = sum(1 for s in spells.values() if len(s["slugs"]) > 1)
    if collapsed:
        print("  collapsed %d spell(s) appearing under multiple slugs "
              "(2014/2024 reprints)." % collapsed)

    srcmap = source_map()
    print("Source map: %d spell names from local books." % len(srcmap))

    # A few spells live in book chapters that are not in the local export at all
    # (e.g. Lost Laboratory of Kwalish Appendix E). data/source-overrides.csv
    # records those explicitly, with the evidence, rather than leaving them
    # generic. Delete a row once the real chapter is available locally.
    ov_path = os.path.join(ROOT, "data", "source-overrides.csv")
    n_ov = 0
    if os.path.exists(ov_path):
        with open(ov_path, encoding="utf-8") as fh:
            for row in csv.DictReader(fh):
                key = norm(row["Name"])
                if key not in srcmap:
                    srcmap[key] = row["Source"]
                    n_ov += 1
        print("Source overrides applied: %d" % n_ov)

    rows = []
    unsourced = 0
    for key, s in spells.items():
        src = srcmap.get(key)
        if not src:
            src = "Other AL-legal (FR)"
            unsourced += 1
        rows.append({"Level": s["Level"], "Name": s["Name"],
                     "School": s["School"], "Source": src, "Restriction": ""})
    rows.sort(key=lambda r: (r["Level"], r["Name"].lower()))

    os.makedirs(os.path.dirname(OUT_CSV), exist_ok=True)
    with open(OUT_CSV, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, ["Level", "Name", "School", "Source", "Restriction"])
        w.writeheader()
        w.writerows(rows)

    os.makedirs(OUT_OPTS, exist_ok=True)
    for lv in range(1, 10):
        names = [r["Name"] for r in rows if r["Level"] == lv]
        with open(os.path.join(OUT_OPTS, "level-%d.txt" % lv), "w",
                  encoding="utf-8") as fh:
            fh.write("\n".join(names) + ("\n" if names else ""))

    print("\nWrote %d leveled wizard spells -> %s" % (len(rows), OUT_CSV))
    print("  unsourced (labelled 'Other AL-legal (FR)'): %d" % unsourced)
    print("  by level :", sorted(Counter(r["Level"] for r in rows).items()))
    for s, c in Counter(r["Source"] for r in rows).most_common():
        print("    %-24s %d" % (s, c))
    return 0 if complete else 1


if __name__ == "__main__":
    sys.exit(main())
