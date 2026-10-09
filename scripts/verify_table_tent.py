"""Print the table tent to PDF in a headless browser and check it fits one A5 page.

    python scripts/verify_table_tent.py

The sign only works printed on ONE sheet. Its CSS is tuned to fit -- the body
padding is dropped in print because @page already supplies the margin -- and a
longer sentence or a bigger QR code can push it onto a second page without
anything looking wrong on screen. So it is printed three ways: the committed
placeholder, a short forms.gle link, and a long un-shortened Forms URL that has
to wrap. Each must be exactly one A5 page, and the two with a link must carry
a clickable link annotation.

The links are made up. Nothing is written into build/.

Needs Chrome, Chromium or Edge. Without one it says SKIPPED and exits 0 -- unless
REQUIRE_BROWSER is set, as it is in CI, where a skip would be a silent pass.
Point CHROME_PATH at a browser if it isn't found.
"""
import os
import pathlib
import re
import shutil
import subprocess
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import make_table_tent  # noqa: E402

sys.stdout.reconfigure(encoding="utf-8")

A5_PT = (419.53, 595.28)          # 148 x 210 mm
TOLERANCE_PT = 2.0

CASES = [
    ("placeholder (as committed)", None),
    # Obviously fake, but the real lengths: a random-looking code could resolve
    # to somebody's actual form, and this file is public.
    ("short forms.gle link", "https://forms.gle/xxxxxxxxxxxxxxxxx"),
    ("long Forms URL that wraps",
     "https://docs.google.com/forms/d/e/EXAMPLE-not-a-real-form-xxxxxxxxxxxxxx"
     "xxxxxxxxxxxxxxxxxxxxxxxxxx/viewform?usp=sf_link"),
]

CANDIDATES = [
    "google-chrome", "google-chrome-stable", "chromium", "chromium-browser",
    "chrome", "msedge",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
]


def find_browser():
    if os.environ.get("CHROME_PATH"):
        return os.environ["CHROME_PATH"]
    for c in CANDIDATES:
        found = shutil.which(c) or (c if os.path.isfile(c) else None)
        if found:
            return found
    return None


def print_pdf(browser, html, workdir, name):
    src = os.path.join(workdir, name + ".html")
    pdf = os.path.join(workdir, name + ".pdf")
    with open(src, "w", encoding="utf-8") as fh:
        fh.write(html)
    args = [browser, "--headless=new", "--disable-gpu", "--no-pdf-header-footer",
            "--print-to-pdf=" + pdf, pathlib.Path(src).as_uri()]
    if sys.platform.startswith("linux"):
        args.insert(1, "--no-sandbox")   # CI containers have no user namespaces
    subprocess.run(args, capture_output=True, timeout=120)
    if not os.path.exists(pdf):
        raise RuntimeError("the browser produced no PDF")
    with open(pdf, "rb") as fh:
        return fh.read()


def main():
    browser = find_browser()
    if not browser:
        msg = "no Chrome, Chromium or Edge found (set CHROME_PATH)"
        if os.environ.get("REQUIRE_BROWSER"):
            print("FAILED: %s" % msg)
            return 1
        print("SKIPPED: %s" % msg)
        return 0
    print("Printing with %s" % browser)

    failures = []

    case = [""]

    def check(label, got, want):
        ok = got == want
        print("  %-52s %-12s %s" % (label, got, "ok" if ok else "EXPECTED %s" % want))
        if not ok:
            failures.append("%s: %s" % (case[0], label))

    with tempfile.TemporaryDirectory() as tmp:
        for i, (label, url) in enumerate(CASES):
            print("\n%s" % label)
            case[0] = label
            pdf = print_pdf(browser, make_table_tent.render(url), tmp, "case%d" % i)
            pages = len(re.findall(rb"/Type\s*/Page(?!s)", pdf))
            check("pages", pages, 1)
            box = re.search(rb"/MediaBox\s*\[\s*0\s+0\s+([\d.]+)\s+([\d.]+)\s*\]", pdf)
            size = (float(box.group(1)), float(box.group(2))) if box else (0.0, 0.0)
            check("A5 page size",
                  all(abs(a - b) <= TOLERANCE_PT for a, b in zip(size, A5_PT)), True)
            if url:
                check("form link is clickable",
                      ("/URI (%s)" % url).encode() in pdf, True)

    print("\n%s" % ("ALL CHECKS PASSED" if not failures
                    else "FAILED: %s" % ", ".join(failures)))
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
