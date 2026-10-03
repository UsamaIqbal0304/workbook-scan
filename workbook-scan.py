#!/usr/bin/env python3
"""What a spreadsheet actually does, read out of the file rather than Excel.

    ./workbook-scan.py BOOK.xlsx [BOOK2.xlsx ...]
    ./workbook-scan.py --json BOOK.xlsx

Why this exists. A workbook that a business runs on is undocumented software.
Somebody wrote it, it accumulated rules for years, and the person who knows
why cell AH44 is 0.87 has usually left. Before anyone offers to replace one,
the honest first step is to read what it already does - and that is a file
format question, not an opinion.

So this opens the .xlsx as what it is, a zip of XML, and reports only things
that are in the bytes:

  cached-error        a formula whose last computed value was #REF!, #DIV/0!,
                      #VALUE!, #N/A, #NAME? or #NUM!. The file was saved in
                      that state, so somebody is looking at it.
  external-link       the workbook reads cells from another file. The path is
                      in the bytes, and it is usually on a machine nobody
                      has any more.
  volatile            NOW, TODAY, RAND or RANDBETWEEN, so the sheet gives a
                      different answer tomorrow with no input changed.
  fragile-reference   INDIRECT or OFFSET, which survive no row insertion.
  approximate-lookup  VLOOKUP or HLOOKUP with the range argument left off or
                      TRUE, which silently returns the wrong row on unsorted
                      data. The single most expensive default in Excel.
  hidden-sheet        a sheet marked hidden or veryHidden.
  unprotected-formulas a sheet with formulas and no sheetProtection element,
                      so a user can overwrite a rule by typing in a cell.
  legacy-macro        .xlsm/.xls content: VBA that no browser-based tool runs.
  defined-name-broken a named range pointing at #REF!.

Nothing is installed and nothing leaves the machine. Python standard library
only - zipfile and xml.etree - because a tool that asks a stranger to install
a dependency before it can read their own file will not get run.
"""
import json
import os
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from collections import Counter

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
ERRORS = ("#REF!", "#DIV/0!", "#VALUE!", "#N/A", "#NAME?", "#NUM!", "#NULL!")
VOLATILE = re.compile(r"\b(NOW|TODAY|RAND|RANDBETWEEN)\s*\(", re.I)
FRAGILE = re.compile(r"\b(INDIRECT|OFFSET)\s*\(", re.I)
# VLOOKUP(x, range, col) with no fourth argument, or an explicit TRUE, is an
# approximate match. Counting commas inside the call is enough to tell, as
# long as nested calls are not miscounted - so only the top level is split.
LOOKUP = re.compile(r"\b([VH]LOOKUP)\s*\(", re.I)


def _args(src, start):
    """The top-level arguments of a call whose '(' is at start."""
    depth, cur, out, i = 0, [], [], start
    while i < len(src):
        c = src[i]
        if c == "(":
            depth += 1
            if depth == 1:
                i += 1
                continue
        elif c == ")":
            depth -= 1
            if depth == 0:
                out.append("".join(cur))
                return out
        elif c == "," and depth == 1:
            out.append("".join(cur))
            cur = []
            i += 1
            continue
        if depth >= 1:
            cur.append(c)
        i += 1
    return out


def scan(path):
    name = os.path.basename(path)
    res = {"file": name, "bytes": os.path.getsize(path), "findings": [],
           "sheets": [], "formulas": 0, "cells": 0}
    add = lambda rule, says, where="": res["findings"].append(
        {"rule": rule, "says": says, "where": where})

    if not zipfile.is_zipfile(path):
        add("legacy-binary", "this is not a zip, so it is the old binary .xls "
            "format - every tool that reads it is guessing at a reverse "
            "engineered layout", name)
        return res

    z = zipfile.ZipFile(path)
    names = z.namelist()

    # sheet names, and which are hidden, come from the workbook part
    try:
        wb = ET.fromstring(z.read("xl/workbook.xml"))
    except KeyError:
        add("not-a-workbook", "there is no xl/workbook.xml in the zip, so "
            "this is not a spreadsheet", name)
        return res
    sheets = []
    for sh in wb.findall(".//m:sheets/m:sheet", NS):
        state = sh.get("state") or "visible"
        sheets.append({"name": sh.get("name"), "state": state})
        if state != "visible":
            add("hidden-sheet", "the sheet is marked %s, so it does not appear "
                "in the tab bar and its rules are invisible to whoever uses "
                "the book" % state, sh.get("name"))
    res["sheets"] = sheets

    for dn in wb.findall(".//m:definedNames/m:definedName", NS):
        if "#REF!" in (dn.text or ""):
            add("defined-name-broken", "the named range points at #REF!, so "
                "every formula using the name is already wrong",
                dn.get("name") or "")

    if any(n.startswith("xl/externalLinks/") for n in names):
        for n in names:
            if n.startswith("xl/externalLinks/_rels/"):
                rel = ET.fromstring(z.read(n))
                for r in rel:
                    tgt = r.get("Target") or ""
                    add("external-link", "the workbook reads cells out of "
                        "another file, so the answer depends on a file that "
                        "may not exist any more", tgt)

    if any(n.startswith("xl/vbaProject") for n in names):
        add("legacy-macro", "the workbook contains a VBA project, so part of "
            "what it does is code that only desktop Excel runs", name)

    # the sheets themselves
    sheet_files = sorted(n for n in names
                         if re.match(r"xl/worksheets/sheet\d+\.xml$", n))
    byname = {}
    for i, sf in enumerate(sheet_files):
        label = sheets[i]["name"] if i < len(sheets) else sf
        root = ET.fromstring(z.read(sf))
        cells = root.findall(".//m:sheetData/m:row/m:c", NS)
        res["cells"] += len(cells)
        has_formula = False
        errs = Counter()
        for c in cells:
            f = c.find("m:f", NS)
            v = c.find("m:v", NS)
            if f is not None:
                has_formula = True
                res["formulas"] += 1
                src = f.text or ""
                if VOLATILE.search(src):
                    add("volatile", "the formula uses a volatile function, so "
                        "the sheet answers differently tomorrow with no input "
                        "changed", "%s!%s" % (label, c.get("r") or ""))
                if FRAGILE.search(src):
                    add("fragile-reference", "the formula uses INDIRECT or "
                        "OFFSET, which breaks silently when a row is inserted",
                        "%s!%s" % (label, c.get("r") or ""))
                for lm in LOOKUP.finditer(src):
                    a = _args(src, lm.end() - 1)
                    if len(a) < 4 or a[3].strip().upper() in ("TRUE", "1"):
                        add("approximate-lookup",
                            "%s is set to approximate match, so on unsorted "
                            "data it returns a neighbouring row instead of no "
                            "answer" % lm.group(1).upper(),
                            "%s!%s" % (label, c.get("r") or ""))
            if c.get("t") == "e" and v is not None and v.text in ERRORS:
                errs[v.text] += 1
        for e, n in errs.items():
            add("cached-error", "%d cell%s in this sheet %s saved holding %s, "
                "so the book was last used in that state"
                % (n, "" if n == 1 else "s", "was" if n == 1 else "were", e),
                label)
        if has_formula and root.find("m:sheetProtection", NS) is None:
            add("unprotected-formulas", "the sheet carries formulas and has no "
                "protection, so anyone can type over a rule and nothing says "
                "so", label)
        byname[label] = zipfile.ZipFile(path).read(sf)
    res["_sheetbytes"] = byname
    return res


def main(argv):
    as_json = "--json" in argv
    paths = [a for a in argv[1:] if not a.startswith("--")]
    if not paths:
        print(__doc__.strip().split("\n\n")[1])
        return 2
    out = []
    for p in paths:
        r = scan(p)
        r.pop("_sheetbytes", None)
        out.append(r)
    if as_json:
        print(json.dumps(out, indent=1))
        return 0
    for r in out:
        print("=" * 70)
        def n(count, word):
            return "%d %s%s" % (count, word, "" if count == 1 else "s")
        print("%s  %s, %s, %s, %s"
              % (r["file"], n(r["bytes"], "byte"), n(len(r["sheets"]), "sheet"),
                 n(r["cells"], "cell"), n(r["formulas"], "formula")))
        print("=" * 70)
        if not r["findings"]:
            print("  nothing this scanner knows how to find.")
            print()
            continue
        counts = Counter(f["rule"] for f in r["findings"])
        for rule, n in counts.most_common():
            print("  %-22s %d" % (rule, n))
        print()
        seen = set()
        for f in r["findings"]:
            key = (f["rule"], f["says"][:40])
            if key in seen:
                continue
            seen.add(key)
            print("  %s  %s" % (f["rule"], f["where"]))
            print("    %s" % f["says"])
        print()
    return 1 if any(r["findings"] for r in out) else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
