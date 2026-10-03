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
  data-connection     a saved query to an external source. Not the same thing
                      as an external link: this one re-runs, and when its path
                      is a mapped drive or a share it refreshes for the person
                      who built it and for nobody else.
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

It reads the old binary .xls too, which is not a zip of XML but an OLE
compound file holding a stream of BIFF records. Both layers are parsed here,
so hidden-sheet, cached-error, unprotected-formulas, external-link and
legacy-macro all work on a .xls, and the file's own created and last-saved
timestamps get printed - a published spreadsheet's last-saved date is often
the most useful fact in it. What does not work on .xls is volatile,
fragile-reference and approximate-lookup: those need the formula text, and
BIFF stores formulas as a token stream keyed by function index. Guessing those
indices would mean reporting findings this tool cannot stand behind, so it
does not claim them.

Nothing is installed and nothing leaves the machine. Python standard library
only - zipfile, xml.etree and struct - because a tool that asks a stranger to
install a dependency before it can read their own file will not get run.
"""
import datetime
import json
import os
import re
import struct
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
# A path that only exists on one machine: a drive letter, a UNC share, or a
# POSIX absolute path. Deliberately not matching http(s), which is a source a
# stranger's copy can actually still reach.
LOCAL_PATH = re.compile(r"^(?:[A-Za-z]:[\\/]|\\\\|/)")


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


# --- the old binary format ----------------------------------------------
# .xls is an OLE compound file holding a stream of BIFF records. Nothing in
# the standard library reads either layer, so both are here: about a hundred
# lines, and it means a published .xls gets a real answer instead of "this is
# not a zip". Four of the five published spreadsheets this tool has been
# pointed at in anger were .xls, so the shrug was most of the tool's output.
OLE_SIG = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
# The error a formula was holding when the file was saved. BIFF stores it as a
# code, not as text, so the mapping is the file format's and not a guess.
BIFF_ERR = {0x00: "#NULL!", 0x07: "#DIV/0!", 0x0F: "#VALUE!", 0x17: "#REF!",
            0x1D: "#NAME?", 0x24: "#NUM!", 0x2A: "#N/A"}
# Property ids in the SummaryInformation stream. Reported, never a finding:
# every Office file carries them, so they say nothing about the workbook's
# quality - they say when it was last true, which is a different question.
SI_PROPS = {4: "author", 8: "last-author", 12: "created", 13: "last-saved",
            18: "application"}


def _ole_streams(raw):
    """{name: bytes} for every stream in an OLE compound file."""
    ssz = 1 << struct.unpack_from("<H", raw, 0x1e)[0]
    msz = 1 << struct.unpack_from("<H", raw, 0x20)[0]
    nfat, dirstart, minicut, ministart, _nmini, difstart, ndif = \
        struct.unpack_from("<ii4xiiiii", raw, 0x2c)

    def at(n, size):
        return 512 + n * size

    # The FAT lives in sectors listed in the header, and past 109 of them in a
    # chain of DIFAT sectors. Small files never need the chain; big ones do.
    fatsecs = list(struct.unpack_from("<109i", raw, 0x4c))
    nxt, left = difstart, ndif
    while nxt >= 0 and left > 0:
        block = struct.unpack_from("<%di" % (ssz // 4), raw, at(nxt, ssz))
        fatsecs += list(block[:-1])
        nxt, left = block[-1], left - 1
    fat = []
    for s in fatsecs[:nfat]:
        if s < 0:
            break
        fat += list(struct.unpack_from("<%di" % (ssz // 4), raw, at(s, ssz)))

    def chain(start, table):
        out, n, seen = [], start, set()
        while n >= 0 and n not in seen and n < len(table):
            seen.add(n)
            out.append(n)
            n = table[n]
        return out

    def cat(start, size, table, sector):
        return b"".join(raw[at(n, sector):at(n, sector) + sector]
                        for n in chain(start, table))[:size]

    dirb = b"".join(raw[at(n, ssz):at(n, ssz) + ssz]
                    for n in chain(dirstart, fat))
    entries = []
    for i in range(0, len(dirb) - 127, 128):
        e = dirb[i:i + 128]
        nlen = struct.unpack_from("<H", e, 64)[0]
        name = e[:max(0, nlen - 2)].decode("utf-16-le", "replace")
        start, size = struct.unpack_from("<iq", e, 116)
        if name:
            entries.append({"name": name, "type": e[66], "start": start,
                            "size": size})
    # Streams under the mini cutoff are packed into one ordinary stream owned
    # by the root entry, indexed by its own allocation table.
    root = next((e for e in entries if e["type"] == 5), None)
    mini, minifat = b"", []
    if root and ministart >= 0:
        mini = b"".join(raw[at(n, ssz):at(n, ssz) + ssz]
                        for n in chain(root["start"], fat))
        for n in chain(ministart, fat):
            minifat += list(struct.unpack_from("<%di" % (ssz // 4), raw,
                                               at(n, ssz)))
    out = {}
    for e in entries:
        if e["type"] != 2:
            continue
        if e["size"] < minicut and minifat:
            out[e["name"]] = b"".join(
                mini[n * msz:(n + 1) * msz]
                for n in chain(e["start"], minifat))[:e["size"]]
        else:
            out[e["name"]] = cat(e["start"], e["size"], fat, ssz)
    return out, [e["name"] for e in entries]


def _props(si):
    """The SummaryInformation property set, as {label: value}."""
    out = {}
    try:
        secoff = struct.unpack_from("<i", si, 44)[0]
        nprops = struct.unpack_from("<i", si, secoff + 4)[0]
    except struct.error:
        return out
    # Property 1 is the code page the byte strings in this set are written in.
    # Reading them as latin-1 instead turned a Polish surname into mojibake,
    # which is the kind of detail that makes a stranger stop trusting output.
    enc = "cp1252"
    for k in range(nprops):
        try:
            pid, off = struct.unpack_from("<ii", si, secoff + 8 + k * 8)
        except struct.error:
            break
        if pid == 1:
            cp = struct.unpack_from("<h", si, secoff + off + 4)[0]
            enc = {65001: "utf-8"}.get(cp % 65536, "cp%d" % (cp % 65536))
            try:
                "".encode(enc)
            except LookupError:
                enc = "cp1252"
    for k in range(nprops):
        try:
            pid, off = struct.unpack_from("<ii", si, secoff + 8 + k * 8)
            base = secoff + off
            typ = struct.unpack_from("<i", si, base)[0]
        except struct.error:
            continue
        label = SI_PROPS.get(pid)
        if not label:
            continue
        if typ == 64:  # FILETIME, 100ns ticks since 1601
            ticks = struct.unpack_from("<Q", si, base + 4)[0]
            if ticks:
                out[label] = (datetime.datetime(1601, 1, 1) +
                              datetime.timedelta(microseconds=ticks / 10)
                              ).strftime("%Y-%m-%d %H:%M:%S")
        elif typ == 30:  # byte string, NUL terminated
            ln = struct.unpack_from("<i", si, base + 4)[0]
            val = si[base + 8:base + 8 + max(0, ln - 1)]
            val = val.decode(enc, "replace").strip("\x00").strip()
            if val:
                out[label] = val
    return out


def scan_xls(path, res, add):
    raw = open(path, "rb").read()
    try:
        streams, names = _ole_streams(raw)
    except (struct.error, IndexError, ValueError) as exc:
        add("unreadable", "the file starts like an OLE compound document but "
            "its allocation tables do not parse, so it is damaged or not "
            "really a spreadsheet", str(exc))
        return res
    book = streams.get("Workbook") or streams.get("Book")
    if book is None:
        add("not-a-workbook", "the compound file holds no Workbook stream, so "
            "whatever it is, it is not an Excel spreadsheet",
            ", ".join(n for n in names if not n.startswith("\x05"))[:120])
        return res

    add("legacy-binary", "this is the pre-2007 binary .xls format, so no "
        "browser-based tool opens it, newer Excel warns before it will, and "
        "every reader of it is working from a reverse engineered layout",
        os.path.basename(path))
    if any(n.startswith("_VBA_PROJECT") or n == "Macros" for n in names):
        add("legacy-macro", "the workbook contains a VBA project, so part of "
            "what it does is code that only desktop Excel runs",
            os.path.basename(path))
    for key in names:
        if key.endswith("SummaryInformation") and "Document" not in key:
            res["props"] = _props(streams[key])

    # One record stream: globals first, then one substream per sheet, each
    # opened by a BOF. Sheet protection is a record inside a sheet's substream,
    # so which substream we are in is what makes the protection check mean
    # anything.
    i, sheet, order = 0, None, []
    cur = {"name": "(workbook globals)", "formulas": 0, "protected": False,
           "errors": Counter()}
    order.append(cur)
    boundsheets = []
    while i + 4 <= len(book):
        t, n = struct.unpack_from("<HH", book, i)
        d = book[i + 4:i + 4 + n]
        if t == 0x0809 and i:  # BOF of a sheet substream
            sheet = (boundsheets.pop(0) if boundsheets else "sheet")
            cur = {"name": sheet, "formulas": 0, "protected": False,
                   "errors": Counter()}
            order.append(cur)
        elif t == 0x0085 and len(d) > 8:  # BOUNDSHEET
            grbit, cch, flags = struct.unpack_from("<HBB", d, 4)
            nm = (d[8:8 + cch * 2].decode("utf-16-le", "replace") if flags & 1
                  else d[8:8 + cch].decode("latin-1"))
            boundsheets.append(nm)
            state = {0: "visible", 1: "hidden", 2: "veryHidden"}.get(
                grbit & 3, "visible")
            res["sheets"].append({"name": nm, "state": state})
            if state != "visible":
                add("hidden-sheet", "the sheet is marked %s, so it does not "
                    "appear in the tab bar and its rules are invisible to "
                    "whoever uses the book" % state, nm)
        elif t == 0x0006 and len(d) >= 20:  # FORMULA
            cur["formulas"] += 1
            res["formulas"] += 1
            res["cells"] += 1
            # A formula's cached result is eight bytes. When the last two are
            # 0xFFFF it is not a number, and the first byte says what it is:
            # 2 means the saved answer was an error.
            if d[12] == 0xFF and d[13] == 0xFF and d[6] == 2:
                cur["errors"][BIFF_ERR.get(d[8], "#?")] += 1
        elif t in (0x00FD, 0x0204, 0x0203, 0x027E, 0x0205, 0x00BE):
            res["cells"] += 1
        elif t == 0x00BD and len(d) >= 6:  # MULRK, one record for many cells
            res["cells"] += (len(d) - 6) // 6
        elif t == 0x0012 and len(d) >= 2:  # PROTECT
            if struct.unpack_from("<H", d, 0)[0]:
                cur["protected"] = True
        elif t == 0x002F:  # FILEPASS
            add("encrypted", "the workbook is password protected, so nothing "
                "below this line could be read out of it",
                os.path.basename(path))
            return res
        elif t == 0x01AE and len(d) >= 4:  # SUPBOOK
            cch = struct.unpack_from("<H", d, 2)[0]
            # 0x0401 is the workbook's reference to itself and 0x3A01 an
            # add-in. Neither is an external file, and treating them as one
            # would report a broken link in every file that has a 3D formula.
            if cch not in (0x0401, 0x3A01):
                tgt = d[4:4 + cch * 2].decode("utf-16-le", "replace")
                tgt = "".join(c if c.isprintable() else "/" for c in tgt)
                add("external-link", "the workbook reads cells out of another "
                    "file, so the answer depends on a file that may not exist "
                    "any more", tgt.strip("/"))
        i += 4 + n

    for s in order:
        for err, count in s["errors"].items():
            add("cached-error", "%d cell%s in this sheet %s saved holding %s, "
                "so the book was last used in that state"
                % (count, "" if count == 1 else "s",
                   "was" if count == 1 else "were", err), s["name"])
        if s["formulas"] and not s["protected"]:
            add("unprotected-formulas", "the sheet carries %d formula%s and "
                "has no protection, so anyone can type over a rule and "
                "nothing says so"
                % (s["formulas"], "" if s["formulas"] == 1 else "s"),
                s["name"])
    return res


def scan(path):
    name = os.path.basename(path)
    res = {"file": name, "bytes": os.path.getsize(path), "findings": [],
           "sheets": [], "formulas": 0, "cells": 0}
    add = lambda rule, says, where="": res["findings"].append(
        {"rule": rule, "says": says, "where": where})

    # The signature decides, not zipfile.is_zipfile: that one hunts for an end
    # of central directory record anywhere in the file, and a 64 KB binary .xls
    # can contain those four bytes by accident. It did, on the first real .xls
    # this tool was pointed at, and the file came back "not a workbook".
    head = open(path, "rb").read(8)
    if head == OLE_SIG:
        return scan_xls(path, res, add)
    if not zipfile.is_zipfile(path):
        add("unreadable", "this is neither a zip nor an OLE compound file, so "
            "it is not any version of an Excel workbook", name)
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

    # A saved query is not an external link: external links read cells out of
    # another workbook, a connection re-runs a query against a source. Both
    # carry a path, and the path is the part that outlives the machine.
    # Found 4 Oct 2026 in six published device profiles that each still named
    # one vendor's mapped drive, so the rule is here because a real file had it.
    if "xl/connections.xml" in names:
        for conn in ET.fromstring(z.read("xl/connections.xml")):
            src = ""
            for child in conn:
                for attr in ("url", "sourceFile", "connection"):
                    if child.get(attr):
                        src = child.get(attr)
                        break
                if src:
                    break
            label = conn.get("name") or ""
            if LOCAL_PATH.match(src):
                add("data-connection", "the workbook carries a saved query to "
                    "a path on one machine - a mapped drive or a share - so it "
                    "refreshes for whoever set it up and for nobody else",
                    "%s -> %s" % (label, src) if label else src)
            elif src:
                add("data-connection", "the workbook carries a saved query to "
                    "an external source, so what it shows depends on something "
                    "outside the file", "%s -> %s" % (label, src) if label
                    else src)

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
        if r.get("props"):
            print("  %s" % "  ".join(
                "%s: %s" % (k, r["props"][k])
                for k in ("created", "last-saved", "last-author", "application")
                if k in r["props"]))
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
