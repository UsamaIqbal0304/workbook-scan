#!/usr/bin/env python3
"""Write the two .xlsx fixtures workbook-scan.py is tested and demonstrated on.

Why generate them instead of committing two binaries: a fixture has to be
readable by the next person, and "here is the XML that produces the finding"
is the only form of that for a zip of XML. It also means the README's captured
run is reproducible by a stranger with no file from us and no file of ours
from a client - which matters, because the one real workbook this tool has
found a defect in belongs to somebody else.

    ./make-workbook-fixtures.py [OUTDIR]

Writes dirty.xlsx (one finding per rule the scanner has), clean.xlsx (a sheet
with formulas, protected, nothing to report) and dirty.xls (the same idea in
the old binary format). Deterministic: same bytes every run, so a diff means a
real change.

One honest limit on dirty.xls. Its OLE container is valid - `file` reads it as
"CDFV2 Microsoft Excel" and 7z lists the Workbook stream out of it - but the
BIFF stream inside carries only the records the scanner's rules are about, and
not the font, format and XF tables a spreadsheet application insists on.
LibreOffice therefore refuses to open it. That is fine for what it is, a test
of the reader, and it is said here so nobody discovers it as a surprise: to
check the .xls path against a file Excel would open, point the scanner at a
real .xls.
"""
import os
import sys
import zipfile

CT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
%s</Types>"""

SHEET_CT = ('<Override PartName="/xl/worksheets/sheet%d.xml" ContentType='
            '"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>\n')

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

M = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


CONNECTIONS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<connections xmlns="%s"><connection id="1" name="RatesTemplate" '
    'type="4" refreshedVersion="0" background="1"><webPr xml="1" '
    'sourceData="1" url="N:\\_Templates\\rates\\RatesTemplate.xml" '
    'htmlTables="1"/></connection></connections>')


def book(sheets, defined_names="", extern=False, connections=False):
    """sheets: list of (name, state, xml). Returns the dict of zip members."""
    tabs = "".join(
        '<sheet name="%s" sheetId="%d" r:id="rId%d"%s/>'
        % (n, i + 1, i + 1, "" if s == "visible" else ' state="%s"' % s)
        for i, (n, s, _) in enumerate(sheets))
    ext_rel = ""
    parts = {}
    if extern:
        # An external link is a relationship to a path, plus the part it names.
        ext_rel = ('<Relationship Id="rIdX" Type="http://schemas.openxmlformats'
                   '.org/officeDocument/2006/relationships/externalLink" '
                   'Target="externalLinks/externalLink1.xml"/>')
        parts["xl/externalLinks/externalLink1.xml"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<externalLink xmlns="%s"><externalBook/></externalLink>' % M)
        parts["xl/externalLinks/_rels/externalLink1.xml.rels"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package'
            '/2006/relationships"><Relationship Id="rId1" Type="http://schemas'
            '.openxmlformats.org/officeDocument/2006/relationships/externalLink'
            'Path" Target="/old-server/finance/rates-2014.xlsx" '
            'TargetMode="External"/></Relationships>')

    conn_ct = conn_rel = ""
    if connections:
        # A saved web query: the part, its content type and its relationship.
        # The path is a mapped drive, which is the case the rule is about.
        conn_rel = ('<Relationship Id="rIdC" Type="http://schemas.openxmlformats'
                    '.org/officeDocument/2006/relationships/connections" '
                    'Target="connections.xml"/>')
        conn_ct = ('<Override PartName="/xl/connections.xml" ContentType='
                   '"application/vnd.openxmlformats-officedocument'
                   '.spreadsheetml.connections+xml"/>\n')
        parts["xl/connections.xml"] = CONNECTIONS % M

    parts["[Content_Types].xml"] = CT % ("".join(
        SHEET_CT % (i + 1) for i in range(len(sheets))) + conn_ct)
    parts["_rels/.rels"] = ROOT_RELS
    parts["xl/workbook.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<workbook xmlns="%s" xmlns:r="%s"><sheets>%s</sheets>%s</workbook>'
        % (M, RNS, tabs, defined_names))
    parts["xl/_rels/workbook.xml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006'
        '/relationships">%s%s</Relationships>'
        % ("".join('<Relationship Id="rId%d" Type="http://schemas.openxmlformats'
                   '.org/officeDocument/2006/relationships/worksheet" '
                   'Target="worksheets/sheet%d.xml"/>' % (i + 1, i + 1)
                   for i in range(len(sheets))), ext_rel + conn_rel))
    for i, (_, _, xml) in enumerate(sheets):
        parts["xl/worksheets/sheet%d.xml" % (i + 1)] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<worksheet xmlns="%s">%s</worksheet>' % (M, xml))
    return parts


def rows(cells, protection=""):
    """cells: (ref, inner_xml) or (ref, inner_xml, cell_type)."""
    out, by_row = [], {}
    for cell in cells:
        ref, body = cell[0], cell[1]
        t = (' t="%s"' % cell[2]) if len(cell) > 2 else ""
        by_row.setdefault(int("".join(c for c in ref if c.isdigit())), []).append(
            (ref, body, t))
    for r in sorted(by_row):
        out.append("<row r=\"%d\">%s</row>"
                   % (r, "".join('<c r="%s"%s>%s</c>' % (ref, t, b)
                                 for ref, b, t in by_row[r])))
    return "<sheetData>%s</sheetData>%s" % ("".join(out), protection)


def write(path, parts):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name in sorted(parts):
            # A fixed date_time keeps the bytes identical between runs.
            info = zipfile.ZipInfo(name, date_time=(2026, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, parts[name])
    print("%-14s %6d bytes, %d members" % (os.path.basename(path),
                                           os.path.getsize(path), len(parts)))


out = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
out = os.path.join(out, "fixtures")
os.makedirs(out, exist_ok=True)

# One finding per rule, each in a cell a reader can go and look at.
DIRTY = rows([
    # cached-error: the saved value is the error, not a number
    ("B2", "<f>SUM(Deleted!A1:A9)</f><v>#REF!</v>", "e"),
    ("B3", "<f>C3/D3</f><v>#DIV/0!</v>", "e"),
    # volatile
    ("B4", "<f>TODAY()-C4</f><v>412</v>"),
    # fragile-reference
    ("B5", '<f>INDIRECT("Rates!B"&amp;C5)</f><v>0.87</v>'),
    # approximate-lookup: no fourth argument, and a nested call so the
    # argument splitter has something to get wrong
    ("B6", "<f>VLOOKUP(A6,Rates!$A$1:$D$99,MATCH(A5,Rates!$A$1:$D$1,0))</f><v>12</v>"),
    # external-link, written as the formula a reader would see
    ("B7", "<f>'/old-server/finance/[rates-2014.xlsx]Sheet1'!$B$2</f><v>1.4</v>"),
])
HIDDEN = rows([("A1", "<v>1</v>")])
CLEAN_IN_DIRTY = rows([("A1", "<f>B1*2</f><v>4</v>"), ("B1", "<v>2</v>")],
                      '<sheetProtection sheet="1"/>')

dirty = book(
    [("Pricing", "visible", DIRTY),
     ("Old rates", "hidden", HIDDEN),
     ("Checks", "visible", CLEAN_IN_DIRTY)],
    defined_names=('<definedNames><definedName name="margin">#REF!'
                   '</definedName></definedNames>'),
    extern=True, connections=True)
write(os.path.join(out, "dirty.xlsx"), dirty)

CLEAN = rows([("A1", "<v>7</v>"), ("B1", "<f>A1*1.2</f><v>8.4</v>")],
             '<sheetProtection sheet="1"/>')
write(os.path.join(out, "clean.xlsx"),
      book([("Sheet1", "visible", CLEAN)]))


# --- the binary fixture --------------------------------------------------
# workbook-scan reads .xls as well, and four of the five published
# spreadsheets it has been pointed at were .xls, so that path needs a fixture
# too. Writing one means writing both layers by hand: an OLE compound file
# holding a stream of BIFF8 records. It is about sixty lines and it is the only
# way a stranger can run the .xls rules without a file of somebody else's.
import struct


def biff(t, data=b""):
    return struct.pack("<HH", t, len(data)) + data


def bof(dt):
    # vers 0x0600 is BIFF8. The build and year fields are what Excel 97 wrote;
    # nothing reads them, but a plausible value beats zeros.
    return biff(0x0809, struct.pack("<HHHHII", 0x0600, dt, 0x0DBB, 0x07CC,
                                    0x000080C9, 0x00000206))


def boundsheet(name, pos, state):
    nm = name.encode("utf-16-le")
    return biff(0x0085, struct.pack("<iHBB", pos, state, len(name), 0x01) + nm)


def formula_error(row, col, code):
    # The cached result is eight bytes: a type byte, the error code, and
    # 0xFFFF in the last two to say "this is not a number".
    result = struct.pack("<BBBxxxBB", 0x02, 0x00, code, 0xFF, 0xFF)
    rgce = struct.pack("<BHH", 0x24, row, col)  # ptgRef, one cell
    return biff(0x0006, struct.pack("<HHH", row, col, 15) + result
                + struct.pack("<HIH", 0x0002, 0, len(rgce)) + rgce)


def xls_stream():
    """The Workbook stream: globals, then one substream per sheet."""
    sheets = [("Prices", 0x0000), ("Old rates", 0x0001)]
    # Sheet 1 holds an unprotected formula whose saved answer was #REF!; sheet
    # 2 is hidden. Between them that is every .xls rule the scanner has.
    bodies = [bof(0x0010) + formula_error(1, 1, 0x17) + biff(0x000A),
              bof(0x0010) + biff(0x0012, struct.pack("<H", 1)) + biff(0x000A)]
    # lbPlyPos is an absolute offset into the stream, so the globals have to be
    # built twice: once to learn their length, once with the real offsets.
    for _ in range(2):
        head = bof(0x0005)
        head += b"".join(boundsheet(n, 0, st) for n, st in sheets)
        head += biff(0x000A)
        base = len(head)
        offs, run = [], base
        for b in bodies:
            offs.append(run)
            run += len(b)
    out = bof(0x0005)
    out += b"".join(boundsheet(n, offs[i], st)
                    for i, (n, st) in enumerate(sheets))
    out += biff(0x000A) + b"".join(bodies)
    return out


def ole(streams):
    """An OLE compound file holding the given {name: bytes}, 512-byte sectors.

    Every stream here is written as a full sector chain rather than through the
    mini stream, which is legal and keeps this readable.
    """
    SZ, FREE, END = 512, -1, -2
    names = list(streams)
    data, starts = [], {}
    for n in names:
        starts[n] = len(data)
        b = streams[n]
        b += b"\0" * (-len(b) % SZ)
        data += [b[i:i + SZ] for i in range(0, len(b), SZ)]
    # Sector order: stream data, then the directory, then the FAT.
    dir_start = len(data)
    entries = [("Root Entry", 5, END, 0)] + [
        (n, 2, starts[n], len(streams[n])) for n in names]
    dirb = b""
    for i, (n, typ, start, size) in enumerate(entries):
        nm = n.encode("utf-16-le") + b"\0\0"
        e = nm + b"\0" * (64 - len(nm))
        e += struct.pack("<HBB", len(nm), typ, 1)        # name len, type, black
        left = 1 if typ == 5 and entries[1:] else FREE   # root's child
        e += struct.pack("<iii", FREE, FREE, left)
        e += b"\0" * 16 + struct.pack("<I", 0) + b"\0" * 16  # clsid, state, ts
        e += struct.pack("<iq", start, size)
        dirb += e + b"\0" * (128 - len(e))
    dirb += b"\0" * (-len(dirb) % SZ)
    dirsecs = [dirb[i:i + SZ] for i in range(0, len(dirb), SZ)]
    fat_start = dir_start + len(dirsecs)

    fat = [END] * (fat_start + 1)
    for n in names:
        s = starts[n]
        count = len(streams[n])
        count = (count + SZ - 1) // SZ
        for k in range(count - 1):
            fat[s + k] = s + k + 1
        fat[s + count - 1] = END
    for k in range(len(dirsecs) - 1):
        fat[dir_start + k] = dir_start + k + 1
    fat[dir_start + len(dirsecs) - 1] = END
    fat[fat_start] = -3  # the FAT sector itself
    fat += [FREE] * (-len(fat) % (SZ // 4))
    fatb = struct.pack("<%di" % len(fat), *fat)
    fatsecs = [fatb[i:i + SZ] for i in range(0, len(fatb), SZ)]

    difat = [fat_start + i for i in range(len(fatsecs))]
    difat += [FREE] * (109 - len(difat))
    head = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\0" * 16
    head += struct.pack("<HHHHHH", 0x003E, 3, 0xFFFE, 9, 6, 0)
    head += b"\0" * 4  # reserved is six bytes from 0x22, two of them above
    head += struct.pack("<iiiiiiii", 0, len(fatsecs), dir_start, 0, 4096,
                        END, 0, END)
    head += struct.pack("<i", 0)
    head += struct.pack("<109i", *difat)
    head += b"\0" * (SZ - len(head))
    return head + b"".join(data) + b"".join(dirsecs) + b"".join(fatsecs)


path = os.path.join(out, "dirty.xls")
with open(path, "wb") as f:
    f.write(ole({"Workbook": xls_stream()}))
print("%-14s %6d bytes, %d members" % ("dirty.xls", os.path.getsize(path), 1))
