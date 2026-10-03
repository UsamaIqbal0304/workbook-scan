#!/usr/bin/env python3
"""Write the two .xlsx fixtures workbook-scan.py is tested and demonstrated on.

Why generate them instead of committing two binaries: a fixture has to be
readable by the next person, and "here is the XML that produces the finding"
is the only form of that for a zip of XML. It also means the README's captured
run is reproducible by a stranger with no file from us and no file of ours
from a client - which matters, because the one real workbook this tool has
found a defect in belongs to somebody else.

    ./make-workbook-fixtures.py [OUTDIR]

Writes dirty.xlsx (one finding per rule the scanner has) and clean.xlsx (a
sheet with formulas, protected, nothing to report). Deterministic: same bytes
every run, so a diff means a real change.
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
