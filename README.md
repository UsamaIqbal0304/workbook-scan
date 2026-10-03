# workbook-scan

What a spreadsheet actually does, read out of the file rather than out of Excel.
Twelve things that are in the bytes: cached errors, external links, saved data
connections, volatile functions, `INDIRECT`/`OFFSET`, approximate `VLOOKUP`,
hidden sheets, formulas nothing protects, broken defined names, and the two
cases where the file is not what its extension says.

It reads the old binary `.xls` as well, which is a different file format
entirely - an OLE compound file holding a stream of BIFF records - so five of
the rules work on a workbook saved in 1997.

One Python file, standard library only - `zipfile`, `xml.etree` and `struct`.
A tool that asks someone to install a dependency before it can read their own
file does not get run.

```
./workbook-scan.py BOOK.xlsx [MORE.xlsx ...]
./workbook-scan.py --json BOOK.xlsx
```

Exit status is 1 when it finds something and 0 when it does not, so it drops
into a build the same way a linter does.

## Why this exists

A workbook a business runs on is undocumented software. Somebody wrote it, it
accumulated rules for years, and the person who knew why cell AH44 is 0.87 has
usually left. Before anyone offers to repair or replace one, the honest first
step is to read what it already does - and that is a file-format question, not
an opinion.

Three of the twelve rules are worth the trouble on their own.

**A cached error is not a cosmetic problem.** When a formula's last computed
value is `#REF!`, Excel saves that error *into the file*, with the cell typed
`e`. It means the book was in that state when somebody last pressed save, so
whatever the sheet is for, it has been answering with a hole in it since.
This program was written after finding exactly that in a published engineering
calculator - the broken cell was the check the calculator exists to perform,
and its single file revision was nine years old. Nobody had opened it with the
question "does this still work" in that time, because nothing asks.

**An approximate `VLOOKUP` is the most expensive default in the format.**
`VLOOKUP(x, range, col)` with the fourth argument left off, or set to `TRUE`,
does not return "not found" on unsorted data. It returns a neighbouring row.
Checking that properly needs real argument counting rather than counting
commas, because `VLOOKUP(A6, Rates!$A$1:$D$99, MATCH(A5, Rates!$A$1:$D$1, 0))`
has four commas and three arguments; this program splits arguments at the top
level only.

**A saved data connection outlives the machine that made it.**
`xl/connections.xml` holds queries the workbook re-runs, and the path is in the
file. This rule exists because six device profiles published for download by
one manufacturer each still carried the same connection, named after a seventh
device, pointing at a mapped drive: `N:\_Templates\...`. Nobody downloading
those files can refresh them, and the path says more about how the files are
built than anyone intended to publish. An external link and a connection are
not the same finding - a link reads cells out of another file, a connection
re-runs a query - so they are reported separately.

## What it reads, and where from

An `.xlsx` is a zip of XML. The program opens it as one and reads:

- `xl/workbook.xml` - the sheet list with each sheet's `state`, and the
  `definedNames` block.
- each `xl/worksheets/sheetN.xml`, resolved through
  `xl/_rels/workbook.xml.rels` rather than by filename order, because sheet
  order and part number are not the same thing.
- every `<c>` element: its `t` attribute, its `<f>` formula text and its `<v>`
  cached value. An error finding requires both `t="e"` and an error string, so
  a cell that merely contains the text `#REF!` is not reported.
- `xl/externalLinks/_rels/*.rels` - the real path of each external workbook,
  which is where the `/old-server/...` paths live.
- `xl/connections.xml` - each saved query's name and source path, with a drive
  letter, a UNC share or an absolute path reported as machine-specific and an
  `https://` source reported as merely external.

It never opens Excel, never evaluates a formula, never writes to the file it
reads, and nothing leaves the machine.

## The old binary `.xls`

A `.xls` is not a zip and `zipfile.is_zipfile` is not how to tell: that
function hunts for an end-of-central-directory signature anywhere in the file,
and a 64 KB binary workbook can contain those four bytes by coincidence. One
did, on the first real `.xls` this program was pointed at, which came back as
"not a workbook". The format is decided by the first eight bytes instead.

What the program then parses, with no library:

- the OLE compound file - header, FAT, the DIFAT chain for files too big for
  the header's 109 entries, the directory, and the mini stream where parts
  under 4096 bytes are packed - to get at the `Workbook` stream.
- the BIFF record stream inside it, segmented by `BOF` so that a sheet
  protection record is attributed to the sheet whose substream it is in. That
  is what makes `unprotected-formulas` mean anything on a `.xls`.
- `BOUNDSHEET` for sheet names and the hidden flag, `FORMULA` for the eight
  byte cached result - where `0xFFFF` in the last two bytes says "not a
  number" and a leading `2` says the saved answer was an error - `PROTECT`,
  `FILEPASS` for an encrypted book, and `SUPBOOK` for an external file. A
  `SUPBOOK` of `0x0401` or `0x3A01` is the workbook referring to itself or to
  an add-in, not an external file, and reporting those would put a broken link
  in every workbook that has a 3-D formula.
- `SummaryInformation` - the `created` and `last-saved` timestamps, the
  application that wrote it and the name it recorded, decoded with the code
  page the property set names rather than assumed to be Latin-1. Those strings
  are printed, never reported as a finding: every Office file has them. For a
  spreadsheet somebody published for download, the last-saved date is usually
  the most useful fact in the file.

## What it does not do

- It does not grade a workbook or give it a score. Twelve rules, each either
  matched in the bytes or not.
- It does not read the *formulas* in a `.xls`, so `volatile`,
  `fragile-reference` and `approximate-lookup` are `.xlsx`-only. BIFF stores a
  formula as a token stream keyed by function index, and guessing those indices
  would mean reporting findings the program cannot stand behind.
- It does not run the VBA in an `.xlsm`. It reports that a VBA project is
  there, which is the part that matters when the question is whether a browser
  can ever do what the workbook does.
- It does not tell you a finding is a defect. A volatile function is correct
  in a sheet that is meant to answer differently today; the finding is that
  the sheet does, not that it should not.

## The fixtures

The repo carries its own test data, generated rather than committed as opaque
binaries, because a fixture nobody can read is not a fixture:

```
./make-workbook-fixtures.py
```

writes `fixtures/dirty.xlsx` - one finding per rule, each in a cell you can go
and look at - `fixtures/clean.xlsx`, which has formulas, is protected, and
reports nothing, and `fixtures/dirty.xls`, which is the same idea written as a
compound file and a BIFF stream by hand. Byte-identical on every run, so a diff
means a real change.

One honest limit on that last one. Its container is valid - `file` reports
`CDFV2 Microsoft Excel` and 7-Zip will list the `Workbook` stream out of it -
but the stream holds only the records these rules are about, and not the font,
format and XF tables a spreadsheet application insists on, so LibreOffice
declines to open it. It is a test of the reader, not a specimen workbook. To
exercise the `.xls` path against something Excel would open, point the program
at a real `.xls`.

## Running it

```
$ ./make-workbook-fixtures.py
dirty.xlsx       3503 bytes, 10 members
clean.xlsx       1584 bytes, 5 members
dirty.xls        2048 bytes, 1 members

$ ./workbook-scan.py fixtures/dirty.xlsx
======================================================================
dirty.xlsx  3503 bytes, 3 sheets, 9 cells, 7 formulas
======================================================================
  cached-error           2
  hidden-sheet           1
  defined-name-broken    1
  external-link          1
  data-connection        1
  volatile               1
  fragile-reference      1
  approximate-lookup     1
  unprotected-formulas   1

  hidden-sheet  Old rates
    the sheet is marked hidden, so it does not appear in the tab bar and its rules are invisible to whoever uses the book
  defined-name-broken  margin
    the named range points at #REF!, so every formula using the name is already wrong
  external-link  /old-server/finance/rates-2014.xlsx
    the workbook reads cells out of another file, so the answer depends on a file that may not exist any more
  data-connection  RatesTemplate -> N:\_Templates\rates\RatesTemplate.xml
    the workbook carries a saved query to a path on one machine - a mapped drive or a share - so it refreshes for whoever set it up and for nobody else
  volatile  Pricing!B4
    the formula uses a volatile function, so the sheet answers differently tomorrow with no input changed
  fragile-reference  Pricing!B5
    the formula uses INDIRECT or OFFSET, which breaks silently when a row is inserted
  approximate-lookup  Pricing!B6
    VLOOKUP is set to approximate match, so on unsorted data it returns a neighbouring row instead of no answer
  cached-error  Pricing
    1 cell in this sheet was saved holding #REF!, so the book was last used in that state
  unprotected-formulas  Pricing
    the sheet carries formulas and has no protection, so anyone can type over a rule and nothing says so

$ ./workbook-scan.py fixtures/clean.xlsx
======================================================================
clean.xlsx  1584 bytes, 1 sheet, 2 cells, 1 formula
======================================================================
  nothing this scanner knows how to find.

$ ./workbook-scan.py fixtures/dirty.xls
======================================================================
dirty.xls  2048 bytes, 2 sheets, 1 cell, 1 formula
======================================================================
  legacy-binary          1
  hidden-sheet           1
  cached-error           1
  unprotected-formulas   1

  legacy-binary  dirty.xls
    this is the pre-2007 binary .xls format, so no browser-based tool opens it, newer Excel warns before it will, and every reader of it is working from a reverse engineered layout
  hidden-sheet  Old rates
    the sheet is marked hidden, so it does not appear in the tab bar and its rules are invisible to whoever uses the book
  cached-error  Prices
    1 cell in this sheet was saved holding #REF!, so the book was last used in that state
  unprotected-formulas  Prices
    the sheet carries 1 formula and has no protection, so anyone can type over a rule and nothing says so
```

## Licence

MIT. Written by Usama Iqbal at [Plantroom Labs](https://plantroomlabs.com),
alongside the free [Niagara and BACnet scanners](https://plantroomlabs.com/tools/).
