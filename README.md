# workbook-scan

What a spreadsheet actually does, read out of the file rather than out of Excel.
Twelve things that are in the bytes: cached errors, external links, saved data
connections, volatile functions, `INDIRECT`/`OFFSET`, approximate `VLOOKUP`,
hidden sheets, formulas nothing protects, broken defined names, and the two
cases where the file is not what its extension says.

One Python file, standard library only - `zipfile` and `xml.etree`. A tool that
asks someone to install a dependency before it can read their own file does not
get run.

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

## What it does not do

- It does not grade a workbook or give it a score. Twelve rules, each either
  matched in the bytes or not.
- It does not read `.xls` or `.xlsm` content. It says so and stops: those are
  a different format and a VBA project, and pretending otherwise would be the
  bug worth avoiding.
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
and look at - and `fixtures/clean.xlsx`, which has formulas, is protected, and
reports nothing. Byte-identical on every run, so a diff means a real change.

## Running it

```
$ ./make-workbook-fixtures.py
dirty.xlsx       3503 bytes, 10 members
clean.xlsx       1584 bytes, 5 members

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
```

## Licence

MIT. Written by Usama Iqbal at [Plantroom Labs](https://plantroomlabs.com),
alongside the free [Niagara and BACnet scanners](https://plantroomlabs.com/tools/).
