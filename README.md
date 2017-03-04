# J

Simple data wrapper that attempts to wrap SheetJS libraries to provide a uniform
way to access data from Excel and other spreadsheet files:

- JS-XLS: [xlsjs on npm](http://npm.im/xlsjs)
- JS-XLSX: [xlsx on npm](http://npm.im/xlsx)
- JS-HARB: [harb on npm](http://npm.im/harb)

Excel files are parsed based on the content (not by filename).  For example, CSV
files can be renamed to .XLS and excel will do the right thing.

Supported Formats:

| Format                                                       | Read  | Write |
|:-------------------------------------------------------------|:-----:|:-----:|
| **Excel Worksheet/Workbook Formats**                         |:-----:|:-----:|
| Excel 2007+ XML Formats (XLSX/XLSM)                          |  :o:  |  :o:  |
| Excel 2007+ Binary Format (XLSB BIFF12)                      |  :o:  |  :o:  |
| Excel 2003-2004 XML Format (XML "SpreadsheetML")             |  :o:  |       |
| Excel 97-2004 (XLS BIFF8)                                    |  :o:  |       |
| Excel 5.0/95 (XLS BIFF5)                                     |  :o:  |       |
| Excel 4.0 (XLS/XLW BIFF4)                                    |  :o:  |       |
| Excel 3.0 (XLS BIFF3)                                        |  :o:  |       |
| Excel 2.0/2.1 (XLS BIFF2)                                    |  :o:  |  :o:  |
| **Excel Supported Text Formats**                             |:-----:|:-----:|
| Delimiter-Separated Values (CSV/TSV/DSV)                     |  :o:  |  :o:  |
| Data Interchange Format (DIF)                                |  :o:  |  :o:  |
| Symbolic Link (SYLK/SLK)                                     |  :o:  |  :o:  |
| Space-Delimited Text (PRN)                                   |  :o:  |       |
| UTF-16 Unicode Text (TXT)                                    |  :o:  |       |
| **Other Workbook/Worksheet Formats**                         |:-----:|:-----:|
| OpenDocument Spreadsheet (ODS)                               |  :o:  |  :o:  |
| Flat XML ODF Spreadsheet (FODS)                              |  :o:  |  :o:  |
| Uniform Office Format Spreadsheet (标文通 UOS1/UOS2)         |  :o:  |       |
| dBASE II/III/IV / Visual FoxPro (DBF)                        |  :o:  |       |
| **Other Common Spreadsheet Output Formats**                  |:-----:|:-----:|
| HTML Tables                                                  |       |  :o:  |
| Markdown Tables                                              |       |  :o:  |
| **Other Output Formats**                                     |:-----:|:-----:|
| XML Data (XML)                                               |       |  :o:  |
| SocialCalc                                                   |  :o:  |  :o:  |


## Installation

```bash
$ npm install -g j
```

## Node Library

```js
var J = require('j');
```

`J.readFile(filename)` opens the file specified by filename and returns an array
whose first object is the parsing object (XLS or XLSX) and whose second object
is the parsed file.

`J.utils` has various helpers that expect an array like those from readFile:

- `to_csv(w) / to_dsv(w, delim)` will generate CSV/DSV respectively
- `to_json(w)` will generate JSON row objects
- `to_html(w)` will generate simple HTML tables
- `to_formulae(w)` will generate lists of formulae
- `to_xml(w)` will generate simple XML
- `to_xlsx(w) / to_xlsm(w) / to_xlsb(w)` will generate XLSX/XLSM/XLSB workbooks
- `to_md(w)` will generate markdown tables

## CLI Tool

The node module ships with a binary `j` which has a help message:

```
$ j --help

  Usage: j.njs [options] <file> [sheetname]

  Options:

    -h, --help               output usage information
    -V, --version            output the version number
    -f, --file <file>        use specified file (- for stdin)
    -s, --sheet <sheet>      print specified sheet (default first sheet)
    -N, --sheet-index <idx>  use specified sheet index (0-based)
    -p, --password <pw>      if file is encrypted, try with specified pw
    -l, --list-sheets        list sheet names and exit
    -o, --output <file>      output to specified file
    -B, --xlsb               emit XLSB to <sheetname> or <file>.xlsb
    -M, --xlsm               emit XLSM to <sheetname> or <file>.xlsm
    -X, --xlsx               emit XLSX to <sheetname> or <file>.xlsx
    -Y, --ods                emit ODS  to <sheetname> or <file>.ods
    -2, --biff2              emit XLS  to <sheetname> or <file>.xls (BIFF2)
    -T, --fods               emit FODS to <sheetname> or <file>.fods (Flat ODS)
    -S, --formulae           print formulae
    -j, --json               emit formatted JSON (all fields text)
    -J, --raw-js             emit raw JS object (raw numbers)
    -A, --arrays             emit rows as JS objects (raw numbers)
    -x, --xml                emit XML
    -H, --html               emit HTML
    -m, --markdown           emit markdown table
    -D, --dif                emit data interchange format (dif)
    -K, --sylk               emit symbolic link (sylk)
    -E, --socialcalc         emit socialcalc
    -F, --field-sep <sep>    CSV field separator
    -R, --row-sep <sep>      CSV row separator
    -n, --sheet-rows <num>   Number of rows to process (0=all rows)
    --sst                    generate shared string table for XLS* formats
    --compress               use compression when writing XLSX/M/B and ODS
    --perf                   do not generate output
    --all                    parse everything
    --dev                    development mode
    --read                   read but do not print out contents
    -q, --quiet              quiet mode
```


## License

Please consult the attached LICENSE file for details.  All rights not explicitly granted by the Apache 2.0 license are reserved by the Original Author.

[![githalytics.com alpha](https://cruel-carlota.pagodabox.com/cb2e495863d0096f50a923515c7331b6 "githalytics.com")](http://githalytics.com/SheetJS/j)

[![Build Status](https://travis-ci.org/SheetJS/j.png?branch=master)](https://travis-ci.org/SheetJS/j)

[![Coverage Status](https://coveralls.io/repos/SheetJS/j/badge.png)](https://coveralls.io/r/SheetJS/j)

## Using J for diffing XLS/XLSB/XLSM/XLSX files

Using git textconv, you can use `J` to generate more meaningful diffs!

One-time configuration (`misc/gitdiff.sh`):

```
#!/bin/bash

# Define a sheetjs diff type that uses j
git config --global diff.sheetjs.textconv "j"

# Configure a user .gitattributes file that maps the xls{,x,m} files
touch ~/.gitattributes
cat <<EOF >>~/.gitattributes
*.xls diff=sheetjs
*.xlsb diff=sheetjs
*.xlsm diff=sheetjs
*.xlsx diff=sheetjs
*.XLS diff=sheetjs
*.XLSB diff=sheetjs
*.XLSM diff=sheetjs
*.XLSX diff=sheetjs
EOF

# Set the .gitattributes to be used for all repos on the system:
git config --global core.attributesfile '~/.gitattributes'
```

If you just want to compare formulae (for example, in a sheet using `NOW`):

```
git config --global diff.sheetjs.textconv "j -S"
```


NOTE: There are some known issues regarding global modules in Windows.  The best
bet is to `npm install j` in your git directory before diffing.

## Notes

Segmentation faults in node v0.10.31 stem from a bug in node.  J will throw an
error if it is running under that version.  Since versions prior to v0.10.30 do
not exhibit the problem, rolling back to a previous version of node is the best
remedy.  See <https://github.com/joyent/node/issues/8208> for more information.
