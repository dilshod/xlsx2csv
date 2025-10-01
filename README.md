
# xlsx2csv

> xlsx to csv converter (http://github.com/dilshod/xlsx2csv)

Converts xlsx files to csv format.
Handles large XLSX files. Fast and easy to use.

## Tested (supported) Python versions:
 - 2.4
 - 2.7
 - 3.4 to 3.14

## Installation:

```sh
sudo easy_install xlsx2csv
```
  or

```sh
pip install xlsx2csv
```


Also works standalone with only the *xlsx2csv.py* script.

**Usage:**
```
 xlsx2csv.py [-h] [-v] [-a] [-c OUTPUTENCODING] [-s SHEETID]
                   [-n SHEETNAME] [-d DELIMITER] [-l LINETERMINATOR]
                   [-f DATEFORMAT] [--floatformat FLOATFORMAT]
                   [-i] [-e] [-p SHEETDELIMITER]
                   [--hyperlinks]
                   [-I INCLUDE_SHEET_PATTERN [INCLUDE_SHEET_PATTERN ...]]
                   [-E EXCLUDE_SHEET_PATTERN [EXCLUDE_SHEET_PATTERN ...]] [-m]
                   xlsxfile [outfile]
```
**positional arguments:**
```
  xlsxfile              xlsx file path, use '-' to read from STDIN
  outfile               output csv file path, or directory if -s 0 is specified
```
**optional arguments:**
```
  -h, --help            show this help message and exit
  -v, --version         show program's version number and exit
  -a, --all             export all sheets
  -c OUTPUTENCODING, --outputencoding OUTPUTENCODING
                        encoding of output CSV **Python 3 only** (default: utf-8)
  -s SHEETID, --sheet SHEETID
                        sheet number to convert, 0 for all
  -n SHEETNAME, --sheetname SHEETNAME
                        sheet name to convert
  -d DELIMITER, --delimiter DELIMITER
                        delimiter - column delimiter in CSV, 'tab' or 'x09'
                        for a tab (default: comma ',')
  -l LINETERMINATOR, --lineterminator LINETERMINATOR
                        line terminator - line terminator in CSV, '\n' '\r\n'
                        or '\r' (default: os.linesep)
  -f DATEFORMAT, --dateformat DATEFORMAT
                        override date/time format (ex. %Y/%m/%d)
  --floatformat FLOATFORMAT
                        override float format (ex. %.15f)
  -i, --ignoreempty     skip empty lines
  -e, --escape          escape \r\n\t characters
  -p SHEETDELIMITER, --sheetdelimiter SHEETDELIMITER
                        sheet delimiter used to separate sheets, pass '' if
                        you do not need a delimiter, or 'x07' or '\\f' for form
                        feed (default: '--------')
  -q QUOTING, --quoting QUOTING
                        field quoting, 'none' 'minimal' 'nonnumeric' or 'all' (default: 'minimal')
  --hyperlinks, --hyperlinks
                        include hyperlinks
  -I INCLUDE_SHEET_PATTERN [INCLUDE_SHEET_PATTERN ...], --include_sheet_pattern INCLUDE_SHEET_PATTERN [INCLUDE_SHEET_PATTERN ...]
                        only include sheets with names matching the given pattern, only
                        affects when -a option is enabled.
  -E EXCLUDE_SHEET_PATTERN [EXCLUDE_SHEET_PATTERN ...], --exclude_sheet_pattern EXCLUDE_SHEET_PATTERN [EXCLUDE_SHEET_PATTERN ...]
                        exclude sheets with names matching the given pattern, only
                        affects when -a option is enabled.
  -m, --merge-cells     merge cells
```

Usage with a folder containing multiple `xlsx` files:
```
    python xlsx2csv.py /path/to/input/dir /path/to/output/dir
```
will output each file in the input directory converted to `.csv` in the output directory. If omitting the output directory, it will output the converted files in the input directory.

Usage from within Python:
```
  from xlsx2csv import Xlsx2csv

  # Recommended: using context manager for proper resource cleanup
  with Xlsx2csv("myfile.xlsx", outputencoding="utf-8") as xlsx2csv:
      xlsx2csv.convert("myfile.csv")

  # Simple usage (but may cause ResourceWarning in modern Python)
  Xlsx2csv("myfile.xlsx", outputencoding="utf-8").convert("myfile.csv")
```

Expat SAX parser is used for XML parsing.

See alternatives:

Perl:

https://metacpan.org/dist/Spreadsheet-Read/view/scripts/xlsx2csv

Bash:
http://kirk.webfinish.com/?p=91

Python:
http://github.com/staale/python-xlsx
http://github.com/leegao/pyXLSX

Ruby:
http://roo.rubyforge.org/

Java:
http://poi.apache.org/


## Meta

  Dilshod Temirkhdojaev – tdilshod@gmail.com

Distributed under the MIT license. See ``LICENSE`` for more information.

[https://github.com/dilshod](https://github.com/dilshod)
