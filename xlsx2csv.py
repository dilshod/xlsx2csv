#!/usr/bin/env python
#
#   Copyright information
#
#	Copyright (C) 2010-2012 Dilshod Temirkhodjaev <tdilshod@gmail.com>
#
#   License
#
#	This program is free software; you can redistribute it and/or modify
#	it under the terms of the GNU General Public License as published by
#	the Free Software Foundation; either version 2 of the License, or
#	(at your option) any later version.
#
#	This program is distributed in the hope that it will be useful,
#	but WITHOUT ANY WARRANTY; without even the implied warranty of
#	MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
#	GNU General Public License for more details.
#
#	You should have received a copy of the GNU General Public License
#	along with this program. If not, see <http://www.gnu.org/licenses/>.

__author__ = "Dilshod Temirkhodjaev <tdilshod@gmail.com>"
__license__ = "GPL-2+"
__version__ = "0.6"

import csv, datetime, zipfile, string, sys, os, re
import xml.parsers.expat
from xml.dom import minidom
try:
    # python2.4
    from cStringIO import StringIO
except:
    pass
try:
    from argparse import ArgumentParser
except:
    # python2.4
    from optparse import OptionParser

# see also ruby-roo lib at: http://github.com/hmcgowan/roo
FORMATS = {
  'general' : 'float',
  '0' : 'float',
  '0.00' : 'float',
  '#,##0' : 'float',
  '#,##0.00' : 'float',
  '0%' : 'percentage',
  '0.00%' : 'percentage',
  '0.00e+00' : 'float',
  'mm-dd-yy' : 'date',
  'd-mmm-yy' : 'date',
  'd-mmm' : 'date',
  'mmm-yy' : 'date',
  'h:mm am/pm' : 'date',
  'h:mm:ss am/pm' : 'date',
  'h:mm' : 'time',
  'h:mm:ss' : 'time',
  'm/d/yy h:mm' : 'date',
  '#,##0 ;(#,##0)' : 'float',
  '#,##0 ;[red](#,##0)' : 'float',
  '#,##0.00;(#,##0.00)' : 'float',
  '#,##0.00;[red](#,##0.00)' : 'float',
  'mm:ss' : 'time',
  '[h]:mm:ss' : 'time',
  'mmss.0' : 'time',
  '##0.0e+0' : 'float',
  '@' : 'float',
  'yyyy\\-mm\\-dd' : 'date',
  'dd/mm/yy' : 'date',
  'hh:mm:ss' : 'time',
  "dd/mm/yy\\ hh:mm" : 'date',
  'dd/mm/yyyy hh:mm:ss' : 'date',
  'yy-mm-dd' : 'date',
  'd-mmm-yyyy' : 'date',
  'm/d/yy' : 'date',
  'm/d/yyyy' : 'date',
  'dd-mmm-yyyy' : 'date',
  'dd/mm/yyyy' : 'date',
  'mm/dd/yy hh:mm am/pm' : 'date',
  'mm/dd/yyyy hh:mm:ss' : 'date',
  'yyyy-mm-dd hh:mm:ss' : 'date',
}
STANDARD_FORMATS = {
  0 : 'general',
  1 : '0',
  2 : '0.00',
  3 : '#,##0',
  4 : '#,##0.00',
  9 : '0%',
  10 : '0.00%',
  11 : '0.00e+00',
  12 : '# ?/?',
  13 : '# ??/??',
  14 : 'mm-dd-yy',
  15 : 'd-mmm-yy',
  16 : 'd-mmm',
  17 : 'mmm-yy',
  18 : 'h:mm am/pm',
  19 : 'h:mm:ss am/pm',
  20 : 'h:mm',
  21 : 'h:mm:ss',
  22 : 'm/d/yy h:mm',
  37 : '#,##0 ;(#,##0)',
  38 : '#,##0 ;[red](#,##0)',
  39 : '#,##0.00;(#,##0.00)',
  40 : '#,##0.00;[red](#,##0.00)',
  45 : 'mm:ss',
  46 : '[h]:mm:ss',
  47 : 'mmss.0',
  48 : '##0.0e+0',
  49 : '@',
}

class XlsxException(Exception):
    pass

class InvalidXlsxFileException(XlsxException):
    pass

class SheetNotFoundException(XlsxException):
    pass

class OutFileAlreadyExistsException(XlsxException):
    pass

class Xlsx2csv:
    """
     Usage: Xlsx2csv("test.xslx", **params).convert("test.csv", sheetid=1)
     parameters:
       sheetid - sheet no to convert (0 for all sheets)
       dateformat - override date/time format
       delimiter - csv columns delimiter symbol
       sheet_delimiter - sheets delimiter used when processing all sheets
       skip_empty_lines - skip empty lines
    """

    def __init__(self, xlsxfile, dateformat=None, delimiter=",", sheetdelimiter="--------", skip_empty_lines=False, escape_strings=False, cmd=False):
        try:
            self.ziphandle = zipfile.ZipFile(xlsxfile)
        except (zipfile.BadZipfile, IOError):
            if cmd:
                sys.stderr.write("Invalid xlsx file: " + xlsxfile + os.linesep)
                sys.exit(1)
            raise InvalidXlsxFileException("Invalid xlsx file: " + xlsxfile)

        self.dateformat = dateformat
        self.delimiter = delimiter
        self.sheetdelimiter = sheetdelimiter
        self.skip_empty_lines = skip_empty_lines
        self.cmd = cmd
        self.py3 = sys.version_info[0] == 3

        self.shared_strings = self._parse(SharedStrings, "xl/sharedStrings.xml")
        self.styles = self._parse(Styles, "xl/styles.xml")
        self.workbook = self._parse(Workbook, "xl/workbook.xml")
        if escape_strings:
            self.shared_strings.escape_strings()

    def convert(self, outfile, sheetid=1):
        """outfile - path to file or filehandle"""
        if sheetid > 0:
            self._convert(sheetid, outfile)
        else:
            if isinstance(outfile, str):
                if not os.path.exists(outfile):
                    os.makedirs(outfile)
                elif os.path.isfile(outfile):
                    if cmd:
                        sys.stderr.write("File " + outfile + " already exists!" + os.linesep)
                        sys.exit(1)
                    raise OutFileAlreadyExistsException("File " + outfile + " already exists!")
            for s in self.workbook.sheets:
                sheetname = s['name']
                if not self.py3:
                    sheetname = sheetname.encode('utf-8')
                of = outfile
                if isinstance(outfile, str):
                    of = os.path.join(outfile, sheetname + '.csv')
                elif self.sheetdelimiter and len(self.sheetdelimiter):
                    of.write(self.sheetdelimiter + " " + str(s['id']) + " - " + sheetname + os.linesep)
                self._convert(s['id'], of)

    def _convert(self, sheetid, outfile):
        closefile = False
        if isinstance(outfile, str):
            outfile = open(outfile, 'w+')
            closefile = True
        try:
            writer = csv.writer(outfile, quoting=csv.QUOTE_MINIMAL, delimiter=self.delimiter, lineterminator=os.linesep)
            sheetfile = self._filehandle("xl/worksheets/sheet%i.xml" % sheetid)
            if not sheetfile:
                if self.cmd:
                    sys.stderr.write("Sheet %s not found!%s" %(sheetid, os.linesep))
                    sys.exit(1)
                raise SheetNotFoundException("Sheet %s not found" %sheetid)
            try:
                sheet = Sheet(self.workbook, self.shared_strings, self.styles, sheetfile)
                sheet.set_dateformat(self.dateformat)
                sheet.set_skip_empty_lines(self.skip_empty_lines)
                sheet.to_csv(writer)
            finally:
                sheetfile.close()
        finally:
            if closefile:
                outfile.close()

    def _filehandle(self, filename):
        for name in filter(lambda f: f.lower() == filename.lower(), self.ziphandle.namelist()):
            # python2.4 fix
            if not hasattr(self.ziphandle, "open"):
                return StringIO(self.ziphandle.read(name))
            return self.ziphandle.open(name, "r")
        return None

    def _parse(self, klass, filename):
        instance = klass()
        filehandle = self._filehandle(filename)
        if filehandle:
            instance.parse(filehandle)
            filehandle.close()
        return instance

class Workbook:
    def __init__(self):
        self.sheets = []
        self.date1904 = False

    def parse(self, filehandle):
        workbookDoc = minidom.parseString(filehandle.read())
        if len(workbookDoc.firstChild.getElementsByTagName("fileVersion")) == 0:
            self.appName = 'unknown'
        else:
            self.appName = workbookDoc.firstChild.getElementsByTagName("fileVersion")[0]._attrs['appName'].value
        try:
            self.date1904 = workbookDoc.firstChild.getElementsByTagName("workbookPr")[0]._attrs['date1904'].value.lower().strip() != "false"
        except:
            pass

        sheets = workbookDoc.firstChild.getElementsByTagName("sheets")[0]
        for sheetNode in sheets.getElementsByTagName("sheet"):
            attrs = sheetNode._attrs
            name = attrs["name"].value
            if self.appName == 'xl':
                if 'r:id' in attrs: id = int(attrs["r:id"].value[3:])
                else: id = int(attrs['sheetId'].value)
            else:
                if 'sheetId' in attrs: id = int(attrs["sheetId"].value)
                else: id = int(attrs['r:id'].value[3:])
            self.sheets.append({'name': name, 'id': id})

class Styles:
    def __init__(self):
        self.numFmts = {}
        self.cellXfs = []

    def parse(self, filehandle):
        styles = minidom.parseString(filehandle.read()).firstChild
        # numFmts
        numFmtsElement = styles.getElementsByTagName("numFmts")
        if len(numFmtsElement) == 1:
            for numFmt in numFmtsElement[0].childNodes:
                if numFmt.nodeType == minidom.Node.ELEMENT_NODE:
                    numFmtId = int(numFmt._attrs['numFmtId'].value)
                    formatCode = numFmt._attrs['formatCode'].value.lower().replace('\\', '')
                    self.numFmts[numFmtId] = formatCode
        # cellXfs
        cellXfsElement = styles.getElementsByTagName("cellXfs")
        if len(cellXfsElement) == 1:
            for cellXfs in cellXfsElement[0].childNodes:
                if cellXfs.nodeType != minidom.Node.ELEMENT_NODE or cellXfs.nodeName != "xf":
                    continue
                if 'numFmtId' in cellXfs._attrs:
                    numFmtId = int(cellXfs._attrs['numFmtId'].value)
                    self.cellXfs.append(numFmtId)
                else:
                    self.cellXfs.append(None)

class SharedStrings:
    def __init__(self):
        self.parser = None
        self.strings = []
        self.si = False
        self.t = False
        self.rPh = False
        self.value = ""

    def parse(self, filehandle):
        self.parser = xml.parsers.expat.ParserCreate()
        self.parser.CharacterDataHandler = self.handleCharData
        self.parser.StartElementHandler = self.handleStartElement
        self.parser.EndElementHandler = self.handleEndElement
        self.parser.ParseFile(filehandle)

    def escape_strings(self):
        for i in range(0, len(self.strings)):
            self.strings[i] = self.strings[i].replace("\r", "\\r").replace("\n", "\\n").replace("\t", "\\t")

    def handleCharData(self, data):
        if self.t:
            self.value+= data

    def handleStartElement(self, name, attrs):
        if name == 'si':
            self.si = True
            self.value = ""
        elif name == 't' and self.rPh:
            self.t = False
        elif name == 't' and self.si:
            self.t = True
        elif name == 'rPh':
            self.rPh = True

    def handleEndElement(self, name):
        if name == 'si':
            self.si = False
            self.strings.append(self.value)
        elif name == 't':
            self.t = False
        elif name == 'rPh':
            self.rPh = False

class Sheet:
    def __init__(self, workbook, sharedString, styles, filehandle):
        self.py3 = sys.version_info[0] == 3
        self.parser = None
        self.writer = None
        self.sharedString = None
        self.styles = None

        self.in_sheet = False
        self.in_row = False
        self.in_cell = False
        self.in_cell_value = False
        self.in_cell_formula = False

        self.columns = {}
        self.rowNum = None
        self.colType = None
        self.s_attr = None
        self.data = None

        self.dateformat = None
        self.skip_empty_lines = False

        self.filehandle = filehandle
        self.workbook = workbook
        self.sharedStrings = sharedString.strings
        self.styles = styles

    def set_dateformat(self, dateformat):
        self.dateformat = dateformat

    def set_skip_empty_lines(self, skip):
        self.skip_empty_lines = skip

    def to_csv(self, writer):
        self.writer = writer
        self.parser = xml.parsers.expat.ParserCreate()
        self.parser.CharacterDataHandler = self.handleCharData
        self.parser.StartElementHandler = self.handleStartElement
        self.parser.EndElementHandler = self.handleEndElement
        self.parser.ParseFile(self.filehandle)

    def handleCharData(self, data):
        if self.in_cell_value:
            self.collected_string+= data
            self.data = self.collected_string
            if self.colType == "s": # shared string
                self.data = self.sharedStrings[int(self.data)]
            elif self.colType == "b": # boolean
                self.data = (int(data) == 1 and "TRUE") or (int(data) == 0 and "FALSE") or data
            elif self.s_attr:
                s = int(self.s_attr)

                # get cell format
                format = None
                xfs_numfmt = self.styles.cellXfs[s]
                if xfs_numfmt in self.styles.numFmts:
                    format = self.styles.numFmts[xfs_numfmt]
                elif xfs_numfmt in STANDARD_FORMATS:
                    format = STANDARD_FORMATS[xfs_numfmt]
                # get format type
                if format and format in FORMATS:
                    format_type = FORMATS[format]
                    try:
                        if format_type == 'date': # date/time
                            if self.workbook.date1904:
                                date = datetime.datetime(1904, 1, 1) + datetime.timedelta(float(self.data))
                            else:
                                date = datetime.datetime(1899, 12, 30) + datetime.timedelta(float(self.data))
                            if self.dateformat:
                                # str(dateformat) - python2.5 bug, see: http://bugs.python.org/issue2782
                                self.data = date.strftime(str(self.dateformat))
                            else:
                                dateformat = format.replace("yyyy", "%Y").replace("yy", "%y"). \
                                  replace("hh:mm", "%H:%M").replace("h", "%H").replace("%H%H", "%H").replace("ss", "%S"). \
                                  replace("d", "%e").replace("%e%e", "%d"). \
                                  replace("mmmm", "%B").replace("mmm", "%b").replace(":mm", ":%M").replace("m", "%m").replace("%m%m", "%m"). \
                                  replace("am/pm", "%p")
                                self.data = date.strftime(str(dateformat)).strip()
                        elif format_type == 'time': # time
                            self.data = str(float(self.data) * 24*60*60)
                        elif format_type == 'float' and ('E' in self.data or 'e' in self.data):
                            self.data = ("%f" %(float(self.data))).rstrip('0').rstrip('.')
                    except (ValueError, OverflowError):
                        # invalid date format
                        pass
        # does not support it
        #elif self.in_cell_formula:
        #    self.formula = data

    def handleStartElement(self, name, attrs):
        if self.in_row and name == 'c':
            self.colType = attrs.get("t")
            self.s_attr = attrs.get("s")
            cellId = attrs.get("r")
            if cellId:
                self.colNum = cellId[:len(cellId)-len(self.rowNum)]
                self.colIndex = 0
            else:
                self.colIndex+= 1
            #self.formula = None
            self.data = ""
            self.in_cell = True
        elif self.in_cell and (name == 'v' or name == 'is'):
            self.in_cell_value = True
            self.collected_string = ""
        #elif self.in_cell and name == 'f':
        #    self.in_cell_formula = True
        elif self.in_sheet and name == 'row' and 'r' in attrs:
            self.rowNum = attrs['r']
            self.in_row = True
            self.columns = {}
            self.spans = None
            if 'spans' in attrs:
                self.spans = [int(i) for i in attrs['spans'].split(":")]
        elif name == 'sheetData':
            self.in_sheet = True

    def handleEndElement(self, name):
        if self.in_cell and name == 'v':
            self.in_cell_value = False
        #elif self.in_cell and name == 'f':
        #    self.in_cell_formula = False
        elif self.in_cell and name == 'c':
            t = 0
            for i in self.colNum: t = t*26 + ord(i) - 64
            self.columns[t - 1 + self.colIndex] = self.data
            self.in_cell = False
        if self.in_row and name == 'row':
            if len(self.columns.keys()) > 0:
                d = [""] * (max(self.columns.keys()) + 1)
                for k in self.columns.keys():
                    val = self.columns[k]
                    if not self.py3:
                        val = val.encode("utf-8")
                    d[k] = val
                if self.spans:
                    l = self.spans[0] + self.spans[1] - 1
                    if len(d) < l:
                        d+= (l - len(d)) * ['']
                # write line to csv
                if not self.skip_empty_lines or d.count('') != len(d):
                    self.writer.writerow(d)
            self.in_row = False
        elif self.in_sheet and name == 'sheetData':
            self.in_sheet = False

def convert_recursive(path, sheetid, kwargs):
    kwargs['cmd'] = False
    for name in os.listdir(path):
        fullpath = os.path.join(path, name)
        if os.path.isdir(fullpath):
            convert_recursive(fullpath, kwargs)
        else:
            if fullpath.lower().endswith(".xlsx"):
                outfilepath = fullpath[:-4] + 'csv'
                print("Converting %s to %s" %(fullpath, outfilepath))
                try:
                    Xlsx2csv(fullpath, **kwargs).convert(outfilepath, sheetid)
                except zipfile.BadZipfile:
                    print("File %s is not a zip file" %fullpath)

if __name__ == "__main__":
    if "ArgumentParser" in globals():
        parser = ArgumentParser(description = "xlsx to csv converter")
        parser.add_argument('infile', metavar='xlsxfile', help="xlsx file path")
        parser.add_argument('outfile', metavar='outfile', nargs='?', help="output csv file path")
        parser.add_argument('-v', '--version', action='version', version='%(prog)s')
        argparser = True
    else:
        parser = OptionParser(usage = "%prog [options] infile [outfile]", version=__version__)
        parser.add_argument = parser.add_option
        argparser = False

    parser.add_argument("-a", "--all", dest="all", default=False, action="store_true",
      help="export all sheets")
    parser.add_argument("-d", "--delimiter", dest="delimiter", default=",",
      help="delimiter - columns delimiter in csv, 'tab' or 'x09' for a tab (default: comma ',')")
    parser.add_argument("-f", "--dateformat", dest="dateformat",
      help="override date/time format (ex. %%Y/%%m/%%d)")
    parser.add_argument("-i", "--ignoreempty", dest="skip_empty_lines", default=False, action="store_true",
      help="skip empty lines")
    parser.add_argument("-e", "--escape", dest='escape_strings', default=False, action="store_true",
      help="Escape \\r\\n\\t characters")
    parser.add_argument("-p", "--sheetdelimiter", dest="sheetdelimiter", default="--------",
      help="sheet delimiter used to separate sheets, pass '' if you do not need delimiter (default: '--------')")
    parser.add_argument("-s", "--sheet", dest="sheetid", default=1, type=int,
      help="sheet number to convert")

    if argparser:
        options = parser.parse_args()
    else:
        (options, args) = parser.parse_args()
        if len(args) < 1:
            parser.print_usage()
            sys.stderr.write("error: too few arguments" + os.linesep)
            sys.exit(1)
        options.infile = args[0]
        options.outfile = len(args) > 1 and args[1] or None

    if len(options.delimiter) == 1:
        delimiter = options.delimiter
    elif options.delimiter == 'tab':
        delimiter = '\t'
    elif options.delimiter == 'comma':
        delimiter = ','
    elif options.delimiter[0] == 'x':
        delimiter = chr(int(options.delimiter[1:]))
    else:
        raise XlsxException("Invalid delimiter")

    kwargs = {
      'delimiter' : delimiter,
      'sheetdelimiter' : options.sheetdelimiter,
      'dateformat' : options.dateformat,
      'skip_empty_lines' : options.skip_empty_lines,
      'escape_strings' : options.escape_strings,
      'cmd' : True
    }
    sheetid = options.sheetid
    if options.all:
        sheetid = 0

    if os.path.isdir(options.infile):
        convert_recursive(options.infile, sheetid, kwargs)
    else:
        xlsx2csv = Xlsx2csv(options.infile, **kwargs)
        outfile = options.outfile or sys.stdout
        xlsx2csv.convert(outfile, sheetid)
