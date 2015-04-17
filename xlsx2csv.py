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
__version__ = "0.7.2"

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
     Input:
       xlsxfile - path to file or filehandle
     options:
       sheetid - sheet no to convert (0 for all sheets)
       dateformat - override date/time format
       delimiter - csv columns delimiter symbol
       sheetdelimiter - sheets delimiter used when processing all sheets
       skip_empty_lines - skip empty lines
       hyperlinks - include hyperlinks
       include_sheet_pattern - only include sheets named matching given pattern
       exclude_sheet_pattern - exclude sheets named matching given pattern
    """

    def __init__(self, xlsxfile, **options):
        # dateformat=None, delimiter=",", sheetdelimiter="--------", skip_empty_lines=False, escape_strings=False, cmd=False
        options.setdefault("delimiter", ",")
        options.setdefault("sheetdelimiter", "--------")
        options.setdefault("dateformat", None)
        options.setdefault("skip_empty_lines", False)
        options.setdefault("escape_strings", False)
        options.setdefault("hyperlinks", False)
        options.setdefault("cmd", False)
        options.setdefault("include_sheet_pattern", ["^.*$"])
        options.setdefault("exclude_sheet_pattern", [])
        options.setdefault("merge_cells", False)

        self.options = options
        try:
            self.ziphandle = zipfile.ZipFile(xlsxfile)
        except (zipfile.BadZipfile, IOError):
            if self.options['cmd']:
                sys.stderr.write("Invalid xlsx file: " + str(xlsxfile) + os.linesep)
                sys.exit(1)
            raise InvalidXlsxFileException("Invalid xlsx file: " + str(xlsxfile))

        self.py3 = sys.version_info[0] == 3

        self.shared_strings = self._parse(SharedStrings, "xl/sharedStrings.xml")
        self.styles = self._parse(Styles, "xl/styles.xml")
        self.workbook = self._parse(Workbook, "xl/workbook.xml")
        self.workbook.relationships = self._parse(Relationships, "xl/_rels/workbook.xml.rels")
        if self.options['escape_strings']:
            self.shared_strings.escape_strings()

    def getSheetIdByName(self, name):
        for s in self.workbook.sheets:
            if s['name'] == name:
                return s['id']
        return None

    def convert(self, outfile, sheetid=1):
        """outfile - path to file or filehandle"""
        if sheetid > 0:
            self._convert(sheetid, outfile)
        else:
            if isinstance(outfile, str):
                if not os.path.exists(outfile):
                    os.makedirs(outfile)
                elif os.path.isfile(outfile):
                    if self.options['cmd']:
                        sys.stderr.write("File " + str(outfile) + " already exists!" + os.linesep)
                        sys.exit(1)
                    raise OutFileAlreadyExistsException("File " + str(outfile) + " already exists!")
            for s in self.workbook.sheets:
                sheetname = s['name']

                # filter sheets by include pattern
                include_sheet_pattern = self.options['include_sheet_pattern']
                if type(include_sheet_pattern) == type(""): # optparser lib fix
                    include_sheet_pattern = [include_sheet_pattern]
                if len(include_sheet_pattern) > 0:
                    include = False
                    for pattern in include_sheet_pattern:
                        include = pattern and len(pattern) > 0 and re.match(pattern, sheetname)
                        if include:
                            break
                    if not include:
                        continue

                # filter sheets by exclude pattern
                exclude_sheet_pattern = self.options['exclude_sheet_pattern']
                if type(exclude_sheet_pattern) == type(""): # optparser lib fix
                    exclude_sheet_pattern = [exclude_sheet_pattern]
                exclude = False
                for pattern in exclude_sheet_pattern:
                    exclude = pattern and len(pattern) > 0 and re.match(pattern, sheetname)
                    if exclude:
                        break
                if exclude:
                    continue

                if not self.py3:
                    sheetname = sheetname.encode('utf-8')
                of = outfile
                if isinstance(outfile, str):
                    of = os.path.join(outfile, sheetname + '.csv')
                elif self.options['sheetdelimiter'] and len(self.options['sheetdelimiter']):
                    of.write(self.options['sheetdelimiter'] + " " + str(s['id']) + " - " + sheetname + os.linesep)
                self._convert(s['id'], of)

    def _convert(self, sheetid, outfile):
        closefile = False
        if isinstance(outfile, str):
            outfile = open(outfile, 'w+')
            closefile = True
        try:
            writer = csv.writer(outfile, quoting=csv.QUOTE_MINIMAL, delimiter=self.options['delimiter'], lineterminator=os.linesep)
            sheetfile = self._filehandle("xl/worksheets/sheet%i.xml" % sheetid)
            if not sheetfile and sheetid == 1:
                sheetfile = self._filehandle("xl/worksheets/sheet.xml")
            if not sheetfile:
                if self.options['cmd']:
                    sys.stderr.write("Sheet %s not found!%s" %(sheetid, os.linesep))
                    sys.exit(1)
                raise SheetNotFoundException("Sheet %s not found" %sheetid)
            try:
                sheet = Sheet(self.workbook, self.shared_strings, self.styles, sheetfile)
                sheet.relationships = self._parse(Relationships, "xl/worksheets/_rels/sheet%i.xml.rels" % sheetid)
                sheet.set_dateformat(self.options['dateformat'])
                sheet.set_skip_empty_lines(self.options['skip_empty_lines'])
                sheet.set_include_hyperlinks(self.options['hyperlinks'])
                sheet.set_merge_cells(self.options['merge_cells'])
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
        if workbookDoc.firstChild.namespaceURI:
            fileVersion = workbookDoc.firstChild.getElementsByTagNameNS(workbookDoc.firstChild.namespaceURI, "fileVersion")
        else:
            fileVersion = workbookDoc.firstChild.getElementsByTagName("fileVersion")
        if len(fileVersion) == 0:
            self.appName = 'unknown'
        else:
            try:
                if workbookDoc.firstChild.namespaceURI:
                    self.appName = workbookDoc.firstChild.getElementsByTagNameNS(workbookDoc.firstChild.namespaceURI, "fileVersion")[0]._attrs['appName'].value
                else:
                    self.appName = workbookDoc.firstChild.getElementsByTagName("fileVersion")[0]._attrs['appName'].value
            except KeyError:
                # no app name
                self.appName = 'unknown'
        try:
            if workbookDoc.firstChild.namespaceURI:
                self.date1904 = workbookDoc.firstChild.getElementsByTagNameNS(workbookDoc.firstChild.namespaceURI, "workbookPr")[0]._attrs['date1904'].value.lower().strip() != "false"
            else:
                self.date1904 = workbookDoc.firstChild.getElementsByTagName("workbookPr")[0]._attrs['date1904'].value.lower().strip() != "false"
        except:
            pass

        if workbookDoc.firstChild.namespaceURI:
            sheets = workbookDoc.firstChild.getElementsByTagNameNS(workbookDoc.firstChild.namespaceURI, "sheets")[0]
        else:
            sheets = workbookDoc.firstChild.getElementsByTagName("sheets")[0]
        if workbookDoc.firstChild.namespaceURI:
            sheetNodes = sheets.getElementsByTagNameNS(workbookDoc.firstChild.namespaceURI, "sheet")
        else:
            sheetNodes = sheets.getElementsByTagName("sheet")
        for sheetNode in sheetNodes:
            attrs = sheetNode._attrs
            name = attrs["name"].value
            if self.appName == 'xl' and len(attrs["r:id"].value) > 2:
                if 'r:id' in attrs: id = int(attrs["r:id"].value[3:])
                else: id = int(attrs['sheetId'].value)
            else:
                if 'sheetId' in attrs: id = int(attrs["sheetId"].value)
                else: id = int(attrs['r:id'].value[3:])
            self.sheets.append({'name': name, 'id': id})

class Relationships:
    def __init__(self):
        self.relationships = {}

    def parse(self, filehandle):
        doc = minidom.parseString(filehandle.read())
        if doc.namespaceURI:
            relationships = doc.getElementsByTagNameNS(doc.namespaceURI, "Relationships")
        else:
            relationships = doc.getElementsByTagName("Relationships")
        if not relationships:
            return
        if doc.namespaceURI:
            relationshipNodes = relationships[0].getElementsByTagNameNS(doc.namespaceURI, "Relationship")
        else:
            relationshipNodes = relationships[0].getElementsByTagName("Relationship")
        for rel in relationshipNodes:
            attrs = rel._attrs
            rId = attrs.get('Id')
            if rId:
                vtype = attrs.get('Type')
                target = attrs.get('Target')
                self.relationships[str(rId.value)] = {
                    "type" : vtype and str(vtype.value) or None,
                    "target" : target and target.value.encode("utf-8") or None
                }

class Styles:
    def __init__(self):
        self.numFmts = {}
        self.cellXfs = []

    def parse(self, filehandle):
        styles = minidom.parseString(filehandle.read()).firstChild
        # numFmts
        if styles.namespaceURI:
            numFmtsElement = styles.getElementsByTagNameNS(styles.namespaceURI, "numFmts")
        else:
            numFmtsElement = styles.getElementsByTagName("numFmts")
        if len(numFmtsElement) == 1:
            for numFmt in numFmtsElement[0].childNodes:
                if numFmt.nodeType == minidom.Node.ELEMENT_NODE:
                    numFmtId = int(numFmt._attrs['numFmtId'].value)
                    formatCode = numFmt._attrs['formatCode'].value.lower().replace('\\', '')
                    self.numFmts[numFmtId] = formatCode
        # cellXfs
        if styles.namespaceURI:
            cellXfsElement = styles.getElementsByTagNameNS(styles.namespaceURI, "cellXfs")
        else:
            cellXfsElement = styles.getElementsByTagName("cellXfs")
        if len(cellXfsElement) == 1:
            for cellXfs in cellXfsElement[0].childNodes:
                if cellXfs.nodeType != minidom.Node.ELEMENT_NODE or not (cellXfs.nodeName == "xf" or cellXfs.nodeName.endswith(":xf")):
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
        # ignore namespace
        i = name.find(":")
        if i >= 0:
            name = name[i + 1:]

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
        # ignore namespace
        i = name.find(":")
        if i >= 0:
            name = name[i + 1:]

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
        self.relationships = None
        self.columns_count = -1

        self.in_sheet = False
        self.in_row = False
        self.in_cell = False
        self.in_cell_value = False
        self.in_cell_formula = False

        self.columns = {}
        self.rowNum = None
        self.colType = None
        self.cellId = None
        self.s_attr = None
        self.data = None

        self.dateformat = None
        self.skip_empty_lines = False

        self.filedata = None
        self.filehandle = filehandle
        self.workbook = workbook
        self.sharedStrings = sharedString.strings
        self.styles = styles

        self.hyperlinks = {}
        self.mergeCells = {}

    def set_dateformat(self, dateformat):
        self.dateformat = dateformat

    def set_skip_empty_lines(self, skip):
        self.skip_empty_lines = skip

    def set_merge_cells(self, mergecells):
        if not mergecells:
            return
        if not self.filedata:
            self.filedata = self.filehandle.read()
        data = str(self.filedata) # python3: convert byte buffer to string

        # find worksheet tag, we need namespaces from it
        start = data.find("<worksheet")
        if start < 0:
            return
        end = data.find(">", start)
        worksheet = data[start : end + 1]

        # find hyperlinks part
        start = data.find("<mergeCells")
        if start < 0:
            # hyperlinks not found
            return
        end = data.find("</mergeCells>")
        data = data[start : end + 13]

        # parse hyperlinks
        doc = minidom.parseString(worksheet + data + "</worksheet>").firstChild

        if doc.namespaceURI:
            mergeCells = doc.getElementsByTagNameNS(doc.namespaceURI, "mergeCell")
        else:
            mergeCells = doc.getElementsByTagName("mergeCell")
        for mergeCell in mergeCells:
            attrs = mergeCell._attrs
            if 'ref' in attrs.keys():
                rangeStr = attrs['ref'].value
                rng = rangeStr.split(":")
                if len(rng) > 1:
                    for cell in self._range(rangeStr):
                        self.mergeCells[cell] = {}
                        self.mergeCells[cell]['copyFrom'] = rng[0]

    def set_include_hyperlinks(self, hyperlinks):
        if not hyperlinks or not self.relationships or not self.relationships.relationships:
            return
        # we must read file first to get hyperlinks, but we don't wont to parse whole file
        if not self.filedata:
            self.filedata = self.filehandle.read()
        data = str(self.filedata) # python3: convert byte buffer to string

        # find worksheet tag, we need namespaces from it
        start = data.find("<worksheet")
        if start < 0:
            return
        end = data.find(">", start)
        worksheet = data[start : end + 1]

        # find hyperlinks part
        start = data.find("<hyperlinks>")
        if start < 0:
            # hyperlinks not found
            return
        end = data.find("</hyperlinks>")
        data = data[start : end + 13]

        # parse hyperlinks
        doc = minidom.parseString(worksheet + data + "</worksheet>").firstChild
        if doc.namespaceURI:
            hiperlinkNodes = doc.getElementsByTagNameNS(doc.namespaceURI, "hyperlink")
        else:
            hiperlinkNodes = doc.getElementsByTagName("hyperlink")
        for hlink in hiperlinkNodes:
            attrs = hlink._attrs
            ref = rId = None
            for k in attrs.keys():
                if k == "ref":
                    ref = str(attrs[k].value)
                if k.endswith(":id"):
                    rId = str(attrs[k].value)
            if not ref or not rId:
                continue
            rel = self.relationships.relationships.get(rId)
            if not rel:
                continue
            target = rel.get('target')
            for cell in self._range(ref):
                self.hyperlinks[cell] = target

    def to_csv(self, writer):
        self.writer = writer
        self.parser = xml.parsers.expat.ParserCreate()
        self.parser.CharacterDataHandler = self.handleCharData
        self.parser.StartElementHandler = self.handleStartElement
        self.parser.EndElementHandler = self.handleEndElement
        if self.filedata:
            self.parser.Parse(self.filedata)
        else:
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
                if not format:
                    return
                format_type = None
                if format in FORMATS:
                    format_type = FORMATS[format]
                elif re.match("^\d+(\.\d+)?$", self.data) and re.match(".*[hsmdyY]", format) and not re.match('.*\[.*[dmhys].*\]', format):
                    # it must be date format
                    if float(self.data) < 1:
                        format_type = "time"
                    else:
                        format_type = "date"
                elif re.match("^\d", self.data):
                    format_type = "float"
                if format_type:
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
                                # ignore ";@", don't know what does it mean right now
                                dateformat = format.replace(";@", ""). \
                                  replace("yyyy", "%Y").replace("yy", "%y"). \
                                  replace("hh:mm", "%H:%M").replace("h", "%H").replace("%H%H", "%H").replace("ss", "%S"). \
                                  replace("d", "%e").replace("%e%e", "%d"). \
                                  replace("mmmm", "%B").replace("mmm", "%b").replace(":mm", ":%M").replace("m", "%m").replace("%m%m", "%m"). \
                                  replace("am/pm", "%p")
                                self.data = date.strftime(str(dateformat)).strip()
                        elif format_type == 'time': # time
                            t = int(float(self.data)%1 * 24*60)
                            self.data = "%.2i:%.2i" %(t / 60, t % 60)  #str(t / 60) + ":" + ('0' + str(t % 60))[-2:]
                        elif format_type == 'float' and ('E' in self.data or 'e' in self.data):
                            self.data = ("%f" %(float(self.data))).rstrip('0').rstrip('.')
                        elif format_type == 'float' and format[0:3] == '0.0':
                            self.data = ("%." + str(len(format.split(".")[1])+(1 if ('%' in format) else 0)) + "f") % float(self.data)

                    except (ValueError, OverflowError):
                        # invalid date format
                        pass
        # does not support it
        #elif self.in_cell_formula:
        #    self.formula = data

    def handleStartElement(self, name, attrs):
        has_namespace = name.find(":") > 0
        if self.in_row and (name == 'c' or (has_namespace and name.endswith(':c'))):
            self.colType = attrs.get("t")
            self.s_attr = attrs.get("s")
            self.cellId = attrs.get("r")
            if self.cellId:
                self.colNum = self.cellId[:len(self.cellId)-len(self.rowNum)]
                self.colIndex = 0
            else:
                self.colIndex+= 1
            #self.formula = None
            self.data = ""
            self.in_cell = True
        elif self.in_cell and ((name == 'v' or name == 'is') or (has_namespace and (name.endswith(':v') or name.endswith(':is')))):
            self.in_cell_value = True
            self.collected_string = ""
        #elif self.in_cell and name == 'f':
        #    self.in_cell_formula = True
        elif self.in_sheet and (name == 'row' or (has_namespace and name.endswith(':row'))) and ('r' in attrs):
            self.rowNum = attrs['r']
            self.in_row = True
            self.columns = {}
            self.spans = None
            if 'spans' in attrs:
                self.spans = [int(i) for i in attrs['spans'].split(":")]

        elif name == 'sheetData' or (has_namespace and name.endswith(':sheetData')):
            self.in_sheet = True
        elif name == 'dimension':
            rng = attrs.get("ref").split(":")
            if len(rng) > 1:
                start = re.match("^([A-Z]+)(\d+)$", rng[0])
                if (start):
                    end = re.match("^([A-Z]+)(\d+)$", rng[1])
                    startCol = start.group(1)
                    #startRow = int(start.group(2))
                    endCol = end.group(1)
                    #endRow = int(end.group(2))
                    self.columns_count = 0
                    for cell in self._range(startCol + "1:" + endCol + "1"):
                        self.columns_count+= 1

    def handleEndElement(self, name):
        has_namespace = name.find(":") > 0
        if self.in_cell and name == 'v':
            self.in_cell_value = False
        #elif self.in_cell and name == 'f':
        #    self.in_cell_formula = False
        elif self.in_cell and (name == 'c' or (has_namespace and name.endswith(':c'))):
            t = 0
            for i in self.colNum: t = t*26 + ord(i) - 64
            d = self.data
            if self.hyperlinks:
                hyperlink = self.hyperlinks.get(self.cellId)
                if hyperlink:
                    d = "<a href='" + hyperlink + "'>" + d + "</a>"
            if self.colNum + self.rowNum in self.mergeCells.keys():
                if 'copyFrom' in self.mergeCells[self.colNum + self.rowNum].keys() and self.mergeCells[self.colNum + self.rowNum]['copyFrom'] == self.colNum + self.rowNum:
                    self.mergeCells[self.colNum + self.rowNum]['value'] = d
                else:
                    d = self.mergeCells[self.mergeCells[self.colNum + self.rowNum]['copyFrom']]['value']

            self.columns[t - 1 + self.colIndex] = d

        if self.in_row and (name == 'row' or (has_namespace and name.endswith(':row'))):
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
                    while len(d) < self.columns_count:
                        d.append("")
                    self.writer.writerow(d)
            self.in_row = False
        elif self.in_sheet and (name == 'sheetData' or (has_namespace and name.endswith(':sheetData'))):
            self.in_sheet = False

    # rangeStr: "A3:C12" or "D5"
    # example: for cell in _range("A1:Z12"): print cell
    def _range(self, rangeStr):
        rng = rangeStr.split(":")
        if len(rng) == 1:
            yield rangeStr
        else:
            start = re.match("^([A-Z]+)(\d+)$", rng[0])
            end = re.match("^([A-Z]+)(\d+)$", rng[1])
            if not start or not end:
                return
            startCol = start.group(1)
            startRow = int(start.group(2))
            endCol = end.group(1)
            endRow = int(end.group(2))
            col = startCol
            while True:
                for row in range(startRow, endRow + 1):
                    yield col + str(row)
                if col == endCol:
                    break
                t = 0
                for i in col: t = t * 26 + ord(i) - 64
                col = ""
                while t >= 0:
                  col = chr(t % 26 + 65) + col
                  t = t // 26 - 1


def convert_recursive(path, sheetid, outfile, kwargs):
    kwargs['cmd'] = False
    for name in os.listdir(path):
        fullpath = os.path.join(path, name)
        if os.path.isdir(fullpath):
            convert_recursive(fullpath, sheetid, outfile, kwargs)
        else:
            # strange code, python2.4 fix
            #outfilepath = outfile if len(outfile) > 0 else ""
            outfilepath = outfile

            if len(outfilepath) == 0 and fullpath.lower().endswith(".xlsx"):
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
        nargs_plus = "+"
        argparser = True
    else:
        parser = OptionParser(usage = "%prog [options] infile [outfile]", version=__version__)
        parser.add_argument = parser.add_option
        nargs_plus = 1
        argparser = False


    if sys.version_info[0] == 2 and sys.version_info[1] < 5:
        inttype = "int"
    else:
        inttype = int
    parser.add_argument("-a", "--all", dest="all", default=False, action="store_true",
      help="export all sheets")
    parser.add_argument("-s", "--sheet", dest="sheetid", default=1, type=inttype,
      help="sheet number to convert")
    parser.add_argument("-n", "--sheetname", dest="sheetname", default=None,
      help="sheet name to convert")
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
    parser.add_argument("--hyperlinks", "--hyperlinks", dest="hyperlinks", action="store_true", default=False,
      help="include hyperlinks")
    parser.add_argument("-I", "--include_sheet_pattern", nargs=nargs_plus, dest="include_sheet_pattern", default="^.*$",
      help="only include sheets named matching given pattern, only effects when -a option is enabled.")
    parser.add_argument("-E", "--exclude_sheet_pattern", nargs=nargs_plus, dest="exclude_sheet_pattern", default="",
      help="exclude sheets named matching given pattern, only effects when -a option is enabled.")
    parser.add_argument("-m", "--merge-cells", dest="merge_cells", default=False, action="store_true",
      help="merge cells")

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
      'hyperlinks' : options.hyperlinks,
      'cmd' : True,
      'include_sheet_pattern' : options.include_sheet_pattern,
      'exclude_sheet_pattern' : options.exclude_sheet_pattern,
      'merge_cells' : options.merge_cells
    }
    sheetid = options.sheetid
    if options.all:
        sheetid = 0

    outfile = options.outfile or sys.stdout
    if os.path.isdir(options.infile):
        convert_recursive(options.infile, sheetid, outfile, kwargs)
    else:
        xlsx2csv = Xlsx2csv(options.infile, **kwargs)
        if options.sheetname:
            sheetid = xlsx2csv.getSheetIdByName(options.sheetname)
            if not sheetid:
                raise XlsxException("Sheet '%s' not found" % options.sheetname)

        xlsx2csv.convert(outfile, sheetid)
