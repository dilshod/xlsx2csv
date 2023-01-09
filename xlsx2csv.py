#!/usr/bin/env python
#
# The MIT License
#
# Copyright (c) 2022 Dilshod Temirkhodjaev
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

__author__ = "Dilshod Temirkhodjaev <tdilshod@gmail.com>"
__license__ = "MIT"
__version__ = "0.8.1"

import csv, datetime, zipfile, sys, os, re, signal
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
    'general': 'float',
    '0': 'float',
    '0.00': 'float',
    '#,##0': 'float',
    '#,##0.00': 'float',
    '0%': 'percentage',
    '0.00%': 'percentage',
    '0.00e+00': 'float',
    'mm-dd-yy': 'date',
    'd-mmm-yy': 'date',
    'd-mmm': 'date',
    'mmm-yy': 'date',
    'h:mm am/pm': 'date',
    'h:mm:ss am/pm': 'date',
    'h:mm': 'time',
    'h:mm:ss': 'time',
    'm/d/yy h:mm': 'date',
    '#,##0 ;(#,##0)': 'float',
    '#,##0 ;[red](#,##0)': 'float',
    '#,##0.00;(#,##0.00)': 'float',
    '#,##0.00;[red](#,##0.00)': 'float',
    'mm:ss': 'time',
    '[h]:mm:ss': 'time',
    'mmss.0': 'time',
    '##0.0e+0': 'float',
    '@': 'float',
    'yyyy\\-mm\\-dd': 'date',
    'dd/mm/yy': 'date',
    'hh:mm:ss': 'time',
    "dd/mm/yy\\ hh:mm": 'date',
    'dd/mm/yyyy hh:mm:ss': 'date',
    'yy-mm-dd': 'date',
    'd-mmm-yyyy': 'date',
    'm/d/yy': 'date',
    'm/d/yyyy': 'date',
    'dd-mmm-yyyy': 'date',
    'dd/mm/yyyy': 'date',
    'mm/dd/yy h:mm am/pm': 'date',
    'mm/dd/yy hh:mm': 'date',
    'mm/dd/yyyy h:mm am/pm': 'date',
    'mm/dd/yyyy hh:mm:ss': 'date',
    'yyyy-mm-dd hh:mm:ss': 'date',
    '#,##0;(#,##0)': 'float',
    '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)': 'float',
    '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)': 'float'
}
STANDARD_FORMATS = {
    0: 'general',
    1: '0',
    2: '0.00',
    3: '#,##0',
    4: '#,##0.00',
    9: '0%',
    10: '0.00%',
    11: '0.00e+00',
    12: '# ?/?',
    13: '# ??/??',
    14: 'mm-dd-yy',
    15: 'd-mmm-yy',
    16: 'd-mmm',
    17: 'mmm-yy',
    18: 'h:mm am/pm',
    19: 'h:mm:ss am/pm',
    20: 'h:mm',
    21: 'h:mm:ss',
    22: 'm/d/yy h:mm',
    37: '#,##0 ;(#,##0)',
    38: '#,##0 ;[red](#,##0)',
    39: '#,##0.00;(#,##0.00)',
    40: '#,##0.00;[red](#,##0.00)',
    45: 'mm:ss',
    46: '[h]:mm:ss',
    47: 'mmss.0',
    48: '##0.0e+0',
    49: '@',
}
CONTENT_TYPES = {
    'shared_strings',
    'styles',
    'workbook',
    'worksheet',
    'relationships',
}

DEFAULT_APP_PATH = "/xl"
DEFAULT_WORKBOOK_PATH = DEFAULT_APP_PATH + "/workbook.xml"

class XlsxException(Exception):
    pass


class InvalidXlsxFileException(XlsxException):
    pass


class SheetNotFoundException(XlsxException):
    pass


class OutFileAlreadyExistsException(XlsxException):
    pass


class XlsxValueError(XlsxException):
    pass


class Xlsx2csv:
    """
     Usage: Xlsx2csv("test.xslx", **params).convert("test.csv", sheetid=1)
     Input:
       xlsxfile - path to file or filehandle
     options:
       sheetid - sheet no to convert (0 for all sheets)
       sheetname - sheet name to convert
       dateformat - override date/time format
       timeformat - override time format
       floatformat - override float format
       quoting - if and how to quote
       delimiter - csv columns delimiter symbol
       sheetdelimiter - sheets delimiter used when processing all sheets
       skip_empty_lines - skip empty lines
       skip_trailing_columns - skip trailing columns
       hyperlinks - include hyperlinks
       include_sheet_pattern - only include sheets named matching given pattern
       exclude_sheet_pattern - exclude sheets named matching given pattern
       exclude_hidden_sheets - exclude hidden sheets
       skip_hidden_rows - skip hidden rows
    """

    def __init__(self, xlsxfile, **options):
        options.setdefault("delimiter", ",")
        options.setdefault("quoting", csv.QUOTE_MINIMAL)
        options.setdefault("sheetdelimiter", "--------")
        options.setdefault("dateformat", None)
        options.setdefault("timeformat", None)
        options.setdefault("floatformat", None)
        options.setdefault("scifloat", False)
        options.setdefault("skip_empty_lines", False)
        options.setdefault("skip_trailing_columns", False)
        options.setdefault("escape_strings", False)
        options.setdefault("no_line_breaks", False)
        options.setdefault("hyperlinks", False)
        options.setdefault("include_sheet_pattern", ["^.*$"])
        options.setdefault("exclude_sheet_pattern", [])
        options.setdefault("exclude_hidden_sheets", False)
        options.setdefault("merge_cells", False)
        options.setdefault("ignore_formats", [''])
        options.setdefault("lineterminator", "\n")
        options.setdefault("outputencoding", "utf-8")
        options.setdefault("skip_hidden_rows", True)

        self.options = options
        try:
            self.ziphandle = zipfile.ZipFile(xlsxfile)
        except (zipfile.BadZipfile, IOError):
            raise InvalidXlsxFileException("Invalid xlsx file: " + str(xlsxfile))

        self.py3 = sys.version_info[0] == 3

        self.content_types = self._parse(ContentTypes, "/[Content_Types].xml")
        self.shared_strings = self._parse(SharedStrings, self.content_types.types["shared_strings"])
        self.styles = self._parse(Styles, self.content_types.types["styles"])
        self.workbook = self._parse(Workbook, self.content_types.types["workbook"])
        workbook_relationships = list(filter(lambda r: "book" in r, self.content_types.types["relationships"]))[0]
        self.workbook.relationships = self._parse(Relationships, workbook_relationships)
        if self.options['no_line_breaks']:
            self.shared_strings.replace_line_breaks()
        elif self.options['escape_strings']:
            self.shared_strings.escape_strings()

    def __del__(self):
        # make sure to close zip file, ziphandler does have a close() method
        self.ziphandle.close()

    def getSheetIdByName(self, name):
        for s in self.workbook.sheets:
            if s['name'] == name:
                return s['index']
        return None

    def convert(self, outfile, sheetid=1, sheetname=None):
        """outfile - path to file or filehandle"""
        if sheetname:
            sheetid = self.getSheetIdByName(sheetname)
            if not sheetid:
                raise XlsxException("Sheet '%s' not found" % sheetname)
        if sheetid > 0:
            self._convert(sheetid, outfile)
        else:
            if isinstance(outfile, str):
                if not os.path.exists(outfile):
                    os.makedirs(outfile)
                elif os.path.isfile(outfile):
                    raise OutFileAlreadyExistsException("File " + str(outfile) + " already exists!")
            for s in self.workbook.sheets:
                sheetname = s['name']
                sheetstate = s['state']

                # filter hidden sheets
                if sheetstate in ('hidden', 'veryHidden') and self.options['exclude_hidden_sheets']:
                    continue

                # filter sheets by include pattern
                include_sheet_pattern = self.options['include_sheet_pattern']
                if type(include_sheet_pattern) == type(""):  # optparser lib fix
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
                if type(exclude_sheet_pattern) == type(""):  # optparser lib fix
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
                    of.write(self.options['sheetdelimiter'] + " " + str(s['index']) + " - " + sheetname + self.options['lineterminator'])
                self._convert(s['index'], of)

    def _convert(self, sheet_index, outfile):
        closefile = False
        if isinstance(outfile, str):
            if sys.version_info[0] == 2:
                outfile = open(outfile, 'wb+')
            elif sys.version_info[0] == 3:
                outfile = open(outfile, 'w+', encoding=self.options['outputencoding'], newline="")
            else:
                raise XlsxException("error: version of your python is not supported: " + str(sys.version_info) + "\n")
            closefile = True
        try:
            writer = csv.writer(outfile, quoting=self.options['quoting'], delimiter=self.options['delimiter'],
                                lineterminator=self.options['lineterminator'])

            sheets_filtered = list(filter(lambda s: s['index'] == sheet_index, self.workbook.sheets))
            if len(sheets_filtered) == 0:
                raise XlsxValueError("Sheet with index %i not found or can't be handled" % sheet_index)

            sheet_path = None
            # using sheet relation information
            if 'relation_id' in sheets_filtered[0] and sheets_filtered[0]['relation_id'] is not None:

                relation_id = sheets_filtered[0]['relation_id']
                if relation_id in self.workbook.relationships.relationships and \
                                'target' in self.workbook.relationships.relationships[relation_id]:
                    relationship = self.workbook.relationships.relationships[relation_id]
                    sheet_path = relationship['target']
                    if not (sheet_path.startswith("/xl/") or sheet_path.startswith("xl/")):
                        sheet_path = "/xl/" + sheet_path

            sheet_file = None
            if sheet_path is None:
                sheet_path = "/xl/worksheets/sheet%i.xml" % sheet_index
                sheet_file = self._filehandle(sheet_path)
                if sheet_file is None:
                    sheet_path = None
            if sheet_path is None:
                sheet_path = "/xl/worksheets/worksheet%i.xml" % sheet_index
                sheet_file = self._filehandle(sheet_path)
                if sheet_file is None:
                    sheet_path = None
            if sheet_path is None and sheet_index == 1:
                sheet_path = self.content_types.types["worksheet"]
                sheet_file = self._filehandle(sheet_path)
                if sheet_file is None:
                    sheet_path = None
            if sheet_file is None and sheet_path is not None:
                sheet_file = self._filehandle(sheet_path)
            if sheet_file is None:
                raise SheetNotFoundException("Sheet %i not found" % sheet_index)
            sheet = Sheet(self.workbook, self.shared_strings, self.styles, sheet_file)
            try:
                relationships_path = os.path.join(os.path.dirname(sheet_path),
                                                  "_rels",
                                                  os.path.basename(sheet_path) + ".rels")
                sheet.relationships = self._parse(Relationships, relationships_path)
                sheet.set_dateformat(self.options['dateformat'])
                sheet.set_timeformat(self.options['timeformat'])
                sheet.set_floatformat(self.options['floatformat'])
                sheet.set_skip_empty_lines(self.options['skip_empty_lines'])
                sheet.set_skip_trailing_columns(self.options['skip_trailing_columns'])
                sheet.set_include_hyperlinks(self.options['hyperlinks'])
                sheet.set_merge_cells(self.options['merge_cells'])
                sheet.set_scifloat(self.options['scifloat'])
                sheet.set_ignore_formats(self.options['ignore_formats'])
                sheet.set_skip_hidden_rows(self.options['skip_hidden_rows'])
                if self.options['escape_strings'] and sheet.filedata:
                    sheet.filedata = re.sub(r"(<v>[^<>]+)&#10;([^<>]+</v>)", r"\1\\n\2",
                                            re.sub(r"(<v>[^<>]+)&#9;([^<>]+</v>)", r"\1\\t\2",
                                                   re.sub(r"(<v>[^<>]+)&#13;([^<>]+</v>)", r"\1\\r\2", sheet.filedata)))
                sheet.to_csv(writer)
            finally:
                sheet_file.close()
                sheet.close()
        finally:
            if closefile:
                outfile.close()

    def _filehandle(self, filename):
        for name in filter(lambda f: filename and f.lower() == filename.lower()[1:], self.ziphandle.namelist()):
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
        self.sheets = list()
        self.date1904 = False

    def parse(self, filehandle):
        workbookDoc = minidom.parseString(filehandle.read())
        if workbookDoc.firstChild.namespaceURI:
            fileVersion = workbookDoc.firstChild.getElementsByTagNameNS(workbookDoc.firstChild.namespaceURI,
                                                                        "fileVersion")
        else:
            fileVersion = workbookDoc.firstChild.getElementsByTagName("fileVersion")
        if len(fileVersion) == 0:
            self.appName = DEFAULT_APP_PATH
        else:
            try:
                if workbookDoc.firstChild.namespaceURI:
                    self.appName = \
                        workbookDoc.firstChild.getElementsByTagNameNS(
                            workbookDoc.firstChild.namespaceURI, "fileVersion")[0]._attrs['appName'].value
                else:
                    self.appName = workbookDoc.firstChild.getElementsByTagName("fileVersion")[0]._attrs['appName'].value
            except KeyError:
                # no app name
                self.appName = DEFAULT_APP_PATH
        try:
            if workbookDoc.firstChild.namespaceURI:
                self.date1904 = \
                    workbookDoc.firstChild.getElementsByTagNameNS(
                        workbookDoc.firstChild.namespaceURI, "workbookPr")[0]._attrs['date1904'].value.lower().strip() \
                    != "false"
            else:
                self.date1904 = \
                    workbookDoc.firstChild.getElementsByTagName("workbookPr")[0] \
                        ._attrs['date1904'].value.lower().strip() \
                    != "false"
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
        for i, sheetNode in enumerate(sheetNodes):
            attrs = sheetNode._attrs
            name = attrs["name"].value
            state = None
            if 'state' in attrs:
                state = attrs["state"].value
            relation_id = None
            if 'r:id' in attrs:
                relation_id = attrs['r:id'].value
            self.sheets.append(
                {
                    'name': name,
                    'relation_id': relation_id,
                    'index': i + 1,
                    'id': i + 1, # remove id starting 0.8.0 version
                    'state': state
                }
            )


class ContentTypes:
    def __init__(self):
        self.types = {}
        for type in CONTENT_TYPES:
            self.types[type] = None

    def parse(self, filehandle):
        types = minidom.parseString(filehandle.read()).firstChild
        if not types:
            return
        if types.namespaceURI:
            overrideNodes = types.getElementsByTagNameNS(types.namespaceURI, "Override")
        else:
            overrideNodes = types.getElementsByTagName("Override")
        for override in overrideNodes:
            attrs = override._attrs
            type = attrs.get('ContentType').value
            name = attrs.get('PartName').value
            if type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml":
                self.types["workbook"] = name
            elif type == "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml":
                self.types["styles"] = name
            elif type == "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml":
                # BUG preserved only last sheet
                self.types["worksheet"] = name
            elif type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml":
                self.types["shared_strings"] = name
            elif type == "application/vnd.openxmlformats-package.relationships+xml":
                if self.types["relationships"] is None:
                    self.types["relationships"] = list()
                self.types["relationships"].append(name)

        if self.types["workbook"] is None:
            self.types["workbook"] = DEFAULT_WORKBOOK_PATH
        if self.types["relationships"] is None:
            self.types["relationships"] = [os.path.dirname(self.types["workbook"]) + "/_rels/" + \
                                           os.path.basename(self.types["workbook"]) + ".rels"]


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
                    "type": vtype and str(vtype.value) or None,
                    "target": target and str(target.value) or None
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

        if styles.namespaceURI:
            cellXfsElement = styles.getElementsByTagNameNS(styles.namespaceURI, "cellXfs")
        else:
            cellXfsElement = styles.getElementsByTagName("cellXfs")
        if len(cellXfsElement) == 1:
            for cellXfs in cellXfsElement[0].childNodes:
                if cellXfs.nodeType != minidom.Node.ELEMENT_NODE or not (
                                cellXfs.nodeName == "xf" or cellXfs.nodeName.endswith(":xf")):
                    continue
                if cellXfs._attrs and 'numFmtId' in cellXfs._attrs:
                    numFmtId = int(cellXfs._attrs['numFmtId'].value)
                    if self.chk_exists(numFmtId) == None:
                        numFmtId = int(cellXfs._attrs['applyNumberFormat'].value)
                    self.cellXfs.append(numFmtId)
                else:
                    self.cellXfs.append(None)

    # When Unknown Numformat ID assign applyNumberFormat
    def chk_exists(self, numFmtId):
        xfs_numfmt = numFmtId
        format_str = None
        if xfs_numfmt in self.numFmts:
            format_str = self.numFmts[xfs_numfmt]
        elif xfs_numfmt in STANDARD_FORMATS:
            format_str = STANDARD_FORMATS[xfs_numfmt]
        return format_str


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

    def replace_line_breaks(self):
        for i in range(0, len(self.strings)):
            self.strings[i] = self.strings[i].replace("\r", " ").replace("\n", " ").replace("\t", " ")

    def handleCharData(self, data):
        if self.t:
            self.value += data

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

        self.columns = {}
        self.lastRowNum = 0
        self.rowNum = None
        self.colType = None
        self.cellId = None
        self.s_attr = None
        self.data = None
        self.max_columns = -1

        self.dateformat = None
        self.timeformat = "%H:%M"  # default time format
        self.floatformat = None
        self.skip_empty_lines = False
        self.skip_trailing_columns = False

        self.filedata = None
        self.filehandle = filehandle
        self.workbook = workbook
        self.sharedStrings = sharedString.strings
        self.styles = styles

        self.hyperlinks = {}
        self.mergeCells = {}
        self.ignore_formats = []
        self.skip_hidden_rows = False

        self.colIndex = 0
        self.colNum = ""

    def close(self):
        # Make sure Worksheet is closed, parsers lib does not have a close() function, so simply delete it
        self.parser = None

    def set_dateformat(self, dateformat):
        self.dateformat = dateformat

    def set_timeformat(self, timeformat):
        if timeformat:
            self.timeformat = timeformat

    def set_floatformat(self, floatformat):
        self.floatformat = floatformat

    def set_skip_empty_lines(self, skip):
        self.skip_empty_lines = skip

    def set_skip_trailing_columns(self, skip):
        self.skip_trailing_columns = skip

    def set_ignore_formats(self, ignore_formats):
        self.ignore_formats = ignore_formats

    def set_skip_hidden_rows(self, skip_hidden_rows):
        self.skip_hidden_rows = skip_hidden_rows

    def set_merge_cells(self, mergecells):
        if not mergecells:
            return
        if not self.filedata:
            self.filedata = self.filehandle.read()
        data = str(self.filedata)  # python3: convert byte buffer to string

        # find worksheet tag, we need namespaces from it
        start = data.find("<worksheet")
        if start < 0:
            return
        end = data.find(">", start)
        worksheet = data[start: end + 1]

        # find hyperlinks part
        start = data.find("<mergeCells")
        if start < 0:
            # hyperlinks not found
            return
        end = data.find("</mergeCells>")
        data = data[start: end + 13]

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

    def set_scifloat(self, scifloat):
        self.scifloat = scifloat

    def set_include_hyperlinks(self, hyperlinks):
        if not hyperlinks or not self.relationships or not self.relationships.relationships:
            return
        # we must read file first to get hyperlinks, but we don't wont to parse whole file
        if not self.filedata:
            self.filedata = self.filehandle.read()
        data = str(self.filedata)  # python3: convert byte buffer to string

        # find worksheet tag, we need namespaces from it
        start = data.find("<worksheet")
        if start < 0:
            return
        end = data.find(">", start)
        worksheet = data[start: end + 1]

        # find hyperlinks part
        start = data.find("<hyperlinks>")
        if start < 0:
            # hyperlinks not found
            return
        end = data.find("</hyperlinks>")
        data = data[start: end + 13]

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
        self.parser.buffer_text = True
        self.parser.CharacterDataHandler = self.handleCharData
        self.parser.StartElementHandler = self.handleStartElement
        self.parser.EndElementHandler = self.handleEndElement
        if self.filedata:
            self.parser.Parse(self.filedata)
        else:
            self.parser.ParseFile(self.filehandle)

    def handleCharData(self, data):
        if self.in_cell_value:
            format_type = None
            format_str = "general"
            self.collected_string += data
            self.data = self.collected_string
            if self.colType == "s":  # shared string
                format_type = "string"
                self.data = self.sharedStrings[int(self.data)]
            elif self.colType == "b":  # boolean
                format_type = "boolean"
                self.data = (int(data) == 1 and "TRUE") or (int(data) == 0 and "FALSE") or data
            elif self.colType == "str" or self.colType == "inlineStr":
                format_type = "string"
                self.data = data
            elif self.s_attr:
                s = int(self.s_attr)

                # get cell format
                xfs_numfmt = None
                if s < len(self.styles.cellXfs):
                    xfs_numfmt = self.styles.cellXfs[s]
                if xfs_numfmt in self.styles.numFmts:
                    format_str = self.styles.numFmts[xfs_numfmt]
                elif xfs_numfmt in STANDARD_FORMATS:
                    format_str = STANDARD_FORMATS[xfs_numfmt]

                # get format type
                if not format_str:
                    raise XlsxValueError("unknown format %s at %d" % (format_str, xfs_numfmt))

                if format_str in FORMATS:
                    format_type = FORMATS[format_str]
                elif re.match(r"^\d+(\.\d+)?$", self.data) and re.match(".*[hsmdyY]", format_str) and not re.match(
                        '.*\[.*[dmhys].*\]', format_str):
                    # it must be date format
                    if float(self.data) < 1:
                        format_type = "time"
                    else:
                        format_type = "date"
                elif re.match(r"^-?\d+(.\d+)?$", self.data) or (
                            self.scifloat and re.match(r"^-?\d+(.\d+)?([eE]-?\d+)?$", self.data)):
                    format_type = "float"
                if format_type == 'date' and self.dateformat == 'float':
                    format_type = "float"
            elif self.colType == "n":
                format_type = "float"

            if format_type and not format_type in self.ignore_formats:
                try:
                    if format_type == 'date':  # date/time
                        if self.workbook.date1904:
                            date = datetime.datetime(1904, 1, 1) + datetime.timedelta(float(self.data))
                        else:
                            date = datetime.datetime(1899, 12, 30) + datetime.timedelta(float(self.data))
                        if self.dateformat:
                            # str(dateformat) - python2.5 bug, see: http://bugs.python.org/issue2782
                            self.data = date.strftime(str(self.dateformat))
                        else:
                            # ignore ";@", don't know what does it mean right now
                            # ignore "[$-409], [$-f409], [$-16001]" and similar format codes
                            dateformat = re.sub(r"\[\$\-[A-z0-9]*\]", "", format_str, 1) \
                                .replace(";@", "").replace("yyyy", "%Y").replace("yy", "%y") \
                                .replace("hh:mm", "%H:%M").replace("h", "%I").replace("%H%H", "%H") \
                                .replace("ss", "%S").replace("dddd", "d").replace("dd", "d").replace("d", "%d") \
                                .replace("am/pm", "%p").replace("mmmm", "%B").replace("mmm", "%b") \
                                .replace(":mm", ":%M").replace("m", "%m").replace("%m%m", "%m")
                            self.data = date.strftime(str(dateformat)).strip()
                    elif format_type == 'time':  # time
                        t = int(round((float(self.data) % 1) * 24 * 60 * 60, 6))  # it should be in seconds
                        d = datetime.time(int((t // 3600) % 24), int((t // 60) % 60), int(t % 60))
                        self.data = d.strftime(self.timeformat)
                    elif format_type == 'float' and ('E' in self.data or 'e' in self.data):
                        self.data = str(self.floatformat or '%f') % float(self.data)
                    # if cell is general, be aggressive about stripping any trailing 0s, decimal points, etc.
                    elif format_type == 'float' and format_str == 'general':
                        self.data = ("%f" % (float(self.data))).rstrip('0').rstrip('.')
                    elif format_type == 'float' and format_str[0:3] == '0.0':
                        if self.floatformat:
                            self.data = str(self.floatformat) % float(self.data)
                        else:
                            L = len(format_str.split(".")[1])
                            if '%' in format_str:
                                L += 1
                            self.data = ("%." + str(L) + "f") % float(self.data)
                    elif format_type == 'float':
                        # unsupported float formatting
                        self.data = ("%f" % (float(self.data))).rstrip('0').rstrip('.')

                except (ValueError, OverflowError):  # this catch must be removed, it's hiding potential problems
                    raise XlsxValueError("Error: potential invalid date format.")

    def handleStartElement(self, name, attrs):
        has_namespace = name.find(":") > 0
        if self.in_row and (name == 'c' or (has_namespace and name.endswith(':c'))):
            self.colType = attrs.get("t")
            self.s_attr = attrs.get("s")
            self.cellId = attrs.get("r")
            if self.cellId:
                self.colNum = self.cellId[:len(self.cellId) - len(self.rowNum)]
                self.colIndex = 0
            else:
                self.colIndex += 1
            self.data = ""
            self.in_cell = True
        elif self.in_cell and (
                    (name == 'v' or name == 'is') or (has_namespace and (name.endswith(':v') or name.endswith(':is')))):
            self.in_cell_value = True
            self.collected_string = ""
        elif self.in_sheet and (name == 'row' or (has_namespace and name.endswith(':row'))) and ('r' in attrs) and not (self.skip_hidden_rows and 'hidden' in attrs and attrs['hidden'] == '1'):
            self.rowNum = attrs['r']
            self.in_row = True
            self.colIndex = 0
            self.colNum = ""
            self.columns = {}
            self.spans = None
            if 'spans' in attrs:
                self.spans = [int(i) for i in attrs['spans'].split(" ")[-1].split(":")]
        elif name == 't':
            # reset collected string
            self.collected_string = ""

        elif name == 'sheetData' or (has_namespace and name.endswith(':sheetData')):
            self.in_sheet = True
        elif name == 'dimension':
            rng = attrs.get("ref").split(":")
            if len(rng) > 1:
                start = re.match(r"^([A-Z]+)(\d+)$", rng[0])
                if (start):
                    end = re.match(r"^([A-Z]+)(\d+)$", rng[1])
                    startCol = start.group(1)
                    endCol = end.group(1)
                    self.columns_count = 0
                    for cell in self._range(startCol + "1:" + endCol + "1"):
                        self.columns_count += 1

    def handleEndElement(self, name):
        has_namespace = name.find(":") > 0
        if self.in_cell and ((name == 'v' or name == 'is' or name == 't') or (
                    has_namespace and (name.endswith(':v') or name.endswith(':is')))):
            self.in_cell_value = False
        elif self.in_cell and (name == 'c' or (has_namespace and name.endswith(':c'))):
            t = 0
            for i in self.colNum: t = t * 26 + ord(i) - 64
            d = self.data
            if self.hyperlinks:
                hyperlink = self.hyperlinks.get(self.cellId)
                if hyperlink:
                    d = "<a href='" + hyperlink + "'>" + d + "</a>"
            if self.colNum + self.rowNum in self.mergeCells.keys():
                if 'copyFrom' in self.mergeCells[self.colNum + self.rowNum].keys() and \
                                self.mergeCells[self.colNum + self.rowNum]['copyFrom'] == self.colNum + self.rowNum:
                    self.mergeCells[self.colNum + self.rowNum]['value'] = d
                else:
                    d = self.mergeCells[self.mergeCells[self.colNum + self.rowNum]['copyFrom']]['value']

            self.columns[t - 1 + self.colIndex] = d
            self.in_cell = False

        if self.in_row and (name == 'row' or (has_namespace and name.endswith(':row'))):
            if len(self.columns.keys()) > 0:
                if min(self.columns.keys()) < 0: # Weird
                    d = []
                    keys = self.columns.keys()
                    keys.sort()
                    for k in keys:
                        val = self.columns[k]
                        if not self.py3:
                            val = val.encode("utf-8")
                        d.append(val)
                else:
                    d = [""] * (max(self.columns.keys()) + 1)
                    for k in self.columns.keys():
                        val = self.columns[k]
                        if not self.py3:
                            val = val.encode("utf-8")
                        d[k] = val
                if self.spans:
                    l = self.spans[1]
                    if len(d) < l:
                        d += (l - len(d)) * ['']

                # write empty lines
                if not self.skip_empty_lines:
                    for i in range(self.lastRowNum, int(self.rowNum) - 1):
                        self.writer.writerow([])
                    self.lastRowNum = int(self.rowNum)

                # write line to csv
                if not self.skip_empty_lines or d.count('') != len(d):
                    while len(d) < self.columns_count:
                        d.append("")

                    if self.skip_trailing_columns:
                        if self.max_columns < 0:
                            self.max_columns = len(d)
                            while len(d) > 0 and d[-1] == "":
                                d = d[0:-1]
                                self.max_columns = self.max_columns - 1
                        elif self.max_columns > 0:
                            d = d[0:self.max_columns]
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
            start = re.match(r"^([A-Z]+)(\d+)$", rng[0])
            end = re.match(r"^([A-Z]+)(\d+)$", rng[1])
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
    for name in os.listdir(path):
        fullpath = os.path.join(path, name)
        if os.path.isdir(fullpath):
            convert_recursive(fullpath, sheetid, outfile, kwargs)
        else:
            outfilepath = outfile
            if isinstance(outfilepath, type(sys.stdout)):
                outfilepath = fullpath[:-4] + 'csv'
            elif os.path.isdir(outfilepath):
                outfilepath = os.path.join(outfilepath, name[:-4] + 'csv')
            elif len(outfilepath) == 0 and fullpath.lower().endswith(".xlsx"):
                outfilepath = fullpath[:-4] + 'csv'

            print("Converting %s to %s" % (fullpath, outfilepath))
            try:
                Xlsx2csv(fullpath, **kwargs).convert(outfilepath, sheetid)
            except zipfile.BadZipfile:
                raise InvalidXlsxFileException("File %s is not a zip file" % fullpath)


def main():
    try:
        signal.signal(signal.SIGPIPE, signal.SIG_DFL)
        signal.signal(signal.SIGINT, signal.SIG_DFL)
    except AttributeError:
        pass

    if "ArgumentParser" in globals():
        parser = ArgumentParser(description="xlsx to csv converter")
        parser.add_argument('infile', metavar='xlsxfile', help="xlsx file path")
        parser.add_argument('outfile', metavar='outfile', nargs='?', help="output csv file path")
        parser.add_argument('-v', '--version', action='version', version=__version__)
        nargs_plus = "+"
        argparser = True
    else:
        parser = OptionParser(usage="%prog [options] infile [outfile]", version=__version__)
        parser.add_argument = parser.add_option
        nargs_plus = 1
        argparser = False

    if sys.version_info[0] == 2 and sys.version_info[1] < 5:
        inttype = "int"
    else:
        inttype = int
    parser.add_argument("-a", "--all", dest="all", default=False, action="store_true",
                        help="export all sheets")
    parser.add_argument("-c", "--outputencoding", dest="outputencoding", default="utf-8", action="store",
                        help="encoding of output csv ** Python 3 only ** (default: utf-8)")
    parser.add_argument("-d", "--delimiter", dest="delimiter", default=",",
                        help="delimiter - columns delimiter in csv, 'tab' or 'x09' for a tab (default: comma ',')")
    parser.add_argument("--hyperlinks", "--hyperlinks", dest="hyperlinks", action="store_true", default=False,
                        help="include hyperlinks")
    parser.add_argument("-e", "--escape", dest='escape_strings', default=False, action="store_true",
                        help="Escape \\r\\n\\t characters")
    parser.add_argument("--no-line-breaks", "--no-line-breaks", dest='no_line_breaks', default=False, action="store_true",
                        help="Replace \\r\\n\\t with space")
    parser.add_argument("-E", "--exclude_sheet_pattern", nargs=nargs_plus, dest="exclude_sheet_pattern", default="",
                        help="exclude sheets named matching given pattern, only effects when -a option is enabled.")
    parser.add_argument("-f", "--dateformat", dest="dateformat",
                        help="override date/time format (ex. %%Y/%%m/%%d)")
    parser.add_argument("-t", "--timeformat", dest="timeformat",
                        help="override time format (ex. %%H/%%M/%%S)")
    parser.add_argument("--floatformat", dest="floatformat",
                        help="override float format (ex. %%.15f)")
    parser.add_argument("--sci-float", dest="scifloat", default=False, action="store_true",
                        help="force scientific notation to float")
    parser.add_argument("-I", "--include_sheet_pattern", nargs=nargs_plus, dest="include_sheet_pattern", default="^.*$",
                        help="only include sheets named matching given pattern, only effects when -a option is enabled.")
    parser.add_argument("--exclude_hidden_sheets", default=False, action="store_true",
                        help="Exclude hidden sheets from the output, only effects when -a option is enabled.")
    parser.add_argument("--ignore-formats", nargs=nargs_plus, type=str, dest="ignore_formats", default=[''],
                        help="Ignores format for specific data types.")
    parser.add_argument("-l", "--lineterminator", dest="lineterminator", default="\n",
                        help="line terminator - lines terminator in csv, '\\n' '\\r\\n' or '\\r' (default: \\n)")
    parser.add_argument("-m", "--merge-cells", dest="merge_cells", default=False, action="store_true",
                        help="merge cells")
    parser.add_argument("-n", "--sheetname", dest="sheetname", default=None,
                        help="sheet name to convert")
    parser.add_argument("-i", "--ignoreempty", dest="skip_empty_lines", default=False, action="store_true",
                        help="skip empty lines")
    parser.add_argument("--skipemptycolumns", dest="skip_trailing_columns", default=False, action="store_true",
                        help="skip trailing empty columns")
    parser.add_argument("-p", "--sheetdelimiter", dest="sheetdelimiter", default="--------",
                        help="sheet delimiter used to separate sheets, pass '' if you do not need delimiter, or 'x07' "
                             "or '\\f' for form feed (default: '--------')")
    parser.add_argument("-q", "--quoting", dest="quoting", default="minimal",
                        help="quoting - fields quoting in csv, 'none' 'minimal' 'nonnumeric' or 'all' (default: minimal)")
    parser.add_argument("-s", "--sheet", dest="sheetid", default=1, type=inttype,
                        help="sheet number to convert")
    parser.add_argument("--include-hidden-rows", dest="include_hidden_rows", default=False, action="store_true",
                        help="include hidden rows")

    if argparser:
        options = parser.parse_args()
    else:
        (options, args) = parser.parse_args()
        if len(args) < 1:
            parser.print_usage()
            sys.exit("error: too few arguments" + os.linesep)
        options.infile = args[0]
        options.outfile = len(args) > 1 and args[1] or None

    if len(options.delimiter) == 1:
        pass
    elif options.delimiter == 'tab' or options.delimiter == '\\t':
        options.delimiter = '\t'
    elif options.delimiter == 'comma':
        options.delimiter = ','
    elif options.delimiter[0] == 'x':
        options.delimiter = chr(int(options.delimiter[1:]))
    else:
        sys.exit("error: invalid delimiter\n")

    if options.quoting == 'none':
        options.quoting = csv.QUOTE_NONE
    elif options.quoting == 'minimal':
        options.quoting = csv.QUOTE_MINIMAL
    elif options.quoting == 'nonnumeric':
        options.quoting = csv.QUOTE_NONNUMERIC
    elif options.quoting == 'all':
        options.quoting = csv.QUOTE_ALL
    else:
        sys.exit("error: invalid quoting\n")

    if options.lineterminator == '\n':
        pass
    elif options.lineterminator == '\\n':
        options.lineterminator = '\n'
    elif options.lineterminator == '\\r':
        options.lineterminator = '\r'
    elif options.lineterminator == '\\r\\n':
        options.lineterminator = '\r\n'
    else:
        sys.exit("error: invalid line terminator\n")

    if options.sheetdelimiter == '--------':
        pass
    elif options.sheetdelimiter == '':
        pass
    elif options.sheetdelimiter == '\\f':
        options.sheetdelimiter = '\f'
    elif options.sheetdelimiter[0] == 'x':
        options.sheetdelimiter = chr(int(options.sheetdelimiter[1:]))
    else:
        sys.exit("error: invalid sheet delimiter\n")

    kwargs = {
        'delimiter': options.delimiter,
        'quoting': options.quoting,
        'sheetdelimiter': options.sheetdelimiter,
        'dateformat': options.dateformat,
        'timeformat': options.timeformat,
        'floatformat': options.floatformat,
        'scifloat': options.scifloat,
        'skip_empty_lines': options.skip_empty_lines,
        'skip_trailing_columns': options.skip_trailing_columns,
        'escape_strings': options.escape_strings,
        'no_line_breaks': options.no_line_breaks,
        'hyperlinks': options.hyperlinks,
        'include_sheet_pattern': options.include_sheet_pattern,
        'exclude_sheet_pattern': options.exclude_sheet_pattern,
        'exclude_hidden_sheets': options.exclude_hidden_sheets,
        'merge_cells': options.merge_cells,
        'outputencoding': options.outputencoding,
        'lineterminator': options.lineterminator,
        'ignore_formats': options.ignore_formats,
        'skip_hidden_rows': not options.include_hidden_rows
    }
    sheetid = options.sheetid
    if options.all:
        sheetid = 0

    outfile = options.outfile or sys.stdout
    try:
        if os.path.isdir(options.infile):
            convert_recursive(options.infile, sheetid, outfile, kwargs)
        else:
            xlsx2csv = Xlsx2csv(options.infile, **kwargs)
            if options.sheetname:
                sheetid = xlsx2csv.getSheetIdByName(options.sheetname)
                if not sheetid:
                    sys.exit("Sheet '%s' not found" % options.sheetname)
            xlsx2csv.convert(outfile, sheetid)
    except XlsxException:
        _, e, _ = sys.exc_info()
        sys.exit(str(e) + "\n")


if __name__ == "__main__":
    main()
