#!/usr/bin/python

__author__ = "Dilshod Temirkhodjaev <tdilshod@gmail.com>"

import csv, datetime, zipfile, sys
import xml.parsers.expat
from xml.dom import minidom

#
# example: xlsx2csv("test.xslx", open("test.csv", "w+"))
#
def xlsx2csv(infilepath, outfile):
    writer = csv.writer(outfile, quoting=csv.QUOTE_MINIMAL)
    ziphandle = zipfile.ZipFile(infilepath)
    shared_strings = SharedStrings(ziphandle.read("xl/sharedStrings.xml"))

    # multisheet:
    #self.workbook = Workbook(ziphandle.read("xl/workbook.xml"))
    #for i in self.workbook.sheets:
    #    SharedStrings(ziphandle.read("xl/worksheets/sheet%s.xml" %(i['id'])))

    Sheet(shared_strings, ziphandle.read("xl/worksheets/sheet1.xml"), writer)
    ziphandle.close()

class Workbook(object):
    sheets = []
    def __init__(self, data):
        workbookDoc = minidom.parseString(data)
        sheets = workbookDoc.firstChild.getElementsByTagName("sheets")[0]
        for sheetNode in sheets.childNodes:
            name = sheetNode._attrs["name"].value
            id = int(sheetNode._attrs["r:id"].value[3:])
            self.sheets.append({'name': name, 'id': id})

class SharedStrings:
    parser = None
    strings = []
    si = False
    t = False
    value = ""

    def __init__(self, data):
        self.parser = xml.parsers.expat.ParserCreate()
        self.parser.CharacterDataHandler = self.handleCharData
        self.parser.StartElementHandler = self.handleStartElement
        self.parser.EndElementHandler = self.handleEndElement
        self.parser.Parse(data)

    def handleCharData(self, data):
        if self.t: self.value+= data

    def handleStartElement(self, name, attrs):
        if name == 'si':
            self.si = True
        elif name == 't' and self.si:
            self.t = True
            self.value = ""

    def handleEndElement(self, name):
        if name == 'si':
            self.si = False
            self.strings.append(self.value)
        elif name == 't':
            self.t = False

class Sheet:
    parser = None
    writer = None
    sharedString = None

    in_sheet = False
    in_row = False
    in_cell = False
    in_cell_value = False
    in_cell_formula = False

    columns = {}
    rowNum = None
    colType = None
    s_attr = None
    data = None

    def __init__(self, sharedString, data, writer):
        self.writer = writer
        self.sharedStrings = sharedString.strings
        self.parser = xml.parsers.expat.ParserCreate()
        self.parser.CharacterDataHandler = self.handleCharData
        self.parser.StartElementHandler = self.handleStartElement
        self.parser.EndElementHandler = self.handleEndElement
        self.parser.Parse(data)

    def handleCharData(self, data):
        if self.in_cell_value:
            if self.colType == "s":
                # shared
                self.data = self.sharedStrings[int(data)]
            #elif self.colType == "b": # boolean
            elif self.s_attr:
                if self.s_attr == '2':
                    # date
                    try:
                        self.data = (datetime.date(1899, 12, 30) + datetime.timedelta(float(data))).strftime("%m/%d/%y")
                    except (ValueError, OverflowError):
                        # invalid date format
                        self.data = data
                elif self.s_attr == '3':
                    # time
                    self.data = str(float(data) * 24*60*60)
                    # datetime
                elif self.s_attr == '1':
                    try:
                        self.data = (datetime.datetime(1899, 12, 30) + datetime.timedelta(float(data))).strftime("%m/%d/%y %H:%M")
                    except (ValueError, OverflowError):
                        # invalid date format
                        self.data = data
                else:
                    self.data = data
            else:
                self.data = data
        # does not support it yet
        #elif self.in_cell_formula:
        #    self.formula = data

    def handleStartElement(self, name, attrs):
        if self.in_row and name == 'c':
            self.colType = attrs.get("t")
            self.s_attr = attrs.get("s")
            cellId = attrs.get("r")
            self.colNum = cellId[:len(cellId)-len(self.rowNum)]
            #self.formula = None
            self.data = ""
            self.in_cell = True
        elif self.in_cell and name == 'v':
            self.in_cell_value = True
        #elif self.in_cell and name == 'f':
        #    self.in_cell_formula = True
        elif self.in_sheet and name == 'row' and attrs.has_key('r'):
            self.rowNum = attrs['r']
            self.in_row = True
            self.columns = {}
        elif name == 'sheetData':
            self.in_sheet = True

    def handleEndElement(self, name):
        if self.in_cell and name == 'v':
            self.in_cell_value = False
        #elif self.in_cell and name == 'f':
        #    self.in_cell_formula = False
        elif self.in_cell and name == 'c':
            t = 0
            for i in self.colNum: t = t*26 + ord(i) - 65
            self.columns[t] = self.data
            self.in_cell = False
        if self.in_row and name == 'row':
            d = [""] * (max(self.columns.keys()) + 1)
            for k in self.columns.keys():
                d[k] = self.columns[k].encode("utf-8")
            self.writer.writerow(d)
            self.in_row = False
        elif self.in_sheet and name == 'sheetData':
            self.in_sheet = False

if __name__ == "__main__":
    if len(sys.argv) == 2:
        xlsx2csv(sys.argv[1], sys.stdout)
    elif len(sys.argv) == 3:
        f = open(sys.argv[2], "w+")
        xlsx2csv(sys.argv[1], f)
        f.close()
    else:
        print "Usage: xlsx2csv <infile> [<outfile>]"
