#!/usr/bin/python2.7
# coding=utf-8
import xlrd
import xlwt
import json
import os.path
import datetime
import os
from glob import glob
import shutil
import urllib
from pip._vendor.six import u
import codecs


def getColNames(sheet):
    rowSize = sheet.row_len (0)
    colValues = sheet.row_values (0, 0, rowSize)
    columnNames = []

    for value in colValues:
        if value is u'':
            pass
        else:
            # unicode(value, "utf8")
            columnNames.append (value)

    return columnNames


def getRowData(row, columnNames):
    rowData = {}
    counter = 0
    columnlen = len (columnNames) - 1
    rowlen = len (row)

    for cell in row:
        # check if it is of date type print in iso format
        if counter <= columnlen:
            if cell.ctype == xlrd.XL_CELL_DATE:
                rowData[columnNames[counter].replace (' ', '_')] = datetime.datetime (
                    *xlrd.xldate_as_tuple (cell.value, 0)).isoformat ()
            else:
                if cell.ctype == xlrd.XL_CELL_TEXT:
                    cell.value = cell.value.strip ()
                else:
                    pass
                rowData[columnNames[counter].replace (' ', '_')] = cell.value
        else:
            pass
        counter += 1

    return rowData


def getSheetData(sheet, columnNames):
    nRows = sheet.nrows
    sheetData = []
    counter = 1

    for idx in range (1, nRows):
        row = sheet.row (idx)
        rowData = getRowData (row, columnNames)
        sheetData.append (rowData)

    return sheetData


def getWorkBookData(workbook):
    # type: (object) -> object
    sheets = workbook.sheets ()
    print type (sheets)
    counter = 0
    # workbookdata = {}
    workbookList = []
    for sheet in sheets:
        sheetName = sheet.name
        columnNames = getColNames (sheet)

        # organize workbooList
        sheetInfo = {}
        aColsData = []
        for columnName in columnNames:
            aColsDict = {}
            aColsDict['sName'] = columnName
            aColsDict['sTechName'] = ""
            aColsData.append(aColsDict)

        sheetInfo["sSheetName"] = sheetName
        sheetInfo["aCols"] = aColsData
        #sheetInfo = sorted(sheetInfo.keys(), reverse=True)
        workbookList.append(sheetInfo)
    return workbookList


def main():
    filename = raw_input ("Enter the path to the filename -> ")
    print type (filename)
    filename = unicode (filename, "utf8")
    if os.path.isfile (filename):
        workbook = xlrd.open_workbook (filename)
        workbookdata = getWorkBookData (workbook)
        file_name = os.path.splitext(filename)[0]
        with codecs.open(file_name+'.json', 'w', 'utf-8') as w:
            w.write(json.dumps(workbookdata, ensure_ascii=False, sort_keys=True, indent=2,  separators=(',', ": ")))
        w.close()
        print "%s was created" % file_name
    else:
        print "Sorry, that was not a valid filename"


main ()
