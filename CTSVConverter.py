#! /usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import datetime
import csv
import re
from pathlib import Path

try:
    import xlsxwriter
except ImportError:
    print("缺少依赖项，正在自动安装...")
    os.system("pip3 install -q -U xlsxwriter")
    print("正在重载工具...")
    os.system(' '.join(sys.argv))
    exit(0)

standardEndodings = [
    'utf8',
    'utf-8-sig',
    '8859',
    'gbk',
    'windows-1252',
    'utf-16-le',
]

def answer(question):
    return input(question)

def ToCSVLine(objArray, textIndexArray = []):
    csvLine = ''
    objIndex = 0
    for obj in objArray:
        if (',' in obj) or ('"' in obj) or ("'" in obj):
            obj = "\"" + obj + "\""
        
        if str(objIndex) in textIndexArray:
            obj = "'" + obj + ","

        csvLine = csvLine + obj + ","
        objIndex += 1
    return csvLine + "\n"

def GetNumberType(number):
    numberType = None
    if '%' in number:
        number.replace("%","")
        numberType = "Percentage"
    elif '.' in number:
        numberType = "Float"
    else:
        numberType = "Int"

    try:
        float(number)
    except ValueError:
        return None

    return numberType

buildInfo = "20220718B"
appName = "CTSV Converter V1.3.1"

encoding = ''
inputTextColumn = ''
inputDatetimeColumn = ''
inputNumberColumn = ''
inputSortedColumn = ''
isKeep = False
isFirst = True

def main():
    os.system("title " + appName)
    print("Build: " + buildInfo)
    print("Author: Meano&Daisy")
    print("Press Ctrl + C to exit")

    print("============================== Step 0 ==============================")
    filePath = ''
    fileSuffix = ''
    while not (os.path.exists(filePath) and (os.path.isdir(filePath) or fileSuffix.lower() == ".csv" or fileSuffix.lower() == ".tsv")):
        filePath = answer("Please input CSV/TSV file or dir path: ").replace("\"", "").replace("\'", "")
        fileSuffix = Path(filePath).suffix

    global isKeep, isFirst
    isKeep = False
    isFirst = True

    if not os.path.isdir(filePath):
        ConvertToXlsx(filePath, fileSuffix)
    else:
        if answer("If use same config for all files? (`Enter` to keep, any other key to config files every time): ") != '':
            isKeep = False
        else:
            isKeep = True

        for file_walk in os.walk(filePath):
            for fileInDir in file_walk[2]:
                fileSuffix = Path(fileInDir).suffix
                if fileSuffix.lower() == ".csv" or fileSuffix.lower() == ".tsv":
                    ConvertToXlsx(file_walk[0] + "\\" + fileInDir, fileSuffix)

def ConvertToXlsx(filePath, fileSuffix):
    print("============================== Step 1 ==============================")
    print("Converting file: " + filePath + "...")

    DatetimeFormatDict = {
        "%Y-%m-%d %H:%M:%S": "yyyy-mm-dd hh:mm:ss",
        "%Y-%m-%d %H:%M": "yyyy-mm-dd hh:mm",
        "%Y-%m-%d": "yyyy-mm-dd",
    }

    fileDelimiter = ','
    if fileSuffix.lower() == ".tsv":
        fileDelimiter = '\t'

    global isFirst, isKeep, encoding, inputTextColumn, inputDatetimeColumn, inputNumberColumn, inputSortedColumn

    print("============================== Step 2 ==============================")
    if not isKeep or isFirst:
        encoding = ''
        while encoding == '':
            print("Please select encoding type of CSV/TSV file:")
            print("0. UTF-8 (Unicode)")
            print("1. UTF-8 With BOM (Unicode)")
            print("2. ISO-8859 (Westen encode)")
            print("3. GBK (Chinese encode)")
            print("4. Windows-1252")
            print("5. UTF-16-LE")
            try:
                encodingIndex = int(answer("Please choose encoding type number(0~5): "))
                encoding = standardEndodings[encodingIndex]
            except:
                print(sys.exc_info())

    csvFile = open(filePath, 'r', encoding = encoding)
    csvReader = csv.DictReader(csvFile, delimiter = fileDelimiter)

    print("============================== Step 3 ==============================")
    if not isKeep or isFirst:
        print("Please input columns format index numbers or names split with ',':")
        columnInfo = ''
        columnIndex = 0
        for columnName in csvReader.fieldnames:
            columnInfo = columnInfo + str(columnIndex) + "." + columnName + (";\n" if columnIndex % 5 == 4 else "; ")
            columnIndex += 1
        print(columnInfo)

        inputTextColumn = answer("Input Text column index / name: ")
        inputDatetimeColumn = answer("Input Datetime column index / name: ")
        inputNumberColumn = answer("Input Number column index / name: ")

    textColumn = inputTextColumn
    datetimeColumn = inputDatetimeColumn
    numberColumn = inputNumberColumn

    for columnIndex, columnName in enumerate(csvReader.fieldnames):
        textColumn = textColumn.replace(columnName, str(columnIndex))
        datetimeColumn = datetimeColumn.replace(columnName, str(columnIndex))
        numberColumn = numberColumn.replace(columnName, str(columnIndex))

    textColumn = re.findall('[0-9]+', textColumn)
    datetimeColumn = re.findall('[0-9]+', datetimeColumn)
    numberColumn = re.findall('[0-9]+', numberColumn)

    if len(textColumn): print("Columns " + str(textColumn) + " will trade as text column.")
    if len(datetimeColumn): print("Columns " + str(datetimeColumn) + " will trade as datetime column.")
    if len(numberColumn): print("Columns " + str(numberColumn) + " will trade as number column.")

    columnIndex = 0
    csvFieldTypeDict = {}
    for csvFieldName in csvReader.fieldnames:
        if textColumn.__contains__(str(columnIndex)):
            csvFieldTypeDict[csvFieldName] = "Text"
        elif datetimeColumn.__contains__(str(columnIndex)):
            csvFieldTypeDict[csvFieldName] = "Datetime"
        elif numberColumn.__contains__(str(columnIndex)):
            csvFieldTypeDict[csvFieldName] = "Number"
        columnIndex += 1

    print("============================== Step 4 ==============================")
    if not isKeep or isFirst:
        print("Please input sorted index split with ',':")

        columnInfo = ''
        columnIndex = 0
        for columnName in csvFieldTypeDict.keys():
            columnInfo = columnInfo + str(columnIndex) + "." + columnName + (";\n" if columnIndex % 5 == 4 else "; ")
            columnIndex += 1
        print(columnInfo)

        inputSortedColumn = answer("Input sorted column index / name: ")

        isFirst = False

    sortedColumn = inputSortedColumn

    for columnIndex, columnName in enumerate(csvFieldTypeDict.keys()):
        sortedColumn = sortedColumn.replace(columnName, str(columnIndex))

    sortedColumn = re.findall('[0-9]+', sortedColumn)

    sortedDict = {}
    csvFieldTypeList = list(csvFieldTypeDict)
    for sortedIndex in sortedColumn:
        itemkey = csvFieldTypeList[int(sortedIndex)]
        sortedDict[itemkey] = csvFieldTypeDict[itemkey]
        csvFieldTypeDict.pop(itemkey)

    csvFieldTypeDict = {**sortedDict, **csvFieldTypeDict}

    print("============================== Converting ==============================")

    xlsxPath = filePath.replace(fileSuffix, ".xlsx")
    print("Will convert file to: " + xlsxPath)
    workbook = xlsxwriter.Workbook(xlsxPath)

    percentageFormat = workbook.add_format({'num_format': '0.00%'})

    sheet = workbook.add_worksheet()
    sheet.write_row(0, 0, csvFieldTypeDict.keys())

    rowIndex = 1
    for csvRow in csvReader:
        colIndex = 0
        for csvFieldName in csvFieldTypeDict.keys():
            item = csvRow[csvFieldName]
            itemType = csvFieldTypeDict[csvFieldName]

            try:
                if itemType == "Text" or item == "":
                    item = (item, None)
                elif itemType.startswith("Datetime"):
                    colDatetimeFormat = itemType.replace("Datetime", "")

                    if rowIndex < 1000 and colDatetimeFormat == "":
                        for datetimeFormat in DatetimeFormatDict.keys():
                            try:
                                datetime.datetime.strptime(item, datetimeFormat)
                                csvFieldTypeDict[csvFieldName] = "Datetime" + datetimeFormat
                                colDatetimeFormat = datetimeFormat
                                if isinstance(DatetimeFormatDict[datetimeFormat], str):
                                    DatetimeFormatDict[datetimeFormat] = workbook.add_format({'num_format': DatetimeFormatDict[datetimeFormat]})
                                print("Datetime format of " + csvFieldName + ":" + item + " is " + datetimeFormat)
                                break
                            except:
                                print(sys.exc_info())
                        if  csvFieldTypeDict[csvFieldName] == "Datetime":
                            print("Datetime format of " + csvRow[csvFieldName] + " is unknow!")

                    item = (datetime.datetime.strptime(csvRow[csvFieldName], colDatetimeFormat), DatetimeFormatDict[colDatetimeFormat])

                elif itemType.startswith("Number"):
                    colNumberFormat = itemType.replace("Number", "")

                    if rowIndex < 1000 and colNumberFormat == "" and (colNumberFormat := GetNumberType(item)) != None:
                        itemType = csvFieldTypeDict[csvFieldName] = "Number" + colNumberFormat

                    if colNumberFormat == "Float" or colNumberFormat == "Int":
                        item = (float(item), None)
                    elif colNumberFormat == "Percentage":
                        item = (float(item.replace("%", "")), percentageFormat)

            except:
                print("Format error: Cell({0}, {1}), Type({2}), Item({3}), Field({4})".format(colIndex, rowIndex, itemType, item, csvFieldName))

            if isinstance(item, str):
                item = (item, None)

            sheet.write(rowIndex, colIndex, item[0], item[1])
            colIndex += 1
        rowIndex += 1
        if rowIndex % 2000 == 0:
            print("Appended rows " + str(rowIndex))
    print("Appended all rows " + str(rowIndex))
    workbook.close()

    print("============================== Convert done ==============================\n\n")

if __name__ == '__main__':
    try:
        while True:
            main()
    except KeyboardInterrupt:
        print("Exit!")
    except Exception as e:
        print(e)
        os.system("pause")
