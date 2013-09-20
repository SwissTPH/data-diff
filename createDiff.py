#!/usr/bin/env python
import csv
import sys
import StringIO
import contextlib
import datetime
import time
import xlwt
import xlrd

tabletDataFile = {'name': 'DemoDataTablet.csv', 'delimiter': ';'}
paperDataFile = {'name': 'DemoDataPaper.csv', 'delimiter': ';'}
mappingFile = {'name': 'mappingDemo.csv', 'delimiter': ';'}
reportTxtFile = 'report.txt'
reportExcelFile = 'report.xls'

keywords = {'IID': 'EID', 'runTime': 'RUN_TIME', 'startTime': 'STIME', 'tabletTimestamp': 'TABLET_TIMESTAMP'}

tabletMappings = []
paperMappings = []

txtReport = open(reportTxtFile, 'wb')

w = xlwt.Workbook()
ws_paper = w.add_sheet('Paper')
ws_tablet = w.add_sheet('Tablet')
ws_diff = w.add_sheet('Diff')


#initialize mappings from tablet and paper data to a master format
def initMappings():
    mappingReader = csv.reader(open(mappingFile['name'], 'rb'), delimiter=mappingFile['delimiter'])
    #skip header
    mappingReader.next()
    #set up all the mappings where a master variable name is present
    varcount = 0
    for row in mappingReader:
        if not row[0] == '':
            varcount += 1
            name = str(varcount).zfill(4) + row[0]
            #RUN_TIME is a keyword in the mapping file, used to calculate run times for a discrepancy analysis
            if name[4:12] == keywords['runTime']:
                row[2] = keywords['tabletTimestamp']
            tabletMappings.append([name, row[1], row[2]])
            paperMappings.append([name, row[3], row[4]])


#consolidate trivial formatting differences
def reformat(data):
    reformatted = data.upper().strip()
    return reformatted


#get the value of a variable by its name
def getVarValue(varName, header, record):
    if varName in header:
        return record[header.index(varName)]
    else:
        return ''


#sum integer variables
def sumIntVariables(varNames, header, record):
    iSum = 0
    for varName in varNames:
        s = getVarValue(varName, header, record).strip()
        if s:
            iSum += int(s)
    return iSum


#sum coded values of multi-choice options
def sumOptions(data):
    if data:
        options = data.split(' ')
        iSum = 0
        for option in options:
            iSum += int(option)
        return iSum
    else:
        return ''


#determine what number a certain location in an option list should have
def getOptionNumber(varNames, header, record, numStartIndex):
    res = ''
    for name in varNames:
        if getVarValue(name, header, record) == 'TRUE':
            res = ' '.join([res, str(int(name[numStartIndex:]))])
    return res


#redirect stdout to a string
@contextlib.contextmanager
def stdoutToString():
    now = sys.stdout
    stdout = StringIO.StringIO()
    sys.stdout = stdout
    yield stdout
    sys.stdout = now


#execute python code in the mapping file, return stdout as a string
def executeMappingRule(data, mappingRule, header, record):
    with stdoutToString() as s:
        exec mappingRule
    return s.getvalue().strip()


#map a single record, store in a dictionary with interview ID as key
def mapRecord(mappings, header, record):
    IID = ""
    mappedRecord = {}
    for mapping in mappings:
        mappedName = mapping[0]
        originalName = mapping[1]
        mappingRule = mapping[2]
        data = getVarValue(originalName, header, record)
        if mappedName[4:] == keywords['IID']:
            IID = data
        else:
            mappedRecord[mappedName] = ['', '']
            #TABLET_TIMESTAMP as mapped above
            if mappingRule == keywords['tabletTimestamp']:
                mappedRecord[mappedName][0] = data
                mappedRecord[mappedName][1] = mappedRecord[mappedName][0]
            else:
                mappedRecord[mappedName][0] = reformat(data)
                mappedRecord[mappedName][1] = mappedRecord[mappedName][0]
                if not mappingRule == '':
                    #TODO: we should keep non-reformatting mappings in a separate record, for this to go into the
                    #resolved sheet, and the non-mapped in to the original
                    mappedRecord[mappedName][1] = executeMappingRule(mappedRecord[mappedName][0],
                                                                     mappingRule, header, record)
    return IID, mappedRecord


#read csv files and create a dictionary of records
def readAndMapRecords(reader, mapping):
    header = reader.next()
    records = {}
    for record in reader:
        IID, data = mapRecord(mapping, header, record)
        records[IID] = data
    return records


#map paper interview ID to tablet interview ID or vice versa
def mapIID(IID):
    if IID.find('F') >= 0:
        IID = IID.replace('F', 'L')
    else:
        IID = IID.replace('L', 'F')
    if IID.find('T') >= 0:
        return IID.replace('T', 'P')
    else:
        return IID.replace('P', 'T')


#look for missing records
def findMissing(tabletKeys, paperKeys):
    txtReport.write('Missing records:\r\n')
    for tIID in tabletKeys:
        if not mapIID(tIID) in paperKeys:
            txtReport.write(tIID + ' not found in paper records, looking for ' + mapIID(tIID) + '\r\n')
    for pIID in paperKeys:
        if not mapIID(pIID) in tabletKeys:
            txtReport.write(pIID + ' not found in tablet records, looking for ' + mapIID(pIID) + '\r\n')


#compare data for a pair of records
def compareFields(tabletIID, tabletData, paperData, IIDIndex):
    ws_paper.write(IIDIndex, 0, mapIID(tabletIID))
    ws_tablet.write(IIDIndex, 0, tabletIID)
    ws_diff.write(IIDIndex, 0, tabletIID[2:])
    startTime = None
    varIndex = 1
    for field in sorted(tabletData.keys()):
        name = field[4:]
        if name == keywords['startTime']:
            startTime = datetime.datetime.strptime(tabletData[field][0], "%I:%M:%S %p")
        if IIDIndex == 1:
            ws_diff.write(0, varIndex, name)
            ws_paper.write(0, varIndex, name)
            ws_tablet.write(0, varIndex, name)
        if name[:8] == keywords['runTime']:
            if tabletData[field][0] == '':
                runtime = 'n.a.'
            else:
                timestamp = datetime.datetime.strptime(tabletData[field][0], "%I:%M:%S %p")
                runtime = (timestamp - startTime).seconds / 60.0
            ws_paper.write(IIDIndex, varIndex, runtime)
            ws_tablet.write(IIDIndex, varIndex, runtime)
            ws_diff.write(IIDIndex, varIndex, runtime)
        else:
            if name != keywords['startTime']:
                ws_paper.write(IIDIndex, varIndex, paperData[field][1])
            ws_tablet.write(IIDIndex, varIndex, tabletData[field][1])
            row = str(IIDIndex + 1)
            col = xlrd.colname(varIndex)
            ws_diff.write(IIDIndex, varIndex,
                          xlwt.Formula('IF(Paper!' + col + row + '=Tablet!' + col + row +
                                       ';"";CONCATENATE(Paper!' + col + row +
                                       ';" --- ";Tablet!' + col + row + '))'))
        varIndex += 1


if __name__ == '__main__':

    initMappings()

    tabletReader = csv.reader(open(tabletDataFile['name'], 'rb'), delimiter=tabletDataFile['delimiter'])
    tabletRecords = readAndMapRecords(tabletReader, tabletMappings)
    paperReader = csv.reader(open(paperDataFile['name'], 'rb'), delimiter=paperDataFile['delimiter'])
    paperRecords = readAndMapRecords(paperReader, paperMappings)

    tabletKeys = tabletRecords.keys()
    paperKeys = paperRecords.keys()
    findMissing(tabletKeys, paperKeys)

    IIDIndex = 1
    for tIID in tabletKeys:
        if mapIID(tIID) in paperKeys:
            compareFields(tIID, tabletRecords[tIID], paperRecords[mapIID(tIID)], IIDIndex)
            IIDIndex += 1

    txtReport.close()
    w.save(reportExcelFile)
