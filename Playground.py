"""
import calendar
from datetime import date, timedelta, datetime

weekDays=["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday", "Sunday"]

def getLastDayOfMonth(yourDate, weekDay):

    day=weekDays.index(weekDay)
    month_range=calendar.monthrange(yourDate.year,yourDate.month)
    lastDateOfMonth = date(your_date.year, your_date.month, month_range[1])
    deltaDays=0

    if datetime.weekday(lastDateOfMonth) < day:
        deltaDays = day-datetime.weekday(lastDateOfMonth)-7
    else:
        deltaDays = day - datetime.weekday(lastDateOfMonth)

    return lastDateOfMonth+timedelta(days=deltaDays)

def getFirstDayOfMonth(yourDate, weekDay):
    day=weekDays.index(weekDay)

    while datetime.weekday(yourDate) != day:
        yourDate = yourDate + timedelta(days = 1)

    return yourDate

weekDay="Friday"
your_date=date(2018,4,1)

print(getFirstDayOfMonth(your_date,weekDay))

print(getLastDayOfMonth(your_date,weekDay))

-------------------------------------------------------------------------------------------------------------------------

#C:\CASTMS\Delivery\data\{b4d0375b-730b-4f1b-bb70-ae43b183cfaf}\index.xml
import os
import re
import sys
import xml.etree.ElementTree as ET
from datetime import date, datetime
from xml.sax.saxutils import escape

fileName=r'C:\CASTMS\Delivery\data\{b4d0375b-730b-4f1b-bb70-ae43b183cfaf}\index.xml'

class DMTVersion(object):
    def __init__(self):
        self.uuid=''
        self.versionDate=date.today()
        self.versionName=''

def parser():
    tree=ET.parse(fileName)
    root=tree.getroot()
    uuidList=[]
    for child in root.iter('entry'):
        if 'uuid' in child.attrib["key"]:
            print(child.text)
            uuidList.append(child.text)

    dmtVersionList=[]

    for uuid in uuidList:
        dmtVersion = DMTVersion()
        dmtVersion.uuid=uuid
        for child in root.iter('entry'):
            if uuid in child.attrib["key"]:
                if 'date' in child.attrib["key"]:
                    dmtVersion.versionDate=datetime.strptime(child.text, "%Y-%m-%d %H:%M:%S")
                if 'name' in child.attrib["key"]:
                    dmtVersion.versionName=child.text
        dmtVersionList.append(dmtVersion)

    sortedDmtVersionList=sorted(dmtVersionList, key=lambda dmtVersion:dmtVersion.versionDate)
    print("After sorting")
    printDmtVersionList(sortedDmtVersionList)

    startDate=datetime(2018,4,18,0,0,0)
    endDate = datetime(2018, 4, 20,23,59,59)

    rangeVersionList=getVersionsBetween(startDate,endDate,dmtVersionList)

    print("Extracted Range")

    printDmtVersionList(rangeVersionList)

    getArchiveDeliveryCommands(rangeVersionList)

def printDmtVersionList(dmtVersionList):
    for version in dmtVersionList:
        print(version.versionName + "|" + str(version.versionDate))
        print("-" * 20)


def getVersionsBetween(startDate, endDate, sortedDmtVersionList):
    tempList=[]
    for dmtVersion in sortedDmtVersionList:
        if startDate <= dmtVersion.versionDate <= endDate:
            tempList.append(dmtVersion)
    return tempList

def getArchiveDeliveryCommands(dmtVersionsList):
    dmtDeliveryToolPath=r'"S:\Program Files\CAST\8.2\DeliveryManagerTool\DeliveryManagerTool-CLI.exe" '
    application='"ADCIS 1655"'
    dmtDeliveryFolder=r'"R:\DMTDelivery823"'
    appGuid='53d90a61-2c36-4f17-b6c0-cadade205764'
    logFilePath='"I:\TEMP\BVN\TO_DELETE\Logs\ADCIS 1655_1507_archiveDMT.log"'
    strCommandString=''
    for version in dmtVersionsList:
        strCommandString = dmtDeliveryToolPath + 'ArchiveDelivery -application '+application+' -version "'+version.versionName+'" -storagePath '+dmtDeliveryFolder+ ' -oneApplicationMode ' +appGuid+' -logFilePath '+logFilePath
        print(strCommandString)

parser()
--------------------------------------------------------------------------------------------------

#import configparser
#
#
#
#parser = configparser.RawConfigParser()
#
#parser.read(r'C:\Users\VRE\Documents\Tooling\TechnoFramework\technoFramework.properties')
#parser._interpolation = configparser.ExtendedInterpolation()
#
#for section_name in parser.sections():
#    print('Section:', section_name)
#    print('  Options:', parser.options(section_name))
#    for name, value in parser.items(section_name):
#        print('  %s = %s' % (name, value))
#    print()

#print(parser.get('connectionSettings', 'WORKSPACE'))

##############################################################################################
import xlsxwriter

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('Expenses02.xlsx')

worksheet = workbook.add_worksheet('MySheet')
worksheet1 = workbook.add_worksheet('MySheet2')

# Add a bold format to use to highlight cells.
#cell_format = workbook.add_format({'bold': True, 'font_color': 'red'})
format = workbook.add_format({'bold': True, 'font_color':'white'})
format.set_pattern(1)
format.set_bg_color('green')


# Add a number format for cells with money.
money = workbook.add_format({'num_format': '$#,##0'})

# Write some data headers.
worksheet.write('A1', 'Item', format)
worksheet.write('B1', 'Cost', format)

# Some data we want to write to the worksheet.
expenses = (
 ['Rent', 1000],
 ['Gas',   100],
 ['Food',  300],
 ['Gym',    50],
)

# Start from the first cell below the headers.
row = 1
col = 0

# Iterate over the data and write it out row by row.
for item, cost in (expenses):
 worksheet.write(row, col,     item)
 worksheet.write(row, col + 1, cost, money)
 row += 1

# Write a total using a formula.
worksheet.write(row, 0, 'Total',       format)
worksheet.write(row, 1, '=SUM(B2:B5)', money)

cell_format = workbook.add_format()
cell_format.set_rotation(45)
cell_format.set_bg_color('#FF5444')
worksheet.set_tab_color('orange')
worksheet.write(9, 9, 'This text is rotated', cell_format)

worksheet1.set_tab_color('blue')
worksheet2 = workbook.add_worksheet('MySheet3')
worksheet2.set_tab_color('green') # set worksheet tab color as green
worksheet2.write(2,2, 'Hello World')
#make worksheet2 as the active sheet
worksheet2.activate()


workbook.close()

"""



