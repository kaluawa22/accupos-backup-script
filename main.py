# Kalu Awa

import pyodbc
import csv
import os
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook

# Function for counting number of entries for logging
def csvNumEntry(filename):
    count = 0
    with open(filename) as f:
        cr = csv.reader(f)
        for data in cr:
            count += 1
    # return cnt - 1 to take account for the first row containing the fields
    return count - 1


# Python module to connect to the Database
inputDir = (r"C:\ScheduledJobs\ISBNSales")
conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=WC-ACCUSERVER\SQLEXPRESS;'
                      'Database=accupos;'
                      'Trusted_Connection=yes;')
crsr = conn.cursor()


# string containing current date to be used as filename (YEAR-MONTH-DAY)
# fileName = datetime.now().strftime("%Y_%m_%d_%I_%M_%S_%p" + ".csv")
fileName = datetime.now().strftime("%Y_%m_%d_%I_%M_%S_%p")
# Declaring relative file path
# HERE = Path(__file__).parent.resolve()


# Setting up Path for CSV File output
# PATH = HERE / fileName

save_path = r"C:\ScheduledJobs\ISBNSales\misc"
directory = os.path.join(save_path, fileName)


# SQL Query Script
sql = """\
declare @lastweek datetime
declare @now datetime
set @now = getdate()
set @lastweek = dateadd(day, -7, @now)
SELECT lines.ItemID as isbn ,lines.Quantity as qty ,items.[item description] ,head.DateEntered
FROM [accupos].[dbo].[apcshead] head ,[accupos].[dbo].[apcsitem] lines ,[accupos].[dbo].[apinms] items ,[accupos].[dbo].[aptill]
where head.DateEntered between @lastweek and @now and lines.HeadKey = head.[Key] and lines.ItemID = items.[Item Id] and head.Ztill = aptill.till and items.[Item Type] in ('Books', 'Bibles') and len(items.[Item Id]) in (10,13) and lines.Ext <> 0
"""
# Ouput rows for CSV
rows = crsr.execute(sql)

with open(directory, 'w', newline='') as csvfile:
    writer = csv.writer(csvfile)
    writer.writerow([x[0] for x in crsr.description])
    for row in rows:
        writer.writerow(row)

# Getting CSV file path for reading and file manipulation
csvFilePath = save_path + os.sep + fileName
reader = csv.reader(csvFilePath)

# first row in csv file cointains the columns of the database so I subtract 1// old code
# numLines = len(list(reader)) - 1
# print(numLines)


# Getting current date and time for logging
now = datetime.now()
today = now.strftime("at %I:%M %p on %m/%d/%Y")


# Tuple containing variables that will be used for logging
logStatement = ((csvNumEntry(csvFilePath)), 'Entries recorded', today)

# Statement to write to log file
with open('log.txt', 'a') as new_file:
    logFile = map(str, logStatement)
    new_file.write(" ".join(logFile) + "\n")

#Code to convert created CSV File to xlsx file to be opened in Excel

wb = Workbook()
ws = wb.active

with open(csvFilePath, 'r') as f:
    for row in csv.reader(f):
        ws.append(row)
    wb.save(r"C:\ScheduledJobs\ISBNSales\WeeklyReports" + '\\' + fileName + '.xlsx')



print('Query Complete!')




