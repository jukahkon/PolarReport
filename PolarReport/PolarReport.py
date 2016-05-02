#
# Main program
#
import FindReports
import ReportReader
import DatabaseHandler
import os
from Settings import PathSettings

dir = PathSettings['ReportDirectory']
files = FindReports.enumReports(dir if dir else os.getcwd())

for file in files:
    report = ReportReader.readReport(file)
    DatabaseHandler.connectToDatabase(PathSettings['Database'])
    DatabaseHandler.insertToDatabase(report)

