#
# Main program
#
import ReportFinder
import ReportReader
import DatabaseHandler
import os
from Settings import PathSettings

dir = PathSettings['ReportDirectory']
files = ReportFinder.enumReports(dir if dir else os.getcwd())

DatabaseHandler.connectToDatabase(PathSettings['Database'])

for file in files:
    report = ReportReader.readReport(file)
    DatabaseHandler.insertToDatabase(report)

