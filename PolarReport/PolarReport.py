#
# Main program
#
import FindReports
import ReportReader
import DatabaseHandler

reportDir = 'c:\Projects\PolarReport\PolarReport\TestData'
db_file = 'c:\projects\polarreport.sqlite'

files = FindReports.enumReports(reportDir)

for file in files:
    report = ReportReader.readReport(file)
    DatabaseHandler.connectToDatabase(db_file)
    DatabaseHandler.insertToDatabase(report)

