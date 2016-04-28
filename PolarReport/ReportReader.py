#
# Monthly report reader
#
import re
import datetime
import sys, errno
from pprint import pprint

try:
    import openpyxl
except:
    exit("ERROR: Importing openpyxl failed")

COVERSHEETNAME = "Yleistiedot"
DATASHEETNAME = "Kuukausi-ilmoitus"
COVER_DATA_COL = 4
COVER_DATE_ROW = 5
COVER_NAME_ROW = 7
COVER_EMAIL_ROW = 8
PROJECT_NUMBER_COL = 2
PROJECT_NAME_COL = 4
ORDER_NUMBER_COL = 6
DATE_ROW = 7

def openExcelWorkbook(fileName):
    try:
        print "Open excel workbook: " + fileName
        wb = openpyxl.load_workbook(fileName, data_only=True)
    except:
        exit("ERROR: Failed to open excel-workbook: " + fileName)
    else:
        return wb

def readData(wb):
    data = {}
    sheet = wb.get_sheet_by_name(COVERSHEETNAME)
    if not sheet:
        print "Virhe: yleistiedot valilehtea ei loydy"
    
    data['report_date'] = sheet.cell(row=COVER_DATE_ROW, column=COVER_DATA_COL).value
    data['employee_name'] = sheet.cell(row=COVER_NAME_ROW, column=COVER_DATA_COL).value
    data['employee_email'] = sheet.cell(row=COVER_EMAIL_ROW, column=COVER_DATA_COL).value

    sheet = wb.get_sheet_by_name(DATASHEETNAME)
    if not sheet:
        print "Virhe: tunti-ilmoitus valilehtea ei loydy"

    project_rows = indexProjects(sheet)
    #pprint(projs)

    date_cols = indexDates(sheet)
    #pprint(dates)

    entries = createWorkEntries(sheet, project_rows, date_cols)
    #pprint(entries)
    data['entries'] = entries
        
    return data

def indexProjects(sheet):
    rows = []
    
    for i in range(1,300):
        txt = sheet.cell(row=i, column=PROJECT_NUMBER_COL).value
        if txt:
            if re.match('\d{4}', unicode(txt)):
                rows.append(i)

    return rows

def indexDates(sheet):
    cols = []

    for i in range(1,50):
        dt = sheet.cell(row=DATE_ROW, column=i).value
        if dt and isinstance(dt, datetime.datetime):
            cols.append(i)

    return cols

def createWorkEntries(sheet,project_rows, date_cols):
    entries = []

    for i in project_rows:
        project_number = sheet.cell(row=i, column=PROJECT_NUMBER_COL).value
        order_number = sheet.cell(row=i, column=ORDER_NUMBER_COL).value
        work_description = sheet.cell(row=i, column=PROJECT_NAME_COL).value

        if not project_number:
            print "Virhe: projektinumero puuttuu"
        
        for j in date_cols:
            entry = {}
            
            fields = ['norm', 'ext50', 'ot50', 'ot100', 'ot150', \
                      'ot200', 'otwk', 'travel', 'km', 'meal', \
                      'allow50', 'allow100' ]

            for k in range(0, len(fields)):
                val = sheet.cell(row=i+k, column=j).value
                if val:
                    entry[fields[k]] = val

            if entry:
                entry['project_num'] = project_number
                entry['order_num'] = order_number
                entry['work_date'] = sheet.cell(row=DATE_ROW, column=j).value
                entry['work_desc'] = work_description
                
                entries.append(entry)

    return entries


# TEST
if __name__ == "__main__":
    file = ".\\TestData\\2016\\March\\report_JoeDoe.xlsx"

    wb = openExcelWorkbook(file)
    if wb:
        data = readData(wb)
        pprint(data)
