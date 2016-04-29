#
# Monthly report parsing
#
import re
import datetime
import sys, errno
from pprint import pprint

try:
    import openpyxl
except:
    exit("Virhe: openpyxl moduulia ei loydy")

COVERSHEETNAME = "Yleistiedot"
DATASHEETNAME = "Kuukausi-ilmoitus"
COVER_DATA_COL = 4
COVER_DATE_ROW = 5
COVER_NAME_ROW = 7
COVER_EMAIL_ROW = 8
PROJECT_NUMBER_COL = 2
WORK_DESC_COL = 4
ORDER_NUMBER_COL = 6
DATE_ROW = 7

def readReport(fileName):
    wb = openExcelWorkbook(fileName)
    return extractData(wb)

def openExcelWorkbook(fileName):
    wb = None

    try:
        print "Avataan Excel-tiedosto: " + fileName
        wb = openpyxl.load_workbook(fileName, data_only=True)
    except:
        exit("Virhe: Tiedoston avaaminen epaonnistui: " + fileName)
    else:
        return wb

def extractData(wb):
    if not wb:
        return

    data = {}
    sheet = wb.get_sheet_by_name(COVERSHEETNAME)
    if not sheet:
        print "Virhe: yleistiedot valilehtea ei loydy"
    
    data['raportointi_pvm'] = sheet.cell(row=COVER_DATE_ROW, column=COVER_DATA_COL).value
    data['henkilo'] = sheet.cell(row=COVER_NAME_ROW, column=COVER_DATA_COL).value
    data['sposti'] = sheet.cell(row=COVER_EMAIL_ROW, column=COVER_DATA_COL).value

    sheet = wb.get_sheet_by_name(DATASHEETNAME)
    if not sheet:
        print "Virhe: tunti-ilmoitus valilehtea ei loydy"

    project_rows = indexProjects(sheet)

    date_cols = indexDates(sheet)

    data['entries'] = createWorkEntries(sheet, project_rows, date_cols)
        
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

def createWorkEntries(sheet, rows, cols):
    entries = []

    for i in rows:
        project_number = sheet.cell(row=i, column=PROJECT_NUMBER_COL).value
        order_number = sheet.cell(row=i, column=ORDER_NUMBER_COL).value
        work_description = sheet.cell(row=i, column=WORK_DESC_COL).value

        if not project_number:
            print "Virhe: tyonumero puuttuu, rivi: " + str(i)
        
        for j in cols:
            entry = {}
            
            fields = ['norm', 'lisatyo', 'ylit50', 'ylit100', 'ylit150', \
                      'ylit200', 'viikkoylit', 'matkat', 'km', 'ateria', \
                      'osapaivar', 'paivaraha' ]

            for k in range(0, len(fields)):
                val = sheet.cell(row=i+k, column=j).value
                if val:
                    entry[fields[k]] = val

            if entry:
                entry['projekti_no'] = project_number
                entry['tilaus'] = order_number
                entry['suorituspaiva'] = sheet.cell(row=DATE_ROW, column=j).value
                entry['tyoselostus'] = work_description
                
                entries.append(entry)

    return entries


# TEST
if __name__ == "__main__":
    file = ".\\TestData\\2016\\March\\report_1.xlsx"

    wb = openExcelWorkbook(file)
    
    data = extractData(wb)
    pprint(data)
