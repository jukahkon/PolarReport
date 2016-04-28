import sqlite3
import os.path
import sys
import datetime

# Current DB connection
conn = None

def connectToDatabase(fileName):
    global conn

    emptyDb = False
    if not os.path.isfile(fileName):
        emptyDb = True

    conn = sqlite3.connect(fileName)
    
    return emptyDb

def createTables():
    c = conn.cursor()

    c.execute('CREATE TABLE "main"."Employee" ("id" INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,\
                                               "name" TEXT, "email" TEXT)')
    c.execute('CREATE TABLE "main"."Project" ("id" INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,\
                                              "number" INTEGER NOT NULL, "description" TEXT)')
    c.execute('CREATE TABLE "main"."Orderx" ("id" INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, \
                                            "number" TEXT NOT NULL, "description" TEXT,\
                                            "project_id" INTEGER)')
    c.execute('CREATE TABLE "main"."WorkDescription" ("id" INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, \
                                                      "description" TEXT)')
    c.execute('CREATE TABLE "main"."WorkEntry" ("id" INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,\
                                                "employee_id" INTEGER, "project_id" INTEGER, "order_id" INTEGER,\
                                                "norm" REAL, "ext50" REAL, "ot50" REAL, "ot100" REAL, "ot150" REAL,\
                                                "ot200" REAL, "otwk" REAL, "travel" REAL, "km" REAL, "meal" INTERGER,\
                                                "allow50" INTERGER, "allow100" INTERGER)')

def insertReportToDatabase(report):
    c = conn.cursor()

    # Insert Employee, if not exists
    c.execute("SELECT id FROM Employee WHERE email=?", (report['employee_email'],))
    emp_id = c.fetchone()

    if not emp_id:
        c.execute("INSERT INTO Employee (name, email) VALUES (?, ?)", (report['employee_name'], report['employee_email'],))

    entries = report['entries']
    
    for entry in entries:
        # Insert Project, if not exists
        c.execute("SELECT id FROM Project WHERE number=?", (entry['project_num'],))
        prj_id = c.fetchone()

        if not prj_id:
            c.execute("INSERT INTO Project (number, description) VALUES (?, ?)", (entry['project_num'], '',))

        # Insert order, if not exists
        if entry['order_num']:
            c.execute("SELECT id FROM Orderx WHERE number=?", (entry['order_num'],))
            ord_id = c.fetchone()

            if not ord_id:
                c.execute("INSERT INTO Orderx (number, description) VALUES (?, ?)", (entry['order_num'], '',))

        # Insert work description
        c.execute("INSERT INTO WorkDescription (description) VALUES (?)", (entry['work_desc'],))

    conn.commit()

    return

# TEST
if __name__ == "__main__":
    db_file = "c:\\projects\\test.sqlite"

    empty = connectToDatabase(db_file)
    if empty:
        createTables()

    # Test report
    report = {
        'report_date' : datetime.datetime(2016,4,2),
        'employee_name' : 'Mikki Hiiri',
        'employee_email' : 'mikki.hiiri@ankkalinna.com',
        'entries' : [
            { 'project_num' : 1234,
              'order_num' : '5555',
              'work_date' : datetime.datetime(2016,3,1),
              'work_desc' : 'Programming',
              'norm' : 7.5,
              'km' : 25,
              'meal' : 1 },
            { 'project_num' : 1234,
              'order_num' : '6666',
              'work_date' : datetime.datetime(2016,3,2),
              'work_desc' : 'Testing',
              'norm' : 5
            }
        ]
    }

    insertReportToDatabase(report)

    conn.close()

