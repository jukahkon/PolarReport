import sqlite3
import os.path
import sys
import datetime

# DB connection
conn = None

def connectToDatabase(fileName):
    global conn
     
    init = False
  
    if not os.path.isfile(fileName):
        init = True

    conn = sqlite3.connect(fileName)
    
    if init:
        initializeDb()

def initializeDb():
    c = conn.cursor()

    c.execute('CREATE TABLE Henkilo ("id" INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,\
                                     "nimi" TEXT NOT NULL,\
                                     "sposti" TEXT NOT NULL)')
    
    c.execute('CREATE TABLE Projekti ("id" INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,\
                                      "numero" INTEGER NOT NULL,\
                                      "kuvaus" TEXT DEFAULT "")')
    
    c.execute('CREATE TABLE Tyosuorite ("id" INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,\
                                        "henkilo_id" INTEGER NOT NULL,\
                                        "projekti_id" INTEGER NOT NULL,\
                                        "tyoselite_id" INTEGER NOT NULL,\
                                        "tilaus" TEXT DEFAULT "",\
                                        "suorituspaiva" DATE NOT NULL,\
                                        "norm" REAL DEFAULT 0.0,\
                                        "lisatyo" REAL DEFAULT 0.0,\
                                        "ylit50" REAL DEFAULT 0.0,\
                                        "ylit100" REAL DEFAULT 0.0,\
                                        "ylit150" REAL DEFAULT 0.0,\
                                        "ylit200" REAL DEFAULT 0.0,\
                                        "viikkoylit" REAL DEFAULT 0.0,\
                                        "matkat" REAL DEFAULT 0.0,\
                                        "km" REAL DEFAULT 0.0,\
                                        "ateria" INTERGER DEFAULT 0,\
                                        "osapaivar" INTERGER DEFAULT 0,\
                                        "paivaraha" INTERGER DEFAULT 0)')
        
    c.execute('CREATE TABLE Tyoselite ("id" INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,\
                                       "selite" TEXT NOT NULL)')


def insertToDatabase(report):
    if not conn:
        print "Virhe: ei tietokanta yhteytta"
        return

    c = conn.cursor()
    count = 0

    # Henkilot
    row = c.execute("SELECT id FROM Henkilo WHERE sposti=?", (report['sposti'],)).fetchone()
    emp_id = None

    if not row:
        c.execute("INSERT INTO Henkilo (nimi, sposti) VALUES (?, ?)", (report['henkilo'], report['sposti'],))
        emp_id = c.lastrowid
    else:
        emp_id = row[0]

    entries = report['entries']
    
    for entry in entries:
        # Projektit
        row = c.execute("SELECT id FROM Projekti WHERE numero=?", (entry['projekti_no'],)).fetchone()
        prj_id = None
        
        if not row:
            c.execute("INSERT INTO Projekti (numero, kuvaus) VALUES (?, ?)", (entry['projekti_no'], '',))
            prj_id = c.lastrowid
        else:
            prj_id = row[0]

        
        # Tyoselite
        row = c.execute("SELECT id FROM Tyoselite WHERE selite=?", (entry['tyoselite'],)).fetchone()
        desc_id = None
        
        if not row:
            c.execute("INSERT INTO Tyoselite (selite) VALUES (?)", (entry['tyoselite'],))
            desc_id = c.lastrowid
        else:
            desc_id = row[0]


        # Tyosuoritteet
        date = entry['suorituspaiva'].date()
        order = entry['tilaus'] if ('tilaus' in entry) else ''
        row = c.execute("SELECT id FROM Tyosuorite WHERE henkilo_id=? AND projekti_id=? AND\
                         tyoselite_id=? AND suorituspaiva=? AND tilaus=?",\
                        (emp_id, prj_id, desc_id, date, order)).fetchone()

        if not row:
            # mandatory fields
            row_data = []
            row_data.append(emp_id)
            row_data.append(prj_id)
            row_data.append(desc_id)
            row_data.append(date)
            
            c.execute("INSERT INTO Tyosuorite (henkilo_id, projekti_id, tyoselite_id, suorituspaiva) VALUES (?,?,?,?)", row_data)
            entry_id = c.lastrowid

            # optional fields
            fields = ['norm', 'lisatyo', 'ylit50', 'ylit100', 'ylit150', \
                      'ylit200', 'viikkoylit', 'matkat', 'km', 'ateria', \
                      'osapaivar', 'paivaraha', 'tilaus']

            for i in range(0,len(fields)):
                if fields[i] in entry:
                    val = entry[fields[i]]
                    query = "UPDATE Tyosuorite SET {cn}=? WHERE id=?".format(cn=fields[i])
                    c.execute(query, (val,entry_id,))
            
            count += 1

    conn.commit()

    print "Lisattiin tietokantaan " + str(count) + " tyosuoritetta!"

    return

# TEST
if __name__ == "__main__":
    db_file = "c:\\projects\\test.sqlite"

    connectToDatabase(db_file)
    
    # Test report
    report = {
        'raportointi_pvm' : datetime.datetime(2016,4,2),
        'henkilo' : 'Mikki Hiiri',
        'sposti' : 'mikki.hiiri@disney.com',
        'entries' : [
            { 'projekti_no' : 1234,
              'tilaus' : '5555',
              'suorituspaiva' : datetime.datetime(2016,3,1),
              'tyoselite' : 'Programming',
              'norm' : 7.5,
              'km' : 25,
              'ateria' : 1 },
            { 'projekti_no' : 1234,
              'tilaus' : '6666',
              'suorituspaiva' : datetime.datetime(2016,3,2),
              'tyoselite' : 'Testing',
              'norm' : 5
            },
            { 'projekti_no' : 1234,
              'tilaus' : '7777',
              'suorituspaiva' : datetime.datetime(2016,3,2),
              'tyoselite' : 'Testing',
              'norm' : 5
            },
            { 'projekti_no' : 4321,
              'tilaus' : '5555',
              'suorituspaiva' : datetime.datetime(2016,3,3),
              'tyoselite' : 'Debugging',
              'norm' : 0.4,
              'lisatyo' : 0.5, 
              'ylit50' : 0.6,
              'ylit100' : 0.7,
              'ylit150' : 0.8,
              'ylit200' : 0.9,
              'viikkoylit' : 1.1,
              'matkat' : 1.2,
              'km' : 100.5,
              'ateria' : 0,
              'osapaivar' : 0,
              'paivaraha' : 1
            }
        ]
    }

    insertToDatabase(report)

    conn.close()

