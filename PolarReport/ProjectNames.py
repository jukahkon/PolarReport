import sqlite3
import os
from Settings import PathSettings

Projects = {
    '0000' : 'Lomat, poissaolot',
    '1601' : 'Tarvikemyynti',
    '1602' : 'Pienet tyot',
    '1603' : 'Outokumpu vuosisopimustyot',
    '1604' : 'Tornion Energia',
    '1605' : 'Konecranes tyot',
    '1606' : 'Efora Veitsiluoto Vianhaku ja ohjelmamu',
    '1607' : 'Kemin Energia tyot 2016',
    '1608' : 'Tapojarvi kaytonjohtajatyot',
    '1609' : 'AEF kotelot',
    '1610' : 'Terrafame kotelot',
    '1611' : 'Metsa board, bms tyot',
    '1612' : 'Tapojarvi tyot 2016',
    '1613' : 'Terrafame SK-siltakuljettimen ja PVA1',
    '1614' : 'Dragon Mining',
    '1615' : 'Kaasumittari kalibroinnit',
    '1630' : 'ELY, Kv-hanke',
    '1631' : 'SMA Mineral Oy',
    '1632' : 'Metsa Fibre',
    '1633' : 'Terrafame PKP10 siirto',
    '1696' : 'Hallinto',
    '1697' : 'Koulutus',
    '1698' : 'Sisaiset tyot',
    '1699' : 'Tarjouslaskenta',

    '1501' : 'Tornion Energia, Tyot 2015',
    '1502' : 'Pienet tyot',
    '1503' : 'Outokumpu Stainless Oy, suunnittelutyot',
    '1504' : 'Tormets',
    '1505' : 'Outokumpu Stainless Oy Vuosisopimus',
    '1506' : 'Konecranes Finland Oy',
    '1507' : 'Talvivaara lastausssuppilo 1 ja 2 muuto',
    '1508' : 'TEVO Lokomo Oy',
    '1509' : 'Vartiuksen Raja-aseman hankintasuunnitt',
    '1510' : 'Imtech SSAB',
    '1511' : 'Metsa Fibre Oy , BMS, Caverion',
    '1512' : 'Stora Enso Kemi, Efora Kemi',
    '1513' : 'Stora Enso Oulu , Efora Oulu',
    '1514' : 'AGA Apu 1 HLL',
    '1515' : 'Tapojarvi, 2015 tyot',
    '1516' : 'Talvivaara, keskikaistan siirto',
    '1517' : 'Talvivaara, primaarin tripperi',
    '1518' : 'AEF Suunnittelu ja Ohjelmointi',
    '1519' : 'AEF Kotelot',
    '1520' : 'Talvivaara Kotelot',
    '1521' : 'Metsa Board, BMS Kemi',
    '1522' : 'Kemin Energia Tyot 2015',
    '1523' : 'Lapin AMK Jaloterasstudio',
    '1524' : 'Terrafame koulutus',
    '1525' : 'Outotec Lihir',
    '1526' : 'Metsa Board, BMS, ohjauspulpetit',
    '1527' : 'Outotec Iran',
    '1528' : 'Kaasumittari kalibroinnit',
    '1529' : 'Tiltek Marine',
    '1530' : 'Terrafame kenttakotelot SK-keskikaistan',
    '1598' : 'Sisaiset tyot',
    '1599' : 'Tarjouslaskenta'
}

# DB connection
conn = None

def connectToDatabase(fileName):
    global conn
     
    if os.path.isfile(fileName):
        conn = sqlite3.connect(fileName)

def updateProjectNames():
    if not conn:
        return

    for (id, name) in Projects.iteritems():
        row = conn.execute("SELECT id FROM Projekti WHERE tyonumero=?", (id,)).fetchone()
    
        if row:
            prj_id = row[0]
            query = "UPDATE Projekti SET tyonimi=? WHERE tyonumero=?"
            conn.execute(query, (name,id,))
            print "Paivitettiin projektinimi: " + id + " " + name

    conn.commit()

if __name__ == "__main__":
    connectToDatabase(PathSettings['Database'])

    updateProjectNames()

    conn.close()
