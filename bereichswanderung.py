# Input: Reporter-Abfrage "Alle Zu- und Abgänge in allen Chören"
#        in Zwischenablage, dann dieses Skript starten

import csv
import io
import win32clipboard
import datetime
from collections import defaultdict

# Ein paar Werte vordefinieren

heute = datetime.datetime.strftime(datetime.datetime.today(), "%Y-%m-%d_%H%M%S")
startjahr = 2017
endejahr = 2022

chor = 0    # Positionen der Werte in der Datentabelle
von = 1
bis = 2

print("Bitte Reporter-Abfrage 'Alle Zu- und Abgänge in allen Chören'")
print("durchführen und Daten in Zwischenablage ablegen.")
input("Bitte ENTER drücken, wenn dies geschehen ist!")

win32clipboard.OpenClipboard()
data = win32clipboard.GetClipboardData()
win32clipboard.CloseClipboard()

if not data.startswith("Person (Nr)\t"):
    print("Fehler: Unerwarteter Inhalt der Zwischenablage!")
    exit()

daten = defaultdict(list)

# Erst mal alle Daten sammeln und den Personen zuordnen (damit die einzelnen Wege durch die Chöre nachvollziehbar werden)

with io.StringIO(data) as infile:
    for person in csv.DictReader(infile, delimiter="\t"):
        von_datum = datetime.datetime.strptime(person["Von"], "%d.%m.%Y") if person["Von"] else None
        bis_datum = datetime.datetime.strptime(person["Bis"], "%d.%m.%Y") if person["Bis"] else None
        daten[person["Person (Nr)"]].append((person["Bereich"], von_datum, bis_datum))

verlauf = []

for person, stationen in daten.items():
    status = {}
    jahr = startjahr-1 # Zustand am Ende des Vorjahres des Auswertezeitraums gesondert ermitteln
    status[jahr] = None
    for station in stationen:
        if station[von] <= datetime.datetime(jahr, 12, 31) and (station[bis] is None or station[bis] > datetime.datetime(jahr, 12, 31)):
            status[jahr] = station[chor]
            break 
    # Jetzt die Jahre ab Startdatum
    for jahr in range(startjahr, endejahr+1):
        status[jahr] = None
        for station in stationen:
            if station[von] <= datetime.datetime(jahr, 1, 1) and (station[bis] is None or station[bis] > datetime.datetime(jahr, 1, 1)):
                status[jahr] = station[chor]
                break
    
    # Jetzt die Werte im Verlauf abgleichen
    pausiert = False    # Der Zustand "pausiert" wird erst erreicht, wenn einmal aktiv gewesen
    for jahr in range(startjahr, endejahr):
        if status[jahr] is None and status[jahr+1] is not None:       # Wechsel in neuen Chor
            verlauf.append([jahr, person, "Pause" if pausiert else "Extern", status[jahr+1]])
            pausiert = False
        elif status[jahr] is not None and status[jahr+1] is None:     # aktuellen Chor verlassen
            if jahr <= endejahr-2 and status[jahr+2] is not None:
                pausiert = True
            verlauf.append([jahr, person, status[jahr], "Pause" if pausiert else "Extern"])
        elif status[jahr] is not None and status[jahr+1] is not None: # bleibt im USC, ggf. anderer Chor
            verlauf.append([jahr, person, status[jahr], status[jahr+1]])
            
            
feldnamen = ["Jahr", "Nr", "Anfang", "Ende"]   

with open(f"Mitgliederwanderung_{heute}.csv", mode="w", newline="", encoding="cp1252") as outfile:
    output = csv.writer(outfile, delimiter=";")
    output.writerow(feldnamen)
    output.writerows(verlauf)

print(f"Fertig! Die Datei Mitgliederwanderung_{heute}.csv wurde im aktuellen Ordner abgelegt.")
