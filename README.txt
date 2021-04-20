# Retourenbericht-BHM
Erstellt aus den in der Zentrale generierten Retourenbericht-Dateien einen Retourenbericht, der auf dem Nadeldrucker (Tally) ausgedruckt werden kann.

''' <summary>
''' ### RETOURENBERICHT ERSTELLEN ###
''' Programm zum Erstellen der Druckdatei für den Retourenbericht.
''' Die Datei wird in Code 850 errzeugt und abgelegt und ist für den Druck auf dem Nadeldrucker gedacht.
''' Der Druck selbst muss aus einer VM erfolgen.
''' </summary>

    ' 1. Berichtsdateien der Filialen in Sicherungsordner kopieren
    ' 2. Berichtsdateien einlesen und in eine einzelne Datei konvertieren (Codepage 850/435)
    '       - Seitenumbrüche entfernen und nur zwischen den Einzelberichten einfügen.
    '       - Prüfen, ob in den Wochenblöcken (Mo-So) Leerzeilen sind. Wenn ja, entfernen.
    '       - Datei ablegen zum Drucken mit Nadeldrucker.
