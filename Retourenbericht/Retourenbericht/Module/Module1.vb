''' <summary>
''' ### RETOURENBERICHT ERSTELLEN ###
''' Programm zum Erstellen der Druckdatei für den Retourenbericht.
''' Die Datei wird in Code 850 errzeugt und abgelegt und ist für den Druck auf dem Nadeldrucker gedacht.
''' Der Druck selbst muss aus einer VM erfolgen.
''' </summary>

Module Module1
    ' 1. Berichtsdateien der Filialen in Sicherungsordner kopieren
    ' 2. Berichtsdateien einlesen und in eine einzelne Datei konvertieren (Codepage 850/435)
    '       - Seitenumbrüche entfernen und nur zwischen den Einzelberichten einfügen.
    '       - Prüfen, ob in den Wochenblöcken (Mo-So) Leerzeilen sind. Wenn ja, entfernen.
    '       - Datei ablegen zum Drucken mit Nadeldrucker.

    '## ToDos
    '# Filialliste variabel machen!
    '   - Über Parameterstart (/setup): Fenster für Konfiguration, Liste ausgeben/einlesen (txt), Einträge hinzufügen
    '   - Sortierreihenfolge bearbeiten

    Dim Arguments() As String = Environment.GetCommandLineArgs  'Kommandozeilenparameter aufnehmen

    Dim conf As New Config
    Dim arrKomplett() As String
    Dim WeekDateiliste As New ArrayList
    Dim arrSortierliste() As Integer = {815, 830, 850, 855, 860, 420, 285, 365, 425, 205, 315, 460, 465,
                                       215, 230, 295, 740, 390, 360, 350, 265, 250, 411, 395, 200, 195,
                                       210, 300, 490, 430, 705, 750, 755, 401, 345, 310, 735, 405, 280,
                                       435, 455, 445, 240, 190, 725, 785, 340, 370, 235, 180, 775, 780,
                                       225, 270, 255, 245, 730, 220, 320, 325, 335, 375, 770, 305}

    Sub Main()

        'TODO Wenn Startparameter = /setup , dann Einrichtungsskripte starten (Filialliste, Reihenfolge)
        'Nach Einrichtungsskripten Programm beenden. Neustart erforderlich

        If Arguments.Count > 1 Then
            ArgumenteVerarbeiten(Arguments)
            Application.Exit()
        End If

        '### Programmparameter festlegen
        If My.Settings.SortierListeFilialen = "" Then
            'arrSortierliste = conf.getFilialliste

        End If

        conf.Quellordner = "H:\Jacoby\Neue\Daten\"
        'conf.Quellordner = IO.Directory.GetCurrentDirectory & "\Dateien\"  'zum Testen
        conf.Sicherungsordner = IO.Directory.GetCurrentDirectory
        conf.NameAusgabedatei = "week.txt"
        conf.SetCodepageDateiEinlesen(1252)
        conf.SetCodepageDateiAusgeben(1252)

        'Redimesionieren von arrKomplett für die Filialen
        ReDim arrKomplett(arrSortierliste.GetUpperBound(0))

        '### Ablauf
        Console.WriteLine("Erstellen der Retourendatei gestartet...")

        WeekDateiliste = DateilisteErstellen(conf.Quellordner, "weekr") 'Erstellt ein ArrayList mit allen weekr*.* Dateien im Quellverzeichnis

        Console.WriteLine(" - Dateien kopieren von " & conf.Quellordner & " nach " & conf.Sicherungsordner & " ...")
        WeekDateienSichern(conf.Sicherungsordner)

        Console.WriteLine(" - Dateien zusammenfassen...")
        WeekDateienZusammenfassen(conf.Sicherungsordner)

        'WeekDateienAbspeichern(conf.Sicherungsordner)
        Console.WriteLine(" - Datei " & conf.NameAusgabedatei & " abspeichern...")
        WeekDateienAbspeichern2(conf.Sicherungsordner)

        Console.WriteLine()
        Console.WriteLine("Erstellung der Retourendatei abgeschlossen.")

    End Sub

    ''' <summary>
    ''' Kopiert alle "weekr*.*" Dateien aus dem Quell- in ein Zielverzeichnis. Gibt true zurück, wenn keine Fehler.
    ''' </summary>
    ''' <param name="Ziel">Zielverzeichnis</param>
    Sub WeekDateienSichern(Ziel As String)
        Try
            For Each Datei In WeekDateiliste
                If InStr(Datei.ToString, "WEEKR") > 0 Then
                    'MsgBox("copy " & Datei.ToString & " nach " & Ziel & IO.Path.GetFileName(Datei.ToString))
                    FileIO.FileSystem.CopyFile(Datei.ToString, Ziel & IO.Path.GetFileName(Datei.ToString), overwrite:=True) 'Kopieren mit überschreiben
                End If
            Next
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' Fasst die "weekr*.*" Dateien im Verzeichnis zusammen und speichert sie als "week.txt" im gleichen Verzeichnis ab. 
    ''' Gibt True zurück, wenn keine Fehler. Sonst False.
    ''' </summary>
    ''' <param name="Verzeichnis">Verzeichnis mit den "weekr*.*" Dateien.</param>
    Sub WeekDateienZusammenfassen(Verzeichnis As String)
        'Jede Datei wird zeilenweise in ein Array eingelesen und auf Leerzeilen und Umbrüche geprüft (entfernt).
        'Anschließend werden alle Zeilen in ein Gesamtarray überführt.
        'Nach jeder Datei wird ein Seitenumbruch eingefügt.
        WeekDateiliste = DateilisteErstellen(conf.Sicherungsordner, "weekr")
        For Each Datei In WeekDateiliste
            Dim Filialnummer As Integer
            Dim arrFilialdatei As New ArrayList
            Dim FileText As String = ""

            arrFilialdatei.Add(vbCrLf)

            'Filialnummer aus Dateiname ermitteln
            If Val(Mid(IO.Path.GetFileNameWithoutExtension(Datei.ToString), 6)) <> 0 Then
                'Stop
                Filialnummer = CInt(Mid(IO.Path.GetFileNameWithoutExtension(Datei.ToString), 6))
                Using sr As New IO.StreamReader(Datei.ToString, conf.CodepageDateiEinlesen)
                    'Datei zeilenweise einlesen OHNE Seitenumbrüche (vbFormFeed)
                    Dim srLine As String

                    While Not sr.EndOfStream
                        srLine = sr.ReadLine
                        Dim chars(srLine.Length - 1) As Char
                        If Not srLine = vbFormFeed Then
                            chars = srLine.ToCharArray

                            For n = 0 To chars.GetUpperBound(0)
                                Select Case Asc(chars(n))
                                    Case = 225 'ß HEXE1
                                        chars(n) = Chr(223) 'HEXDF
                                    Case = 81 'ü
                                        chars(n) = Chr(252)
                                    Case = 148 'ö HEX94
                                        chars(n) = Chr(246) 'HEXF6
                                    Case = 132 'ä HEX84
                                        chars(n) = Chr(228) 'HEXE4
                                    Case = 196 '- HEXC4
                                        chars(n) = Chr(45) 'HEX2D
                                End Select
                            Next n
                            'Stop
                            srLine = chars & vbCrLf

                            FileText += srLine
                        End If
                    End While
                    'FileText += vbFormFeed 'Am Ende der Dateizeilen einen Seitenumbruch einfügen (vbFormFeed)
                End Using

                'arrFilialdatei an der entsprechenden SortierStelle des arrkomplett speichern
                For i = 0 To arrSortierliste.GetUpperBound(0)
                    If CInt(Filialnummer) = arrSortierliste(i) Then
                        arrKomplett(i) = FileText
                        Exit For
                    End If
                Next

            End If
        Next Datei

    End Sub

    ''' <summary>
    ''' Speichert den Inhalt aller weekr*.* Dateien in einer Textdatei "week.txt" ab. 
    ''' Die Datei wird 
    ''' </summary>
    ''' <param name="Zielverzeichnis"></param>
    Sub WeekDateienAbspeichern2(Zielverzeichnis As String)
        Dim Dateiname As String = conf.NameAusgabedatei
        Dim enc As Text.Encoding = conf.CodepageDateiAusgeben
        Using sw As IO.StreamWriter = New IO.StreamWriter(Zielverzeichnis & Dateiname, False, conf.CodepageDateiAusgeben)
            For i = 0 To arrKomplett.GetUpperBound(0)
                sw.WriteLine(arrKomplett(i))
                sw.WriteLine(vbFormFeed)
            Next i
            sw.WriteLine(vbFormFeed)
        End Using
        'Stop
    End Sub


    ''' <summary>
    ''' Erstellt ein Arraylist mit allen Dateien des genannten Verzeichnisses, die den gegebenen Namen oder Namensteil enthalten.)
    ''' </summary>
    ''' <param name="Verzeichnis">Verzeichnis, welches die zu sammelnden Dateien enthält.</param>
    ''' <param name="Name">Name oder Namensteil der Dateien, die in die Liste aufgenommen werden sollen. 
    ''' Wird nichts angegeben, werden alle Dateien des Verzeichnisses zurückgegeben.</param>
    Function DateilisteErstellen(Verzeichnis As String, Optional Name As String = "") As ArrayList
        Dim Dateiliste As New ArrayList

        If Name = "" Then
            Dateiliste.AddRange(IO.Directory.GetFiles(Verzeichnis)) 'Komplettes Verzeichnis einlesen
        Else
            For Each Datei In IO.Directory.GetFiles(Verzeichnis)    'Nur Dateien mit Namensteil "Name" einlesen
                If InStr(IO.Path.GetFileName(Datei).ToUpper, Name.ToUpper) > 0 Then
                    Dateiliste.Add(Datei)
                End If
            Next
        End If
        Return Dateiliste
    End Function

    Sub ArgumenteVerarbeiten(Arguments() As String)

        Select Case Arguments(1)
            Case "-?"
                ZeigeHilfe()

            Case "-getlist"
                'ErstelleTextdatei(arrSortierliste)

            Case "-setlist"



        End Select



    End Sub


    ''' <summary>
    ''' Zeigt eine Konsolenausgabe mit allen Parametern und deren Funktion an.

    ''' </summary>
    Sub ZeigeHilfe()
        Console.WriteLine()
        Console.WriteLine("*** ""Retourenbericht"" Parameterliste ***")
        Console.WriteLine("------------------------------------------")
        Console.WriteLine()
        Console.WriteLine("Syntax:")
        Console.WriteLine(">> WeekTXT_erstellen.exe [Parameter]")
        Console.WriteLine("Es wird nur der erste Parameter ausgewertet. Eine Angabe mehrerer")
        Console.WriteLine("Parameter ist nicht möglich.")
        Console.WriteLine()
        Console.WriteLine("Parameter  -    Beschreibung")
        Console.WriteLine(" -?        - Zeigt diese Hilfeseite an.")
        Console.WriteLine(" -getlist  - Gibt eine Textdatei mit der Sortierreihenfolge der Filial-")
        Console.WriteLine("             nummern aus. Nach Änderung kann sie wieder mit -setlist")
        Console.WriteLine("             eingelesen werden. Die Textdatei wird dann gelöscht.")
        Console.WriteLine(" -setlist  - Liest eine mit -getlist erstellte Textdatei mit der Reihen-")
        console.writeline("             folge der Filialnummern ein und speichert sie. Die Textdatei")
        Console.WriteLine("             wird gelöscht.")
        Console.WriteLine("")
        Console.WriteLine("-- Ende Parameterliste --")

    End Sub

    ''' <summary>
    ''' Schreibt die SortierListeFilialen in eine Textdatei. Jede Filialnummer in eine Zeile
    ''' </summary>
    ''' <param name="Filialliste"></param>
    Sub writeSortierListeFilialen(Filialliste As String())

        Try
            Using sw As New IO.StreamWriter(IO.Directory.GetCurrentDirectory & "\Liste.txt")
                For i = 0 To Filialliste.Count - 1
                    With sw
                        .WriteLine(Filialliste(i))
                    End With
                Next
            End Using
        Catch ex As Exception

        End Try
    End Sub


    ''' <summary>
    ''' Liest die Textdatei mit der Sortierreihenfolge der Filialen in My.Settings ein und löscht die Datei
    ''' </summary>
    Sub readSortierListeFilialen()

        Try
            'Datei öffnen und in String einlesen

            'String splitten in String-Array


            'strArray in intArray umwandeln
            ' .tostring


        Catch ex As Exception

        End Try
    End Sub
End Module
