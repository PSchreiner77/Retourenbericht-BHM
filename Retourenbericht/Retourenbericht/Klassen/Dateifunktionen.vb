''' <summary>
''' Sammlung von Dateifunktionen und -operationen für BHM. 
''' (Speichern, Schreiben, Auslesen in Array, Suchen in, ...)
''' </summary>
Public Class Dateifunktionen


    Function DateiInArrayEinlesen(Dateipfad As String) As String()
        Return DateiInArrayEinlesen(Dateipfad, 0)
    End Function


    ''' <summary>
    ''' Liest eine Textdatei zeilenweise in ein ArrayList ein und gibt dieses zurück. 
    ''' Wird die Datei nicht gefunden oder ist sie leer, wird ein leeres Arraylist zurückgegeben.
    ''' </summary>
    ''' <param name="Dateipfad">Pfad zur Datei, die eingelesen werden soll.</param>
    ''' <returns>Gibt ein Arraylist mit den Zeilen der Textdatei zurück.</returns>
    Function DateiInArrayEinlesen(Dateipfad As String, Codepage As Integer) As String()
        Dim Dateiinhalt As String()
        Dim enc As Text.Encoding

        Try
            enc = Text.Encoding.GetEncoding(Codepage)

        Catch ex As Exception

        End Try


        Try
            'Wenn Datei existiert, auslesen. Anonsten leere Arraylist zurück.
            If IO.File.Exists(Dateipfad) Then
                'Datei zeilenweise in Arraylist einlesen.

                Using sr As New IO.StreamReader(Dateipfad)
                    Dateiinhalt = Split(sr.ReadToEnd, vbCrLf)

                End Using

            Else
                'nichts machen, Arraylist bleibt leer
            End If

        Catch ex As Exception

        End Try

    End Function


    ''' <summary>
    ''' Erstellt ein Arraylist mit allen weekr*.* Dateien des genannten Verzeichnisses)
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

End Class