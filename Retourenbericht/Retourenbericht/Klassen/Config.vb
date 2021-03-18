Public Class Config

    Private p_Quellordner As String
    Private p_Sicherungsordner As String
    Dim p_WeekDateiliste As New ArrayList
    Dim CPInput As Integer
    Dim CPOutput As Integer

#Region "Properties"

    Public Property NameAusgabedatei() As String

    Public Property Quellordner() As String
        Set(value As String)
            If Mid(value, Len(value), 1) <> "\" Then
                p_Quellordner = value & "\"
            Else
                p_Quellordner = value
            End If
        End Set
        Get
            Return p_Quellordner
        End Get
    End Property

    Public Property Sicherungsordner() As String
        Set(value As String)
            If Mid(value, Len(value)) <> "\" Then
                p_Sicherungsordner = value & "\"
            Else
                p_Sicherungsordner = value
            End If
        End Set
        Get
            Return p_Sicherungsordner
        End Get
    End Property

    ReadOnly Property CodepageDateiEinlesen() As Text.Encoding
        Get
            Return Text.Encoding.GetEncoding(CPInput)
        End Get
    End Property

    ReadOnly Property CodepageDateiAusgeben() As Text.Encoding
        Get
            Return Text.Encoding.GetEncoding(CPOutput)
        End Get
    End Property

    'ReadOnly Property GetWeekDateiliste As ArrayList
    '    Get
    '        Return p_WeekDateiliste
    '    End Get
    'End Property

#End Region

#Region "Methoden"

    ''' <summary>
    ''' Ruft die akuelle Sortierliste aus My.Settings ab.
    ''' </summary>
    ''' <returns></returns>
    Public Function getSortierListeFilialen() As Integer()

        If My.Settings.SortierListeFilialen = "" Then

            My.Settings.SortierListeFilialen = My.Settings.SortierlisteFilialenStandard
        End If

        Dim Liste As String() = Split(My.Settings.SortierListeFilialen, ";")
        Dim intListe(Liste.Count - 1) As Integer
        For i = 0 To Liste.Count - 1
            Integer.TryParse(Liste(i), intListe(i)) 'Wandelt die Werte in Ingeger um
        Next

        Return intListe
    End Function

    ''' <summary>
    ''' Schreibt die Integer()-Filialliste als ;-String in die aktuellen Programmsettings
    ''' </summary>
    ''' <param name="Liste"></param>
    Public Sub setSortierListeFilialen(arrListe As Integer())
        Dim Liste As String
        For i = 0 To arrListe.Count - 1
            Liste = Liste & arrListe(i) & ";"
        Next
        Liste = Mid(Liste, 1, Len(Liste) - 1)
        My.Settings.SortierListeFilialen = Liste

    End Sub


    Public Sub SetCodepageDateiEinlesen(page As Integer)
        CPInput = page
    End Sub

    Public Sub SetCodepageDateiAusgeben(page As Integer)
        CPOutput = page
    End Sub

#End Region



End Class
