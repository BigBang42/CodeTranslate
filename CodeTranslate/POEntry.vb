Public Class POEntry
    Private Enum EntryStatus
        Empty = 0
        ExtractFiles = 1
        MsgIDExtracted = 2
        MsgIDPluralExtracted = 3
        ExtractionComplete = 4
    End Enum

    Private _Files As New List(Of String)
    Private _BaseEntry As New BaseEntry
    Private Status As EntryStatus = EntryStatus.Empty

    Private Shared _FuzzyEntryIsProcessed As Boolean = False

    ReadOnly Property Files As List(Of String)
        Get
            Return _Files
        End Get
    End Property

    Property MsgID As String
        Get
            Return _BaseEntry.MsgID
        End Get
        Set(value As String)
            _BaseEntry.MsgID = value
        End Set
    End Property

    Property MsgStr As String
        Get
            Return _BaseEntry.MsgStr
        End Get
        Set(value As String)
            _BaseEntry.MsgStr = value
        End Set
    End Property

    ReadOnly Property BaseEntry As BaseEntry
        Get
            Return _BaseEntry
        End Get
    End Property

    Private Shared Function ExtractFile(Line As String)
        Return Line.Substring(Line.IndexOf(":") + 2, Line.LastIndexOf(":") - Line.IndexOf(":") - 2)
    End Function

    Private Shared Function ExtractMsgID(Line As String) As String
        Return Line.Substring(Line.IndexOf(Chr(34)) + 1, Line.LastIndexOf(Chr(34)) - Line.IndexOf(Chr(34)) - 1)
    End Function

    Private Shared Function ExtractMsgStr(Line As String) As String
        Return Line.Substring(Line.IndexOf(Chr(34)) + 1, Line.LastIndexOf(Chr(34)) - Line.IndexOf(Chr(34)) - 1)
    End Function

    Private Shared Function ImportPOLine(POLine As String, POEntry As POEntry) As Boolean
        'Static Dim _FuzzyEntryIsProcessed As Boolean = False

        Select Case True
            Case POLine.StartsWith("#:")
                If {EntryStatus.Empty, EntryStatus.ExtractFiles}.Contains(POEntry.Status) Then
                    POEntry.Status = EntryStatus.ExtractFiles
                    POEntry.Files.Add(ExtractFile(POLine))
                Else
                    Return False
                End If

            Case POLine.StartsWith("msgid_plural")
                If POEntry.Status = EntryStatus.MsgIDExtracted Then
                    'Wird erstmal ignoriert, da nur in einem File bis jetzt vorkommend
                Else
                    Return False
                End If

            Case POLine.StartsWith("msgid")
                If POEntry.Status = EntryStatus.ExtractFiles Then 'Ohne zu bearbeitende Files kein Eintrag
                    POEntry.MsgID = ExtractMsgID(POLine)
                    POEntry.Status = EntryStatus.MsgIDExtracted
                ElseIf POEntry.Status = EntryStatus.Empty And Not _FuzzyEntryIsProcessed Then  'Am Anfang jeder po.-Datei ist eine sog. Fuzzy Entry mit Leerstrings als MsgId und MsgStr
                    POEntry.Status = EntryStatus.MsgIDExtracted
                    _FuzzyEntryIsProcessed = True   'Wird diese gerade bearbeitet, ist die Files-Collection natürlich leer, was in diesem Falle kein Fehler ist
                Else
                    Return False
                End If

            Case POLine.StartsWith("msgstr[1]")
                If POEntry.Status = EntryStatus.Empty Then
                    'Wird erstmal ignoriert, da nur in einem File bis jetzt vorkommend
                Else
                    Return False
                End If
            Case POLine.StartsWith("msgstr")
                If POEntry.Status = EntryStatus.MsgIDExtracted Then
                    POEntry.MsgStr = ExtractMsgStr(POLine)
                    POEntry.Status = EntryStatus.ExtractionComplete
                Else
                    Return False
                End If

        End Select

        Return True

    End Function

    Public Function ImportPOLine(POLine As String) As Boolean
        Return ImportPOLine(POLine, Me)
    End Function

    Public Function IsComplete() As Boolean
        Return Status = EntryStatus.ExtractionComplete
    End Function

    Public Shared Sub Reset()
        _FuzzyEntryIsProcessed = False
    End Sub
End Class
