Public Class POEntry

    Private _Files As New List(Of String)
    Private _BaseEntries As New List(Of BaseEntry)
    Private _Properties As New EntryProperty

    Private _RawEntryLines As New List(Of String)

    Private _SearchString As String
    Private _ReplaceString As String

    ReadOnly Property Files As List(Of String)
        Get
            Return _Files
        End Get
    End Property

    ReadOnly Property BaseEntries As List(Of BaseEntry)
        Get
            Return _BaseEntries
        End Get
    End Property

    ReadOnly Property SearchString As String
        Get
            Return _SearchString
        End Get
    End Property

    ReadOnly Property ReplaceString As String
        Get
            Return _ReplaceString
        End Get
    End Property

    ReadOnly Property RawEntryLines As List(Of String)
        Get
            Return _RawEntryLines
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
                ElseIf POEntry.Status = EntryStatus.Empty Then  'Am Anfang jeder po.-Datei ist eine sog. Fuzzy Entry mit Leerstrings als MsgId und MsgStr
                    POEntry.Status = EntryStatus.MsgIDExtracted
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

    Private Shared Function Analyze(POEntry As POEntry) As Boolean

    End Function

    Public Sub AddRawEntryLine(RawEntryLine As String)
        _RawEntryLines.Add(RawEntryLine)
    End Sub

    Public Function Analyze() As Boolean
        Return Analyze(Me)
    End Function

    Public Function ImportPOLine(POLine As String) As Boolean
        Return ImportPOLine(POLine, Me)
    End Function
End Class
