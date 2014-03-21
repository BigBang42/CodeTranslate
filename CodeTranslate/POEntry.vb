Public Class POEntry
    Private Enum AnalyzeStatus
        Start = 0
        CreatingFileList = 1
        CreatingMsgID = 2
        CreatingMsgStr = 3
    End Enum

    Private _Files As New List(Of String)
    'Private _BaseEntries As New List(Of BaseEntry)
    Private _Properties As New EntryProperty

    Private _RawEntryLines As New List(Of String)

    Private _MsgID As String
    Private _MsgID_Plural As String
    Private _MsgStr As String
    Private _MsgStr_Plural As String

    Private _SearchString As String
    Private _ReplaceString As String

    Private _IsFuzzyEntry As Boolean = False

    ReadOnly Property Files As List(Of String)
        Get
            Return _Files
        End Get
    End Property

    'ReadOnly Property BaseEntries As List(Of BaseEntry)
    '    Get
    '        Return _BaseEntries
    '    End Get
    'End Property

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

    Property IsFuzzyEntry As Boolean
        Get
            Return _IsFuzzyEntry
        End Get
        Set(value As Boolean)
            _IsFuzzyEntry = value
        End Set
    End Property

    Property Properties As EntryProperty
        Get
            Return _Properties
        End Get
        Set(value As EntryProperty)
            _Properties = value
        End Set
    End Property

    Property MsgID As String
        Get
            Return _MsgID
        End Get
        Set(value As String)
            _MsgID = value
        End Set
    End Property

    Property MsgID_Plural As String
        Get
            Return _MsgID_Plural
        End Get
        Set(value As String)
            _MsgID_Plural = value
        End Set
    End Property

    Property MsgStr As String
        Get
            Return _MsgStr
        End Get
        Set(value As String)
            _MsgStr = value
        End Set
    End Property

    Property MsgStr_Plural As String
        Get
            Return _MsgStr_Plural
        End Get
        Set(value As String)
            _MsgStr_Plural = value
        End Set
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

    Private Shared Function Analyze(POEntry As POEntry) As Boolean
        Dim Status As AnalyzeStatus = AnalyzeStatus.Start

        For Each RawEntryLine As String In POEntry.RawEntryLines

            Select Case True
                Case RawEntryLine.StartsWith("#:")
                    If {AnalyzeStatus.Start, AnalyzeStatus.CreatingFileList}.Contains(Status) Then
                        POEntry.Files.Add(ExtractFile(RawEntryLine))
                        Status = AnalyzeStatus.CreatingFileList
                    Else
                        Return False
                    End If

                Case RawEntryLine.StartsWith("#,")
                    If Status = AnalyzeStatus.CreatingFileList Then
                        POEntry.IsFuzzyEntry = True
                    Else
                        Return False
                    End If

                Case RawEntryLine.StartsWith("msgid_plural")
                    If Status = AnalyzeStatus.CreatingMsgID Then
                        POEntry.Properties.Type = EntryType.PluralText
                        POEntry.MsgID_Plural = ExtractMsgID_Plural(RawEntryLine)
                    Else
                        Return False
                    End If

                Case RawEntryLine.StartsWith("msgid")
                    If Status = AnalyzeStatus.CreatingFileList Then 'Ohne zu bearbeitende Files kein Eintrag
                        POEntry.MsgID = ExtractMsgID(RawEntryLine)
                        Status = AnalyzeStatus.CreatingMsgID
                    Else
                        Return False
                    End If

                Case RawEntryLine.StartsWith("msgstr[0]")
                    If Status = AnalyzeStatus.CreatingMsgID Then
                        POEntry.MsgStr = ExtractMsgStr(RawEntryLine)
                        Status = AnalyzeStatus.CreatingMsgStr
                    Else
                        Return False
                    End If

                Case RawEntryLine.StartsWith("msgstr[1]")
                    If Status = AnalyzeStatus.CreatingMsgStr Then
                        POEntry.MsgStr_Plural = ExtractMsgStr_Plural(RawEntryLine)
                    Else
                        Return False
                    End If

                Case RawEntryLine.StartsWith("msgstr")
                    If Status = AnalyzeStatus.CreatingMsgID AndAlso POEntry.Properties.Type = EntryType.SingularText Then
                        POEntry.MsgStr = ExtractMsgStr(RawEntryLine)
                        Status = AnalyzeStatus.CreatingMsgStr
                    Else
                        Return False
                    End If

                Case RawEntryLine.StartsWith(Chr(32))   'Dann gehört diese Zeile wohl zu einem MsgID bzw. MsgStr-Eintrag
                    If {AnalyzeStatus.CreatingMsgID, AnalyzeStatus.CreatingMsgStr}.Contains(Status) Then

                    Else
                        Return False
                    End If
            End Select

        Next
    End Function

    Public Sub AddRawEntryLine(RawEntryLine As String)
        _RawEntryLines.Add(RawEntryLine)
    End Sub

    Public Function Analyze() As Boolean
        Return Analyze(Me)
    End Function
End Class
