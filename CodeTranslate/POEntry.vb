Public Class POEntry

    Private _Files As New List(Of String)
    Private _Properties As New EntryProperty

    Private _MsgID As String
    Private _MsgID_Plural As String
    Private _MsgStr As String
    Private _MsgStr_Plural As String

    Private _SearchString As String
    Private _ReplaceString As String

    Private _IsFuzzyEntry As Boolean = False

    Private Shared _ImportStatus As AnalyzeStatus = AnalyzeStatus.Start
    Private Shared _CurrentPOLineMessagePrefix As String = String.Empty

    ReadOnly Property Files As List(Of String)
        Get
            Return _Files
        End Get
    End Property

    Property SearchString As String
        Get
            Return _SearchString
        End Get
        Set(value As String)
            _SearchString = value
        End Set
    End Property

    Property ReplaceString As String
        Get
            Return _ReplaceString
        End Get
        Set(value As String)
            _ReplaceString = value
        End Set
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

    ReadOnly Property IsEmpty As Boolean
        Get
            Return _ImportStatus = AnalyzeStatus.Start
        End Get
    End Property

    Private Shared Function IsValid(POEntry As POEntry) As Boolean
        If _ImportStatus = AnalyzeStatus.CreatingMsgStr Then
            If String.IsNullOrEmpty(POEntry.MsgID) OrElse String.IsNullOrEmpty(POEntry.MsgStr) Then Return False
            If POEntry.Properties.Type = EntryType.PluralText AndAlso (String.IsNullOrEmpty(POEntry.MsgID_Plural) OrElse String.IsNullOrEmpty(POEntry.MsgStr_Plural)) Then Return False
            Return True
        Else
            Return False
        End If
    End Function
    Private Shared Function ExtractFile(Line As String)
        Return Line.Substring(Line.IndexOf(":") + 2, Line.LastIndexOf(":") - Line.IndexOf(":") - 2)
    End Function

    'Private Shared Function ExtractMsgID(Line As String) As String
    '    Return Line.Substring(Line.IndexOf(Chr(34)) + 1, Line.LastIndexOf(Chr(34)) - Line.IndexOf(Chr(34)) - 1)
    'End Function

    'Private Shared Function ExtractMsgStr(Line As String) As String
    '    Return Line.Substring(Line.IndexOf(Chr(34)) + 1, Line.LastIndexOf(Chr(34)) - Line.IndexOf(Chr(34)) - 1)
    'End Function

    Private Shared Sub ExtractMessage(POEntry As POEntry, POLine As String, POLinePrefix As String)
        Dim MessageContent As String = POLine.Substring(POLine.IndexOf(Chr(34)) + 1, POLine.LastIndexOf(Chr(34)) - POLine.IndexOf(Chr(34)) - 1)

        Select Case POLinePrefix
            Case "msgid_plural"
                POEntry.MsgID_Plural = MessageContent
            Case "msgid"
                POEntry.MsgID = MessageContent
            Case "msgstr[0]"
                POEntry.MsgStr = MessageContent
            Case "msgstr[1]"
                POEntry.MsgStr_Plural = MessageContent
            Case "msgstr"
                POEntry.MsgStr = MessageContent
        End Select
    End Sub

    Private Shared Function ImportPOLine(POEntry As POEntry, POLine As String) As Boolean

        Select Case True
            Case POLine.StartsWith("#:")
                If {AnalyzeStatus.Start, AnalyzeStatus.CreatingFileList}.Contains(_ImportStatus) Then
                    POEntry.Files.Add(ExtractFile(POLine))
                    _ImportStatus = AnalyzeStatus.CreatingFileList
                Else
                    Return False
                End If

            Case POLine.StartsWith("#,")
                If _ImportStatus = AnalyzeStatus.CreatingFileList Then
                    POEntry.IsFuzzyEntry = True
                Else
                    Return False
                End If

            Case POLine.StartsWith("msgid_plural")
                If _ImportStatus = AnalyzeStatus.CreatingMsgID Then
                    POEntry.Properties.Type = EntryType.PluralText
                    _CurrentPOLineMessagePrefix = "msgid_plural"
                    ExtractMessage(POEntry, POLine, _CurrentPOLineMessagePrefix)
                Else
                    Return False
                End If

            Case POLine.StartsWith("msgid")
                If _ImportStatus = AnalyzeStatus.CreatingFileList Then 'Ohne zu bearbeitende Files kein Eintrag
                    POEntry.Properties.Type = EntryType.SingularText
                    _CurrentPOLineMessagePrefix = "msgid"
                    ExtractMessage(POEntry, POLine, _CurrentPOLineMessagePrefix)
                    _ImportStatus = AnalyzeStatus.CreatingMsgID
                Else
                    Return False
                End If

            Case POLine.StartsWith("msgstr[0]")
                If _ImportStatus = AnalyzeStatus.CreatingMsgID AndAlso POEntry.Properties.Type = EntryType.PluralText Then
                    _CurrentPOLineMessagePrefix = "msgstr[0]"
                    ExtractMessage(POEntry, POLine, _CurrentPOLineMessagePrefix)
                    _ImportStatus = AnalyzeStatus.CreatingMsgStr
                Else
                    Return False
                End If

            Case POLine.StartsWith("msgstr[1]")
                If _ImportStatus = AnalyzeStatus.CreatingMsgStr AndAlso POEntry.Properties.Type = EntryType.PluralText Then
                    _CurrentPOLineMessagePrefix = "msgstr[1]"
                    ExtractMessage(POEntry, POLine, _CurrentPOLineMessagePrefix)
                Else
                    Return False
                End If

            Case POLine.StartsWith("msgstr")
                If _ImportStatus = AnalyzeStatus.CreatingMsgID AndAlso POEntry.Properties.Type = EntryType.SingularText Then
                    _CurrentPOLineMessagePrefix = "msgstr"
                    ExtractMessage(POEntry, POLine, _CurrentPOLineMessagePrefix)
                    _ImportStatus = AnalyzeStatus.CreatingMsgStr
                Else
                    Return False
                End If

            Case POLine.StartsWith(Chr(32))   'Dann gehört diese Zeile wohl zu einem MsgID bzw. MsgStr-Eintrag
                If {AnalyzeStatus.CreatingMsgID, AnalyzeStatus.CreatingMsgStr}.Contains(_ImportStatus) Then
                    ExtractMessage(POEntry, POLine, _CurrentPOLineMessagePrefix)
                Else
                    Return False
                End If
        End Select

        Return True
    End Function

    Public Function ImportPOLine(POLine As String) As Boolean
        Return ImportPOLine(Me, POLine)
    End Function

    Public Function IsValid() As Boolean
        Return IsValid(Me)
    End Function

    Public Sub New()
        _CurrentPOLineMessagePrefix = String.Empty
        _ImportStatus = AnalyzeStatus.Start
    End Sub
End Class
