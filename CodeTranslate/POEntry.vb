Public Class POEntry

    Dim _Files As New List(Of String)
    Dim _BaseEntry As New BaseEntry

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
        Return Line.Replace("#: ", String.Empty)
    End Function

    Private Shared Function ExtractMsgID(Line As String) As String
        Return Line.Substring(Line.IndexOf(Chr(34)) + 1, Line.LastIndexOf(Chr(34)) - Line.IndexOf(Chr(34)))
    End Function

    Private Shared Function ExtractMsgStr(Line As String) As String
        Return Line.Substring(Line.IndexOf(Chr(34)) + 1, Line.LastIndexOf(Chr(34)) - Line.IndexOf(Chr(34)))
    End Function

    Private Shared Function ImportPOLine(POLine As String, POEntry As POEntry) As Boolean

        Select Case True
            Case POLine.StartsWith("#:")
                If String.IsNullOrEmpty(POEntry.MsgID) And String.IsNullOrEmpty(POEntry.MsgStr) Then
                    POEntry.Files.Add(ExtractFile(POLine))
                Else
                    Return False
                End If

            Case POLine.StartsWith("msgid")
                If POEntry.Files.Count > 0 Then
                    POEntry.MsgID = ExtractMsgID(POLine)
                Else
                    Return False
                End If

            Case POLine.StartsWith("msgstr")
                If Not String.IsNullOrEmpty(POEntry.MsgID) Then
                    POEntry.MsgStr = ExtractMsgStr(POLine)
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
        If _Files.Count > 0 AndAlso Not String.IsNullOrEmpty(_BaseEntry.MsgID) AndAlso Not String.IsNullOrEmpty(_BaseEntry.MsgStr) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub Clear()
        _Files.Clear()
        _BaseEntry.MsgID = Nothing
        _BaseEntry.MsgStr = Nothing
    End Sub
End Class
