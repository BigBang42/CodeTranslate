Public Class BaseEntry
    Private _MsgID As String
    Private _MsgStr As String
    Private _IsPlural As Boolean = False

    Property MsgID As String
        Get
            Return _MsgID
        End Get
        Set(value As String)
            _MsgID = value
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

End Class
