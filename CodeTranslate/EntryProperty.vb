Public Class EntryProperty
    Private _EntryType As EntryType
    Private _Prefix As String

    Property Type As EntryType
        Get
            Return _EntryType
        End Get
        Set(value As EntryType)
            _EntryType = value
            Select Case _EntryType
                Case EntryType.SingularText
                    _Prefix = "__("
                Case EntryType.PluralText
                    _Prefix = "__n("
                Case EntryType.Numeric
                Case EntryType.DateTime
                Case EntryType.Currency
            End Select
        End Set
    End Property

    ReadOnly Property Prefix As String
        Get
            Return _Prefix
        End Get
    End Property

End Class
