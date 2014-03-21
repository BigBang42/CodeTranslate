Public Class FileToTranslate
    Dim _Name As String
    Dim _Entries As New List(Of POEntry)

    Property Name As String
        Get
            Return _Name
        End Get
        Set(value As String)
            _Name = value
        End Set
    End Property

    Property Entries As List(Of POEntry)
        Get
            Return _Entries
        End Get
        Set(value As List(Of POEntry))
            _Entries = value
        End Set
    End Property
End Class
