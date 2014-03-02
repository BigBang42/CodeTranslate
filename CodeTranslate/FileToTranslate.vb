Public Class FileToTranslate
    Dim _Name As String
    Dim _Entries As New List(Of BaseEntry)

    Property Name As String
        Get
            Return _Name
        End Get
        Set(value As String)
            _Name = value
        End Set
    End Property

    Property Entries As List(Of BaseEntry)
        Get
            Return _Entries
        End Get
        Set(value As List(Of BaseEntry))
            _Entries = value
        End Set
    End Property
End Class
