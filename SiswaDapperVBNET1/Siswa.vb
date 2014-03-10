Public Class Siswa

    Private _nis As String
    Public Property Nis() As String
        Get
            Return _nis
        End Get
        Set(ByVal value As String)
            _nis = value
        End Set
    End Property


    Private _nama As String
    Public Property Nama() As String
        Get
            Return _nama
        End Get
        Set(ByVal value As String)
            _nama = value
        End Set
    End Property

End Class
