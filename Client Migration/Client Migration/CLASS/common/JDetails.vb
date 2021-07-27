
Public Class JDetails

    Private _JID As String
    Public Property JID() As String
        Get
            Return _JID
        End Get
        Set(ByVal value As String)
            _JID = value
        End Set
    End Property

    Private _nameId As String
    Public Property NameID() As String
        Get
            Return _nameId
        End Get
        Set(ByVal value As String)
            _nameId = value
        End Set
    End Property


    Private _stat As String
    Public Property Status() As String
        Get
            Return _stat
        End Get
        Set(ByVal value As String)
            _stat = value
        End Set
    End Property
    Private _cv As String
    Public Property CVNO() As String
        Get
            Return _cv
        End Get
        Set(ByVal value As String)
            _cv = value
        End Set
    End Property

End Class