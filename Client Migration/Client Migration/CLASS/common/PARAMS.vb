Public Class PARAMS

    Private _ACCID As Integer
    Public Property ACCOUNTID() As Integer
        Get
            Return _ACCID
        End Get
        Set(ByVal value As Integer)
            _ACCID = value
        End Set
    End Property
    Private _DATE As Date
    Public Property DATEFROM() As Date
        Get
            Return _DATE
        End Get
        Set(ByVal value As Date)
            _DATE = value
        End Set
    End Property

    Private _DATET As Date
    Public Property DATETO() As Date
        Get
            Return _DATE
        End Get
        Set(ByVal value As Date)
            _DATE = value
        End Set
    End Property

    Private _CONS As Integer
    Public Property CONSO() As Integer
        Get
            Return _CONS
        End Get
        Set(ByVal value As Integer)
            _CONS = value
        End Set
    End Property


    Private _FLID As Integer
    Public Property FLID() As Integer
        Get
            Return _FLID
        End Get
        Set(ByVal value As Integer)
            _FLID = value
        End Set
    End Property


    Private _FLINDEX As Integer
    Public Property FL_INDEX() As Integer
        Get
            Return _FLINDEX
        End Get
        Set(ByVal value As Integer)
            _FLINDEX = value
        End Set
    End Property



    Private _COMPANYID As String
    Public Property COMPANYID() As String
        Get
            Return _COMPANYID
        End Get
        Set(ByVal value As String)
            _COMPANYID = value
        End Set
    End Property

    Private _OFFICEID As Integer
    Public Property OFFICEID() As Integer
        Get
            Return _OFFICEID
        End Get
        Set(ByVal value As Integer)
            _OFFICEID = value
        End Set
    End Property

    Private _GID As String
    Public Property GID() As String
        Get
            Return _GID
        End Get
        Set(ByVal value As String)
            _GID = value
        End Set
    End Property

    Private _GID1 As String
    Public Property GID1() As String
        Get
            Return _GID1
        End Get
        Set(ByVal value As String)
            _GID1 = value
        End Set
    End Property
    Sub New(Optional ByVal ACID As Integer = 0, Optional ByVal DF As Date = Nothing, Optional ByVal DT As Date = Nothing, Optional ByVal CONS As Integer = 0, Optional ByVal FLINDX As Integer = 0, Optional ByVal FL_ID As Integer = 0,
           Optional ByVal COMPID As Integer = 0, Optional ByVal OFFICE_ID As Integer = 0, Optional ByVal G_ID As String = "", Optional ByVal G_ID1 As String = "")
        ACCOUNTID = ACID
        DATEFROM = DF
        DATETO = DT
        CONSO = CONS
        FLID = FL_ID
        FL_INDEX = FLINDX
        COMPANYID = COMPID
        OFFICEID = OFFICE_ID
        GID = G_ID
        GID1 = G_ID1
    End Sub
End Class
