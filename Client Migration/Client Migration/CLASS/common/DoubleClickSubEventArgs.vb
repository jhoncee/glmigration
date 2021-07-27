Public Class EventArgs_
    Private _dep As String
    Public Property DEPARTMENT() As String
        Get
            Return _dep
        End Get
        Set(ByVal value As String)
            _dep = value
        End Set
    End Property

    Private ID As String
    Public Property SUBID() As String
        Get
            Return ID
        End Get
        Set(ByVal value As String)
            ID = value 
        End Set
    End Property

    Private _DATEF As Date
    Public Property DATE_FROM() As Date
        Get
            Return _DATEF
        End Get
        Set(ByVal value As Date)
            _DATEF = value
        End Set
    End Property
    Private _chrTypes As String
    Public Property Chrtypes() As String
        Get
            Return _chrTypes
        End Get
        Set(ByVal value As String)
            _chrTypes = value
        End Set
    End Property

    Private _loc As String
    Public Property Location() As String
        Get
            Return _loc
        End Get
        Set(ByVal value As String)
            _loc = value
        End Set
    End Property


    Private _title As String
    Public Property TITLEID() As String
        Get
            Return _title
        End Get
        Set(ByVal value As String)
            _title = value
        End Set
    End Property

    Private _dateto As Date
    Public Property DATE_TO() As Date
        Get
            Return _dateto
        End Get
        Set(ByVal value As Date)
            _dateto = value
        End Set
    End Property

    Private _funLoc As String
    Public Property Fund_Location() As String
        Get
            Return _funLoc
        End Get
        Set(ByVal value As String)
            _funLoc = value
        End Set
    End Property 

    Private _dtl As String
    Public Property DTLID() As String
        Get
            Return _dtl
        End Get
        Set(ByVal value As String)
            _dtl = value
        End Set
    End Property

    Sub New(ByVal _SUBID As String, ByVal DF As Date, ByVal DT As Date, ByVal ACID As String, ByVal _LOCACTION As String, ByVal _TITLE As String, Optional ByVal FL As Integer = 0, Optional ByVal FLINDEX_ As Integer = 0, Optional ByVal DEPARTMNT As String = "", Optional ByVal dtl As String = "%")
        SUBID = _SUBID
        DATE_FROM = DF
        DATE_TO = DT
        DTLID = dtl
        ACCID = ACID
        TITLEID = _TITLE
        Location = _LOCACTION
        Fund_Location = FL
        DEPARTMENT = DEPARTMNT
        FLINDEX = FLINDEX_
    End Sub


    Sub New(ByVal _SUBID As String, ByVal DF As Date, ByVal DT As Date, ByVal ACID As String, ByVal _LOCACTION As String, ByVal _TITLE As String, Optional ByVal _chr_Types As String = "%")
        SUBID = _SUBID
        DATE_FROM = DF
        DATE_TO = DT
        ACCID = ACID
        TITLEID = _TITLE
        Location = _LOCACTION
        Fund_Location = _LOCACTION
        Chrtypes = _chr_Types
    End Sub
    Private _FLINDEX As Integer
    Public Property FLINDEX() As Integer
        Get
            Return _FLINDEX
        End Get
        Set(ByVal value As Integer)
            _FLINDEX = value
        End Set
    End Property

    Private _accID As String
    Public Property ACCID() As String
        Get
            Return _accID
        End Get
        Set(ByVal value As String)
            _accID = value
        End Set
    End Property 
End Class
