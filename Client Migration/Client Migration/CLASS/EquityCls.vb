Imports System.ComponentModel

Public Class EquityCls
    Implements INotifyPropertyChanged
    Property ARDate As Date
    Property ARNum As String
    Property PRDate As Date
    Property PRNum As String
    Property ORDate As Date
    Property ORNum As String
    Property CRDate As Date
    Property CRNum As String
    Property RefDate As Date
    Property RefNum As String

    Property Check As Boolean
    Property DueDate As Date
    Property Unit As String
    Property UNITID As Integer
    Property PRJID As Integer
    Property ChargeName As String
    Property Description As String
    Property DueAmount As Decimal 'AR
    Private _AmountPaid As Decimal

    Private _bank As String
    Public Property Bank() As String
        Get
            Return _bank
        End Get
        Set(ByVal value As String)
            _bank = value
        End Set
    End Property

    Private _CheckNum As String
    Public Property CheckNum() As String
        Get
            Return _CheckNum
        End Get
        Set(ByVal value As String)
            _CheckNum = value
        End Set
    End Property
    Property CheckDate As Date
    Public Property AmountPaid() As Decimal
        Get
            Return _AmountPaid
        End Get
        Set(ByVal value As Decimal)
            _AmountPaid = value
            OnPropertyChanged("AmountPaid")
        End Set
    End Property

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Private Sub EquityCls_PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Handles Me.PropertyChanged
        If e.PropertyName = "AmountPaid" Then
            If AmountPaid > DueAmount Then
                AmountPaid = DueAmount
            End If
        End If
        '  If e.PropertyName = "Bank" And Bank <> "" Then If Check = False Then Check = True
        '   If e.PropertyName = "CheckNum" And CheckNum <> "" Then If Check = False Then Check = True
    End Sub
    Protected Sub OnPropertyChanged(propertyName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    End Sub
    Sub New(Chk As Boolean, DueDate As Date, PRJID As Integer, Unit As String, UNITID As Integer, ChargeName As String, Description As String, DueAmt As Decimal)
        Try
            Me.Check = Chk
            Me.DueDate = DueDate
            Me.PRJID = PRJID
            Me.UNITID = UNITID
            Me.Unit = Unit
            Me.ChargeName = ChargeName
            Me.Description = Description
            Me.DueAmount = DueAmt
            Me.AmountPaid = DueAmount
            Me.CheckNum = ""
            Me.Bank = ""
            Me.CheckDate = DueDate
        Catch ex As Exception
        End Try
    End Sub

    Sub New(Chk As Boolean, DueDate As Date, PRJID As Integer, Unit As String, UNITID As Integer, ChargeName As String, Description As String, DueAmt As Decimal, Bank As String, CheckNum As String, CheckDate As Date, ARDate As Date, ARNum As String, PRDate As Date, PRNum As String, ORDate As Date, ORNum As String, CRDate As Date, CRNum As String, RefDate As Date, RefNUm As String)
        Try
            Me.Check = Chk
            Me.PRJID = PRJID
            Me.UNITID = UNITID
            Me.Unit = Unit
            Me.ChargeName = ChargeName
            Me.Description = Description
            Me.DueAmount = DueAmt
            Me.AmountPaid = DueAmount
            Me.CheckNum = CheckNum
            Me.Bank = Bank

            Me.ARNum = ARNum
            Me.PRNum = PRNum
            Me.ORNum = ORNum
            Me.CRNum = CRNum
            Me.RefNum = RefNUm

            Me.DueDate = DueDate
            Me.CheckDate = CheckDate
            Me.ARDate = ARDate
            Me.PRDate = PRDate
            Me.ORDate = ORDate
            Me.CRDate = CRDate
            Me.RefDate = RefDate
        Catch ex As Exception
        End Try
    End Sub
End Class
