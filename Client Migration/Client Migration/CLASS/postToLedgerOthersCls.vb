Imports System.ComponentModel
Public Class postToLedgerOthersCls
    Implements INotifyPropertyChanged
    Property Check As Boolean
    Property DueDate As Date
    Property _PRJID As Integer
    Property Unit As String
    Property UNITID As Integer
    Property ChargeName As String
    Property Description As String
    Property Amount As Decimal
    Private _AppliedPayment As Decimal
    Public Property AppliedPayment() As Decimal
        Get
            Return _AppliedPayment
        End Get
        Set(ByVal value As Decimal)
            _AppliedPayment = value
            OnPropertyChanged("AppliedPayment")
        End Set
    End Property

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Private Sub postToLedgerOthersCls_PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Handles Me.PropertyChanged
        If e.PropertyName = "AppliedPayment" Then
            If AppliedPayment > 0 Then
                If Check = False Then
                    Check = True
                End If
            End If
            If AppliedPayment > Amount Then
                AppliedPayment = Amount
            End If
        End If
    End Sub
    Protected Sub OnPropertyChanged(propertyName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    End Sub
    Sub New(chk As Boolean, DueDate As Date, PRJID As Integer, Unit As String, UNITID As Integer, ChargeName As String, Description As String, Amt As Decimal, appliedPay As Decimal)
        Try
            Me.Amount = Amt
            AppliedPayment = appliedPay
            Check = chk
            Me.ChargeName = ChargeName
            Me.Description = Description
            Me._PRJID = PRJID
            Me.Unit = Unit
            Me.UNITID = UNITID
            Me.DueDate = DueDate
        Catch ex As Exception
        End Try
    End Sub
End Class
