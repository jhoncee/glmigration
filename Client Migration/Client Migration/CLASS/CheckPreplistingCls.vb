

Imports System.ComponentModel
Public Class CheckPreplistingCls
    Implements INotifyPropertyChanged
    Property ACID As Integer
    Property Check As Boolean
    Property APV As String
    Property Dates As Date
    Property Remarks As String
    Property Ledger As String
    Property SUBID As Integer
    Property Bal As Decimal
    Private _Amount As Decimal

    Private _Defered As Decimal
    Public Property DeferredAmt() As Decimal
        Get
            Return _Defered
        End Get
        Set(ByVal value As Decimal)
            _Defered = value
        End Set
    End Property
    Public Property Amount() As Decimal
        Get
            Return _Amount
        End Get
        Set(ByVal value As Decimal)
            _Amount = value
            If _Amount <> Bal Then
                OnPropertyChanged("Amount")
            End If
        End Set
    End Property
    Property GUID As String
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Protected Sub OnPropertyChanged(propertyName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    End Sub

    Sub New(chk As Boolean, ACID As Integer, SUBTYPE As String, SUBID As Integer, APV As String, Dates As Date, Amount As Decimal, GUID As String, remarks As String, _DeferedAMT As Decimal)
        Me.Check = chk
        Me.APV = APV
        Me.Dates = Dates
        Me.Bal = Amount
        Me.Ledger = SUBTYPE
        Me.Amount = Amount
        Me.GUID = GUID
        Me.Remarks = remarks
        Me.ACID = ACID
        Me.SUBID = SUBID
        Me.DeferredAmt = _DeferedAMT
    End Sub

    Private Sub ApvListing_PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Handles Me.PropertyChanged
        Try
            If e.PropertyName = "Amount" Then
                Check = True
                If Amount > Bal Then
                    Amount = Bal
                ElseIf Amount < 0 Then
                    Amount = 0
                    Check = False
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub
End Class
