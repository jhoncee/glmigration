

Imports System.ComponentModel
Public Class ApvListing
    Implements INotifyPropertyChanged
    Property Check As Boolean
    Property Ref As String
    Property IsVat As String
    Property Dates As Date
    Property Bal As Decimal
    Private _Amount As Decimal
    Public Property Amount() As Decimal
        Get
            Return _Amount
        End Get
        Set(ByVal value As Decimal)
            _Amount = value
            OnPropertyChanged("Amount")
        End Set
    End Property
    Property PROJECTID As Integer
    Property GUID As String
    Property TransName As String
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Protected Sub OnPropertyChanged(propertyName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    End Sub

    Sub New(IsVat As String, Ref As String, Dates As Date, Amount As Decimal, ProjectID As Integer, GUID As String, TransName As String, Optional _check As Boolean = False)
        Me.Check = _check
        Me.Dates = Dates
        Me.Bal = Amount
        Me.Amount = Amount
        Me.PROJECTID = ProjectID
        Me.GUID = GUID
        Me.Ref = Ref
        Me.TransName = TransName
        Me.IsVat = IsVat
    End Sub

    Private Sub ApvListing_PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Handles Me.PropertyChanged
        Try
            If e.PropertyName = "Amount" Then
                If Amount > Bal Then
                    Amount = Bal
                    Check = True
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub
End Class
