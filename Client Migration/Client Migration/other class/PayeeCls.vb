Imports System.ComponentModel
Public Class PayeeCls
    Implements INotifyPropertyChanged
    Property NO As Integer
    Property CODE As String
    Property FULLNAME As String
    Property PAYEENAME As String

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Private Sub PayeeCls_PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Handles Me.PropertyChanged

    End Sub
    Protected Sub OnPropertyChanged(propertyName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    End Sub
    Sub New(no_ As Integer, code_ As String, fullname_ As String, payee_ As String)
        Try
            Me.NO = no_
            Me.CODE = code_
            Me.FULLNAME = fullname_
            Me.PAYEENAME = payee_
        Catch ex As Exception
            MessageBoxError2(ex)
        End Try
    End Sub
End Class
