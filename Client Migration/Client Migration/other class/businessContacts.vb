Imports System.ComponentModel
Public Class businessContacts
    Implements INotifyPropertyChanged
    Property NO As Integer
    Property CODE As String
    Property FULLNAME As String
    Property CONTACTTYPE As String
    Property COUNTRYCODE As String
    Property AREACODE As String
    Property CONTACTNUMBER As String
    Property CONTACTNAME As String
    Property POSITION As String
    Property CONTACTEMAIL As String

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Private Sub businessContacts_PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Handles Me.PropertyChanged

    End Sub
    Protected Sub OnPropertyChanged(propertyName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    End Sub
    Sub New(no_ As Integer, code_ As String, fullName As String, contactType_ As String, countrycode_ As String, areacode_ As String,
            contactNumber_ As String, contactName_ As String, pos As String, contactEmail_ As String)
        Try
            Me.NO = no_
            Me.CODE = code_
            Me.FULLNAME = fullName
            Me.CONTACTTYPE = contactType_
            Me.COUNTRYCODE = countrycode_
            Me.AREACODE = areacode_
            Me.CONTACTNUMBER = contactNumber_
            Me.CONTACTNAME = contactName_
            Me.POSITION = pos
            Me.CONTACTEMAIL = contactEmail_
        Catch ex As Exception
            MessageBoxError2(ex)
        End Try
    End Sub
End Class
