Imports System.ComponentModel
Public Class businessListDtl
    Implements INotifyPropertyChanged
    Property NO As Integer
    Property CODE As String
    Property FULLNAME As String
    Property TYPEOFADDRESS As String
    Property STREET As String
    Property BARANGAY As String
    Property CITY As String
    Property PROVINCE As String
    Property ZIP As String
    Property COUNTRY As String


    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Private Sub businessListDtl_PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Handles Me.PropertyChanged

    End Sub
    Protected Sub OnPropertyChanged(propertyName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    End Sub
    Sub New(no_ As Integer, code_ As String, fullname_ As String, typeOfAdd_ As String, street_ As String, barangay_ As String, city_ As String, province_ As String, zip_ As String, country_ As String)
        Try
            Me.NO = no_
            Me.CODE = code_
            Me.FULLNAME = fullname_
            Me.TYPEOFADDRESS = typeOfAdd_
            Me.STREET = street_
            Me.BARANGAY = barangay_
            Me.CITY = city_
            Me.PROVINCE = province_
            Me.ZIP = zip_
            Me.COUNTRY = country_
        Catch ex As Exception
            MessageBoxError2(ex)
        End Try
    End Sub

End Class
