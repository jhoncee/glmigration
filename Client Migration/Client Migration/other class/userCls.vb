Public Class userCls
    Property NO As Integer
    Property EMPLOYEEID As String
    Property LASTNAME As String
    Property FIRSTNAME As String
    Property MI As String
    Property USERNAME As String
    Property PASSWORD As String
    Property EMAILADDRESS As String
    Sub New(no_ As Integer, employeeID_ As String, ln As String, fn As String, mi_ As String, userN As String, pword As String, emailAd As String)
        Try
            Me.NO = no_
            Me.EMPLOYEEID = employeeID_
            Me.LASTNAME = ln
            Me.FIRSTNAME = fn
            Me.MI = mi_
            Me.USERNAME = userN
            Me.PASSWORD = pword
            Me.EMAILADDRESS = emailAd
        Catch ex As Exception
            MessageBoxError2(ex)
        End Try
    End Sub
End Class
Public Class profileCls
    Property NO As Integer
    Property PROPERTYNAME As String
    Sub New(no_ As Integer, propertyN As String)
        Try
            Me.NO = no_
            Me.PROPERTYNAME = propertyN
        Catch ex As Exception
            MessageBoxError2(ex)
        End Try
    End Sub
End Class
