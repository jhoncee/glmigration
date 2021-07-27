Public Class employeeCls
    Property NO As Integer
    Property IDNUMBER As String
    Property LASTNAME As String
    Property FIRTSNAME As String
    Property MI As String
    Property DEPARTMENT As String
    Property ADDRESS As String

    Sub New(no_ As Integer, idnum As String, ln As String, fn As String, mi_ As String, dept As String, add_ As String)
        Try
            Me.NO = no_
            Me.IDNUMBER = idnum
            Me.LASTNAME = ln
            Me.FIRTSNAME = fn
            Me.MI = mi_
            Me.DEPARTMENT = dept
            Me.ADDRESS = add_
        Catch ex As Exception
            MessageBoxError2(ex)
        End Try
    End Sub
End Class
