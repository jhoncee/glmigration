'Public Class UserCls
'    Inherits Basic.UserClass 
'    Private _mpID As String 

'    Sub Loguser_Out(ByVal TIMEIN As String)
'        Try
'            DataSource("update `user_logs` SET `TimeOut`='" & Format(ServerDateTime(), "H:mm:ss") & "' where UID='" & USER_LOGIN.ID & "' and TimeIn='" & TIMEIN & "'")
'        Catch ex As Exception
'        End Try
'    End Sub
'    Sub Loguser_IN(ByVal TIMEIN As String)
'        Try
'            DataSource("INSERT INTO `user_logs` SET UID='" & USER_LOGIN.ID & "',`TimeIn`='" & TIMEIN & "'")
'        Catch ex As Exception
'        End Try
'    End Sub

'End Class 