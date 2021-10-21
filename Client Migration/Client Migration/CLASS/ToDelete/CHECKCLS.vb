Public Class CHECKCLS
    Public Shared Function GET_CHECK_INFO(ByVal GUID As String) As DataTable
        Return DataSource("call GetCheckingAccountbyGUID('" & GUID & "')")
    End Function
    Public Shared Function GET_CHECK_INFO_TR(ByVal JID As Integer) As DataTable
        Return DB.QueryandReturnTable("call GetCheckingAccountbyJID('" & JID & "')")
    End Function


End Class
