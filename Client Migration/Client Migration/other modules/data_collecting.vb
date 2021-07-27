Module data_collecting
    Function GetBusinessList(ByVal type_ As String) As DataTable
        Try
            Return DB.QueryandReturnTable("SELECT * FROM general.`business_list` WHERE imported_from_excel =1 AND `type` = '" & type_ & "'")
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Function GetPoPerCompany(ByVal compId As Integer) As DataTable
        Try
            Return DB.QueryandReturnTable("SELECT * FROM procurement.`purchase_order_main` WHERE imported_from_excel = 1 AND title_id = '" & compId & "'")
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
End Module
