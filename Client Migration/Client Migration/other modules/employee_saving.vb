Module employee_saving


    Public Sub save_employee(ByVal empID As String, ByVal fullN As String, ByVal fn As String, ByVal ln As String, ByVal mi As String,
                             ByVal address_ As String, ByVal dept_ As String, bid_ As Integer, PersonType As String)
        Try
            Dim emp As New Dictionary(Of String, String)
            With emp
                .Add("PersonType", PersonType)
                .Add("Name_Empl", fullN)
                .Add("FirstName_Empl", fn)
                .Add("MiddleName_Empl", mi)
                .Add("LastName_Empl", ln)
                .Add("Address_Empl", address_)
                .Add("Department_Empl", dept_)
                .Add("Status_Empl", "ACTIVE")
                .Add("Group_Empl", 1)
                .Add("title_id", 1)
                .Add("BID", bid_)
            End With
            DB.Insert(HRDBASE & ".employees", emp)
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Sub
End Module
