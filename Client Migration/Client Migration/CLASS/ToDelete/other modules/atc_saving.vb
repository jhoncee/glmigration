Module atc_saving
    Public Sub saving_atc(ByVal desc_ As String, ByVal type_ As String, ByVal atc_ As String, ByVal rate_ As String, ByVal bir_ As String)
        Try
            Dim atc As New Dictionary(Of String, String)
            With atc
                .Add("DESCRIPTION", desc_)
                .Add("Goodservice", type_)
                .Add("ATC", atc_)
                .Add("RATE", rate_)
                .Add("ACID", "44")
                .Add("CACID", "20")
                .Add("CID", 1)
                .Add("BIR_FORMS", bir_)
                .Add("imported_from_excel", "1")
            End With
            DB.Insert(ACCOUNTINGDB & ".atc_setup", atc)
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Sub
    Public Sub save_atc_rate(atc_ As String, rate_ As String)
        Try
            Dim rate As New Dictionary(Of String, String)
            With rate
                .Add("ATC", atc_)
                .Add("Rate", rate_)
            End With
            DB.Insert(ACCOUNTINGDB & ".atc_tax_rate", rate)
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Sub
End Module
