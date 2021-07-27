Module mods
    Function DeleteBusinessList(ByVal types_ As String)
        Try
            DB.QueryExecNonQuery("DELETE FROM general.`business_list` WHERE imported_from_excel =1 AND `type` = '" & types_ & "'")
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Function
    Function DeleteBusinessListDtl(ByVal bid As Integer)
        Try
            DB.QueryExecNonQuery("DELETE FROM general.`business_list_detail` WHERE `BusinessNumber` = '" & bid & "'")
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Function
    Function DeleteBusinessListContact(ByVal bid As Integer)
        Try
            DB.QueryExecNonQuery("DELETE FROM general.`businesscontacts` WHERE `BusinessNumber` = '" & bid & "'")
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Function
    Function DeletePayee(ByVal bid As Integer)
        Try
            DB.QueryExecNonQuery("DELETE FROM accounting.`payee` WHERE BID =  '" & bid & "'")
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Function
    Function DeleteEmployee(ByVal bid As Integer)
        Try
            DB.QueryandReturnTable("DELETE FROM humanresource.`employees` WHERE BID = '" & bid & "'")
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Function
    Function DeleteProperty(ByVal types_ As String)
        Try
            DB.QueryExecNonQuery("DELETE a,b FROM general.`setup_project` AS a INNER JOIN propmanagement.`tbl_property` AS b ON a.`project_id` = b.`projectid` WHERE a.`imported_from_excel` = 1 AND a.`project_type` = '" & types_ & "'")
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Function

    Function DeleteUnit()
        Try
            DB.QueryExecNonQuery("DELETE FROM propmanagement.`tbl_property_unit` WHERE imported_from_excel =1")
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Function
    Function DeleteFloor()
        Try
            DB.QueryExecNonQuery("DELETE FROM propmanagement.`tbl_floor` WHERE imported_from_excel = 1")
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Function
    Function DeleteCountered(ByVal types_ As String)
        Try
            DB.QueryExecNonQuery("DELETE FROM accounting.`countered_rr` WHERE imported_from_excel =1 AND `Types` = '" & types_ & "'")
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Function
    Function DeleteAtc()
        Try
            DB.QueryExecNonQuery("DELETE FROM accounting.`atc_setup` WHERE imported_from_excel = 1")
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Function
    Function DeleteItemMaster(ByVal types_ As String)
        Try
            DB.QueryExecNonQuery("DELETE FROM procurement.setup_group_detail WHERE imported_from_excel = 1 AND inventory_type = '" & types_ & "'")
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Function
    Function DeletePoDtl(ByVal poid As Integer)
        Try
            DB.QueryExecNonQuery("DELETE FROM procurement.`purchase_order_details` WHERE po_id = '" & poid & "'")
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Function
    Function DeletePoMain(ByVal compID As Integer)
        Try
            DB.QueryExecNonQuery("DELETE FROM procurement.`purchase_order_main` WHERE imported_from_excel = 1 AND title_id  = '" & compID & "'")
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Function

    Function DeleteTrialBalance()
        Try
            DB.QueryExecNonQuery("SELECT * FROM accounting.`journal_detail` a INNER JOIN accounting.`journal_main` b USING (journal_id) WHERE b.`imported_from_excel` =1 AND b.`FrmCaption` = 'trialBalance'")
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Function
    Function DeleteAPBeginning()
        Try
            DB.QueryExecNonQuery("DELETE a,b,c FROM accounting.`apv_records` a INNER JOIN accounting.`journal_main` b ON a.`JID` = b.`journal_id` INNER JOIN accounting.`journal_detail` c ON b.`journal_id` = c.`journal_id` WHERE a.`imported_from_excel` = 1")
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Function
    Function DeleteUnitType()
        Try
            DB.QueryExecNonQuery("DELETE FROM propmanagement.`tbl_property_unit_type` WHERE imported_from_excel = 1")
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Function
    Function DeleteUnitLocation()
        Try
            DB.QueryExecNonQuery("DELETE FROM propmanagement.`tbl_property_unit_location` WHERE imported_from_excel = 1")
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Function
End Module
