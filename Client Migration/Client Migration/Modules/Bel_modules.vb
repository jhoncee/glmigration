Module Bel_modules

    Public Function GET_ALL_SALES_CLIENT(ByVal unitId As Integer) As DataTable
        Return Datasource("SELECT DISTINCT business_id 'ID', BUYERFULLNAME 'NAME', id 'tbl_id', agent_id, broker_id, tbl_buyer_information.total_contract_price, tbl_property_unit.`vatable` FROM " & PropmanagementdBASE & ".tbl_buyer_information INNER JOIN " & PropmanagementdBASE & ".`tbl_property_unit` ON (tbl_property_unit.`UNITID` = tbl_buyer_information.`unit_id`) WHERE tbl_buyer_information.`unit_id` = " & unitId & "  ")
    End Function

    Public Function GET_broker_list(ByVal agent_id As String) As DataTable
        Try
            Return Datasource("SELECT `business_number`,`business_fullname` FROM " & GENERALDB & ".`business_list` WHERE `type` = 'REALTY' and `status` = 'ACTIVE'AND business_number = " & agent_id & "")
        Catch ex As Exception
            Return Datasource("SELECT `business_number`,`business_fullname` LIMIT 0")
        End Try
    End Function

    Public Function GET_agent_list(ByVal agent_id As String) As DataTable
        Try
            Return Datasource("SELECT `business_number`,`business_fullname` FROM " & GENERALDB & ".`business_list` WHERE `type` = 'AGENT' AND `status` = 'ACTIVE' AND business_number = " & agent_id & "")
        Catch ex As Exception
            Return Datasource("SELECT `business_number`,`business_fullname` LIMIT 0")
        End Try
    End Function


End Module
