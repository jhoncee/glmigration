Module modforID
    Function GetBusinessTinNum(bid As Integer) As String
        Try
            Dim x = DB.QueryandReturnTable("SELECT tin_number FROM general.`business_list` WHERE business_number = '" & bid & "'")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function
    Function GetBusinessAddress(bid As Integer) As String
        Try
            Dim x = DB.QueryandReturnTable("SELECT CONCAT(street,SPACE(1),barangay,', ',city,', ',province,', ',zip,', ',country)'Address' FROM general.`business_list_detail` WHERE BusinessNumber  = '" & bid & "'")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function
    Function get_businessListID_by_NAME_and_CODE(ByVal accNum As String, ByVal businessN As String, ByVal type_ As String) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT IFNULL(business_number,0)'ID' FROM general.`business_list` WHERE ACCOUNTNUMBER = '" & accNum & "' or business_fullname = '" & businessN & "' AND `type` = '" & type_ & "'")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function

    Function get_businessListID_by_NAME(ByVal businessN As String) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT IFNULL(business_number,0)'ID' FROM general.`business_list` WHERE business_fullname = '" & businessN.RSQ & "'")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function


    Function get_businessID_by_Code(ByVal code_ As String) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT IFNULL(business_number,0)'ID' FROM general.`business_list` WHERE ACCOUNTNUMBER = '" & code_ & "'")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function



    Function get_all_businessAddress(ByVal bid As Integer) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT IFNULL(ID,0)'ID' FROM general.`business_list_detail` WHERE BusinessNumber = '" & bid & "'")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Function get_all_businessPayee(ByVal bid As Integer) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT IFNULL(ID,0)'ID' FROM accounting.`payee` WHERE BID = '" & bid & "'")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function

    Function get_atc_id(ByVal atc_ As String) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT ID FROM accounting.`atc_setup` WHERE ATC = '" & atc_ & "'")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Function get_item_with_type_and_projectID_id(ByVal itemCode_ As String, ByVal itemName As String, ByVal companyID_ As Integer, ByVal types_ As String, ByVal projectID_ As Integer) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT `group_detail_id` FROM procurement.`setup_group_detail` WHERE title_id1 ='" & companyID_ & "' AND (item_code = '" & itemCode_ & "' or description = '" & itemName & "') AND inventory_type = '" & types_ & "' AND title_id = '" & projectID_ & "'")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Function get_item_with_type_id(ByVal itemCode_ As String, ByVal itemName As String, ByVal companyID_ As Integer, ByVal types_ As String) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT `group_detail_id` FROM procurement.`setup_group_detail` WHERE title_id1 ='" & companyID_ & "' AND (item_code = '" & itemCode_ & "' or description = '" & itemName & "') AND inventory_type = '" & types_ & "'")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Function get_item_id(ByVal itemCode_ As String, ByVal itemName As String, ByVal companyID_ As Integer) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT `group_detail_id` FROM procurement.`setup_group_detail` WHERE title_id1 ='" & companyID_ & "' AND (item_code = '" & itemCode_ & "' or description = '" & itemName & "')")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Function get_company_ID(ByVal companyCode As String) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT title_id FROM general.`project_title` WHERE title_code = '" & companyCode & "'")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Function GetCompanyIDbyPrjID(ByVal PrjID As String) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT title_id FROM general.`setup_project` WHERE project_id = '" & PrjID & "'")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function

    Function get_setupGroupID(ByVal compID_ As Integer, ByVal type_ As String) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT group_id FROM procurement.`setup_group` WHERE title_id = '" & compID_ & "' AND group_description = '" & type_ & "'")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
            Return 0
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Function get_branchID(ByVal compID As Integer) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT project_id FROM general.`setup_project` WHERE Main_office_id = '" & compID & "' ")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Function get_empID_with_Name(ByVal empID As String, ByVal empN As String) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT IFNULL(SysPK_Empl,0)'ID' FROM humanresource.`employees` WHERE Name_Empl = '" & empN & "' OR UserID_Empl = '" & empID & "'")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Function get_deptID(ByVal dept_ As String) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT ifnull(SysPK_Dept,0)'ID' FROM humanresource.`department` WHERE Name_Dept = '" & dept_ & "'")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Function get_mainID(ByVal compID As Integer) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT project_id FROM general.setup_project WHERE title_id='" & compID & "' AND project='MAIN OFFICE'")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Function get_projectNameWithoutCompID(ByVal proj_no As String) As String
        Try
            Dim x = DB.QueryandReturnTable("SELECT IFNULL(project_name,'')'ProjectN' FROM general.`setup_project` WHERE project_no = '" & proj_no & "'")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function
    Function get_projectName(ByVal proj_no As String, ByVal compID As Integer) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT IFNULL(project_name,'')'ProjectN' FROM general.`setup_project` WHERE project_no = '" & proj_no & "' and title_id = '" & compID & "'")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Function get_projectIDwithCompID(ByVal proj_no As String, ByVal compID As Integer) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT IFNULL(project_id,0)'ID' FROM general.`setup_project` WHERE project_no = '" & proj_no & "' and title_id = '" & compID & "'")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Function get_projectID(ByVal proj_no As String) As Integer
        Try
            Dim xX As New DataTable
            xX = DB.QueryandReturnTable("SELECT project_id 'ID' FROM general.`setup_project` WHERE project_no = '" & proj_no & "'")
            If xX.Rows.Count > 0 Then
                Return xX(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Function get_propertyID_with_Name(ByVal propertyNo As String, ByVal propertyName As String) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT IFNULL(Proid,0)'ID' FROM propmanagement.`tbl_property` WHERE PropertyNo = '" & propertyNo & "' OR PropertyName = '" & propertyName & "'")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Function GetUnitTypeID(ByVal unitType_ As String) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT IFNULL(id,0)'ID' FROM propmanagement.`tbl_property_unit_type` WHERE unit_type = '" & unitType_ & "'")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Function GetUnitLocID(ByVal unitLoc_ As String) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT IFNULL(id,0)'ID' FROM propmanagement.`tbl_property_unit_location` WHERE location = '" & unitLoc_ & "'")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Function GetFloorID(ByVal floorN As String, ByVal prjid_ As Integer) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT IFNULL(id,0)'ID' FROM propmanagement.`tbl_floor` WHERE floor_name = '" & floorN & "' AND PRJID = '" & prjid_ & "'")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Function GetFloorNumber(ByVal floorN As String, ByVal prjid_ As Integer) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT IFNULL(floor,0)'ID' FROM propmanagement.`tbl_floor` WHERE floor_name = '" & floorN & "' AND PRJID = '" & prjid_ & "'")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Function GetUnitID(ByVal unitN As String) As Integer
        Try
            Dim xx As New DataTable
            xx = DB.QueryandReturnTable("SELECT IFNULL(UNITID,0)'ID' FROM propmanagement.`tbl_property_unit` WHERE UNITNO = '" & unitN.Trim & "'")
            If xx.Rows.Count > 0 Then
                Return xx(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function

    Function GetEmpID(ByVal empNum As String) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT IFNULL(SysPK_Empl,0)'ID' FROM humanresource.`employees` WHERE UserID_Empl = '" & empNum & "'")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Function GetEmpName(ByVal empNum As String) As String
        Try
            Dim x = DB.QueryandReturnTable("SELECT IFNULL(Name_Empl,'')'Name' FROM humanresource.`employees` WHERE UserID_Empl = '" & empNum & "'")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function
    Function GetUserId(ByVal userN As String, ByVal empID As Integer) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT IFNULL(id,0)'ID' FROM general.`users` WHERE UserName_User = '" & userN & "' AND Employee_id = '" & empID & "'")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Function GetEmployeeDepartment(ByVal empNum As String) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT IFNULL(Department_Empl,0)'ID' FROM humanresource.`employees` WHERE UserID_Empl = '" & empNum & "'")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function
    Function GetCompanyIdWithName(ByVal code_ As String, ByVal titleN As String) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT IFNULL(title_id,0)'ID' FROM general.`project_title` WHERE title_code = '" & code_ & "' OR title_name = '" & titleN & "'")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function
    Function GetCompanyIdWithCode(ByVal code_ As String) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT IFNULL(title_id,0)'ID' FROM general.`project_title` WHERE title_code = '" & code_ & "' ")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Function GetProfileID(ByVal profileN As String) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT IFNULL(ID,0)'ID' FROM general.`systemprofile` WHERE Profilename = '" & profileN & "' ")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Function GetCounterIDbyReference(ByVal ref As String) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT IFNULL(ID,0)'ID' FROM accounting.countered_rr WHERE Counter_Number ='" & ref & "' ")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function
    'Function GetCounterIDbySI(ByVal ref As String) As Integer
    '    Try
    '        Dim x = DB.QueryandReturnTable("SELECT IFNULL(ID,0)'ID' FROM accounting.countered_rr WHERE SI ='" & ref & "' ")
    '        If x.Rows.Count > 0 Then
    '            Return x(0)(0)
    '        End If
    '    Catch ex As Exception
    '        Return 0
    '    End Try
    'End Function
    Function GetPoIdbyRef(ByVal ref As String, COMPID As Integer) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT IFNULL(po_id,0)'ID' FROM procurement.`purchase_order_main` WHERE po_number = '" & ref & "' AND title_id='" & COMPID & "'")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Function GetPoAmntbyRef(ByVal ref As String) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT RRAMT FROM procurement.purchase_order_main LEFT JOIN (SELECT SUM(received_qty * receiving_details.net) 'RRAMT',po_id ,receipt_id  FROM procurement.receiving_details WHERE STATUS='Active'   AND     countered='NC'  GROUP BY receipt_id ) X ON x.po_id=purchase_order_main.po_id WHERE purchase_order_main.po_number ='" & ref & "'  GROUP BY purchase_order_main.po_id   ")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Function GetCounterGUID(ByVal ref As String) As String
        Try
            Dim x = DB.QueryandReturnTable("SELECT GUID FROM accounting.`countered_rr` WHERE Counter_number ='" & ref & "'")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function
    Function GetPayeeIDbyName(ByVal Name_ As String) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT ID FROM accounting.`payee` WHERE PAYEE_NAME ='" & Name_ & "' ")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Function GetAtcDescription(ByVal code_ As String) As String
        Try
            Dim x = DB.QueryandReturnTable("SELECT Description FROM accounting.`atc_setup` WHERE ATC ='" & code_ & "'")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function
    Function GetAtcRate(ByVal code_ As String) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT rate FROM accounting.`atc_setup` WHERE ATC = '" & code_ & "' ")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Function GetAccountIDbyCode(ByVal code_ As String) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT IFNULL(account_id,0)'ID' FROM accounting.`account_setup` WHERE account_code ='" & code_ & "' ")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Function GetItemCategory(ByVal itemID As Integer) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT category_id FROM procurement.`setup_group_detail` WHERE group_detail_id ='" & itemID & "' ")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Function GetItemSubCategory(ByVal itemID As Integer) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT category_id FROM procurement.`setup_group_detail` WHERE group_detail_id ='" & itemID & "' ")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Function GetItemUoM(ByVal itemID As Integer) As String
        Try
            Dim x = DB.QueryandReturnTable("SELECT UOM FROM procurement.`setup_qty_packing` WHERE item_id ='" & itemID & "' ")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function
    Function GetPoLogId(ByVal poID As Integer) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT unit_measure FROM procurement.`setup_group_detail` WHERE group_detail_id ='" & poID & "' ")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Function GetAcId(ByVal code_ As String) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT account_id FROM account_setup WHERE ActCode = '" & code_ & "' ")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Function GetAcIdbyName(ByVal desc As String) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT account_id FROM account_setup WHERE account_description = '" & desc & "' ")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Function GetAcIdbyShortCode(ByVal Code As String) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT account_id FROM account_setup WHERE ActCode = '" & Code.Trim.RSQ & "' ")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Function GetPayeeId(ByVal bid As Integer) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT IFNULL(ID,0)'ID' FROM accounting.`payee` WHERE BID = '" & bid & "' ")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Function GetCompanyIdByName(ByVal name_ As String) As Integer
        Try
            Dim x = DB.QueryandReturnTable("SELECT title_id FROM general.`project_title` WHERE title_name LIKE '%" & name_ & "%' ")
            If x.Rows.Count > 0 Then
                Return x(0)(0)
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function
End Module
