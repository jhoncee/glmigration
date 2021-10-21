Module saving
    Public Sub SaveProperty(ByVal propertyN As String, ByVal propertyNo As String, ByVal compID As Integer, ByVal remarks_ As String,
                             ByVal projectID_ As Integer, ByVal reservation_ As Decimal, MAINID As Integer)
        Try
            Dim property_setup As New Dictionary(Of String, String)
            With property_setup
                .Add("PropertyName", propertyN)
                .Add("PropertyNo", propertyNo)
                .Add("TitleID", compID)
                .Add("MAINID", MAINID)
                .Add("Remarks", remarks_)
                .Add("projectid", projectID_)
                .Add("PropertySubType", "4")
                .Add("ReservationAMt", reservation_)
                .Add("imported_from_excel", "1")
            End With
            DB.Insert(PropmanagementdBASE & ".tbl_property", property_setup)
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Sub

    Public Sub SaveSetupProject(ByVal projectN As String, ByVal remarks_ As String, ByVal projectNo As String, ByVal compID As Integer, ByVal mainID As Integer,
                                    ByVal projects_ As String, _location As String)
        Try
            Dim project_ As New Dictionary(Of String, String)
            With project_
                .Add("project", projects_)
                .Add("project_name", projectN)
                .Add("project_type", projects_)
                .Add("project_remarks", remarks_)
                .Add("project_no", projectNo)
                .Add("TITLE_ID", compID)
                .Add("Main_office_id", mainID)
                .Add("project_location", _location)
                .Add("imported_from_excel", "1")
            End With
            DB.Insert(GENERALDBASE & ".setup_project", project_)
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Sub
    Public Sub update_property(ByVal prjid As Integer, ByVal no_ As String, ByVal street_ As String, ByVal brgy_ As String, ByVal city_ As String,
                               ByVal zipcode_ As String, ByVal province_ As String)
        Try
            Dim property_setup As New Dictionary(Of String, String)
            With property_setup
                .Add("No", no_)
                .Add("Street", street_)
                .Add("Brgy", brgy_)
                .Add("City", city_)
                .Add("Zipcode", zipcode_)
                .Add("Province", province_)
            End With
            DB.Update(PropmanagementdBASE & ".tbl_property", property_setup, "projectid = '" & prjid & "'")
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Sub

    Public Sub update_project_address(ByVal projID As Integer, ByVal address_ As String)
        Try
            Dim project_ As New Dictionary(Of String, String)
            With project_
                .Add("project_location", address_)
            End With
            DB.Update(GENERALDBASE & ".setup_project", project_, "project_id = '" & projID & "'")
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Sub

    Public Sub save_unitType(ByVal unitType_ As String, ByVal propID As Integer, ByVal userID As Integer)
        Try
            Dim unitType As New Dictionary(Of String, String)
            With unitType
                .Add("unit_type", unitType_)
                .Add("property_id", propID)
                .Add("user_id", userID)
                .Add("imported_from_excel", "1")
            End With
            DB.Insert(PropmanagementdBASE & ".tbl_property_unit_type", unitType)
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Sub
    Public Sub save_unitLocation(ByVal unitLoc_ As String, ByVal propID As Integer, ByVal userID As Integer)
        Try
            Dim unitLoc As New Dictionary(Of String, String)
            With unitLoc
                .Add("location", unitLoc_)
                .Add("property_id", propID)
                .Add("user_id", userID)
                .Add("imported_from_excel", "1")
            End With
            DB.Insert(PropmanagementdBASE & ".tbl_property_unit_location", unitLoc)
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Sub

    Public Sub save_floor(ByVal prjid_ As Integer, ByVal floors_ As String, ByVal floorName As String, ByVal floorArea_ As String)
        Try
            Dim floor_ As New Dictionary(Of String, String)
            With floor_
                .Add("PRJID", prjid_)
                .Add("floor", floors_)
                .Add("floor_name", floorName)
                .Add("FloorArea", floorArea_)
                .Add("imported_from_excel", "1")
            End With
            DB.Insert(PropmanagementdBASE & ".tbl_floor", floor_)
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Sub
    Public Sub save_users(ByVal userN As String, ByVal pword As String, ByVal empName As String,
                          ByVal empID As String, ByVal deptID As String, ByVal email_ As String)
        Try
            Dim user_ As New Dictionary(Of String, String)
            With user_
                .Add("UserName_User", userN)
                .Add("Password_User", pword)
                .Add("UserFull_name", empName)
                .Add("Employee_id", empID)
                .Add("department_id", deptID)
                .Add("email_address", email_)
            End With
            DB.Insert(GENERALDBASE & ".users", user_)
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Sub
    Public Sub save_Company(ByVal titleN As String, ByVal address_ As String, ByVal contactN_ As String, ByVal code_ As String, ByVal tin_ As String, ByVal zip_ As String)
        Try
            Dim comp As New Dictionary(Of String, String)
            With comp
                .Add("title_name", titleN)
                .Add("title_desc1", address_)
                .Add("title_desc2", contactN_)
                .Add("title_code", code_)
                .Add("tin", tin_)
                .Add("ZIP", zip_)
            End With
            DB.Insert(GENERALDBASE & ".project_title", comp)
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Sub

    Public Sub save_department(deptCode As String, deptName As String, projid As Integer, compID As Integer)
        Try
            Dim dept As New Dictionary(Of String, String)
            With dept
                .Add("DeptCode_Dept", deptCode)
                .Add("Name_Dept", deptName)
                .Add("project_id", projid)
                .Add("titleid", compID)
            End With
            DB.Insert(HRDBASE & ".department", dept)
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Sub

    Public Sub save_profile(profileN As String, uid As Integer)
        Try
            Dim profile As New Dictionary(Of String, String)
            With profile
                .Add("Profilename", profileN)
                .Add("AddedByUID", uid)
            End With
            DB.Insert(GENERALDBASE & ".systemprofile", profile)
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Sub
    Public Sub save_countering(counterNum_ As String, counterDate_ As Date, siDate_ As Date, siNum_ As String, supID_ As Integer,
                               vendorName_ As String, guid_ As String, user_ID_ As Integer, visible_ As String, POID_ As Integer,
                               poNum_ As String, poAmnt_ As Decimal, isVat_ As String, siAmnt As Decimal, titleID_ As Integer, Prjid_ As Integer,
                               GS_ As String, remarks_ As String, type_ As String, Sheet As String)
        Try
            Dim countering As New Dictionary(Of String, String)
            With countering
                .Add("COUNTER_NUMBER", counterNum_)
                .Add("COUNTER_DATES", counterDate_.ToMysqlFormat)
                .Add("SI_Date", siDate_.ToMysqlFormat)
                .Add("SI", siNum_)
                .Add("SIAMT", siAmnt)
                .Add("SUPID", supID_)
                .Add("VendorName", vendorName_)
                .Add("GUID", guid_)
                .Add("USER_ID", user_ID_)
                .Add("IsVisibleInTranmittal", visible_)
                .Add("POID", POID_)
                .Add("PONUM", poNum_)
                .Add("POAMT", poAmnt_)
                .Add("IsVat", isVat_)
                .Add("TITLEID", titleID_)
                .Add("PRJID", Prjid_)
                .Add("GS", GS_)
                .Add("Remarks", remarks_)
                .Add("imported_from_excel", "1")
                .Add("Types", type_)
                .Add("TransType", type_)
                .Add("Sheet", Sheet)
                'Auto Transmit 
                'transmittal 
                .Add("TransmitGUID", guid_)
                .Add("TRANS_REF_NUM", "")
                .Add("transmital_datetime", Now.ToMysqlFormat)
                .Add("STATUS", "Posted")
                .Add("TRANSMITBY", 1)
            End With
            DB.Insert(ACCOUNTINGDB & ".countered_rr", countering)
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Sub
    Public Sub save_apBeginning(apDate_ As Date, apNum As String, duedate_ As Date, sias_ As String, siDate_ As Date, amnt As Decimal, atc_ As String,
                                rate_ As String, atc_description_ As String, tinNum_ As String, supID_ As Integer, vendorN_ As String,
                                payeeID_ As Integer, guid_ As String, jid_ As Integer, titleID_ As Integer, prjid_ As Integer, uid_ As Integer,
                                sourceGUID_ As String, transN_ As String, address_ As String)
        Try
            Dim ap As New Dictionary(Of String, String)
            With ap
                .Add("APV_DATE", apDate_.ToMysqlFormat)
                .Add("APVNUM", apNum)
                .Add("DUEDATE", duedate_.ToMysqlFormat)
                .Add("SIAS", sias_)
                .Add("SIDate", siDate_.ToMysqlFormat)
                .Add("AMT", amnt)
                '.Add("ATC", atc_)
                '.Add("RATE", rate_)
                '.Add("ATC_DESCRIPTION", atc_description_)
                .Add("TINNUM", tinNum_)
                .Add("SUPID", supID_)
                .Add("VendorName", vendorN_)
                .Add("PAYEEID", payeeID_)
                .Add("GUID", guid_)
                .Add("JID", jid_)
                .Add("TITLEID", titleID_)
                .Add("PRJID", prjid_)
                .Add("UID", uid_)
                .Add("SourceGUID", sourceGUID_)
                .Add("TRANSNAME", transN_)
                .Add("Address", address_)
                .Add("imported_from_excel", "1")
            End With
            DB.Insert(ACCOUNTINGDB & ".apv_records", ap)
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Sub

    Public Sub save_journal_main(seriesN As String, transType_ As String, type_ As String, transDate_ As Date, status_ As String,
                                 uid_ As Integer, nameID As Integer, guid_ As String, vendorID_ As Integer, caption As String, Sheet As String)
        Try
            Dim jm As New Dictionary(Of String, String)
            With jm
                .Add("SERIESNUMBER", seriesN)
                .Add("trans_type", transType_)
                .Add("FrmCaption", caption)
                .Add("Type", type_)
                .Add("trans_date", transDate_.ToMysqlFormat)
                .Add("STATUS", status_)
                .Add("userid", uid_)
                .Add("name_id", nameID)
                .Add("GroupEntry", guid_)
                .Add("VENDORID", vendorID_)
                .Add("imported_from_excel", "1")
                .Add("Sheet", Sheet.RSQ)
            End With
            DB.Insert(ACCOUNTINGDB & ".journal_main", jm)
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Sub

    Public Sub save_journal_detail(jid As Integer, acid As Integer, subsidiary_ As String, subsidiaryType As String,
                                   subaccountID_ As Integer, guid_ As String, cmp As Integer, brn As Integer, chr As Integer,
                                   invRef As String, debit As Decimal, credit As Decimal)
        Try
            Dim jd As New Dictionary(Of String, String)
            With jd
                .Add("journal_id", jid)
                .Add("account_id", acid)
                .Add("subsidiary", subsidiary_)
                .Add("SUBACCOUNTID", subaccountID_)
                .Add("subsidiary_type", subsidiaryType)
                .Add("GroupEntryD", guid_)
                .Add("_CMP", cmp)
                .Add("_BRN", brn)
                .Add("_CHR", chr)
                .Add("INV_REF", invRef)
                .Add("DEBITS", debit)
                .Add("CREDITS", credit)
                '.Add("DueDate", dueDate_.ToMysqlFormat)
                'If atcCode_ <> "" Then
                '    .Add("ATC_CODE", atcCode_)
                '    .Add("ATC_RATE", atcRate)
                '    .Add("ATC_DESCRIPTION", atcDesc)
                'End If
                .Add("imported_from_excel", "1")
            End With
            DB.Insert(ACCOUNTINGDB & ".journal_detail", jd)
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Sub
    Public Sub save_vendor_atc(desc_ As String, atc_ As String, acc_ As Integer, caccid_ As Integer, cid_ As Integer, bid_ As Integer)
        Try
            Dim atcV As New Dictionary(Of String, String)
            With atcV
                .Add("DESCRIPTION", desc_)
                .Add("ATC", atc_)
                .Add("RATE", acc_)
                .Add("ACID", caccid_)
                .Add("CACID", cid_)
                .Add("CID", bid_)
                .Add("BID", bid_)
            End With
            DB.Insert(ACCOUNTINGDB & ".atc_vendor_default", atcV)
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Sub
End Module
