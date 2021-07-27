Module business_saving
    Public Sub save_business_list(ByVal fullN As String, ByVal fn As String, ByVal mn As String, ByVal ln As String, ByVal suf As String,
                                  ByVal tinNumber As String, ByVal termType As String, ByVal vat_ As String, ByVal vnv_ As String,
                                  ByVal creditLimit As String, ByVal type_ As String, ByVal accNum As String, ByVal class_ As String,
                                  ByVal categoryType As String, ByVal secN As String, ByVal _2303_ As String, ByVal atc_ As String,
                                  ByVal repTeam As String, ByVal inhouse_ As String, ByVal managerID_ As Integer, ByVal brokerID_ As Integer,
                                  ByVal licNum_ As String, ByVal hlurb_ As String, ByVal accredited As String, ByVal email_ As String,
                                  ByVal btype As String, ByVal vat_percent As String, contractorType As String)
        Try
            Dim business_list As New Dictionary(Of String, String)
            With business_list
                .Add("business_fullname", fullN)
                .Add("trade_nickname", fullN)
                .Add("tin_number", tinNumber)
                .Add("term_type", termType)
                .Add("vat", If(vat_ = "", "Non-Vat", vat_))
                .Add("vat_percentage", vat_percent)
                .Add("vnv", vnv_)
                .Add("title_id", 1)
                .Add("type", type_)
                .Add("ACCOUNTNUMBER", accNum)
                .Add("EmailAdd", email_)
                .Add("imported_from_excel", "1")
                ' .Add("Contractor_Type", contractorType)
                If type_ = "CLIENT" Then
                    .Add("Business_type", btype)
                End If
                If type_ <> "SUPPLIER" And btype <> "Corporate" Then
                    If type_ <> "CONTRACTOR" Then
                        .Add("firstname", fn)
                        .Add("middlename", mn)
                        .Add("lastname", ln)
                        .Add("suffix", suf)
                    End If
                End If

                If type_ = "MANAGER" Then
                    .Add("rep_team", repTeam)
                End If

                If type_ = "MANAGER" Or type_ = "BROKER" Then
                    .Add("IsAccredited", accredited)
                    .Add("LicNum", brokerID_)
                    .Add("HLURB", hlurb_)
                End If

                If type_ = "AGENT" Then
                    .Add("ManagerID", managerID_)
                    .Add("BrokerID", brokerID_)
                End If

                If type_ = "MANAGER" Or type_ = "AGENT" Then
                    .Add("inhouse", inhouse_)
                End If

                If type_ = "SUPPLIER" Or type_ = "CONTRACTOR" Then
                    .Add("categoryType", categoryType)
                End If

                If type_ = "SUPPLIER" Then
                    .Add("classification", class_)
                    .Add("secnumber", secN)
                    .Add("_2303", _2303_)
                    .Add("credit_limit", If(creditLimit = "", 0, creditLimit))
                End If

                If type_ = "SUPPLIER" Or type_ = "BROKER" Then
                    .Add("ATC", atc_)
                End If
            End With
            DB.Insert(GENERALDBASE & ".business_list", business_list)

        Catch ex As Exception
            DB.RollBackAction()
            MSG(ex.Message)
        End Try
    End Sub
    Public Sub save_business_dtl(ByVal bussinesNumber As Integer, ByVal typeOfAddress As String, ByVal street_ As String,
                                 ByVal barangay_ As String, ByVal city_ As String, ByVal province_ As String,
                                 ByVal zip_ As String, ByVal country_ As String, ByVal defaultAdd As String)
        Try
            Dim business_dtl As New Dictionary(Of String, String)
            With business_dtl
                .Add("BusinessNumber", bussinesNumber)
                .Add("TypeOfAddress", typeOfAddress)
                .Add("street", street_)
                .Add("barangay", barangay_)
                .Add("city", city_)
                .Add("province", province_)
                .Add("zip", zip_)
                .Add("country", country_)
                .Add("DefaultAddress", defaultAdd)
                .Add("imported_from_excel", "1")
            End With
            DB.Insert(GENERALDBASE & ".business_list_detail", business_dtl)
        Catch ex As Exception
            DB.RollBackAction()
            MSG(ex.Message)
        End Try
    End Sub
    Public Sub save_business_contact(ByVal businessNumber As Integer, ByVal contactType_ As String, ByVal country_code As String,
                                     ByVal areacode_ As String, ByVal contactN As String, ByVal number_ As String, ByVal pos_ As String,
                                     ByVal email_ As String)
        Try
            Dim business_contact As New Dictionary(Of String, String)
            With business_contact
                .Add("BusinessNumber", businessNumber)
                .Add("contactType", contactType_)
                .Add("CountryCode", country_code)
                .Add("AreaCode", areacode_)
                .Add("ContactName", contactN)
                .Add("Number", number_)
                .Add("Position", pos_)
                .Add("EmailAdd", email_)
                .Add("imported_from_excel", "1")
            End With
            DB.Insert(GENERALDBASE & ".businesscontacts", business_contact)
        Catch ex As Exception
            DB.RollBackAction()
            MSG(ex.Message)
        End Try
    End Sub
    Public Sub save_payee(ByVal businessNumber As Integer, ByVal payeeN As String, ByVal auto_ As Integer)
        Try
            Dim payee_ As New Dictionary(Of String, String)
            With payee_
                .Add("BID", businessNumber)
                .Add("PAYEE_NAME", payeeN)
                .Add("Auto", auto_)
            End With
            DB.Insert(ACCOUNTINGDB & ".payee", payee_)
        Catch ex As Exception
            DB.RollBackAction()
            MSG(ex.Message)
        End Try
    End Sub
    Public Sub update_managerBroker(ByVal bid_ As Integer, ByVal brok_ As Integer)
        Try
            Dim mb As New Dictionary(Of String, String)
            With mb
                .Add("BrokerID", brok_)
            End With
            DB.Update(GENERALDBASE & ".business_list", mb, "business_number = '" & bid_ & "'")
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Sub

    Public Sub save_buyerInfo(guid_ As String, bid_ As String, buyerN As String, buyerFirstN As String, cv As String, uid As String, cs As String)
        Try
            Dim bi As New Dictionary(Of String, String)
            With bi
                .Add("GUID", guid_)
                .Add("business_id", bid_)
                .Add("BUYERFULLNAME", buyerN)
                .Add("buyer_firstname", buyerFirstN)
                .Add("citizenship", cs)
                .Add("civil_status", cv)
                .Add("user_id", uid)
                .Add("imported_from_excel", "1")
            End With
            DB.Insert(PropmanagementdBASE & ".buyersinfomain", bi)
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Sub

    Public Sub save_buyerProject(guid_ As String, prjid As Integer, prjN As String, floorID As Integer, floorNum As String, unitID As Integer, unitN As String, totalContractP As Decimal)
        Try
            Dim bp As New Dictionary(Of String, String)
            With bp
                .Add("GUID", guid_)
                .Add("project_id", prjid)
                .Add("project", prjN)
                .Add("FloorNum", floorID)
                .Add("FloorName", floorNum)
                .Add("unit_id", unitID)
                .Add("unit", unitN)
                .Add("total_contract_price", totalContractP)
            End With
            DB.Insert(PropmanagementdBASE & ".buyersinfoproject", bp)
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Sub
    Public Sub save_Manager_Team(team_ As String)
        Try
            Dim teams As New Dictionary(Of String, String)
            With teams
                .Add("team_name", team_)
            End With
            DB.Insert(GENERALDBASE & ".business_team", teams)
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Sub
End Module
