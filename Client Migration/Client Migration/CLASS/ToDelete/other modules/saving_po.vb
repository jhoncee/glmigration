Module saving_po
    Public Sub save_poMain(poNum_ As String, poDate As Date, dateNeeded As Date, supplierID_ As Integer, deliverTerm_ As String,
                           title_id_ As Integer, project_id_ As Integer, tinNum As String, draftNum As String,
                           PrimaryID As Integer, secondaryID As Integer, vat_ As String, itemCat As Integer,
                           orderDate_ As Date, budgetDate_ As Date, Sheet As String)
        Try
            Dim po As New Dictionary(Of String, String)
            With po
                .Add("po_number", poNum_)
                .Add("supplierID", supplierID_)
                .Add("deliverTerm", deliverTerm_)
                .Add("title_id", title_id_)
                .Add("project_id", project_id_)
                .Add("Status", "APPROVED")
                .Add("approvedate", Now.ToMysqlFormat)
                .Add("tin_number", tinNum)
                .Add("draft_number", "")
                .Add("deliverIDPrimary", PrimaryID)
                .Add("deliverIDSecondary", secondaryID)
                .Add("ISVAT", vat_)
                .Add("item_cat_id", itemCat)
                .Add("imported_from_excel", "1")
                .Add("Sheet", Sheet)
                If IsDate(poDate) Then
                    .Add("po_date", poDate.ToMysqlFormat)
                End If
                If IsDate(dateNeeded) Then
                    .Add("date_needed", dateNeeded.ToMysqlFormat)
                End If
                If IsDate(orderDate_) Then
                    .Add("orderdate", orderDate_.ToMysqlFormat)
                End If
                If IsDate(budgetDate_) Then
                    .Add("budgetdate", budgetDate_.ToMysqlFormat)
                End If

            End With
            DB.Insert(PROCUREMENTDBASE & ".purchase_order_main", po)
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Sub
    Public Sub save_poDtl(poid As Integer, itemNo_ As Integer, uom_ As String, quantity_ As String, unitCost As Decimal,
                          totalUnitCost As Decimal, gross_ As Decimal, vat_ As Decimal, net_ As Decimal, subCategory As Integer)
        Try
            Dim podtl As New Dictionary(Of String, String)
            With podtl
                .Add("po_id", poid)
                .Add("itemNo", itemNo_)
                .Add("uom", uom_)
                .Add("quantity", quantity_)
                .Add("unit_cost", unitCost)
                .Add("total_unitcost", totalUnitCost)
                .Add("GROSS", gross_)
                .Add("VAT", vat_)
                .Add("NET", net_)
                .Add("sub_category", subCategory)
            End With
            DB.Insert(PROCUREMENTDBASE & ".purchase_order_details", podtl)
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Sub
    Public Sub save_for_accountingEntry(poid_ As Integer, totalPO As Decimal, podate As Date)
        Try
            Dim acc As New Dictionary(Of String, String)
            With acc
                .Add("poid", poid_)
                .Add("total_po", totalPO)
                .Add("po_date", podate)
            End With
            DB.Insert(PROCUREMENTDBASE & ".tbl_accounting_entry", acc)
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Sub
    Public Sub saveTransMain(poid As Integer, reference As String, transDate_ As Date, supID As Integer,
                                        compID As Integer, prjid As Integer, budgetby_ As Integer,
                                        remarks_ As String, uid As Integer, ipAddress_ As String, hostN As String)
        Try
            Dim transM As New Dictionary(Of String, String)
            With transM
                .Add("TRANSID", poid)
                .Add("TRANSNO", reference)
                .Add("TRANSDATE", transDate_.ToMysqlFormat)
                .Add("TRANSTYPE", "PO")
                .Add("DRAFNO", reference)
                .Add("DRAFDATE", transDate_.ToMysqlFormat)
                .Add("SUPPLIERID", supID)
                .Add("TITLEID", compID)
                .Add("PROJID", prjid)
                .Add("BUDGETEDBY", budgetby_)
                .Add("REMARKS", remarks_)
                .Add("USERID", uid)
                .Add("IPADDRESS", ipAddress_)
                .Add("HOSTNAME", hostN)
            End With
            DB.Insert(PROCUREMENTDBASE & ".tbl_transaction_data_logs", transM)
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Sub
    Public Sub saveTransDtl(mainID_ As Integer, itemID_ As Integer, uom_ As String, qty As String,
                            ppu As Decimal, gross As Decimal, vat_ As Decimal, net_ As Decimal)
        Try
            Dim transDtl As New Dictionary(Of String, String)
            With transDtl
                .Add("MAINID", mainID_)
                .Add("ITEMID", itemID_)
                .Add("UOM", uom_)
                .Add("QTY", qty)
                .Add("PPU", ppu)
                .Add("GROSS", gross)
                .Add("VAT", vat_)
                .Add("NET", net_)
                .Add("TRANSTYPE", "PO")
            End With
            DB.Insert(PROCUREMENTDBASE & ".tbl_transaction_data_dtl_logs", transDtl)
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Sub
End Module
