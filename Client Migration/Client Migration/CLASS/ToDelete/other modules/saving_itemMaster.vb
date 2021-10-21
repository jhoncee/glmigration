Module saving_itemMaster
    Public Sub save_itemMaster(ByVal itemCode As String, group_ As Integer, category_ As Integer, desc_ As String, type_ As String,
                               ByVal titleID As Integer, ByVal branch_ As Integer, Sheet As String, uom As String)
        Try
            Dim items_ As New Dictionary(Of String, String)
            With items_
                .Add("item_code", itemCode)
                .Add("group_id", group_)
                .Add("base1_uom", uom)
                .Add("unit_measure", uom)
                .Add("category_id", category_)
                .Add("description", desc_)
                .Add("title_id", branch_)
                .Add("inventory_type", type_)
                .Add("title_id1", titleID)
                .Add("imported_from_excel", "1")
                .Add("Sheet", Sheet)
            End With
            DB.Insert(PROCUREMENTDBASE & ".setup_group_detail", items_)
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Sub
    Public Sub save_setup_qty_packing(ByVal itemId As Integer, ByVal uom_ As String)
        Try
            Dim qty As New Dictionary(Of String, String)
            With qty
                .Add("item_id", itemId)
                .Add("BASE", "Base 1")
                .Add("UOM", uom_)
                .Add("QTY", "1")
            End With
            DB.Insert(PROCUREMENTDBASE & ".setup_qty_packing", qty)
        Catch ex As Exception
            DB.RollBackAction(ex.Message)
        End Try
    End Sub
End Module
