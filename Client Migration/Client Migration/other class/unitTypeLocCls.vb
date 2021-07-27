Public Class unitTypeCls
    Property NO As Integer
    Property UNITTYPE As String

    Sub New(no_ As Integer, unittype_ As String)
        Try
            Me.NO = no_
            Me.UNITTYPE = unittype_
        Catch ex As Exception
            MessageBoxError2(ex)
        End Try
    End Sub
End Class
Public Class unitLocCls
    Property NO As Integer
    Property UNITLOCATION As String
    Sub New(no_ As Integer, unitLoc As String)
        Try
            Me.NO = no_
            Me.UNITLOCATION = unitLoc
        Catch ex As Exception
            MessageBoxError2(ex)
        End Try
    End Sub
End Class
Public Class unitSetupCls
    Property NO As Integer
    Property PROPERTYCODE As String
    Property PROPERTYNAME As String
    Property FLOORNAME As String
    Property UNITCODE As String
    Property FLOORAREA As String
    Property UNITAREA As String
    Property PRICEPERSQM As Decimal
    Property UNITLOCATION As String
    Property UNITTYPE As String
    Property VAT As String
    Property ZONALVALUE As Decimal
    Property FAIRMARKETVALUE As Decimal
    Property FITOUT As Decimal
    Property MOVEINFEES As String
    Property TRANSFERCHARGE As String
    Property CONTRACTPRICE As Decimal
    Property SETUPFUND As Decimal
    Sub New(no_ As Integer, propertycode_ As String, propertyName_ As String, floor_ As String, unitCode_ As String,
            floorArea_ As String, unitArea_ As String, pricepersqm_ As Decimal, unitLoc As String, unitType_ As String, vat_ As String,
            zonalValue_ As Decimal, fairmarketvalue_ As Decimal, fitout_ As Decimal, moveinfees_ As String, transferCharge_ As String,
            contractPrice_ As Decimal, setupFund_ As Decimal)
        Try
            Me.NO = no_
            Me.PROPERTYCODE = propertycode_
            Me.PROPERTYNAME = propertyName_
            Me.FLOORNAME = floor_
            Me.UNITCODE = unitCode_
            Me.FLOORAREA = floorArea_
            Me.UNITAREA = unitArea_
            Me.PRICEPERSQM = pricepersqm_
            Me.UNITLOCATION = unitLoc
            Me.UNITTYPE = unitType_
            Me.VAT = vat_
            Me.ZONALVALUE = zonalValue_
            Me.FAIRMARKETVALUE = fairmarketvalue_
            Me.FITOUT = fitout_
            Me.MOVEINFEES = moveinfees_
            Me.TRANSFERCHARGE = transferCharge_
            Me.CONTRACTPRICE = contractPrice_
            Me.SETUPFUND = setupFund_
        Catch ex As Exception
            MessageBoxError2(ex)
        End Try
    End Sub
End Class