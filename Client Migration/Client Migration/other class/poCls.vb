Public Class poCls
    Property NO As Integer
    Property PONUMBER As String
    Property PODATE As Date
    Property ISVAT As String
    Property PROJECTNAME As String
    Property PROJECTCODE As String
    Property VENDORNAME As String
    Property VENDORCODE As String
    Private _DELIVERYTYPE As String
    Public Property DELIVERYTYPE() As String
        Get
            Return _DELIVERYTYPE
        End Get
        Set(ByVal value As String)
            If value = "" Then
                _DELIVERYTYPE = "Until"
            Else
                _DELIVERYTYPE = value
            End If
        End Set
    End Property
    Property DELIVERYDATE As Date
    Property REQUESTEDBY As String
    Property PRIMARYRECEIVER As String
    Property REMARKS As String
    Property SECONDARYRECEIVER As String
    Property ITEMCODE As String
    Property DESCRIPTION As String
    Property SPECIFICATION As String
    Property QTY As String
    Property UNITPRICE As Decimal
    Property SUBTOTAL As Decimal
    Property COMPID As Integer
    Property PRJID As Integer
    Sub New(no_ As Integer, poNum As String, poDate_ As Date, vat_ As String, projectN As String, projectC As String,
            vendorN As String, vendorC As String, deliveryType_ As String, deliveryDate As Date, requestBy As String,
            primaryReceiver_ As String, remarks_ As String, secondaryReceiver_ As String, itemC As String, description_ As String,
            specification_ As String, qty_ As String, unitP As Decimal, subTotal_ As Decimal, COMPID As Integer, PRJID As Integer)
        Try
            Me.NO = no_
            Me.PONUMBER = poNum
            Me.PODATE = poDate_
            Me.ISVAT = vat_
            Me.PROJECTNAME = projectN
            Me.PROJECTCODE = projectC
            Me.VENDORNAME = vendorN
            Me.VENDORCODE = vendorC
            Me.DELIVERYTYPE = deliveryType_
            Me.DELIVERYDATE = deliveryDate
            Me.REQUESTEDBY = requestBy
            Me.PRIMARYRECEIVER = primaryReceiver_
            Me.REMARKS = remarks_
            Me.SECONDARYRECEIVER = secondaryReceiver_
            Me.ITEMCODE = itemC
            Me.DESCRIPTION = description_
            Me.SPECIFICATION = specification_
            Me.QTY = qty_
            Me.UNITPRICE = unitP
            Me.SUBTOTAL = subTotal_
            Me.COMPID = COMPID
            Me.PRJID = PRJID
        Catch ex As Exception
            MessageBoxError2(ex)
        End Try
    End Sub
End Class
Public Class PoAmntCls
    Property NO As Integer
    Property PONUMBER As String
    Property GROSSAMOUNT As Decimal
    Property CHARGES As Decimal
    Property DISCOUNT As String
    Property NETAMOUNT As Decimal
    Sub New(no_ As Integer, poNum As String, grossAmnt_ As Decimal, charges_ As Decimal, discount_ As String, netAmnt_ As Decimal)
        Try
            Me.NO = no_
            Me.PONUMBER = poNum
            Me.GROSSAMOUNT = grossAmnt_
            Me.CHARGES = charges_
            Me.DISCOUNT = discount_
            Me.NETAMOUNT = netAmnt_
        Catch ex As Exception
            MessageBoxError2(ex)
        End Try
    End Sub
End Class
