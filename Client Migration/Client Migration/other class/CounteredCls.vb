Public Class CounteredCls
    Property NO As Integer
    Property COUNTERDATE As Date
    Property VENDORCODE As String
    Property VENDORNAME As String
    Property SIDATE As Date
    Property SINUMBER As String
    Property SIAMOUNT As Decimal
    Property PONUMBER As String
    Property COMPID As Integer
    Property PRJID As Integer

    Sub New(no_ As Integer, counterdate_ As Date, vendorCode_ As String, vendorName_ As String, siDate_ As Date,
            siNumber_ As String, siAmount_ As Decimal, poNum_ As String, COMPID As Integer, PRJID As Integer)
        Try
            Me.NO = no_
            Me.COUNTERDATE = counterdate_
            Me.VENDORCODE = vendorCode_
            Me.VENDORNAME = vendorName_
            Me.SIDATE = siDate_
            Me.SINUMBER = siNumber_
            Me.SIAMOUNT = siAmount_
            Me.PONUMBER = poNum_
            Me.COMPID = COMPID
            Me.PRJID = PRJID
        Catch ex As Exception
            MessageBoxError2(ex)
        End Try
    End Sub
End Class
