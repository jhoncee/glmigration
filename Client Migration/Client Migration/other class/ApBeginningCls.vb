Public Class ApBeginningCls
    Property NO As Integer
    Property COUNTEREDDATE As String
    Property COUNTERSLIPNUM As String
    Property SALESINVOICENO As String
    Property SALESINVOICEDATE As String
    Property APVNUMBER As String
    Property APVDATE As String
    Property DUEDATE As String
    Property VATTYPE As String
    Property TOTALAMOUNT As Decimal
    Property COMPANYNAME As String
    Property PROJECTCODE As String
    Property PROJECTNAME As String
    Property VENDORCODE As String
    Property VENDORNAME As String
    Property PAYEENAME As String
    Property APCODE As String
    Property APACOUNT As String
    Property APAMOUNT As Decimal
    Property ATCCODE As String
    Property ATCRATE As String
    Property WTAXCODE As String
    Property WTAXACCOUNT As String
    Property WTAXAMOUNT As Decimal
    Property TAXCODE As String
    Property TAXACCOUNT As String
    Property TAXAMOUNT As Decimal
    Property COMPID As Integer
    Property PRJID As Integer
    Sub New(no_ As Integer, counterDate_ As Date, counterSlipNum_ As String, salesInvoice_ As String, salesInvoiceDate_ As Date,
            apNumber_ As String, apDate_ As Date, dueDate_ As Date, vat_ As String, totalAmnt_ As Decimal, companyName_ As String,
            projectCode_ As String, projectN_ As String, vendorCode_ As String, vendorName_ As String, payee_ As String,
            apCode_ As String, apAccount_ As String, apAmnt As Decimal, atcCode_ As String, atcRate_ As String, wtaxCode_ As String, wtaxtAccount_ As String,
            wtaxAmnt As Decimal, taxCOde_ As String, taxAccount_ As String, taxAmnt_ As Decimal, COMPID As Integer, PRJID As Integer)
        Try
            Me.NO = no_
            Me.COUNTEREDDATE = counterDate_
            Me.COUNTERSLIPNUM = counterSlipNum_
            Me.SALESINVOICENO = salesInvoice_
            Me.SALESINVOICEDATE = salesInvoiceDate_
            Me.APVNUMBER = apNumber_
            Me.APVDATE = apDate_
            Me.DUEDATE = dueDate_
            Me.VATTYPE = vat_
            Me.TOTALAMOUNT = totalAmnt_
            Me.COMPANYNAME = companyName_
            Me.PROJECTCODE = projectCode_
            Me.PROJECTNAME = projectN_
            Me.VENDORCODE = vendorCode_
            Me.VENDORNAME = vendorName_
            Me.PAYEENAME = payee_
            Me.APCODE = apCode_
            Me.APACOUNT = apAccount_
            Me.APAMOUNT = apAmnt
            Me.ATCCODE = atcCode_
            Me.ATCRATE = atcRate_
            Me.WTAXCODE = wtaxCode_
            Me.WTAXACCOUNT = wtaxtAccount_
            Me.WTAXAMOUNT = wtaxAmnt
            Me.TAXCODE = taxCOde_
            Me.TAXACCOUNT = taxAccount_
            Me.TAXAMOUNT = taxAmnt_
            Me.COMPID = COMPID
            Me.PRJID = PRJID
        Catch ex As Exception
            MessageBoxError2(ex)
        End Try
    End Sub
End Class
