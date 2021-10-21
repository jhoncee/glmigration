Public Class validFieldNameBuyer
    Private Sub validFieldNameBuyer_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim FieldName As String() = New String() {"Unit Code *", "Buyer Name *", "List Price *", "Vat", "Total Selling Price", "Discount 1",
            "Discount 2", "Add Vat", "Net Selling Price", "Reservation", "Move-In",
            "Transfer Charge", "Loan Takeout", "Buyer Status"}

        For Each i In FieldName
            MemoEdit1.Text &= i & vbNewLine
        Next
    End Sub
End Class