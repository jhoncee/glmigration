Public Class ValidFieldName
    Public UseDefault As Boolean = True
    Private Sub ValidFieldName_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If UseDefault Then
            Dim FieldName As String() = New String() {"Unit Code *", "Buyer Name *", "Payment Scheme",
                         "Charge Name *", "Particulars *", "Due Date *", "Due Amount *", "MOP *",
                         "AR Date *", "AR # *", "Bank", "Check Date ",
                         "Check #", "Amount", "Account Ref ", "OR Date", "OR #", "CR Date",
                         "CR #", "Amount2", "Payments Status", "Check Status", "Posting Date"}

            MemoEdit1.Text = ""
            For Each i In FieldName
                MemoEdit1.Text = MemoEdit1.Text & i & vbNewLine
            Next
        End If
    End Sub
End Class