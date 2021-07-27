Public Class beginningCls
    Property NO As Integer
    Property ACCOUNTCODE As String
    Property ACCOUNTNAME As String
    Property DEBIT As Decimal
    Property CREDIT As Decimal
    Sub New(no_ As Integer, code_ As String, name_ As String, debit_ As Decimal, credit_ As Decimal)
        Try
            Me.NO = no_
            Me.ACCOUNTCODE = code_
            Me.ACCOUNTNAME = name_
            Me.DEBIT = debit_
            Me.CREDIT = credit_

        Catch ex As Exception
            MessageBoxError2(ex)
        End Try
    End Sub

End Class
