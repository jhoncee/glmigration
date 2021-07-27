Public Class Defaults
    Class Properties
        Property [Property] As String
        Property ACID As Integer
        Sub New([Property] As String, ACID As Integer)
            Me.Property = [Property]
            Me.ACID = ACID
        End Sub
    End Class
    Private Account As List(Of Properties)
    Public Sub LoadDefaultsAccount()
        Try
            Account = New List(Of Properties)
            Dim T As New DataTable
            T = Datasource("select * from accounting.defaults")
            For Each i As DataRow In T.Rows
                Account.Add(New Properties(i("Property").ToString, CInt(i("ACID"))))
            Next
        Catch ex As Exception
        End Try
    End Sub
    Sub New()
        LoadDefaultsAccount()
    End Sub
    Public Function GetAccount(PropertyName As String) As Integer
        Try
            Dim F = From i In Account
                    Where i.Property.ToLower = PropertyName.ToLower
                    Select i.ACID

            If F.Count = 0 Then Return 0
            Return F(0)
        Catch ex As Exception
            Return 0
        End Try
    End Function
End Class
