Public Class managerBrokerCls
    Property NO As Integer
    Property MANAGERCODE As String
    Property BROKERCODE As String
    Property BROKERNAME As String

    Sub New(no_ As Integer, managerCode_ As String, brokerCode_ As String, brokerN As String)
        Try
            Me.NO = no_
            Me.MANAGERCODE = managerCode_
            Me.BROKERCODE = brokerCode_
            Me.BROKERNAME = brokerN
        Catch ex As Exception
            MessageBoxError2(ex)
        End Try
    End Sub
End Class
Public Class teamCls
    Property NO As Integer
    Property TEAM As String
    Sub New(no_ As Integer, team_ As String)
        Try
            Me.NO = no_
            Me.TEAM = team_
        Catch ex As Exception
            MessageBoxError2(ex)
        End Try
    End Sub
End Class
