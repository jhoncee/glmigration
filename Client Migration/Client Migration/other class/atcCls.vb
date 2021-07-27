Public Class atcCls
    Property NO As Integer
    Property ATC As String
    Property RATE As String
    Property DESCRIPTION As String
    Property TYPES As String
    Property BIRFORMS As String

    Sub New(no_ As Integer, atc_ As String, rate_ As String, description_ As String, type_ As String, bir_ As String)
        Try
            Me.NO = no_
            Me.ATC = atc_
            Me.RATE = rate_
            Me.DESCRIPTION = description_
            Me.TYPES = type_
            Me.BIRFORMS = bir_
        Catch ex As Exception
            MessageBoxError2(ex)
        End Try
    End Sub
End Class
