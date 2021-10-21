Public Class ClsCalculator

    Private Shared NUMS() As String
    Private Shared res As Double
    Public Shared Function GET_COMPUTED_VALUE(ByVal NUMBER As String) As Decimal
        Try
            NUMBER = NUMBER.Trim
            If NUMBER.Contains("*") Then
                NUMS = NUMBER.Split("*")
                Return NUMS(0) * NUMS(1)
            ElseIf NUMBER.Contains("+") Then
                NUMS = NUMBER.Split("+")
                Return NUMS(0) + NUMS(1)
            ElseIf NUMBER.Contains("-") Then
                NUMS = NUMBER.Split("-")
                Return NUMS(0) - NUMS(1)
            ElseIf NUMBER.Contains("/") Then
                NUMS = NUMBER.Split("/")
                Return NUMS(0) / NUMS(1)
            ElseIf NUMBER.Contains("\") Then
                NUMS = NUMBER.Split("\")
                Return NUMS(0) / NUMS(1)
            ElseIf NUMBER.Contains("X") Then
                NUMS = NUMBER.Split("*")
                Return NUMS(0) * NUMS(1)
            Else
                Return NUMBER
            End If
        Catch ex As Exception
            Return NUMBER
        End Try
    End Function
End Class
