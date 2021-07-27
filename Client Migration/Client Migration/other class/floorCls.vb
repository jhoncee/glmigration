Public Class floorCls
    Property NO As Integer
    Property PROPERTYCODE As String
    Property PROPERTYNAME As String
    Property BASEMENT As String
    Property FLOORNAME As String
    Property FLOORAREA As String
    Sub New(no_ As Integer, propertyCode_ As String, propertyName_ As String, basement_ As String, floorName_ As String, floorArea_ As String)
        Try
            Me.NO = no_
            Me.PROPERTYCODE = propertyCode_
            Me.PROPERTYNAME = propertyName_
            Me.BASEMENT = basement_
            Me.FLOORNAME = floorName_
            Me.FLOORAREA = floorArea_
        Catch ex As Exception
            MessageBoxError2(ex)
        End Try
    End Sub
End Class
