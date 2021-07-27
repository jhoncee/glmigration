Public Class itemMasterCls
    Property NO As Integer
    Property ITEMCODE As String
    Property DESCRIPTION As String
    Private _UOM As String
    Public Property UOM() As String
        Get
            Return _UOM
        End Get
        Set(ByVal value As String)
            If value = "" Then
                _UOM = "Piece(s)"
            Else
                _UOM = value
            End If
        End Set
    End Property

    Property TYPES As String
    Property COMPID As Integer
    Property PRJID As Integer

    Sub New(no_ As Integer, itemCode_ As String, desc_ As String, uom_ As String, type_ As String, COMPID As Integer, PRJID As Integer)
        Try
            Me.NO = no_
            Me.ITEMCODE = itemCode_
            Me.DESCRIPTION = desc_
            Me.UOM = uom_
            Me.TYPES = type_
            Me.COMPID = COMPID
            Me.PRJID = PRJID
        Catch ex As Exception
            MessageBoxError2(ex)
        End Try
    End Sub
End Class
