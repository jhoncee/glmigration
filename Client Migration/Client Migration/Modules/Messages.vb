Module Messages

    Function MessageContinueSave() As Boolean
        Dim res = MsgBox("Are you sure you want to save this data?", MessageBoxButtons.YesNo, "Continue")
        If res = MsgBoxResult.Yes Then Return True
        Return False
    End Function
    Function MessageContinueUpdate() As Boolean
        Dim res = MsgBox("Are you sure you want to update this data?", MessageBoxButtons.YesNo, "Continue")
        If res = MsgBoxResult.Yes Then Return True
        Return False
    End Function
    Function MessageContinueRemove() As Boolean
        Dim res = MsgBox("Are you sure you want to delete the this data?", MessageBoxButtons.YesNo, "Continue")
        If res = MsgBoxResult.Yes Then Return True
        Return False
    End Function

    Sub MessageBoxStr(ByVal str As String)
        MessageXtra(str)
    End Sub
    Sub MessageXtra(ByVal msg As String)
        MsgBox(msg, MessageBoxButtons.OK, "Information")
    End Sub
    Sub MessageBoxError2(ByVal ex As Exception)
        MsgBox("Exception occur..." & vbNewLine & ex.Message, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error")
    End Sub
End Module
