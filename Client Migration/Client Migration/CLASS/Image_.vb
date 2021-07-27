Imports System.IO
Imports System.Drawing.Imaging
Imports MySql.Data.MySqlClient
Imports DevExpress.XtraGrid

Public Class image_
    Private Shared cmd As New MySql.Data.MySqlClient.MySqlCommand
    Public Shared Function Img_To_Byte_Array(ByVal img As Image) As Byte()
        Try
            Using M As New MemoryStream
                img.Save(M, img.RawFormat)
                Return M.ToArray
            End Using
        Catch ex As Exception
            My.Computer.Clipboard.SetText(ex.Message)
            Return Nothing
        End Try
    End Function
    Public Shared Function ResizeImage(ByVal InputImage As Image) As Image
        Return New Bitmap(InputImage, New Size(400, 450))
    End Function

    Public Shared Sub savephoto(ByVal rawData As Byte(), ByVal _SQL As String)
        Try
            Dim cmd As New MySql.Data.MySqlClient.MySqlCommand
            If DB.Connection.State = ConnectionState.Closed Then DB.Connection.Open()
            cmd.Connection = DB.Connection
            cmd.CommandText = _SQL
            cmd.Parameters.AddWithValue("?xx", rawData)
            cmd.ExecuteNonQuery()
        Catch EEX As MySql.Data.MySqlClient.MySqlException
            DB.RollBackAction(EEX.Message)
        Catch ex As Exception

            MessageBoxStr(ex.Message)
        End Try
    End Sub

End Class
