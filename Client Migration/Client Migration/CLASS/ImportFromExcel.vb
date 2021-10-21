Imports System.Data.OleDb

Public Class ImportFromExcel
    Private ExcelConnection As New OleDbConnection
    Public IsConnected As Boolean
    Public Sub New(ByVal ExcelFilePath As String)
        Try
            Dim ConStr As String = ""
            Dim Ext As String = ExcelFilePath.Substring(ExcelFilePath.LastIndexOf("."))
            If Ext.Length <= 4 Then
                ConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ExcelFilePath + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1';"
            Else
                ConStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + ExcelFilePath + ";Extended Properties='Excel 12.0 xml;HDR=YES;IMEX=1';"
            End If
            ExcelConnection = New OleDbConnection(ConStr)
            IsConnected = True
        Catch ex As Exception
            IsConnected = False
        End Try
    End Sub
    Public Function GetExcelSheets() As String()
        Try
            Dim dt As New DataTable
            dt.Clear()
            If ExcelConnection.State = ConnectionState.Open Then ExcelConnection.Close()
            ExcelConnection.Open()
            dt = ExcelConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
            If dt Is Nothing Then
                Return Nothing
            End If
            Dim excelSheetNames As [String]() = New [String](dt.Rows.Count - 1) {}
            Dim i As Integer = 0
            For Each row As DataRow In dt.Rows
                excelSheetNames(i) = row("TABLE_NAME").ToString()
                i += 1
            Next
            ExcelConnection.Close()
            Return excelSheetNames
        Catch ex As Exception
            IsConnected = False
            Return Nothing
        End Try
    End Function
    Private Function OpenSheet(ByVal sheet As String) As DataTable
        Try
            Dim dt As New DataTable
            dt.Clear()
            Dim cmd As New OleDbCommand("Select * from [" + sheet + "]", ExcelConnection)
            If ExcelConnection.State = ConnectionState.Open Then ExcelConnection.Close()
            ExcelConnection.Open()
            dt.Load(cmd.ExecuteReader())
            ExcelConnection.Close()
            Return dt
        Catch ex As Exception
            IsConnected = False
            Return Nothing
        End Try
    End Function
    Private Function ImportDirectFirstSheet() As DataTable
        Try
            Dim c() As String = GetExcelSheets()
            Dim excelSheetNames As String = c(0)
            Return OpenSheet(excelSheetNames)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Private Function ImportDirectFirstSheet(ByVal SheetName As String) As DataTable
        Try
            Dim excelSheetNames As String = ""
            If SheetName.ToString.Contains("$") Then
                If SheetName.Trim = "$" Then
                    Return ImportDirectFirstSheet()
                Else
                    excelSheetNames = SheetName
                End If
            Else
                excelSheetNames = SheetName & "$"
            End If
            Return OpenSheet(excelSheetNames)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function OpenExcelSheet(Optional ByVal SheetName As String = "") As DataTable
        Try
            If SheetName = "" Then
                Return ImportDirectFirstSheet()
            Else
                Return ImportDirectFirstSheet(SheetName)
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

End Class
