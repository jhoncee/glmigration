Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraGrid
Imports DevExpress.XtraGrid.Columns
Imports DevExpress

Module GRID_LABELS
    Sub MessageBoxError(ex As Exception)
        Basic.Messages.Msg(ex.Message)
    End Sub
    Function MessageBoxContinuesave() As Boolean
        Return Basic.Messages.ContinueSave()
    End Function
    Function MessageBoxContinueUpdate() As Boolean
        Return Basic.Messages.ContinueUpdate()
    End Function
    Function MessageWhat(str As String) As Boolean
        Return MsgIfwhat(str)
    End Function

    Function MsgIfwhat(str As String) As Boolean
        If MsgBox(str, MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            Return True
        Else
            Return False
        End If
    End Function

End Module
Module ModMethods
    Sub AutoStart(ByVal auto As Boolean)
        Dim applicationName As String = Application.ProductName
        Dim applicationPath As String = Application.ExecutablePath

        If auto Then
            Dim regKey As Microsoft.Win32.RegistryKey
            regKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", True)
            regKey.SetValue(applicationName, """" & applicationPath & """")
            regKey.Close()
        Else
            Dim regKey As Microsoft.Win32.RegistryKey
            regKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", True)
            regKey.DeleteValue(applicationName, False)
            regKey.Close()
        End If
    End Sub
    Function GET_GUID() As String
        Try
            Return Guid.NewGuid().ToString()
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Public Sub ExportGrid(ByVal grid As GridControl, ByVal Rep_name As String)
        Try
            If Not grid.IsPrintingAvailable Then
                MessageBox.Show("The 'XtraPrinting' library is not found", "Error")
                Exit Sub
            End If
            If My.Computer.FileSystem.DirectoryExists(My.Computer.FileSystem.SpecialDirectories.Desktop & "\AccountingReports") = False Then
                My.Computer.FileSystem.CreateDirectory(My.Computer.FileSystem.SpecialDirectories.Desktop & "\AccountingReports")
            End If
            Dim filenames As String = My.Computer.FileSystem.SpecialDirectories.Desktop.ToString & "\AccountingReports" & "\" & Rep_name & Now.Hour & Now.Minute & Now.Second & ".xls"
            grid.ExportToXls(filenames)
            MsgBox("Exported to " & filenames)
            Process.Start(filenames)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Sub ExportGrid(ByVal grid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal Rep_name As String)
        Try
            Try
                If My.Computer.FileSystem.DirectoryExists(My.Computer.FileSystem.SpecialDirectories.Desktop & "\AccountingReports") = False Then
                    My.Computer.FileSystem.CreateDirectory(My.Computer.FileSystem.SpecialDirectories.Desktop & "\AccountingReports")
                End If
            Catch exsdf As Exception
            End Try
            Dim filenames As String = My.Computer.FileSystem.SpecialDirectories.Desktop.ToString & "\AccountingReports" & "\" & Rep_name & Now.Hour & Now.Minute & Now.Second & ".xls"
            grid.ExportToExcel(filenames)
            MsgBox("Exported to " & filenames)
            Process.Start(filenames)
        Catch ex As Exception
            MsgBox(ex.Message)
            '' MsgBox("Check if previously exported file is not open.")
        End Try
    End Sub
    Public Sub ExportGrid(ByVal grid As GridView, ByVal Rep_name As String)
        Try
            Dim f As New SaveFileDialog
            f.Filter = "Excel Files|*.xlsx"
            If f.ShowDialog = DialogResult.OK Then
                Dim filenames As String = f.FileName
                grid.ExportToXlsx(filenames)
                Process.Start(filenames)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Module
