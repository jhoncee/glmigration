Imports System.Runtime.CompilerServices

Module rvn_modCommon

    Public str_report_name As String
    Public str_report_print_date As String = Date.Now
    Public str_report_generated As String = ""
    Public frmLoad_IsOpen As Boolean = False
    Public str_report_date As String
    Public CHECKS As CHECKCLS
    'Public empnumber As String
    Public logo_name As String
    Function GetDefaultID(Identifier As String) As Integer
        Try
            Return Dataobject("SELECT IFNULL((SELECT IDS  FROM defaultids WHERE IDENTIFIER='" & Identifier & "' limit 1),0) 'ID'")
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Public Function logo_header() As Image
        If logo_name <> "" Then
            If IO.File.Exists(String.Format("{0}/Logo/{1}/", My.Application.Info.DirectoryPath, logo_name)) Then
                Return System.Drawing.Image.FromFile(String.Format("{0}/Logo/{1}", My.Application.Info.DirectoryPath, logo_name))
            End If
        End If
    End Function

End Module
