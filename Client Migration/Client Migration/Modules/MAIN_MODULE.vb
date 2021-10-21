Imports DevExpress.Utils
Imports DevExpress.XtraBars.Docking2010.Customization
Imports DevExpress.XtraBars.Docking2010.Views.WindowsUI
Imports DevExpress.XtraSplashScreen
Imports MySql.Data.MySqlClient
Module MAIN_MODULE
    Public CompanyLogo As Image
    'Public Function Datasource(ByVal command As String) As DataTable
    '    Return DB.Datasource(command)
    'End Function
    'Public Function Dataobject(ByVal command As String, Optional ByVal columnIndex As Integer = 0) As Object
    '    Return DB.DatasourceRetField(command, columnIndex)
    'End Function 
    Private Timer As Timer
    Public Class MyFlyoutDialog
        Inherits FlyoutDialog
        Public Sub New(owner As Form, action As FlyoutAction, control As UserControl)
            MyBase.New(owner, action)
            FlyoutControl = control
        End Sub
    End Class

End Module
