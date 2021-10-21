Imports DevExpress.Utils
Imports DevExpress.XtraBars.Docking2010.Customization
Imports DevExpress.XtraBars.Docking2010.Views.WindowsUI
Imports DevExpress.XtraSplashScreen
Imports MySql.Data.MySqlClient
Module MAIN_MODULE
    Public CompanyLogo As Image
    Property DB As Basic.MYSQL
    Private Timer As Timer
    Public Class MyFlyoutDialog
        Inherits FlyoutDialog
        Public Sub New(owner As Form, action As FlyoutAction, control As UserControl)
            MyBase.New(owner, action)
            FlyoutControl = control
        End Sub
    End Class

End Module
