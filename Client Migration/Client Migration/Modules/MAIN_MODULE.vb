Imports DevExpress.Utils
Imports DevExpress.XtraBars.Docking2010.Customization
Imports DevExpress.XtraBars.Docking2010.Views.WindowsUI
Imports DevExpress.XtraSplashScreen
Imports MySql.Data.MySqlClient
Module MAIN_MODULE

    Public _ACCES As Boolean = True
    Public WithEvents DB As Basic.MYSQL
    Public QUERY_WITH_ERROR As Boolean = False
    Public TBL As DataTable
    Public ACCOUNTINGDB As String = " accounting"
    Public HRDBASE As String = " humanresource"
    Public PROCUREMENTDBASE As String = " PROCUREMENT"
    Public GENERALDBASE As String = " GENERAL"
    Public MarketingdBASE As String = " marketing"
    Public PropmanagementdBASE As String = " propmanagement"
    Public REPOSITORYDB As String = " repository"
    Public GENERALDB As String = "General"
    Public CopyRightCaption As String = "Copyright © 2016 Informatika Teknologia De Quantum"
    Public JEDATE As Date
    Public autosign As Struct.Autosigning
    Public POSITION_NAME As Struct.Position
    Public SIGNATORY_TYPE As Struct.Signatorytype
    Public FORM_TYPE As Struct.FORMTYPE
    Public CAN_MOD_TRANS_DATE As Boolean = True 'IF FALSE NOT ALLOW TO MODIFY TRANSACTION INTO OTHER MONTH SO THAT JE SERIES WILL NOT BE MOVED.

    Public Toload As Boolean = True
    Public CompanyLogo As Image



    Public Function Datasource(ByVal command As String) As DataTable
        Return DB.Datasource(command)
    End Function

    Public Function Dataobject(ByVal command As String, Optional ByVal columnIndex As Integer = 0) As Object
        Return DB.DatasourceRetField(command, columnIndex)
    End Function

    Sub VIEW_FORM_Dialog(FRM As Form, Optional Parent As Form = Nothing)
        Try
            DevExpress.XtraSplashScreen.SplashScreenManager.ShowDefaultWaitForm()
            FRM.StartPosition = FormStartPosition.CenterScreen
            DevExpress.XtraSplashScreen.SplashScreenManager.CloseDefaultWaitForm()
            FRM.ShowDialog(Parent)
        Catch ex As Exception
            DevExpress.XtraSplashScreen.SplashScreenManager.CloseDefaultWaitForm()
        End Try
    End Sub


    Private Function canCloseFunc(ByVal parameter As DialogResult) As Boolean
        Return parameter <> DialogResult.Cancel
    End Function



    Public Sub successfullsave()
        MessageXtra("Save")
    End Sub

    Public Sub MSG(ByVal MSG As String)
        MessageXtra(MSG)
    End Sub

    Public Sub successfullyupdate()
        MessageXtra("Updated")
    End Sub
    Public Sub successfullydeleted()
        MessageXtra("Deleted")

    End Sub
    Public Sub SHOWLOADINGOPEN()
        SplashScreenManager.ShowDefaultWaitForm()

    End Sub
    Public Sub SHOWLOADINGCLOSE()
        SplashScreenManager.CloseDefaultWaitForm()
    End Sub

    Public Function _get_id(ByVal gridview As DevExpress.XtraGrid.Views.Grid.GridView, ByVal _COLUMN As Integer)
        Return gridview.Columns.View.GetFocusedRow(_COLUMN)
    End Function
    Public Sub _get_format_to_double(ByVal gridview As DevExpress.XtraGrid.Views.Grid.GridView, ByVal _COLUMN As Object)
        Try
            gridview.Columns(_COLUMN).DisplayFormat.FormatType = FormatType.Numeric
            gridview.Columns(_COLUMN).DisplayFormat.FormatString = "{0:n2}"
        Catch ex As Exception

        End Try
    End Sub

    Public Sub _GETNAME(ByVal CMB As SergeUtils.EasyCompletionComboBox)


        CMB.DataSource = PERSONNAME
        CMB.ValueMember = "pp_person_code"
        CMB.DisplayMember = "pp_fullname"
        CMB.SelectedIndex = -1
    End Sub

    Public Sub _GETSUPPLIER(ByVal CMB As SergeUtils.EasyCompletionComboBox)


        CMB.DataSource = Datasource("SELECT * FROM `general`.`business_list` WHERE TYPE = 'SUPPLIER'")
        CMB.ValueMember = "business_number"
        CMB.DisplayMember = "business_fullname"
        CMB.SelectedIndex = -1
    End Sub


    Public PERSONNAME As New DataTable
    Public LOADPERSON As New DataTable

    Public Sub _LOAD_DATA()
        PERSONNAME = Datasource("SELECT * FROM qpsii_ehrm.`hr_person_profile` ORDER BY `pp_fullname` ASC; ")
        LOADPERSON = Datasource("SELECT
`pp_person_code` AS 'ID'
,`pp_prefix` AS 'PREFIX'
,`pp_lastname` AS 'LAST NAME'
,`pp_firstname` AS 'FIRST NAME'
,`pp_middlename` AS 'MIDDLE NAME'
,`pp_suffix` AS 'SUFFIX'
,`pp_dob` AS 'DOB'
,`pp_birthplace` AS 'BIRTHPLACE'
,`pp_sex` AS 'SEX'
,`pp_civilstatus` AS 'CIVIL STATUS'
,`pp_citizenship` AS 'CITIZINSHIP'
,`pp_P_street` AS 'PUROK/STREET'
,`pp_P_brgy` AS 'BARANGAY'
,`brgy_id` 
,`sitio_id`
,`pp_fullname`
FROM `qpsii_ehrm`.`hr_person_profile`
WHERE STATUS ='ACTIVE'")

    End Sub
    Public Sub _GETBRGY(ByVal CMB As SergeUtils.EasyCompletionComboBox)

        CMB.DataSource = Datasource("SELECT BRGY_ID, BRGY_NAME FROM `qpsii_lgusystem`.`lgis_brgy` WHERE REF_STATUS='ACTIVE' ORDER BY BRGY_NAME ASC;")
        CMB.ValueMember = "BRGY_ID"
        CMB.DisplayMember = "BRGY_NAME"
        CMB.SelectedIndex = 0
    End Sub



    Sub MessageBoxStr(ByVal str As String)
        MSG(str)
    End Sub


    Sub MessageXtra(ByVal msg As String)
        Basic.Messages.Msg(msg)
    End Sub


    Public Sub showflyouts(ByVal mdi As Form, ByVal COntrol As UserControl)

        Dim action As New FlyoutAction()
        action.Caption = Nothing
        Dim dialog As New MyFlyoutDialog(mdi, action, COntrol)

    End Sub
    Private Timer As Timer
    Public Class MyFlyoutDialog
        Inherits FlyoutDialog
        Public Sub New(owner As Form, action As FlyoutAction, control As UserControl)
            MyBase.New(owner, action)

            FlyoutControl = control

        End Sub

    End Class
    Private Sub DB_OnError(ByVal sender As Object, ByVal e As MySql.Data.MySqlClient.MySqlException) Handles DB.OnError
        Try
            QUERY_WITH_ERROR = True
            If DB.SHOW_ERRORS = Basic.ENUMS.REPORT_STATUS.Show Then
                MsgBox(e.Message, MsgBoxStyle.Exclamation, "Exception")
                'If ACCESSLEVEL = Basic.ENUMS.COMPANY_ACCESS.ALL Then

                'Else
                '    MsgBox("EXCEPTION OCCUR.PLEASE DOUBLE CHECK THE INPUTS.", MsgBoxStyle.Exclamation, "Exception")
                'End If
            Else
                MsgBox(e.Message, MsgBoxStyle.Exclamation, "Exception")
            End If
            DB.RollBackAction()
        Catch ex As Exception
        End Try
    End Sub
    Private Sub DB_OnErrorNTR(ByVal e As MySql.Data.MySqlClient.MySqlException) Handles DB.OnErrorNTR
        QUERY_WITH_ERROR = True
        If DB.SHOW_ERRORS = Basic.ENUMS.REPORT_STATUS.Show Then
            MessageBoxStr(e.Message)
            'If ACCESSLEVEL = Basic.ENUMS.COMPANY_ACCESS.ALL Then

            'Else
            '    If e.Number = 1042 Then
            '        MessageBoxStr("Cannot connect to the server.")
            '    ElseIf e.Number = 1044 Then
            '        MessageBoxStr("Invalid Credentials")
            '    ElseIf e.Number = 1064 Then
            '        MessageBoxStr("Syntax error report to developer.")
            '    ElseIf e.Number = 1305 Then
            '        MessageBoxStr("Stored not found.")
            '    ElseIf e.Number = 1054 Then
            '        MessageBoxStr("Column not found.")
            '    ElseIf e.Number = 1052 Then
            '        MessageBoxStr("Ambigous Field")
            '    Else
            '        MessageBoxStr(e.Message)
            '    End If
            'End If
        End If

    End Sub
    Private Sub DB_OnFinish(ByVal msg As Basic.ENUMS.IsSuccess, ByVal e As String) Handles DB.OnFinish
        If DB.REPORTSTAT = Basic.ENUMS.REPORT_STATUS.Show Then
            If msg = Basic.ENUMS.IsSuccess.Success Then
                If e <> "" Then
                    MessageBoxStr(e)
                Else
                    MessageBoxStr("TRANSACTION COMPLETED SUCCESSFULLY")
                    DB.Connection.Close()
                End If
            Else
                MessageBoxStr("TRANSACTION COMPLETED WITH ERROR.DATA MAY NOT BE SAVED.")
                DB.Connection.Close()
            End If
        End If
    End Sub
    Private Sub DB_OnRolledBack(ByVal Msg As String) Handles DB.OnRolledBack
        If DB.REPORTSTAT = Basic.ENUMS.REPORT_STATUS.Show Then
            If Msg <> "" Then
                MessageBoxStr(Msg)
            Else
                MessageBoxStr("TRANSACTION WAS CANCELLED" & Environment.NewLine & Msg)
            End If
        End If
    End Sub

    Private Sub DB_OnStart(ByVal sender As Object, ByVal e As System.EventArgs) Handles DB.OnStart
        QUERY_WITH_ERROR = False
    End Sub
    Private Sub DB_REPORT_STAT(ByVal MSG As String) Handles DB.REPORT_STAT
        '  DB_DISPLAY_WINDOW.Invoke(MSG)
    End Sub
    'Private je As New Journal_EntryUpdate
    'Sub OPEN_ENTRY_DETAILS(ByVal JID As Integer)
    '    Try
    '        je = New Journal_EntryUpdate
    '        je.GET_DATA(JID)
    '        'je.Activate()
    '        'je.BringToFront()
    '        je.ShowDialog(FRM_MAIN)
    '    Catch ex As Exception
    '    End Try
    'End Sub
    'Sub OPEN_ENTRY_DETAILS(ByVal GRP As String)
    '    Try
    '        'je = New Journal_EntryUpdate
    '        'je.Show(FRM_MAIN)
    '        'je.GET_DATA(GRP)
    '        'je.Activate()
    '        'je.BringToFront() 
    '        je = New Journal_EntryUpdate
    '        je.GET_DATA(GRP)
    '        'je.Activate()
    '        'je.BringToFront()
    '        je.ShowDialog(FRM_MAIN)
    '    Catch ex As Exception
    '    End Try
    'End Sub


    '    Sub CANCEL_ENTRY(ByVal JID As String, ByVal SEC_FRM As String)
    '        Try
    '            '' If USER_LOGIN.CHECK_ACCESS(CLSENUMS.ACCESS.DELETE, SEC_FRM) = False Then Exit Sub
    '            Dim TBLX = New DataTable
    '            TBLX = GET_JOURNALS(JID)
    '            Dim reasons As String = ""
    '            Dim GRP As String = ""
    '            Dim STAT As String = ""
    'again:

    '            Dim POPMSG As New OPTION_REMOVE_ENTRY
    '            If TBLX.Rows.Count > 0 Then
    '                GRP = TBLX(0)("GROUPENTRY")
    '                STAT = TBLX(0)("STATUS")
    '                If STAT <> "ACTIVE" Then
    '                    MessageBoxStr("Only active transaction can be cancelled.")

    '                    Exit Sub
    '                Else
    '                    If POPMSG.ShowDialog = DialogResult.OK Then
    '                        reasons = POPMSG.TextBoxX1.Text
    '                        If reasons.Trim = "" Then
    '                            MessageXtra("Please enter reason")
    '                            GoTo again
    '                        End If
    '                        DB.QueryExecNonQuery("Update journal_main set status='CANCELLED',`vc_reason`='" & reasons & "',`CancelledDate`=CURDATE() where  groupentry='" & GRP & "'")
    '                        DB.QueryExecNonQuery("UPDATE  `debits_actual_deducted` SET `status`='CANCELLED' WHERE `GUID`='" & GRP & "'")
    '                        DB.QueryExecNonQuery("UPDATE `po_actual_deducted` SET `STATUS`='CANCELLED' WHERE `GUID`='" & GRP & "'")
    '                        DB.QueryExecNonQuery("UPDATE `po_offset_payment` SET `STATUS`='CANCELLED' WHERE `GUID`='" & GRP & "'")
    '                        DB.QueryExecNonQuery("UPDATE undeposited_collection SET STAT='INCLUDE',offsetGID='Notffset' where offsetGID='" & GRP & "'")
    '                        DB.Delete("offsetted_trans_record", "OFFSET_GID='" & GRP & "'")

    '                        'DB.QueryExecNonQuery("INSERT INTO `journal_acctivity_log` SET  `JID`='" & JID i("journal_id") & "',`UID`='" & USER_LOGIN.ID & "',`username`='" & USER_LOGIN.UserName & "',`UserFullName`='" & USER_LOGIN.UserFullName & "',`Reason`='" & reasons & "',`ActionMade`='CANCELLED TRANSACTION'")
    '                        DB.QueryExecNonQuery("UPDATE `advances_from_customer` SET `status`='CANCELLED' WHERE `JID`='" & JID & "'")
    '                        DB.QueryExecNonQuery("UPDATE `advance_to_supplier` SET `status`='CANCELLED' WHERE `jid`='" & JID & "'")
    '                        DB.QueryExecNonQuery("UPDATE `casl_liq` SET `status`='CANCELLED' WHERE `jid`='" & JID & "'")

    '                        DB.QueryExecNonQuery("UPDATE undeposited_collection SET STAT='INCLUDE',offsetGID='Notffset' where offsetGID='" & GRP & "'")
    '                        DB.Delete("offsetted_trans_record", "OFFSET_GID='" & GRP & "'")

    '                        If My.Settings.CLIENTS <> CLSENUMS.CLIENTS.LGU Then
    '                            DB.QueryExecNonQuery("update " & My.Settings.HRDB & ".`cash_advances` set JE_status='CANCELLED' WHERE `JID`='" & JID & "'")
    '                        Else
    '                            DB.QueryExecNonQuery("DELETE from " & My.Settings.HRDB & ".`tbl_emp_otherdeduction`  where  JID='" & JID & "' and JID >0")
    '                        End If
    '                        DB.WriteToDB()
    '                    Else
    '                        Exit Sub
    '                    End If
    '                End If
    '            End If
    '        Catch ex As Exception
    '            DB.RollBackAction()
    '        End Try
    '    End Sub
    'Public Function PrintChecks(ByVal GUID_ As String) As CLSENUMS.Sucess
    '    Try
    '        CHECKS = New CHECKCLS
    '        TBL = CHECKCLS.GET_CHECK_INFO(GUID_)
    '        If TBL(0)("STATUS") = "CANCELLED" Then
    '            MessageXtra("This is a cancelled transaction cannot print..")
    '            Return CLSENUMS.Sucess.NotSuccess
    '            'ElseIf TBL(0)("STATUS") <> "APPROVED" Then
    '            '    MessageXtra("This is a unapproved transaction cannot print..")
    '            '    Return CLSENUMS.Sucess.NotSuccess
    '        ElseIf TBL.Rows.Count > 0 And (TBL(0)("CHECKNUMBER") = "" Or TBL(0)("CHECKNUMBER") = "--" Or TBL(0)("CHECKNUMBER") = "-") Then
    '            MessageXtra("Check number is not yet set")
    '            Return CLSENUMS.Sucess.NotSuccess
    '        End If
    '        If TBL.Rows.Count > 0 Then
    '            Dim CV As String = TBL(0)("DVNUM") 'CHECKS.GENERATE_CV(JID, TBL(0)("CHECKNUMBER"), TBL(0)("acid"), TBL(0)("groupentry"), Officeid)
    '            Dim sql As String
    '            sql = "call display_check_('" & GUID_ & "')"
    '            Dim newReport As New fzzReportViewerForm
    '            newReport.UID = USER_LOGIN.ID
    '            newReport.ACID = TBL(0)("ACID")
    '            newReport.CHECK = TBL(0)("CHECKNUMBER")
    '            newReport.Groups = TBL(0)("GroupEntry")
    '            Dim CKOP As New CHK_OPTIONS
    '            If CKOP.ShowDialog = DialogResult.OK Then
    '                If CKOP.RadioButton1.Checked = True Then
    '                    newReport.AttachReport(sql, "CHECK") = New rpt_check_bdo_LAN With {.CHKNUMBER = TBL(0)("CHECKNUMBER"), .CV = CV, .M = CKOP.M, .ShowLogo = CKOP.CheckBox2.Checked, .AllowCross = CKOP.CheckBox1.Checked}
    '                    VIEW_FORM(newReport)
    '                    Return CLSENUMS.Sucess.success
    '                ElseIf CKOP.RadioButton3.Checked = True Then
    '                    newReport.AttachReport(sql, "CHECK") = New rpt_check_bdo_LAN_TOP With {.CHKNUMBER = TBL(0)("CHECKNUMBER"), .CV = CV, .M = CKOP.M, .ShowLogo = CKOP.CheckBox2.Checked, .AllowCross = CKOP.CheckBox1.Checked}
    '                    VIEW_FORM(newReport)
    '                    Return CLSENUMS.Sucess.success
    '                Else
    '                    newReport.AttachReport(sql, "CHECK") = New rpt_check_bdo With {.CHKNUMBER = TBL(0)("CHECKNUMBER"), .CV = CV, .M = CKOP.M, .ShowImage = CKOP.CheckBox2.Checked, .AllowCross = CKOP.CheckBox1.Checked}
    '                    VIEW_FORM(newReport)
    '                    Return CLSENUMS.Sucess.success
    '                End If
    '            End If
    '        Else
    '            MessageXtra("Only check payment transaction can be printed..")
    '            Return CLSENUMS.Sucess.NotSuccess
    '        End If
    '    Catch ex As Exception
    '        Return CLSENUMS.Sucess.NotSuccess
    '    End Try
    'End Function

    Function GET_SIGNATORY_BY_POSTION(ByVal POSITIONNAME As String) As String
        Try
            Dim T As New DataTable
            T = Datasource("SELECT CONCAT(PREFIX, ' ', `EMPNAME`,' ',SUFIX) 'NAME' FROM `employee_with_positions` WHERE `POSITION` LIKE '" & POSITIONNAME & "'")
            If T.Rows.Count > 0 Then
                Return T(0)(0)
            Else
                Return ""
            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function
    Sub MessageBoxError2(ByVal ex As Exception)
        Dim s As String = "Exception occur..."
        s.ToLower.Speak()
        MsgBox("Exception occur..." & vbNewLine & ex.Message, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error")
    End Sub

    Public Function _LASTDAY(ByVal _DATE As Date) As Date

        Return CDate(Format(_DATE.AddMonths(+1), "yyyy-MM-01")).AddDays(-1)

    End Function



End Module
