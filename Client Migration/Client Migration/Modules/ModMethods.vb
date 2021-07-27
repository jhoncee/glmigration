Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraGrid
Imports DevExpress.XtraGrid.Columns
Imports DevExpress

Module GRID_LABELS
    Public Enum PMT_TYPE
        Cash
        Check
        Online
        CreditCard
    End Enum
    Function GET_PMT_ACCOUNTS(Pmt As PMT_TYPE) As Integer()
        Try
            Dim PmtType As String = ""
            If Pmt = PMT_TYPE.Cash Then
                PmtType = "Cashier Cash Payment"
            ElseIf Pmt = PMT_TYPE.Check Then
                PmtType = "Cashier Check Payment"
            ElseIf Pmt = PMT_TYPE.Online Then
                PmtType = "Cashier Online"
            ElseIf Pmt = PMT_TYPE.CreditCard Then
                PmtType = "Cashier Credit Card"
            End If

            Dim Accounts(1) As Integer
            Accounts(0) = 0
            Accounts(1) = 0
            Dim T As New DataTable
            T = Datasource("SELECT ACID, DRCR  FROM  propmanagement.tbl_default_account_new WHERE Description='" & PmtType & "' ORDER BY drcr DESC LIMIT 2")
            If T.Rows.Count > 0 Then
                If T.Rows.Count = 1 Then GoTo XX
                Accounts(0) = T(0)(0) 'DR
                Accounts(1) = T(1)(0) 'CR 
                Return Accounts
            Else
XX:
                Return Nothing
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Function GetCashierRef() As String
        Try
            Dim Ref As String = ""
            Dim cnt As Integer = 0
            Ref = Dataobject("SELECT IFNULL(MAX(PMTRREF),0) 'REF'  FROM propmanagement.cashierpayment group by GUID")
            If Ref Is Nothing Then
                cnt = 1
            ElseIf Ref = "" Then
                cnt = 1
            ElseIf Ref = 0 Then
                cnt = 1
            ElseIf IsNumeric(Ref) = False Then
                cnt = 1
            ElseIf IsNumeric(Ref) Then
                cnt = CInt(Ref) + 1
            End If
            Return cnt.ToString("000000")
        Catch ex As Exception
            Return "-000001"
        End Try
    End Function

    Private Delegate Sub LoadG(ByVal G As DevExpress.XtraGrid.GridControl, ByVal T As DataTable)
    Private Delegate Sub LoadC1(ByVal G As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal T As DataTable)
    Private Delegate Sub LoadL(ByVal L As DevExpress.XtraEditors.LabelControl, ByVal Txt As String)

    Function GETPWD(str As String) As String
        Try
            Return Basic.UTILS.Encrypt(str)
        Catch ex As Exception
        End Try
    End Function


    Function GET_ATC(Optional Goodservice As String = "%", Optional vendorID As Integer = 0) As DataTable
        Try
            Return Datasource("SELECT ATC,RATE ,Description,Goodservice 'Type',ID FROM atc_vendor_default WHERE BID ='" & vendorID & "' AND Goodservice LIKE '" & Goodservice & "'")
        Catch ex As Exception
            Return Datasource("select '' as ATC ,0 AS 'RATE','' as 'Description','' as 'Type',0 LIMIT 0")
        End Try
    End Function
    Function GET_BEG_SERIES(ByVal BT As Basic.ENUMS.BTMODE, cmp As Integer) As String
        Try
            Dim Prefix As String = "JEV#"
            Dim t As New DataTable
            Dim c As Integer = 0
            t = DB.QueryandReturnTable(" SELECT IFNULL(MAX(SUBSTR(SERIESNUMBER,10)),0)  FROM  journal_main  INNER JOIN journal_detail USING(journal_id)  WHERE SERIESNUMBER LIKE '%BEG#%' AND  _CMP ='" & cmp & "' AND YEAR(trans_date)=(SELECT YEAR(CURDATE())) ;")
            If BT = Basic.ENUMS.BTMODE.SAVE Then
                c = t(0)(0) + 1
                Return Prefix & "-" & ServerDateTime.Year & c.ToString("0000")
            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function


    Public Sub filterGRIDVIEW(ByVal Gridx As DevExpress.XtraGrid.Views.Grid.GridView)
        Dim cond As AutoFilterCondition
        cond = AutoFilterCondition.Contains
        For Each col As GridColumn In Gridx.Columns
            col.OptionsFilter.AutoFilterCondition = cond
        Next col
    End Sub
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


    Sub LoadGridData(ByVal grd As DevExpress.XtraGrid.GridControl, ByVal dataTBl As DataTable)
        If grd.InvokeRequired Then
            grd.BeginInvoke(New LoadG(AddressOf LoadGridData), grd, dataTBl)
        Else
            grd.DataSource = dataTBl
        End If
    End Sub

    Sub SetLabel(ByVal frm As Form, ByVal s As String)
        frm.Text = s
    End Sub
    Sub LoadGridDataC1(ByVal grd As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal dataTBl As DataTable)
        If grd.InvokeRequired Then
            grd.Invoke(New LoadC1(AddressOf LoadGridDataC1), grd, dataTBl)
        Else
            grd.DataSource = dataTBl
            grd.Update()
        End If
    End Sub

    Sub LoadLabel(ByVal ctrl As DevExpress.XtraEditors.LabelControl, ByVal Txt As String)
        Try
            If ctrl.InvokeRequired Then
                ctrl.Invoke(New LoadL(AddressOf LoadLabel), ctrl, Txt)
            Else
                ctrl.Text = FormatCurrency(CDec(Txt), 2, TriState.True, TriState.True)
                ctrl.Update()
            End If
        Catch ex As Exception
            Try
                ctrl.Text = "0.00"
                ctrl.Update()
            Catch exx As Exception
            End Try
        End Try
    End Sub
End Module
Module ModMethods
    Public Sub FilterChange(ByVal tdbGrid As C1.Win.C1TrueDBGrid.C1TrueDBGrid, ByVal DT As DataTable)
        Dim sb As New System.Text.StringBuilder()
        Dim dc As C1.Win.C1TrueDBGrid.C1DataColumn
        For Each dc In tdbGrid.Columns
            If dc.FilterText.Length > 0 Then
                If sb.Length > 0 Then
                    sb.Append(" AND ")
                End If
                If dc.FilterText.Contains("~") Then
                    sb.Append("(Convert(" & (dc.DataField + " ,System.String) >= '" & dc.FilterText.Substring(0, dc.FilterText.IndexOf("~")) & "'") & " AND " & "Convert(" & (dc.DataField + " ,System.String) <= '" & dc.FilterText.Substring(dc.FilterText.IndexOf("~") + 1, (dc.FilterText.Length - 1) - dc.FilterText.IndexOf("~")) & "')"))
                Else
                    sb.Append("(Convert(" & (dc.DataField + " ,System.String) like '" + dc.FilterText + "')"))
                End If
            End If
        Next dc
        Try
            'sb.Append(dc.FilterText + "%")
            'DS.Tables(0).DefaultView.RowFilter = sb.ToString()
            DT.DefaultView.RowFilter = sb.ToString
        Catch ee As EvaluateException
            ' MsgBox("modMain..FilterChange - " & ee.Message & vbCrLf & vbCrLf & ee.StackTrace, MsgBoxStyle.Critical, Application.ProductName & " - Error Encountered!")
        End Try
    End Sub

    Function ISDRCRBALANCE(GRV As GridView) As Boolean
        Try
            Dim DR As Double = GRV.Columns("DEBIT").SummaryItem.SummaryValue.ToString()
            Dim CR As Double = GRV.Columns("CREDIT").SummaryItem.SummaryValue.ToString()
            If DR = CR Then
                Return True
            End If
            Return False
        Catch ex As Exception
            Return False
        End Try
    End Function



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

    Function ServerDateTime() As DateTime
        Try
            Dim servtime As New DataTable
            servtime = Datasource("SELECT CURRENT_TIMESTAMP()")
            Return CType(servtime(0)(0), DateTime)
        Catch ex As Exception
            Return Now
        End Try
    End Function
    Function GET_START_DATE() As DateTime
        Try
            Dim t As New DataTable
            t = Datasource("SELECT Start_Date FROM fiscalyear")
            If t.Rows.Count > 0 Then
                Return CDate(t(0)(0))
            Else
                Return CDate("1-1-2015")
            End If
        Catch ex As Exception
            Return CDate("1-1-2015")
        End Try
    End Function

    Function GET_SIGNAME_ERP(ByVal FRM As String, ByVal TYPES As String) As DataTable
        Try
            Dim STR As New DataTable
            STR = Datasource(" SELECT DISPNAME 'name' FROM signatories WHERE frmname='" & FRM & "' AND  SIGTYPE='" & TYPES & "'")
            Return STR
        Catch ex As Exception
            Return Nothing
        End Try
    End Function


    Function GETACCOUNTBYID(ByVal ID As Integer) As DataTable
        Try
            Return Datasource("select  account_code , account_description ,t_account  from account_setup where account_id='" & ID & "' ")
        Catch ex As Exception
            Return Nothing
        End Try
    End Function


    Function Get_CompanyName(ByVal TID As Integer) As String
        Try
            Return Basic.DataProvider.GET_COMPANYNAME(TID)
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Function GET_ALL_SYSTEM_USERS() As DataTable
        Try
            Return Datasource("SELECT * FROM " & GENERALDBASE & ".users WHERE STATUS='ACTIVE'")
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Enum _SIGTYPE
        PREPARED = 1
        REVIEWED = 2

    End Enum
    Public Sub GET_ALL_SYSTEM_USERS_BY_FORM(ByVal _TYPE As String, FRM As Form, ByVal cmb As SergeUtils.EasyCompletionComboBox)
        Try
            cmb.DataSource = Datasource("Select * FROM " & HRDBASE & ".`employees` INNER JOIN `signatories` ON(employees.`SysPK_Empl` = signatories.`EMPID`) WHERE   `SIGTYPE` = '" & _TYPE & "' AND `frmname`= '" & FRM.Name & "'")
            cmb.DisplayMember = "DISPNAME"
            cmb.ValueMember = "EMPID"
        Catch ex As Exception
        End Try
    End Sub

    Public Sub _FORM_SIGNATORY(ByVal _TYPE As String, FRM As Form, ByVal cmb As SergeUtils.EasyCompletionComboBox)
        Try
            cmb.DataSource = Datasource("SELECT `ID`,CONCAT(`DISPNAME`,'/',IFNULL(`positions`,'')) AS 'NAME' FROM `signatories` WHERE   `SIGTYPE` = '" & _TYPE & "' AND `frmname`= '" & FRM.Name & "'")
            cmb.DisplayMember = "NAME"
            cmb.ValueMember = "ID"
        Catch ex As Exception
        End Try
    End Sub

    Function GET_ALL_ACCOUNTS_TITLE(ByVal COMPANYID As Integer, Optional ByVal ONLY_VISIBLE As Boolean = True, Optional ByVal short_description As String = "") As DataTable
        Try
            Dim VIS As Integer = IIf(ONLY_VISIBLE = True, 1, 0)
            Return Datasource("CALL ALL_ACCOUNTS('" & COMPANYID & "','" & short_description & "','" & VIS & "')")
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function GET_UNITS_LIST(Optional Propertyid As String = "%", Optional _custid As String = "%") As DataTable
        Try
            Return Datasource("SELECT  property_id 'ID',unit_id 'UNITID',project_id 'PRJID' ,PROPERTY,UNIT FROM " & PropmanagementdBASE & ".tbl_payment_schedule WHERE property_id  LIKE '" & Propertyid & "' and business_id like '" & _custid & "'")
        Catch ex As Exception
            Return Datasource("SELECT 0 AS 'ID',0 AS 'UNITID',0 AS 'PRJID','' AS 'PROPERTY','' AS 'UNIT' LIMIT 0")
        End Try
    End Function


    Function GET_ALL_GL_ACCOUNTS(ByVal COMPANYID As Integer, Optional ByVal short_description As String = "") As DataTable
        Try
            Return Datasource("CALL ALL_GL_ACCOUNTS('" & COMPANYID & "','" & short_description & "')")
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Function GET_ALL_PPE_ACCOUNTS(Optional ByVal PPE_LOCATION As Integer = 0) As DataTable
        Try
            If PPE_LOCATION = 0 Then
                Return Datasource("	SELECT account_id ,account_code ,account_description FROM `fa_setup` INNER JOIN ACCOUNT_SETUP ON  account_id=`PPE_Account_id` ")
            Else
                Return Datasource("	SELECT account_id ,account_code ,account_description FROM `fa_setup` INNER JOIN ACCOUNT_SETUP ON  account_id=`PPE_Account_id` WHERE Location='" & PPE_LOCATION & "'  ")
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Function GET_GUID() As String
        Try
            Return Basic.UTILS.GET_GUID
        Catch ex As Exception
            Return ""
        End Try
    End Function
    Public Sub filterGRID(ByVal GridView5 As DevExpress.XtraGrid.Views.Grid.GridView)
        Dim cond As AutoFilterCondition
        cond = DevExpress.XtraGrid.Columns.AutoFilterCondition.Contains
        For Each col As XtraGrid.Columns.GridColumn In GridView5.Columns
            col.OptionsFilter.AutoFilterCondition = cond
        Next col
    End Sub

    Function GET_EMPLOYEE_ALL() As DataTable

        Return Datasource("CALL GET_ALL_HR_EMP()")

    End Function


    Public Sub GET_LOCATION(ByVal cmb As SergeUtils.EasyCompletionComboBox)

        cmb.DataSource = Datasource("SELECT * FROM " & GENERALDBASE & ".`setup_project` WHERE STATUS = 'ACTIVE'")
        cmb.DisplayMember = "project_name"
        cmb.ValueMember = "project_id"
    End Sub
    Function GET_LOCATIONS(Optional _TITLE_ID As Integer = 0) As DataTable
        Try
            Return Basic.DataProvider.GET_LOCATIONS(_TITLE_ID)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Function GET_PPE_Locations(ByVal TITLE As Integer) As DataTable
        Try
            Return Datasource("SELECT  DISTINCT project_id 'ID',CONCAT(`project`,'- ',IF(`Officealias` = '',`Project_name`,Officealias))  'Name', project,`Main_office_id` 'MAINID' ,  setup_project.title_id  FROM " & GENERALDBASE & ".setup_project inner join fa_setup on project_id=location")
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Function GET_PROJECT_SITES(ByVal TITLE As Integer, Optional ByVal OFFICE As Integer = 0) As DataTable
        Try
            Return Basic.DataProvider.GET_PROJECT_SITES(TITLE, OFFICE)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Function GET_GROUP_BY_JID(ByVal JID As Integer) As String
        Try
            Dim T As New DataTable
            T = Datasource("select groupentry from  journal_main where journal_id='" & JID & "'")
            If T.Rows.Count > 0 Then
                Return T(0)(0)
            Else
                Return ""
            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Function GET_POID_BY_JID(ByVal JID As Integer) As String
        Try
            Dim T As New DataTable
            T = Datasource("select PO_ID from  journal_main where journal_id='" & JID & "'")
            If T.Rows.Count > 0 Then
                Return T(0)(0)
            Else
                Return 0
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Function Get_Office_Address(ByVal Locations As Integer) As String
        Try
            Dim LOc As DataTable = Datasource("SELECT `project_location` FROM " & GENERALDBASE & ".`setup_project` WHERE  `project_id`='" & Locations & "'")
            'Dim LOc As DataTable = DataSource("SELECT  concat(IF(`OfficeAlias`='',`project_name`,OfficeAlias),' - ',`project_location`)  FROM `setup_project` WHERE  `project_id`='" & Locations & "'")
            If LOc.Rows.Count > 0 Then
                Return LOc(0)(0)
            Else
                Return ""
            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function


    Function Get_CompanyAddress(ByVal TitleID As Integer)
        Try
            Dim CAddress As DataTable = Datasource("SELECT `title_desc1` FROM " & GENERALDBASE & ".`project_title` WHERE `title_id`='" & TitleID & "'")
            If CAddress.Rows.Count > 0 Then
                Return CAddress(0)(0)
            Else
                Return ""
            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function
    Function Get_Company_Contact(ByVal title_id As Integer) As String
        Try

            Dim LOc As DataTable = Datasource("SELECT `Contact` FROM  " & GENERALDBASE & ".`project_title` WHERE  title_id='" & title_id & "'")
            If LOc.Rows.Count > 0 Then
                Return LOc(0)(0)
            Else
                Return ""
            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Public Sub SAVE_DATA(ByVal VALIDATION As IACTIONS.VALID, ByVal CMD As IACTIONS.METHODS)
        CMD.Invoke(VALIDATION.Invoke)
    End Sub
    Public Sub DISPLAY_DATA(ByVal Implementor As IVIEWALL_INFO, ByVal P As PARAMS)
        Implementor.EXECUTE_COMMAND(P)
    End Sub
    Public Sub MODIFY_DATA(ByVal cmd As IACTIONS.Modify)
        cmd.Invoke()
    End Sub
    Public Sub SAVE(ByVal cmd As IACTIONS)
        cmd.COMMAND()
    End Sub



    Public Function GET_ALL_CUSTOMER() As DataTable
        Dim BP_DEFAULT As String = "BUSINESS"
        Return Datasource("CALL BUSINESS_PERSON()")
    End Function

    Function GET_PAYEE_LIST(BID As Integer) As DataTable
        Try
            Return Datasource("SELECT ID,BID,PAYEE_NAME 'NAME' FROM accounting.payee WHERE BID='" & BID & "'")
        Catch ex As Exception
            Return Datasource("SELECT  0 'ID','' AS  'NAME' LIMIT 0")
        End Try
    End Function
    Public Function GETVENDORS(Optional ByVal TYPES As String = "%") As DataTable
        Return Basic.DataProvider.GETBUSINESSES(Basic.DataProvider.Business_type.Supplier)
    End Function
    Public Function GetStatus(ByVal JID As Integer) As String
        Try
            Dim stat As New DataTable
            stat = Datasource("select status from  journal_main where journal_id='" & JID & "'")
            If stat.Rows.Count > 0 Then
                Return stat(0)(0).ToString
            Else
                Return "LOCKED"
            End If
        Catch ex As Exception
            Return "LOCKED"
        End Try
    End Function


    Public Function GET_CASH_ACCOUNT(ByVal TITLEID As Integer, Optional ByVal ISBANK As CLSENUMS.CASH_TYPE = CLSENUMS.CASH_TYPE.CASH_ONLY) As DataTable
        Try
            If ISBANK = CLSENUMS.CASH_TYPE.ALL Then
                Return Datasource("SELECT Account_id,Bank,Account_description,ledger,ACCOUNT_CODE FROM Account_setup  WHERE subclassname like '%Cash%'  and visibleinentry=1")
            ElseIf ISBANK = CLSENUMS.CASH_TYPE.BANK_ONLY Then
                Return Datasource("SELECT Account_id,Bank,Account_description,ledger,ACCOUNT_CODE FROM Account_setup  WHERE bank='True' and subclassname  like  '%Cash%' and visibleinentry=1")
            Else
                Return Datasource("SELECT Account_id,Bank,Account_description,ledger,ACCOUNT_CODE FROM Account_setup  WHERE bank='False' and subclassname like  '%Cash%' and visibleinentry=1")
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    ''' <summary>
    ''' This is to get all accounts that is bank type
    ''' </summary>
    ''' <param name="TITLEID"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GET_ACCOUNT_TITLE_ALL_BANK_TYPE(ByVal TITLEID As Integer) As DataTable
        Try
            Return GET_CASH_ACCOUNT(TITLEID, CLSENUMS.CASH_TYPE.BANK_ONLY)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public FS As System.IO.FileStream
    Sub MakeRoundedForm(ByRef f As Form)
        Try
            Dim p As New Drawing2D.GraphicsPath()
            p.StartFigure()
            p.AddArc(New Rectangle(0, 0, 10, 10), 180, 90)
            p.AddLine(10, 0, f.Width - 10, 0)
            p.AddArc(New Rectangle(f.Width - 10, 0, 10, 10), -90, 90)
            p.AddLine(f.Width, 10, f.Width, f.Height - 10)
            p.AddArc(New Rectangle(f.Width - 10, f.Height - 10, 10, 10), 0, 90)
            p.AddLine(f.Width - 10, f.Height, 10, f.Height)
            p.AddArc(New Rectangle(0, f.Height - 10, 10, 10), 90, 90)
            p.CloseFigure()
            f.Region = New Region(p)
        Catch ex As Exception
        End Try
    End Sub
    Public Function GetLogo() As Image
        Try
            If DB.IsConnected Then
                Dim arrImage() As Byte
                Dim S As String = "SELECT  header_img as 'SIGIMG' FROM  " & GENERALDBASE & ".project_title   order by title_id limit 1"
                arrImage = Dataobject(S)
                Dim mstream As New System.IO.MemoryStream(arrImage)
                Return Image.FromStream(mstream)

            Else
                Return Nothing
            End If
        Catch ex As Exception
            Return Nothing
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

    Sub CHECK_UPDATER_METHODS(ByVal JID As Integer)
        Try
            ''  If USER_LOGIN.CHECK_ACCESS(CLSENUMS.ACCESS.FORM, "UPDATE CHECK INFO") = False Then Exit Sub
            Dim T As New DataTable
            T = Datasource("SELECT STATUS,JID,ACID,CHECKNUMBER,CHECKDATES,PAYEEID,Groups,CVNUMBER,_brn as 'Location' from journal_main inner join journal_detail using(journal_id) INNER JOIN checkingaccountlog on JID=journal_id  where JID='" & JID & "'")
            If T.Rows.Count > 0 Then
                If T(0)("STATUS").ToString.ToUpper = "ACTIVE" Then
                    'Dim UPCHK As New CHANGE_CHECK_DATE_ With {.Loc = T(0)("location"), .NameID = T(0)("PAYEEID"), .CHKDATE = T(0)("CHECKDATES"), .GROUP = T(0)("Groups"), .JID = JID, .OLDCHECK = T(0)("CHECKNUMBER"), .ACID = T(0)("ACID")}
                    'UPCHK.ShowDialog()
                Else
                    MessageBoxStr("Cannot change " & T(0)("STATUS").ToString.ToUpper & " transaction")
                End If
            Else
                MessageBoxStr("This is not a bank disbursement transaction")
            End If
        Catch ex As Exception
        End Try
    End Sub


    Public Function GET_JE_STATUS_BY_JID(ByVal JID As Integer) As String
        Try
            Dim stat As String
            stat = Dataobject("select status from  journal_main where journal_id='" & JID & "'")
            If stat = Nothing Then
                Return ""
            Else
                Return stat
            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function
    Sub Print2307(GUID As String)
        Try
            Dim JID As Integer = GET_JID_BY_GRP(GUID)
            Print2307(JID)
        Catch ex As Exception
        End Try
    End Sub


    Function GET_JOURNAL(ByVal JID As Integer) As DataTable
        Try
            Return Datasource("select  *  from  journal_main  where  journal_id='" & JID & "'")
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function GET_JOURNAL(ByVal GRP As String) As DataTable
        Try
            Return Datasource("select  *  from  journal_main WHERE  groupentry='" & GRP & "'")
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function GET_JE_STATUS_BY_GRP(ByVal GRP As String) As String
        Try
            Dim stat As New DataTable
            stat = Datasource("select status from  journal_main where  Groupentry='" & GRP & "'")
            If stat.Rows.Count > 0 Then
                Return stat(0)(0).ToString
            Else
                Return "LOCKED"
            End If
        Catch ex As Exception
            Return "LOCKED"
        End Try
    End Function
    Public Function GET_JID_BY_GRP(ByVal GRP As String) As Integer
        Try
            Dim JID As Integer
            JID = Dataobject("select  journal_id from  journal_main where  Groupentry='" & GRP & "' limit 1")
            If JID > 0 Then
                Return JID
            Else
                Return -1
            End If
        Catch ex As Exception
            Return -1
        End Try
    End Function

End Module
