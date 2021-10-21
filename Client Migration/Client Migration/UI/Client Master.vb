Imports DevExpress.XtraSplashScreen
Partial Public Class Client_Master
	Private ImportEx As ImportFromExcel
	Private FileName As String = ""
	Private CLIENTID As Integer
	Private UNITID As Integer
	Private PRJID As Integer
	Private TransGuid As String
	Private ChargeID As Integer
	Private Tdate As Date
	Private CashierPK As Integer = 0
	Public WithEvents LoginFrms As Configurations.LOGINFRM
	Public WithEvents NewConnectionFrm As Configurations.Connection
	Private Config As New Configurations.DBconfigurations

	Private Units As New List(Of UnitModel)
	Private Projects As New List(Of ProjectModel)
	Private Clients As New List(Of BusinessModel)
	Private Buyers As New List(Of BuyersModel)
	Private paymentschedmainRecords As New List(Of PaymentScheduleMainModel)

	'  Private PmtCols() As String = New String() {"Unit Code *", "Buyer Name *",
	'"Check #", "Amount", "Account Ref ",
	'"OR Date", "OR #", "Amount2", "Check Status", "Posting Date"}


	'  Private BuyerCols As String() = New String() {"Unit Code *", "Buyer Name *", "Vat", "Gross",
	'      "Discount 1", "Discount 2", "Add Vat", "TCP", "Reservation", "Move-In", "Transfer Charge",
	'      "Loan Takeout"}

	Public Sub New()
		InitializeComponent()
	End Sub

	Private Sub SimpleButton1_Click(sender As Object, e As EventArgs) Handles SimpleButton1.Click
		Try
			Dim F As New OpenFileDialog
			F.Title = "Open Excel files"
			F.Filter = "All files|*.*"
			If F.ShowDialog = DialogResult.OK Then
				FileName = F.FileName
				ImportEx = New ImportFromExcel(F.FileName)
				ComboBox1.DataSource = ImportEx.GetExcelSheets
				ComboBox1.SelectedIndex = 0
			End If
		Catch ex As Exception
		End Try
	End Sub

	Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
		Try
			SplashScreenManager.ShowDefaultWaitForm()
			GridView1.Columns.Clear()
			Dim T As New DataTable
			If ImportEx.IsConnected = True Then
				T = ImportEx.OpenExcelSheet(ComboBox1.Text)
				T.Columns.Add("Row")
				For i As Integer = 0 To T.Rows.Count - 1
					T(i)("Row") = i + 1
				Next
				GridControl1.DataSource = T
				GridView1.Columns("Row").VisibleIndex = 0
			End If
			SplashScreenManager.CloseDefaultWaitForm()
			GridView1.BestFitColumns()
			GridView1.OptionsView.ColumnAutoWidth = False
			'GridView1.OptionsBehavior.ReadOnly = True
			'GridView1.OptionsBehavior.Editable = False
		Catch ex As Exception
			MessageBoxError2(ex)
		End Try
	End Sub
	Function IsClientExist(ByVal ClientName As String) As Integer
		Try
			Dim IsClient = clients.SingleOrDefault(Function(x) x.Name.ToLower().Trim = ClientName.ToLower().Trim And x.Type.ToLower.Trim = "client")
			If IsClient?.ID > 0 Then Return IsClient.ID
			Return 0
			'Dim x = DB.QueryandReturnTable("SELECT IFNULL(business_number,0) 'ID' FROM general.`business_list` WHERE business_fullname = '" & ClientName.Trim & "' AND `TYPE`='Client'")
			'If x.Rows.Count > 0 Then
			'	Return x(0)(0)
			'End If
		Catch ex As Exception
			Return 0
		End Try
	End Function
	Function SAVE_BUSINESS(db As UCommand, _type As String, _FullName As String) As Integer
		Try
			'SAVE BUSINESS LIST
			'new Buyer
			Dim DIC As New Dictionary(Of String, String)
			DIC = New Dictionary(Of String, String)
			DIC.Add("business_fullname", _FullName.Trim.RSQ)
			DIC.Add("`TYPE`", _type)
			DIC.Add("`Sheet`", ComboBox1.Text.RSQ)
			db.Insert("general.business_list", DIC)
			Return db.LastPK
		Catch ex As Exception
			Return 0
		End Try
	End Function

	Sub SAVE_CLIENT(db As UCommand)
		Dim cntr As Integer = 0
		Try
			Dim view = GridView1
			Dim DIC As New Dictionary(Of String, String)
			Remove(db) '//remove data
			For I As Integer = 0 To GridView1.RowCount - 1
				cntr = I
				TransGuid = Guid.NewGuid().ToString()
				If view.GetRowCellValue(I, "Buyer Name *").ToString.Trim = "" Then
					DB.QueryExecNonQuery(" UPDATE propmanagement.tbl_property_unit SET unit_status='AVAILABLE',CurrentBuyerID='0' WHERE UNITNO='" & view.GetRowCellValue(I, "Unit Code *").ToString.Trim.RSQ & "'")
					Continue For
				End If
				If view.GetRowCellValue(I, "Unit Code *").ToString.Trim = "" Then
					Continue For
				End If
				If view.GetRowCellValue(I, "Buyer Name *").ToString.Trim = "Unsold" Or view.GetRowCellValue(I, "Buyer Name *").ToString.Trim = "UNSOLD" Or view.GetRowCellValue(I, "Buyer Name *").ToString.Trim = "NOT USED" Or view.GetRowCellValue(I, "Buyer Name *").ToString.Trim = "NotUsed" Or view.GetRowCellValue(I, "Buyer Name *").ToString.Trim = "" Then
					db.QueryExecNonQuery(" UPDATE propmanagement.tbl_property_unit SET unit_status='AVAILABLE',CurrentBuyerID='0' WHERE UNITNO='" & view.GetRowCellValue(I, "Unit Code *").ToString.Trim.RSQ & "'")
					Continue For
				End If
				If view.GetRowCellValue(I, "Buyer Name *").ToString.Trim = "Un-Sold" Then
					db.QueryExecNonQuery(" UPDATE propmanagement.tbl_property_unit SET unit_status='AVAILABLE',CurrentBuyerID='0' WHERE UNITNO='" & view.GetRowCellValue(I, "Unit Code *").ToString.Trim.RSQ & "'")
					Continue For
				End If
				CLIENTID = IsClientExist(view.GetRowCellValue(I, "Buyer Name *").ToString.Trim.RSQ)
				If CLIENTID > 0 Then
				Else
					'SAVE BUSINESS /CLIENT TO BUSINESS LIST
					CLIENTID = SAVE_BUSINESS(db, "CLIENT", view.GetRowCellValue(I, "Buyer Name *").ToString)
				End If

				'SAVE BUYER INFO  main
				DIC = New Dictionary(Of String, String)
				DIC.Add("GUID", TransGuid)
				DIC.Add("BUYERFULLNAME", view.GetRowCellValue(I, "Buyer Name *").ToString.RSQ)
				DIC.Add("business_id", CLIENTID)
				DIC.Add("buyer_firstname", view.GetRowCellValue(I, "Buyer Name *").ToString.RSQ)
				DIC.Add("citizenship", view.GetRowCellValue(I, "Citizenship").ToString.RSQ)
				' DIC.Add("Status", view.GetRowCellValue(I, "Buyer Status").ToString.RSQ)
				DIC.Add("Sheets", ComboBox1.Text.RSQ.Trim)
				db.Insert("propmanagement.buyersinfomain", DIC)
				'insert all business that is not in payee table..
				db.QueryExecNonQuery("INSERT INTO accounting.payee ( BID,PAYEE_NAME)  SELECT business_number,business_fullname FROM  general.business_list  WHERE business_number NOT IN (SELECT BID FROM  accounting.payee)")

				'SAVE AGENT 
				DIC = New Dictionary(Of String, String)
				DIC.Add("GUID", TransGuid)
				DIC.Add("AgentID", GetAgentInfo(db, view.GetRowCellValue(I, "Agent *").ToString.Trim))
				DIC.Add("Agent", view.GetRowCellValue(I, "Agent *").ToString.RSQ)
				DIC.Add("ManagerID", GetManagerInfo(db, view.GetRowCellValue(I, "Manager").ToString.Trim))
				DIC.Add("Manager", view.GetRowCellValue(I, "Manager").ToString.RSQ)
				DIC.Add("BrokerID", GetBrokerInfo(db, view.GetRowCellValue(I, "Broker").ToString.Trim))
				DIC.Add("Broker", view.GetRowCellValue(I, "Broker").ToString.RSQ)
				db.Insert("propmanagement.buyersinfoagents", DIC)

				'GET PROJECT AND UNIT INFO
				'//////////////////
				Dim PrjNo As String = ""
				Dim str = view.GetRowCellValue(I, "Unit Code *").ToString.Trim.Split("-")
				PrjNo = str(0)
				PRJID = get_projectID(db, PrjNo)
				If PRJID = 0 Then
					db.isError = True
					MessageBoxStr($"Project Code [{PrjNo}] not Found. Line:" & I + 1)
				End If
				UNITID = GetUnitID(db, view.GetRowCellValue(I, "Unit Code *").ToString)
				If UNITID = 0 Then
					db.isError = True
					MessageBoxStr($"Unit Code [{view.GetRowCellValue(I, "Unit Code *").ToString}] not Found. Line:" & I + 1)
				End If
				'///////////////// 
				'SAVE buyer project 
				DIC = New Dictionary(Of String, String)
				DIC.Add("GUID", TransGuid)
				DIC.Add("unit_id", UNITID)
				DIC.Add("project_id", PRJID)
				If IsNumeric(view.GetRowCellValue(I, "Net Selling Price")) Then
					DIC.Add("total_contract_price", CDbl(view.GetRowCellValue(I, "Net Selling Price")))
				End If
				db.Insert("propmanagement.buyersinfoproject", DIC)
				'SAVE SCHEDULE
				SAVE_SCHEDULE(db, I)
				'update Unit
				db.QueryExecNonQuery(" UPDATE propmanagement.tbl_property_unit SET unit_status='SOLD OUT',CurrentBuyerID='" & CLIENTID & "' WHERE UNITID='" & UNITID & "'")
			Next
		Catch ex As Exception
		End Try
	End Sub
	Function GetCompanyIDbyPrjID(ByVal PrjID As String) As Integer
		Try
			Dim proj = Projects.SingleOrDefault(Function(x) x.PROJECTID = PrjID)
			Return proj.COMPANYID
			'DB.QueryandReturnTable("SELECT title_id FROM general.`setup_project` WHERE project_id = '" & PrjID & "'")
			'If x.Rows.Count > 0 Then
			'	Return x(0)(0)
			'End If
		Catch ex As Exception
			Return 0
		End Try
	End Function

	Function GetUnitID(db As UCommand, ByVal unitN As String) As Integer
		Try
			Dim prj = units.SingleOrDefault(Function(x) x.UNITNO.ToLower.Trim = unitN.ToLower.Trim)
			Return prj.UNITID
			'Dim xx As New DataTable
			'xx = db.dats("SELECT IFNULL(UNITID,0)'ID' FROM propmanagement.`tbl_property_unit` WHERE UNITNO = '" & unitN.Trim & "'")
			'If xx.Rows.Count > 0 Then Return xx(0)(0)
			'Return 0
		Catch ex As Exception
			Return 0
		End Try
	End Function

	Function get_projectID(db As UCommand, ByVal proj_no As String) As Integer
		Try
			Dim prj = projects.SingleOrDefault(Function(x) x.PROJNO.ToLower.Trim = proj_no.ToLower.Trim)
			Return prj?.PROJECTID
			'Dim xX As New DataTable
			'xX = db.Datasource("SELECT project_id 'ID' FROM general.`setup_project` WHERE project_no = '" & proj_no & "'")
			'If xX.Rows.Count > 0 Then
			'	Return xX(0)(0)
			'End If
		Catch ex As Exception
			Return 0
		End Try
	End Function

	'AGENT
	Function GetAgentInfo(db As UCommand, AgentName As String) As Integer
		Try
			If AgentName.Trim = "" Then Return 0
			Dim agent = clients.Where(Function(x) x.Name.ToLower.Trim = AgentName.ToLower.Trim And x.Type.ToLower = "agent").FirstOrDefault()
			If agent IsNot Nothing Then
				Return agent.ID
			Else
				'SAVE AGENT 
				Dim DIC As New Dictionary(Of String, String)
				DIC = New Dictionary(Of String, String)
				DIC.Add("business_fullname", AgentName.Trim.RSQ)
				DIC.Add("`TYPE`", "AGENT")
				DIC.Add("`Sheet`", ComboBox1.Text.RSQ)
				db.Insert("general.business_list", DIC)
				Dim mId = db.LastPK
				clients.Add(New BusinessModel() With {.ID = mId, .Name = AgentName.Trim, .Sheet = ComboBox1.Text, .Type = "AGENT"})
				Return mId
			End If

			'Dim T As New DataTable
			'T = db.QueryandReturnTable("select business_number,business_fullname FROM  general.business_list where `Type`='Agent' and business_fullname='" & AgentName.Trim.RSQ & "'")
			'If T.Rows.Count > 0 Then
			'	Return T(0)("business_number")
			'Else
			'	'SAVE AGENT 
			'	Dim DIC As New Dictionary(Of String, String)
			'	DIC = New Dictionary(Of String, String)
			'	DIC.Add("business_fullname", AgentName.Trim.RSQ)
			'	DIC.Add("`TYPE`", "AGENT")
			'	DIC.Add("`Sheet`", ComboBox1.Text.RSQ)
			'	db.Insert("general.business_list", DIC)
			'	Return db.LastPK
			'End If
		Catch ex As Exception
			Return 0
		End Try
	End Function
	'MANAGER
	Function GetManagerInfo(db As UCommand, ManagerName As String) As Integer
		Try
			If ManagerName.Trim = "" Then Return 0
			Dim manager = clients.Where(Function(x) x.Name.ToLower.Trim = ManagerName.ToLower.Trim And x.Type.ToLower = "manager").FirstOrDefault()
			If manager IsNot Nothing Then
				Return manager.ID
			Else
				Dim DIC As New Dictionary(Of String, String)
				DIC = New Dictionary(Of String, String)
				DIC.Add("business_fullname", ManagerName.Trim.RSQ)
				DIC.Add("`TYPE`", "MANAGER")
				DIC.Add("`Sheet`", ComboBox1.Text.RSQ)
				db.Insert("general.business_list", DIC)
				Dim mId = db.LastPK
				clients.Add(New BusinessModel() With {.ID = mId, .Name = ManagerName.Trim, .Sheet = ComboBox1.Text, .Type = "MANAGER"})
				Return mId
			End If

			'Dim T As New DataTable
			'T = db.Datasource("select business_number,business_fullname FROM  general.business_list where `Type`='Manager' and business_fullname='" & ManagerName.Trim.RSQ & "'")
			'If T.Rows.Count > 0 Then
			'	Return T(0)("business_number")
			'Else
			'	'SAVE AGENT 
			'	Dim DIC As New Dictionary(Of String, String)
			'	DIC = New Dictionary(Of String, String)
			'	DIC.Add("business_fullname", ManagerName.Trim.RSQ)
			'	DIC.Add("`TYPE`", "MANAGER")
			'	DIC.Add("`Sheet`", ComboBox1.Text.RSQ)
			'	db.Insert("general.business_list", DIC)
			'	Return db.LastPK
			'End If
		Catch ex As Exception
			Return 0
		End Try
	End Function
	'BROKER
	Function GetBrokerInfo(db As UCommand, BrokerName As String) As Integer
		Try
			If BrokerName.Trim = "" Then Return 0
			Dim broker = clients.Where(Function(x) x.Name.ToLower.Trim = BrokerName.ToLower.Trim And x.Type.ToLower = "broker").FirstOrDefault()
			If broker IsNot Nothing Then
				Return broker.ID
			Else
				Dim DIC As New Dictionary(Of String, String)
				DIC = New Dictionary(Of String, String)
				DIC.Add("business_fullname", BrokerName.Trim.RSQ)
				DIC.Add("`TYPE`", "BROKER")
				DIC.Add("`Sheet`", ComboBox1.Text.RSQ)
				db.Insert("general.business_list", DIC)
				Dim mId = db.LastPK
				clients.Add(New BusinessModel() With {.ID = mId, .Name = BrokerName.Trim, .Sheet = ComboBox1.Text, .Type = "BROKER"})
				Return mId
			End If

			'If BrokerName.Trim = "" Then Return 0
			'Dim T As New DataTable
			'T = DB.QueryandReturnTable("select business_number,business_fullname FROM  general.business_list where `Type`='Broker' and business_fullname='" & BrokerName.Trim.RSQ & "'")
			'If T.Rows.Count > 0 Then
			'	Return T(0)("business_number")
			'Else
			'	'SAVE AGENT 
			'	Dim DIC As New Dictionary(Of String, String)
			'	DIC = New Dictionary(Of String, String)
			'	DIC.Add("business_fullname", BrokerName.Trim.RSQ)
			'	DIC.Add("`TYPE`", "BROKER")
			'	DIC.Add("`Sheet`", ComboBox1.Text.RSQ)
			'	DB.Insert("general.business_list", DIC)
			'	Return DB.LastPK
			'End If

		Catch ex As Exception
			Return 0
		End Try
	End Function
	Sub SAVE_SCHEDULE(db As UCommand, i As Integer)
		Try
			Dim Discount1 As Decimal = 0, Discount2 As Decimal = 0
			Dim VIEW = GridView1
			Dim DIC As New Dictionary(Of String, String)
			DIC.Add("PRJID", PRJID)
			DIC.Add("UNITID", UNITID)
			DIC.Add("GUID", TransGuid)
			DIC.Add("CLIENTID", CLIENTID)
			DIC.Add("BUYERGUID", TransGuid)
			DIC.Add("trans_date", Tdate.ToMysqlFormat)
			DIC.Add("no_months", 0) ' VIEW.GetRowCellValue(i, "Deferred Months"))

			'list price
			If IsNumeric(VIEW.GetRowCellValue(i, "List Price *")) Then
				DIC.Add("original_sell_price", CDbl(VIEW.GetRowCellValue(i, "List Price *")))
			End If
			'Vat 1
			If IsNumeric(VIEW.GetRowCellValue(i, "Vat")) Then
				DIC.Add("IsVat", If(CDbl(VIEW.GetRowCellValue(i, "Vat")) > 0, "V", "NV"))
				DIC.Add("vat_percentage", 12)
				DIC.Add("vat_amount", CDbl(VIEW.GetRowCellValue(i, "Vat")))
			End If

			'Gross
			If IsNumeric(VIEW.GetRowCellValue(i, "Gross")) Then
				DIC.Add("total_selling_price", CDbl(VIEW.GetRowCellValue(i, "Gross")))
			End If

			'Discounts
			If IsNumeric(VIEW.GetRowCellValue(i, "Discount 1")) Then
				DIC.Add("discount1_amount", CDbl(VIEW.GetRowCellValue(i, "Discount 1")))
				Discount1 = CDbl(VIEW.GetRowCellValue(i, "Discount 1"))
			End If

			If IsNumeric(VIEW.GetRowCellValue(i, "Discount 2")) Then
				DIC.Add("discount2_amount", CDbl(VIEW.GetRowCellValue(i, "Discount 2")))
				Discount2 = CDbl(VIEW.GetRowCellValue(i, "Discount 2"))
			End If

			'Net Price/Sub Total
			If IsNumeric(VIEW.GetRowCellValue(i, "List Price *")) Then
				DIC.Add("SubTotal", CDbl(VIEW.GetRowCellValue(i, "List Price *")) - Discount1 - Discount2)
			End If

			'add vat >0
			If IsNumeric(VIEW.GetRowCellValue(i, "Add Vat")) Then
				If CDbl(VIEW.GetRowCellValue(i, "Add Vat")) > 0 Then
					DIC.Add("net_vat_amount", CDbl(VIEW.GetRowCellValue(i, "Add Vat")))
				End If
			End If
			'TCP
			If IsNumeric(VIEW.GetRowCellValue(i, "TCP")) Then
				DIC.Add("total_net_selling_price", CDbl(VIEW.GetRowCellValue(i, "TCP")))
			End If

			'DIC.Add("equity_percentage",)
			'DIC.Add("equity_amount",)  
			Try
				If IsNumeric(VIEW.GetRowCellValue(i, "Transfer Charge")) And IsNumeric(VIEW.GetRowCellValue(i, "TCP")) Then
					DIC.Add("transfer_charges_percentage", CDbl(FormatNumber(CDbl(VIEW.GetRowCellValue(i, "Transfer Charge")) / CDbl(VIEW.GetRowCellValue(i, "TCP")) * 100, 2)))
					DIC.Add("transfer_amount", CDbl(VIEW.GetRowCellValue(i, "Transfer Charge")))
				End If
			Catch ex As Exception
			End Try
			Try
				If IsNumeric(VIEW.GetRowCellValue(i, "Move-In")) And IsNumeric(VIEW.GetRowCellValue(i, "TCP")) Then
					DIC.Add("move_in_fees_percentage", CDbl(FormatNumber(CDbl(VIEW.GetRowCellValue(i, "Move-In")) / CDbl(VIEW.GetRowCellValue(i, "TCP")) * 100, 2)))
					DIC.Add("movein_amount", CDbl(VIEW.GetRowCellValue(i, "Move-In")))
				End If
			Catch ex As Exception
			End Try

			'DIC.Add("equity_amount", 0) ' CDbl(VIEW.GetRowCellValue(i, "Equity Amount")))
			'DIC.Add("equity_percentage", 0) ' CDbl(VIEW.GetRowCellValue(i, "Equity  Pecentage")))
			Try
				If IsNumeric(VIEW.GetRowCellValue(i, "Loan Takeout")) Then
					DIC.Add("take_out_loan_amount", CDbl(VIEW.GetRowCellValue(i, "Loan Takeout")))
				End If
			Catch ex As Exception
			End Try
			DIC.Add("trans_type", "EMI")
			db.Insert("propmanagement.paymentschedmain", DIC)
		Catch ex As Exception
		End Try
	End Sub
	Function GetSchedGUID() As String
		Try
			Dim item = paymentschedmainRecords.SingleOrDefault(Function(x) x.CLIENTID = CLIENTID And x.UNITID = UNITID)
			Return item.GUID
			'Return db.DataObject("SELECT GUID FROM   propmanagement.paymentschedmain WHERE CLIENTID='" & CLIENTID & "' AND UNITID='" & UNITID & "'")
		Catch ex As Exception
			Return ""
		End Try
	End Function
	'Sub SAVE_SETUP()
	'	Try
	'		Dim View = GridView1
	'		Dim DIC As New Dictionary(Of String, String)

	'		For i As Integer = 0 To View.RowCount - 1
	'			UNITID = GetUnitID(View.GetRowCellValue(i, "Unit Code *"))
	'			If UNITID = 0 Then
	'				MessageXtra("Unit Code not Found. Line:" & i + 1)
	'				Exit Sub
	'			End If
	'			CLIENTID = IsClientExist(View.GetRowCellValue(i, "Buyer Name *").ToString.Trim.RSQ)
	'			If CLIENTID = 0 Then
	'				MessageXtra("Client Not Found. line: " & i + 1)
	'				Exit Sub
	'			End If
	'			TransGuid = GetSchedGUID()
	'			If TransGuid = "" Then
	'				MessageXtra("Cannot Locate Payment Schedule. line: " & i + 1)
	'				Exit Sub
	'			End If
	'			DIC = New Dictionary(Of String, String)
	'			DIC.Add("`FROM`", View.GetRowCellValue(i, "From Month").ToString.RSQ)
	'			DIC.Add("`TO`", View.GetRowCellValue(i, "From Month").ToString.RSQ)
	'			DIC.Add("Percent", View.GetRowCellValue(i, "From Month").ToString.RSQ)
	'			DIC.Add("Amount", CDbl(View.GetRowCellValue(i, "From Month")))
	'			DIC.Add("GUID", TransGuid)
	'			DIC.Add("Sheet", ComboBox1.Text.Trim.RSQ)
	'			DB.Insert(" propmanagement.stepupsetup", DIC)
	'		Next
	'		DB.WriteToDB()
	'	Catch ex As Exception
	'		DB.RollBackAction(ex.Message)
	'	End Try
	'End Sub
	'Function GetUnitInfo(db As UCommand, UNIT As String) As DataTable
	'	Return db.QueryandReturnTable("select * from  propmanagement.tbl_property_unit WHERE UNITNO='" & UNIT.RSQ & "'")
	'End Function
	Function GetBuyer(db As UCommand) As Integer
		Try
			Dim buyer = Buyers.SingleOrDefault(Function(x) x.CLIENTID = CLIENTID)
			Return buyer.ID
			'Return DataObject("SELECT ID FROM propmanagement.buyersinfomain WHERE business_id='" & CLIENTID & "'")
		Catch ex As Exception
			Return 0
		End Try
	End Function

	Function FirstValidationCheck() As Boolean
		Try
			Dim FieldName As String() = New String() {"Unit Code *", "Buyer Name *", "List Price *", "Vat", "Gross", "Discount 1",
			"Discount 2", "Add Vat", "TCP", "Reservation", "Move-In",
			"Transfer Charge", "Loan Takeout", "Buyer Status"}

			If ComboBox1.Text.Trim = "'Buyers Info$'" Then
				MessageBoxStr("Please rename sheet Name " & vbNewLine & "Example: Buyers Info P05-05F")
				Return False
			End If

			Dim ExcelCols As New List(Of String)
			ExcelCols.Clear()

			For i = 0 To GridView1.Columns.Count - 1
				Dim str As String = GridView1.Columns(i).FieldName
				ExcelCols.Add(New String(str))
			Next

			For Each i In FieldName
				Dim f = ExcelCols.Where(Function(x) x = i)
				If f.Count = 0 Then
					MessageBoxStr(String.Format("Column Name {0} is  not found.Please check excels Column Name", i))
					Dim c As New validFieldNameBuyer
					c.ShowDialog(Me)
					Return False
				End If
			Next
			Return True
		Catch ex As Exception
			Return False
		End Try
	End Function
	Function Validation() As Boolean
		Try
			Dim View = GridView1
			If View.RowCount = 0 Then
				MessageBoxStr("Nothing to save")
				Return False
			End If
			If ComboBox1.Text.Trim.ToString.ToLower.Contains("buyer") Then
				If FirstValidationCheck() = False Then
					Return False
				End If
			End If
			If ComboBox1.Text.Trim.ToString.ToLower.Contains("payment") Then
				If SecondValidationCheck() = False Then
					Return False
				End If
			End If

			For i As Integer = 0 To GridView1.RowCount - 1
				If View.GetRowCellValue(i, "Unit Code *").ToString.Trim = "" Then
					MessageBoxStr("Invalid Unit Code Line:" & i + 1)
					Return False
				End If
				'need to have a client Name
				If ComboBox1.Text.Trim.ToString.ToLower.Contains("payment") Then
					If ComboBox1.Text.Trim.ToString.ToLower.Contains("buyer") And View.GetRowCellValue(i, "Buyer Name *").ToString.Trim = "" Then
						MessageBoxStr("Invalid Client Name Line:" & i + 1)
						Return False
					End If
				End If
			Next
			Return True
		Catch ex As Exception
			MessageBoxStr("Please check Column name")
			Return False
		End Try
	End Function
	Sub SAVE_CHARGES(db As UCommand)
		Try
			Dim cNTR As Integer = 0
			Dim View = GridView1
			Remove(db)
			Dim DIC As New Dictionary(Of String, String)
			cNTR = 0
			'SAVE CHARGES AND PAYMENTS

			For I As Integer = 0 To GridView1.RowCount - 1
				cNTR += 1
				Dim PrjNo As String = ""
				CashierPK = 0
				Dim str = View.GetRowCellValue(I, "Unit Code *").ToString.Trim.Split("-")
				PrjNo = str(0)
				If PrjNo.Trim = "" Then
					db.isError = True
					MessageBoxStr("Invalid project . line: " & I + 1)
					Exit Sub
				End If
				If View.GetRowCellValue(I, "Charge Name *").ToString.Trim = "" Then
					db.isError = True
					MessageBoxStr("Invalid Charge Name. line: " & I + 1)
					Exit Sub
				End If
				If View.GetRowCellValue(I, "Particulars *").ToString.Trim = "" Then
					db.isError = True
					MessageBoxStr("Invalid Particulars. line: " & I + 1)
					Exit Sub
				End If
				PRJID = get_projectID(db, PrjNo)
				If PRJID = 0 Then
					db.isError = True
					MessageBoxStr("Project Code not Found. Line:" & I + 1)
					Exit Sub
				End If
				UNITID = GetUnitID(db, View.GetRowCellValue(I, "Unit Code *"))
				If UNITID = 0 Then
					db.isError = True
					MessageBoxStr("Unit Code not Found. Line:" & I + 1)
					Exit Sub
				End If
				CLIENTID = IsClientExist(View.GetRowCellValue(I, "Buyer Name *").ToString.Trim.RSQ)
				If CLIENTID = 0 Then
					db.isError = True
					MessageBoxStr("Client Not Found. line: " & I + 1)
					Exit Sub
				End If
				Dim BuyerID As Integer = GetBuyer(db)
				If BuyerID = 0 Then
					db.isError = True
					MessageBoxStr("Invalid Buyer- Line:  " & I + 1)
					Exit Sub
				End If
				If View.GetRowCellValue(I, "MOP *").ToString.Trim <> "" Then
					If View.GetRowCellValue(I, "Check Status").ToString.Trim = "" Then
						db.isError = True
						MessageBoxStr("Invalid Check Status @Line:  " & I + 1)
						Exit Sub
					End If
				End If

				If View.GetRowCellValue(I, "MOP *").ToString.Trim <> "" Then
					Dim Amt As Double = 0
					If IsNumeric(GridView1.GetRowCellValue(I, "Amount2")) Then
						Amt = CDbl(View.GetRowCellValue(I, "Amount2"))
					End If
					If Amt = 0 Then
						If IsNumeric(GridView1.GetRowCellValue(I, "Amount")) Then
							Amt = CDbl(View.GetRowCellValue(I, "Amount"))
						End If
					End If
					If Amt = 0 Then
						db.isError = True
						MessageBoxStr("Invalid amount @line:  " & I + 1)
						Exit Sub
					End If
				End If

				If View.GetRowCellValue(I, "Charge Name *").ToString.Trim <> "" And View.GetRowCellValue(I, "Particulars *").ToString.Trim = "" Then
					db.isError = True
					MessageBoxStr("Invalid Particulars @line:  " & I + 1)
					Exit Sub
				End If
				TransGuid = GetSchedGUID()
				If TransGuid = "" Then
					db.isError = True
					MessageBoxStr("Cannot Locate Payment Schedule. line: " & I + 1)
					Exit Sub
				End If

				DIC = New Dictionary(Of String, String)
				DIC.Add("GUID", TransGuid)
				DIC.Add("BUYERID", BuyerID)
				DIC.Add("BID", CLIENTID)
				DIC.Add("BuyerName", GridView1.GetRowCellValue(I, "Buyer Name *").ToString.RSQ)
				DIC.Add("UNITID", UNITID)
				DIC.Add("UNITNO", GridView1.GetRowCellValue(I, "Unit Code *").ToString.RSQ.Trim)
				DIC.Add("PRJID", PRJID)
				DIC.Add("TITLEID", GetCompanyIDbyPrjID(PRJID))

				DIC.Add("Date_description", GridView1.GetRowCellValue(I, "Particulars *").ToString.Trim.RSQ)
				DIC.Add("PMTTYPE", GridView1.GetRowCellValue(I, "Payment Scheme").ToString.Trim.RSQ)
				DIC.Add("VisibleInReceivedPmt", 1)
				DIC.Add("Sheet", ComboBox1.Text.Trim.RSQ)
				DIC.Add("ChargeName", GridView1.GetRowCellValue(I, "Charge Name *").ToString.Trim.RSQ)

				If IsNumeric(GridView1.GetRowCellValue(I, "Due Amount *")) Then
					DIC.Add("DueAmount", CDbl(GridView1.GetRowCellValue(I, "Due Amount *")))
				End If
				If IsDate(GridView1.GetRowCellValue(I, "Due Date *")) Then
					DIC.Add("DueDate", CDate(GridView1.GetRowCellValue(I, "Due Date *")).ToMysqlFormat)
				End If

				''' for PMT status 
				Dim DueAmt As Double = 0
				Dim Amt1 As Double = 0
				Dim Amt2 As Double = 0
				If IsNumeric(GridView1.GetRowCellValue(I, "Due Amount *")) Then
					DueAmt = GridView1.GetRowCellValue(I, "Due Amount *")
				End If
				If IsNumeric(GridView1.GetRowCellValue(I, "Amount2")) Then
					Amt2 = GridView1.GetRowCellValue(I, "Amount2")
				End If

				If IsNumeric(GridView1.GetRowCellValue(I, "Amount")) Then
					Amt1 = GridView1.GetRowCellValue(I, "Amount")
				End If

				'If DueAmt > 0 Then
				'    If Amt2 > 0 Then
				'        If DueAmt <= Amt2 Then
				'            DIC.Add("PMTSTAT", "Paid")
				'        ElseIf Amt2 > 0 And Amt2 < DueAmt Then
				'            DIC.Add("PMTSTAT", "Partially Paid")
				'        End If
				'    ElseIf Amt1 > 0 Then
				'        If DueAmt <= Amt1 Then
				'            DIC.Add("PMTSTAT", "Paid")
				'        ElseIf Amt1 > 0 And Amt1 < DueAmt Then
				'        End If
				'    End If
				'End If
				Try
					DIC.Add("PMTSTAT", GridView1.GetRowCellValue(I, "Payments Status"))
				Catch ex As Exception
					db.isError = True
					MessageBoxStr("Invalid Payment Status @line:  " & I + 1)
					Exit Sub
				End Try

				db.Insert("propmanagement.allcharges", DIC)
				ChargeID = db.LastPK

				''' save schedule 
				SAVE_SCHED_DETAILS(db, I)
				'//Update loan takeout amount
				If IsNumeric(GridView1.GetRowCellValue(I, "Due Amount *")) Then
					If GridView1.GetRowCellValue(I, "Charge Name *").ToString.ToLower.Contains("balance") Then
						db.QueryExecNonQuery("Update propmanagement.paymentschedmain set take_out_loan_amount='" & CDbl(GridView1.GetRowCellValue(I, "Due Amount *")) & "' where UNITID='" & UNITID & "' and CLIENTID='" & CLIENTID & "' ")
					ElseIf GridView1.GetRowCellValue(I, "Charge Name *").ToString.ToLower.Contains("loan") Then
						db.QueryExecNonQuery("Update propmanagement.paymentschedmain set take_out_loan_amount='" & CDbl(GridView1.GetRowCellValue(I, "Due Amount *")) & "' where UNITID='" & UNITID & "' and CLIENTID='" & CLIENTID & "' ")
					ElseIf GridView1.GetRowCellValue(I, "Charge Name *").ToString.ToLower.Contains("takeout") Then
						db.QueryExecNonQuery("Update propmanagement.paymentschedmain set take_out_loan_amount='" & CDbl(GridView1.GetRowCellValue(I, "Due Amount *")) & "' where UNITID='" & UNITID & "' and CLIENTID='" & CLIENTID & "' ")
					End If
				End If
				Save_to_cashier(db, I)
				SAVE_APPLIED_PAYMENT(db, I)
				SAVE_RESERVATION(db, I)
			Next
		Catch ex As Exception
		End Try
	End Sub
	Sub SAVE_RESERVATION(db As UCommand, I As Integer)
		Try
			Dim View = GridView1
			If View.GetRowCellValue(I, "Charge Name *").ToString.ToLower.Contains("RF") Or GridView1.GetRowCellValue(I, "Charge Name *").ToString.ToLower.Contains("RESERVATION") Or GridView1.GetRowCellValue(I, "Charge Name *").ToString.ToLower.Contains("RF FEE") Or GridView1.GetRowCellValue(I, "Charge Name *").ToString.ToLower.Contains("RF FEES") Then
				If IsDate(View.GetRowCellValue(I, "Posting Dates")) = False Then Exit Sub
				'IF AMOUNT IS GREATER THAN ZERO 
				Dim Amt As Double = 0
				If IsNumeric(GridView1.GetRowCellValue(I, "Amount2")) Then
					Amt = CDbl(View.GetRowCellValue(I, "Amount2"))
				End If
				If Amt = 0 Then
					If IsNumeric(GridView1.GetRowCellValue(I, "Amount")) Then
						Amt = CDbl(View.GetRowCellValue(I, "Amount"))
					End If
				End If
				If Amt = 0 Then Exit Sub
				db.QueryExecNonQuery("INSERT INTO  reservation SET GUID='" & TransGuid & "', AMT='" & Amt & "', UNITID='" & UNITID & "', UNITNO='" & View.GetRowCellValue(I, "Unit Code *").ToString.Trim.RSQ & "', CLIENTID='" & CLIENTID & "', CustName='" & View.GetRowCellValue(I, "Buyer Name *").ToString.Trim.RSQ & "', RSDate='" & CDate(View.GetRowCellValue(I, "Posting Date")).ToMysqlFormat & "'")
			End If
		Catch ex As Exception
		End Try
	End Sub
	Sub SAVE_APPLIED_PAYMENT(db As UCommand, I As Integer)
		Try
			Dim VIEW = GridView1
			Dim Applied As New Dictionary(Of String, String)
			If VIEW.GetRowCellValue(I, "MOP *").ToString.Trim = "" Then Exit Sub
			If VIEW.GetRowCellValue(I, "Check Status").ToString.Trim.ToLower <> "cleared" Then Exit Sub
			'IF AMOUNT IS GREATER THAN ZERO
			Dim Amt As Double = 0
			Dim Bank As String = ""
			If IsNumeric(GridView1.GetRowCellValue(I, "Amount2")) Then
				Amt = CDbl(VIEW.GetRowCellValue(I, "Amount2"))
			End If
			If Amt = 0 Then
				If IsNumeric(GridView1.GetRowCellValue(I, "Amount")) Then
					Amt = CDbl(VIEW.GetRowCellValue(I, "Amount"))
				End If
			End If
			If Amt = 0 Then Exit Sub

			Applied = New Dictionary(Of String, String)
			Applied.Add("CashierGUID", TransGuid)
			Applied.Add("CashierID", CashierPK)
			Applied.Add("GUID", TransGuid)
			Applied.Add("CMPID", GetCompanyIDbyPrjID(PRJID))
			Applied.Add("PROJECTID", PRJID)
			Applied.Add("UNITID", UNITID)
			Applied.Add("UNITNO", VIEW.GetRowCellValue(I, "Unit Code *").ToString.RSQ)
			Applied.Add("CUSTOMERID", CLIENTID)
			Applied.Add("Customer", VIEW.GetRowCellValue(I, "Buyer Name *").ToString.RSQ)
			Applied.Add("Description", GridView1.GetRowCellValue(I, "Particulars *").ToString.Trim.RSQ)
			Applied.Add("MOP", VIEW.GetRowCellValue(I, "MOP *"))
			Applied.Add("ORNum", VIEW.GetRowCellValue(I, "OR #").ToString.RSQ)
			Applied.Add("CRNum", VIEW.GetRowCellValue(I, "CR #").ToString.RSQ)
			Applied.Add("ARNum", VIEW.GetRowCellValue(I, "AR # *").ToString.RSQ)
			Applied.Add("ChargeName", VIEW.GetRowCellValue(I, "Charge Name *").ToString.Trim.RSQ)
			Applied.Add("CheckNumber", VIEW.GetRowCellValue(I, "Check #").ToString.RSQ.Trim)
			Applied.Add("ALLCHARGESID", ChargeID)

			If VIEW.GetRowCellValue(I, "MOP *").ToString.Trim = "Cash" Then
				Bank = VIEW.GetRowCellValue(I, "Bank").ToString.Trim & " " & VIEW.GetRowCellValue(I, "Account Ref ").ToString.Trim
			Else
				Bank = VIEW.GetRowCellValue(I, "MOP *").ToString.Trim
			End If

			If VIEW.GetRowCellValue(I, "Bank").ToString.RSQ.Trim.ToLower <> "" Then
				Bank = VIEW.GetRowCellValue(I, "Bank").ToString.Trim
			ElseIf VIEW.GetRowCellValue(I, "Account Ref ").ToString.RSQ.Trim.ToLower <> "" Then
				Bank = VIEW.GetRowCellValue(I, "Account Ref ").ToString.Trim
			Else
				Bank = VIEW.GetRowCellValue(I, "MOP *").ToString.Trim
			End If

			Applied.Add("Bank", Bank.RSQ)
			Applied.Add("Amount", Amt)
			If IsDate(VIEW.GetRowCellValue(I, "OR Date")) Then
				Applied.Add("PostingDate", CDate(VIEW.GetRowCellValue(I, "OR Date")).ToMysqlFormat)
				Applied.Add("DocDate", CDate(VIEW.GetRowCellValue(I, "OR Date")).ToMysqlFormat)
			ElseIf IsDate(VIEW.GetRowCellValue(I, "CR Date")) Then
				Applied.Add("PostingDate", CDate(VIEW.GetRowCellValue(I, "CR Date")).ToMysqlFormat)
				Applied.Add("DocDate", CDate(VIEW.GetRowCellValue(I, "CR Date")).ToMysqlFormat)
			ElseIf IsDate(VIEW.GetRowCellValue(I, "Posting Date")) Then
				Applied.Add("PostingDate", CDate(VIEW.GetRowCellValue(I, "Posting Date")).ToMysqlFormat)
				Applied.Add("DocDate", CDate(VIEW.GetRowCellValue(I, "Posting Date")).ToMysqlFormat)
			End If

			'///////////////
			If IsDate(VIEW.GetRowCellValue(I, "Due Date *")) Then
				Applied.Add("DueDate", CDate(VIEW.GetRowCellValue(I, "Due Date *")).ToMysqlFormat)
			End If
			If IsDate(VIEW.GetRowCellValue(I, "AR Date *")) Then
				Applied.Add("ARDate", CDate(VIEW.GetRowCellValue(I, "AR Date *")).ToMysqlFormat)
			End If
			If IsDate(VIEW.GetRowCellValue(I, "Check Date ")) Then
				Applied.Add("CheckDate", CDate(VIEW.GetRowCellValue(I, "Check Date ")).ToMysqlFormat)
			End If
			db.Insert("propmanagement.appliedpayment", Applied)
		Catch ex As Exception
		End Try
	End Sub
	Sub SAVE_SCHED_DETAILS(db As UCommand, I As Integer)
		Try
			Dim DIC = New Dictionary(Of String, String)
			DIC = New Dictionary(Of String, String)
			DIC.Add("GUID", TransGuid)
			DIC.Add("ChargeName", GridView1.GetRowCellValue(I, "Charge Name *").ToString.Trim.RSQ)
			DIC.Add("date_description", GridView1.GetRowCellValue(I, "Particulars *").ToString.Trim.RSQ)
			If IsDate(GridView1.GetRowCellValue(I, "Due Date *")) Then
				DIC.Add("payment_date", CDate(GridView1.GetRowCellValue(I, "Due Date *")).ToMysqlFormat)
			End If
			If IsNumeric(GridView1.GetRowCellValue(I, "Due Amount *")) Then
				DIC.Add("payment_amount", CDbl(GridView1.GetRowCellValue(I, "Due Amount *")))
			End If
			db.Insert("propmanagement.tbl_payment_details", DIC)
			db.QueryExecNonQuery("Update propmanagement.allcharges set PaymentDetailsID='" & db.LastPK & "' where ID='" & ChargeID & "'")
		Catch ex As Exception
		End Try
	End Sub

	Sub Save_to_cashier(db As UCommand, I As Integer)
		Try
			Dim VIEW = GridView1
			If VIEW.GetRowCellValue(I, "MOP *").ToString.Trim = "" Then Exit Sub
			'IF AMOUNT IS GREATER THAN ZERO
			Dim Amt As Double = 0
			Dim Bank As String = ""
			If IsNumeric(GridView1.GetRowCellValue(I, "Amount2")) Then
				Amt = CDbl(VIEW.GetRowCellValue(I, "Amount2"))
			End If
			If Amt = 0 Then
				If IsNumeric(GridView1.GetRowCellValue(I, "Amount")) Then
					Amt = CDbl(VIEW.GetRowCellValue(I, "Amount"))
				End If
			End If
			If Amt = 0 Then Exit Sub

			If VIEW.GetRowCellValue(I, "MOP *").ToString.Trim = "Cash" Then
				Bank = VIEW.GetRowCellValue(I, "Bank").ToString.Trim & " " & VIEW.GetRowCellValue(I, "Account Ref ").ToString.Trim
			Else
				Bank = VIEW.GetRowCellValue(I, "MOP *").ToString.Trim
			End If

			If VIEW.GetRowCellValue(I, "Bank").ToString.RSQ.Trim.ToLower <> "" Then
				Bank = VIEW.GetRowCellValue(I, "Bank").ToString.Trim
			ElseIf VIEW.GetRowCellValue(I, "Account Ref ").ToString.RSQ.Trim.ToLower <> "" Then
				Bank = VIEW.GetRowCellValue(I, "Account Ref ").ToString.Trim
			Else
				Bank = VIEW.GetRowCellValue(I, "MOP *").ToString.Trim
			End If

			Select Case VIEW.GetRowCellValue(I, "MOP *").ToString.ToLower.Trim
				Case "check", "cash", "online", "credit card"
				Case Else
					db.isError = True
					MessageBoxStr(String.Format("{0} is not identified @ row{1}", VIEW.GetRowCellValue(I, "MOP *").ToString.Trim, I))
			End Select

			Dim cashier As New Dictionary(Of String, String)
			cashier = New Dictionary(Of String, String)
			With cashier
				.Add("PMTRREF", "")
				.Add("Description", GridView1.GetRowCellValue(I, "Particulars *").ToString.Trim.RSQ)
				.Add("UNITID", UNITID)
				.Add("UNITNO", VIEW.GetRowCellValue(I, "Unit Code *").ToString.Trim.RSQ)
				.Add("PROJECTID", PRJID)
				.Add("CUSTOMERID", CLIENTID)
				.Add("Customer", VIEW.GetRowCellValue(I, "Buyer Name *").ToString.Trim.RSQ)
				.Add("CMPID", GetCompanyIDbyPrjID(PRJID))
				.Add("ARNum", VIEW.GetRowCellValue(I, "AR # *").ToString.Trim.RSQ)
				.Add("GUID", TransGuid)
				.Add("AllChargeGUID", TransGuid)
				.Add("ALLCHARGESID", ChargeID)
				.Add("JID", 0)
				.Add("STATUS", "Transmitted")
				.Add("MOP", VIEW.GetRowCellValue(I, "MOP *").ToString.RSQ.Trim)
				.Add("CheckNumber", VIEW.GetRowCellValue(I, "Check #").ToString.RSQ.Trim)
				.Add("ChargeName", GridView1.GetRowCellValue(I, "Charge Name *").ToString.Trim.RSQ)
				.Add("Amount", Amt)
				.Add("OrigAmt", Amt)
				.Add("Bank", Bank.RSQ)

				If GridView1.GetRowCellValue(I, "Check Status").ToString.Trim.ToLower = "unclear" Or GridView1.GetRowCellValue(I, "Check Status").ToString.Trim.ToLower = "uncleared" Then
					.Add("BankStatus", "Un-Cleared")
				ElseIf VIEW.GetRowCellValue(I, "Check Status").ToString.Trim = "" Then
					.Add("BankStatus", "Un-Cleared")
				Else
					.Add("BankStatus", GridView1.GetRowCellValue(I, "Check Status").ToString.Trim.RSQ)
				End If

				'PmtStatus
				If VIEW.GetRowCellValue(I, "Check Status").ToString.Trim.ToUpper = "DAIF" Then
					.Add("PmtStatus", "Cancelled")
				ElseIf VIEW.GetRowCellValue(I, "Check Status").ToString.Trim.ToUpper = "DAUD" Then
					.Add("PmtStatus", "Cancelled")
				ElseIf VIEW.GetRowCellValue(I, "Check Status").ToString.Trim.ToLower.Contains("alteration") Then
					.Add("PmtStatus", "Cancelled")
				Else
					.Add("PmtStatus", "Active")
				End If

				If IsDate(VIEW.GetRowCellValue(I, "Due Date *")) Then
					.Add("DueDate", CDate(GridView1.GetRowCellValue(I, "Due Date *")).ToMysqlFormat)
				End If

				If IsDate(VIEW.GetRowCellValue(I, "Posting Date")) Then
					.Add("PmtDate", CDate(VIEW.GetRowCellValue(I, "Posting Date")).ToMysqlFormat)
				Else
					.Add("PmtDate", Now.ToMysqlFormat)
				End If
				If IsDate(VIEW.GetRowCellValue(I, "AR Date *")) Then
					.Add("ARDate", CDate(VIEW.GetRowCellValue(I, "AR Date *")).ToMysqlFormat)
				End If
				If IsDate(VIEW.GetRowCellValue(I, "Check Date ")) Then
					.Add("CheckDate", CDate(VIEW.GetRowCellValue(I, "Check Date ")).ToMysqlFormat)
				End If

			End With
			db.Insert("propmanagement.cashierpayment_Bk", cashier)
			db.Insert("propmanagement.cashierpayment", cashier)
			CashierPK = db.LastPK
		Catch ex As Exception
			db.isError = True
		End Try
	End Sub

	Private Sub BarButtonItem1_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarButtonItem1.ItemClick
		Try
			SplashScreenManager.ShowDefaultWaitForm()
		Catch ex As Exception
		End Try
		Try
			loadData()
			TransGuid = Guid.NewGuid().ToString()
			Dim View = GridView1
			View.ClearColumnsFilter()
			If Validation() = False Then Exit Sub
			If ComboBox1.Text.ToLower.Contains("buyers") OrElse ComboBox1.Text.ToLower.Contains("buyer") Then
				Using db As New UCommand
					SAVE_CLIENT(db)
					db.SaveChanges()
				End Using
			ElseIf ComboBox1.Text.ToLower.Contains("payment") Then
				Using db As New UCommand
					SAVE_CHARGES(db)
					db.QueryExecNonQuery("update propmanagement.cashierpayment set PmtStatus='Cancelled' WHERE BankStatus  LIKE '%DAIF%';")
					db.QueryExecNonQuery("UPDATE propmanagement.cashierpayment SET PmtStatus='Cancelled' WHERE BankStatus  LIKE '%DAUD%';")
					db.QueryExecNonQuery("update propmanagement.cashierpayment set PmtStatus='Cancelled' WHERE BankStatus  LIKE '%W/%';")
					db.QueryExecNonQuery("update propmanagement.cashierpayment set PmtStatus='Cancelled' WHERE BankStatus  LIKE '%PULL%';")
					db.QueryExecNonQuery("UPDATE propmanagement.cashierpayment SET BankStatus='Un-Cleared' WHERE BankStatus  LIKE '%UNCLEARED%';")
					db.QueryExecNonQuery("UPDATE propmanagement.cashierpayment SET BankStatus='Un-Cleared' WHERE BankStatus  LIKE '%Unclear%';")
					db.QueryExecNonQuery("UPDATE propmanagement.cashierpayment SET BankStatus='Cleared' WHERE BankStatus  LIKE 'Clear';")
					db.QueryExecNonQuery("UPDATE propmanagement.cashierpayment SET PmtStatus='Active' WHERE BankStatus  LIKE 'Cleared';")
					db.QueryExecNonQuery("UPDATE propmanagement.paymentschedmain SET SubTotal=original_sell_price-discount1_amount-discount2_amount;")
					db.SaveChanges()
				End Using
			Else
				MessageBox.Show("Template not Fetch")
			End If
		Catch ex As Exception
		End Try
		Try
			SplashScreenManager.CloseDefaultWaitForm()
		Catch ex As Exception
		End Try
	End Sub
	Function SecondValidationCheck() As Boolean
		Try
			'Field validations
			Dim FieldName As String() = New String() {"Unit Code *", "Buyer Name *", "Payment Scheme",
				"Charge Name *", "Particulars *", "Due Date *", "Due Amount *", "MOP *",
				"AR Date *", "AR # *", "Bank", "Check Date ",
				"Check #", "Amount", "Account Ref ", "OR Date", "OR #", "CR Date",
				"CR #", "Amount2", "Payments Status", "Check Status", "Posting Date"}

			Dim ChargeName As String() = New String() {"EQUITY", "RESERVATION", "LOAN TAKE OUT", "MOVE-IN", "OTHERS", "RPT",
				 "TRANSFER CHARGE", "FIT OUT", "SETUP FUND"}

			Dim MOP As String() = New String() {"CASH", "CHECK", "CREDIT CARD", "ONLINE"}

			Dim ExcelCols As New List(Of String)
			ExcelCols.Clear()
			For i = 0 To GridView1.Columns.Count - 1
				ExcelCols.Add(GridView1.Columns(i).FieldName)
			Next
			If ComboBox1.Text.Trim = "'Payment Schedule $'" Then
				MessageBox.Show("Please rename Sheet Name to its corresponding Project and Floors Name." & vbNewLine & "Needed a unique template name" & vbNewLine & "Example: Payment Schedule P05-05F")
				Return False
			End If
			For Each i In FieldName
				Dim f = ExcelCols.Where(Function(x) x = i)
				If f.Count = 0 Then
					MessageBoxStr(String.Format("Column Name {0} is  not found.Please check excels Columns", i))
					MessageBoxStr("All Excel column Name is case sensitive.")
					Dim c As New ValidFieldName
					c.ShowDialog(Me)
					Return False
				End If
			Next

			'Check existing Charge Names   

			Dim excelChargeName As New List(Of String)
			Dim EXcelMOP As New List(Of String)
			For i = 1 To GridView1.RowCount - 1
				excelChargeName.Add(New String(GridView1.GetRowCellValue(i, "Charge Name *").ToString))
				EXcelMOP.Add(New String(GridView1.GetRowCellValue(i, "MOP *").ToString))
			Next

			'cHECK cHARGENAMES
			For Each i In excelChargeName
				If i.Trim = "" Then
					MessageBoxStr("Charge Name canot be empty.")
					Return False
				Else
					Dim f = ChargeName.Where(Function(x) x.ToUpper.Trim = i.ToUpper.Trim)
					If f.Count = 0 Then
						MessageBoxStr(String.Format("Charge Name [ {0} ] is not valid.", i))
						Dim valid = New ValidFieldName
						valid.UseDefault = False
						valid.Text = "Valid Charge Name"
						valid.MemoEdit1.Text = ""
						For Each ii In ChargeName
							valid.MemoEdit1.Text &= ii & vbNewLine
						Next
						valid.ShowDialog()
						Return False
					End If
				End If
			Next
			Dim validationcounter As Integer = 0
			'check MOps
			For Each i In EXcelMOP
				validationcounter += 1

				If i.Trim = "" Then
					Continue For
				End If
				Dim f = MOP.Where(Function(x) x.ToUpper = i.ToUpper)
				If f.Count = 0 Then
					MessageBoxStr(String.Format("Invalid MOP [ {0} ]", i & "   @line " & validationcounter))
					Dim valid = New ValidFieldName
					valid.UseDefault = False
					valid.Text = "Valid MOP's"
					valid.MemoEdit1.Text = ""
					For Each ii In MOP
						valid.MemoEdit1.Text &= ii & vbNewLine
					Next
					valid.ShowDialog()
					Return False
				End If
			Next
			Return True
		Catch ex As Exception
			MessageBoxStr(ex.Message)
			Return False
		End Try
	End Function
	'Function GetAcid(S As String)
	'	Return DB.QueryandReturnObject("SELECT  account_id FROM accounting.account_setup WHERE account_code='" & S.Trim.RSQ & "'")
	'End Function

	'Function GetAcccountName(S As String)
	'	Return DB.QueryandReturnObject("SELECT  account_description FROM accounting.account_setup WHERE account_code='" & S.Trim.RSQ & "'")
	'End Function
	Sub Remove(db As UCommand)
		Try
			If ComboBox1.Text.ToLower.Contains("buyers") OrElse ComboBox1.Text.ToLower.Contains("buyer") Then
				db.Delete("general.business_list", "Sheet='" & ComboBox1.Text.RSQ & "'")
				db.QueryExecNonQuery("DELETE FROM propmanagement.buyersinfomain WHERE Sheets='" & ComboBox1.Text.Trim.RSQ & "' and business_id NOT IN (SELECT business_number FROM general.business_list)")
				db.QueryExecNonQuery("DELETE FROM propmanagement.buyersinfoagents WHERE GUID NOT IN (SELECT GUID FROM propmanagement.buyersinfomain)")
				db.QueryExecNonQuery("DELETE FROM propmanagement.buyersinfoproject WHERE GUID NOT IN (SELECT GUID FROM propmanagement.buyersinfomain)")
				db.QueryExecNonQuery("DELETE FROM propmanagement.paymentschedmain WHERE BUYERGUID NOT IN (SELECT GUID FROM propmanagement.buyersinfomain)")
			ElseIf ComboBox1.Text.ToLower.Contains("payment") Then
				db.Delete("propmanagement.allcharges", "Sheet='" & ComboBox1.Text.Trim.RSQ & "'")
				db.Delete("propmanagement.cashierpayment", "AllChargeGUID not in (select GUID from propmanagement.allcharges)")
				db.Delete("propmanagement.cashierpayment_bk", "GUID NOT IN (SELECT guid FROM propmanagement.cashierpayment)")
				db.Delete("propmanagement.appliedpayment", "CashierGUID not in (select GUID from propmanagement.cashierpayment)")
				db.Delete("propmanagement.tbl_payment_details", "GUID not in (select GUID from propmanagement.allcharges)")
				db.Delete("propmanagement.tbl_othercharges_schedule", "GUID not in (select GUID from propmanagement.allcharges)")
				db.Delete("propmanagement.reservation", "GUID not in (select GUID from propmanagement.allcharges)")
			End If
		Catch ex As Exception
		End Try
	End Sub
	Private Sub SimpleButton2_Click(sender As Object, e As EventArgs) Handles SimpleButton2.Click
		Try
			If Not MessageContinueRemove() Then Exit Sub
			Using db As New UCommand
				Remove(db)
				db.SaveChanges()
			End Using
		Catch ex As Exception
		End Try
	End Sub

	Sub loadData()
		Try
			Units.Clear()
			Projects.Clear()
			Clients.Clear()
			Buyers.Clear()
			paymentschedmainRecords.Clear()

			Using Command As New UCommand
				'//LOAD UNITS
				Dim unitdata = Command.Datasource($"SELECT  UNITID,PRJID,TITLE_ID 'COMPANYID',UNITNO,UNITNAME  ,PROJECT_NO 'PROJECTNO'  FROM propmanagement.tbl_property_unit INNER JOIN general.setup_project  ON PROJECT_ID=PRJID")
				Dim projectData = Command.Datasource($" SELECT  project_id,title_id,Main_office_id,project,project_no,project_name ,STATUS FROM general.setup_project")
				Dim CLIENTDATA = Command.Datasource($"SELECT IFNULL(business_number,0) 'ID',business_fullname ,Type ,Sheet FROM general.`business_list` WHERE `TYPE`='Client'")
				Dim BUYERSDATA = Command.Datasource($"SELECT ID,business_id,BUYERFULLNAME FROM propmanagement.buyersinfomain")
				Dim PAYMENTSCHEMDATA = Command.Datasource($"select ID,GUID,UNITID,PRJID,CLIENTID,BUYERGUID from propmanagement.paymentschedmain")

				If unitdata.Rows.Count > 0 Then
					For Each i In unitdata.Rows
						Dim item As New UnitModel()
						item.UNITID = i("UNITID")
						item.PROJECTID = i("PRJID")
						item.COMPANYID = i("COMPANYID")
						item.UNITNO = i("UNITNO")
						item.UNITNAME = i("UNITNAME")
						item.PROJECTNO = i("PROJECTNO")
						Units.Add(item)
					Next
				End If
				If projectData.Rows.Count > 0 Then
					For Each i In projectData.Rows
						Dim item As New ProjectModel()
						item.PROJECTID = i("project_id")
						item.COMPANYID = i("title_id")
						item.PROJMAINID = i("Main_office_id")
						item.PROJTYPE = i("project")
						item.PROJNO = i("project_no")
						item.PROJNAME = i("project_name")
						item.STATUS = i("STATUS")
						Projects.Add(item)
					Next
				End If
				If CLIENTDATA.Rows.Count > 0 Then
					For Each i In CLIENTDATA.Rows
						Dim item As New BusinessModel()
						item.ID = i("ID")
						item.Name = i("business_fullname")
						item.Sheet = i("Sheet")
						item.Type = i("Type")
						Clients.Add(item)
					Next
				End If
				If BUYERSDATA.Rows.Count > 0 Then
					For Each i In BUYERSDATA.Rows
						Dim item As New BuyersModel()
						item.ID = i("ID")
						item.CLIENTID = i("business_id")
						item.BUYERFULLNAME = i("BUYERFULLNAME")
						Buyers.Add(item)
					Next
				End If
				If PAYMENTSCHEMDATA.Rows.Count > 0 Then
					For Each i In PAYMENTSCHEMDATA.Rows
						Dim item As New PaymentScheduleMainModel()
						item.ID = i("ID")
						item.GUID = i("GUID")
						item.UNITID = i("UNITID")
						item.PRJID = i("PRJID")
						item.CLIENTID = i("CLIENTID")
						item.BUYERGUID = i("BUYERGUID")
						item.CLIENTID = i("CLIENTID")
						paymentschedmainRecords.Add(item)
					Next
				End If


			End Using
		Catch ex As Exception
			MessageBoxError2(ex)
		End Try
	End Sub
	Private Sub Client_Master_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		Try
			KeyPreview = True
			DB = New Basic.MYSQL()
			DB.CONNECT("MigrationTools")
			Config.SET_CONNECTION(DB, "MigrationTools", "Accounting")
			If DB.IsConnected Then
				LoginFrms = New Configurations.LOGINFRM With {.DefaulCompany = My.Settings.DefaultCompany}
				LoginFrms.DefaulCompany = My.Settings.DefaultCompany
				LoginFrms.ShowDialog(Me)
			Else
				MessageXtra("Unable to connect to the server." & vbNewLine & "Please check your network connection.")
				LoginFrms = New Configurations.LOGINFRM With {.DefaulCompany = My.Settings.DefaultCompany}
				LoginFrms.DefaulCompany = My.Settings.DefaultCompany
				LoginFrms.ShowDialog(Me)
			End If
		Catch ex As Exception
			MessageBoxStr(ex.Message)
		End Try
	End Sub
	Private Sub SimpleButton3_Click(sender As Object, e As EventArgs) Handles SimpleButton3.Click
		Try
			If Not MessageContinueUpdate() Then Exit Sub
			Using db As New UCommand
				db.Datasource("call propmanagement.ChargeNameUpdater()")
				db.SaveChanges()
			End Using
		Catch ex As Exception
		End Try
	End Sub
	Private Sub SimpleButton4_Click(sender As Object, e As EventArgs) Handles SimpleButton4.Click
        Try
            Dim c As New Check_List
            c.Show(Me)
        Catch ex As Exception
        End Try
    End Sub
    Private Sub BarButtonItem2_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarButtonItem2.ItemClick
        ExportGrid(GridView1, Me.Text)
    End Sub

	Private Sub LoginFrms_CloseME(_exit As Boolean) Handles LoginFrms.CloseME
        Try
            If _exit Then
                Me.Close()

            End If
        Catch ex As Exception
        End Try
    End Sub
	Private Sub Client_Master_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
		Try
			If e.KeyCode = Keys.F2 Then
				LoginFrms = New Configurations.LOGINFRM With {.DefaulCompany = My.Settings.DefaultCompany}
				LoginFrms.ShowDialog(Me)
			End If
		Catch ex As Exception
		End Try
	End Sub

	Private Sub SimpleButton5_Click(sender As Object, e As EventArgs) Handles SimpleButton5.Click
		ComboBox1_SelectedIndexChanged(Nothing, Nothing)
	End Sub

	Private Sub LoginFrms_Success(UserInfo As Basic.UserClass, _DB As Basic.MYSQL) Handles LoginFrms.Success
		Try
			loadData()
			My.Settings.DefaultCompany = UserInfo.CompanyID
			My.Settings.Save()
			LoginFrms.Hide()
		Catch ex As Exception
		End Try
	End Sub
End Class
