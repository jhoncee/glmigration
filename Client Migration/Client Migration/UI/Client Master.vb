Imports DevExpress.XtraSplashScreen
Partial Public Class Client_Master
	Private ImportEx As ImportFromExcel
	Private FileName As String = ""
	Private TransGuid As String
	Private ChargeID As Integer
	Private Tdate As Date
	Private CashierPK As Integer = 0
	Public WithEvents LoginFrms As Configurations.LOGINFRM
	Public WithEvents NewConnectionFrm As Configurations.Connection
	Private Config As New Configurations.DBconfigurations

	Private Units As New List(Of UnitModel)
	'Private Projects As New List(Of ProjectModel)
	Private Businesses As New List(Of BusinessModel)
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
	Function GET_BUSINESS(ByVal EntityName As String) As BusinessModel
		Try
			Dim Client = Businesses.Where(Function(x) x.Name.Trim = EntityName.Trim And x.Type = "CLIENT").LastOrDefault()
			Return Client
			'Dim x = DB.QueryandReturnTable("SELECT IFNULL(business_number,0) 'ID' FROM general.`business_list` WHERE business_fullname = '" & ClientName.Trim & "' AND `TYPE`='Client'")
			'If x.Rows.Count > 0 Then
			'	Return x(0)(0)
			'End If
		Catch ex As Exception
			Return Nothing
		End Try
	End Function
	Sub SAVE_BUSINESS(db As UCommand, NewClient As BusinessModel)
		Try
			Dim DIC As New Dictionary(Of String, String)
			DIC.Add("business_fullname", NewClient.Name.Trim.RSQ)
			DIC.Add("`TYPE`", NewClient.Type)
			DIC.Add("`Sheet`", ComboBox1.Text.RSQ)
			db.Insert("general.business_list", DIC)
			NewClient.ID = db.LastPK
			Businesses.Add(NewClient)
		Catch ex As Exception
		End Try
	End Sub

	Sub SAVE_CLIENT(db As UCommand)
		Try
			Dim view = GridView1
			Dim DIC As New Dictionary(Of String, String)

			For I As Integer = 0 To GridView1.RowCount - 1
				TransGuid = Guid.NewGuid().ToString()
				Dim clientName As String = If(String.IsNullOrWhiteSpace(view.GetRowCellValue(I, "Buyer Name *").ToString), "", view.GetRowCellValue(I, "Buyer Name *").ToString.Trim)
				Dim unitCode = If(String.IsNullOrWhiteSpace(view.GetRowCellValue(I, "Unit Code *").ToString.Trim), "", view.GetRowCellValue(I, "Unit Code *").ToString.Trim)
				'If view.GetRowCellValue(I, "Unit Code *").ToString.Trim = "P05-ET2-02" Then
				'	Dim tracing = 0
				'End If

				If unitCode = "" Then Continue For
				If clientName.ToLower = "unsold" OrElse clientName.ToLower = "not used" Or clientName.ToLower = "un-sold" OrElse clientName.ToLower = "notused" OrElse clientName = "" Then
					db.QueryExecNonQuery(" UPDATE propmanagement.tbl_property_unit SET unit_status='AVAILABLE',CurrentBuyerID='0' WHERE UNITNO='" & unitCode.Trim.RSQ & "'")
					Continue For
				End If

				Dim CLIENT = GET_BUSINESS(view.GetRowCellValue(I, "Buyer Name *").ToString)
				If CLIENT?.ID > 0 Then
				Else
					CLIENT = New BusinessModel
					CLIENT.Name = view.GetRowCellValue(I, "Buyer Name *").ToString
					CLIENT.Type = "CLIENT"
					CLIENT.Sheet = ComboBox1.Text
					SAVE_BUSINESS(db, CLIENT)
				End If

				'GET PROJECT AND UNIT INFO
				Dim Unit = GET_UNIT(unitCode)

				'SAVE BUYER INFO  main
				DIC = New Dictionary(Of String, String)
				DIC.Add("GUID", TransGuid)
				DIC.Add("BUYERFULLNAME", CLIENT.Name.RSQ)
				DIC.Add("buyer_firstname", CLIENT.Name.RSQ)
				DIC.Add("business_id", CLIENT.ID)
				DIC.Add("PropUnitId", Unit.UNITID)
				DIC.Add("citizenship", view.GetRowCellValue(I, "Citizenship").ToString.RSQ)
				' DIC.Add("Status", view.GetRowCellValue(I, "Buyer Status").ToString.RSQ)
				DIC.Add("Sheets", ComboBox1.Text.RSQ.Trim)
				db.Insert("propmanagement.buyersinfomain", DIC)
				'insert all business that is not in payee table..
				db.QueryExecNonQuery("INSERT INTO accounting.payee ( BID,PAYEE_NAME)  SELECT business_number,business_fullname FROM  general.business_list  WHERE business_number NOT IN (SELECT BID FROM  accounting.payee)")

				Dim agent = GetAgentInfo(db, view.GetRowCellValue(I, "Agent *").ToString.Trim)
				Dim manager = GetAgentInfo(db, view.GetRowCellValue(I, "Manager").ToString().Trim)
				Dim broker = GetAgentInfo(db, view.GetRowCellValue(I, "Broker").ToString().Trim)

				'SAVE AGENT 
				DIC = New Dictionary(Of String, String)
				DIC.Add("GUID", TransGuid)
				DIC.Add("AgentID", agent?.ID)
				DIC.Add("Agent", agent?.Name.RSQ)
				DIC.Add("ManagerID", manager?.ID)
				DIC.Add("Manager", manager?.Name.RSQ)
				DIC.Add("BrokerID", broker?.ID)
				DIC.Add("Broker", broker?.Name.RSQ)
				db.Insert("propmanagement.buyersinfoagents", DIC)

				'///////////////// 
				'SAVE buyer project 
				DIC = New Dictionary(Of String, String)
				DIC.Add("GUID", TransGuid)
				DIC.Add("unit_id", Unit.UNITID)
				DIC.Add("project_id", Unit.PROJECTID)
				Dim nsp As Double = 0
				Double.TryParse(view.GetRowCellValue(I, "Net Selling Price"), nsp)
				DIC.Add("total_contract_price", nsp)
				db.Insert("propmanagement.buyersinfoproject", DIC)
				'SAVE SCHEDULE
				SAVE_SCHEDULE(db, I, CLIENT, Unit)
				'update Unit
				db.QueryExecNonQuery(" UPDATE propmanagement.tbl_property_unit SET unit_status='SOLD OUT',CurrentBuyerID='" & CLIENT.ID & "' WHERE UNITID='" & Unit.UNITID & "'")
			Next
		Catch ex As Exception
			db.isError = True
			MessageBoxError2(ex)
		End Try
	End Sub
	'Function GetCompanyIDbyPrjID(ByVal PrjID As String) As Integer
	'	Try
	'		Dim proj = Projects.SingleOrDefault(Function(x) x.PROJECTID = PrjID)
	'		Return proj.COMPANYID
	'		'DB.QueryandReturnTable("SELECT title_id FROM general.`setup_project` WHERE project_id = '" & PrjID & "'")
	'		'If x.Rows.Count > 0 Then
	'		'	Return x(0)(0)
	'		'End If
	'	Catch ex As Exception
	'		Return 0
	'	End Try
	'End Function

	Function GET_UNIT(ByVal unitN As String) As UnitModel
		Try
			Dim UNIT = Units.LastOrDefault(Function(x) x.UNITNO.ToLower.Trim = unitN.ToLower.Trim)
			Return UNIT
		Catch ex As Exception
			Return Nothing
		End Try
	End Function

	'Function GET_PROJ_ID(ByVal proj_no As String) As Integer
	'	Try
	'		Dim prj = Projects.SingleOrDefault(Function(x) x.PROJNO.ToLower.Trim = proj_no.ToLower.Trim)
	'		Return prj?.PROJECTID
	'		'Dim xX As New DataTable
	'		'xX = db.Datasource("SELECT project_id 'ID' FROM general.`setup_project` WHERE project_no = '" & proj_no & "'")
	'		'If xX.Rows.Count > 0 Then
	'		'	Return xX(0)(0)
	'		'End If
	'	Catch ex As Exception
	'		Return 0
	'	End Try
	'End Function

	'AGENT
	Function GetAgentInfo(db As UCommand, AgentName As String) As BusinessModel
		Try
			If AgentName.Trim = "" Then Return New BusinessModel
			Dim agent = Businesses.Where(Function(x) x.Name.Trim = AgentName.Trim And x.Type.ToLower.Trim = "Agent").LastOrDefault
			If agent IsNot Nothing Then
				Return agent
			Else
				agent = New BusinessModel() With {.Name = AgentName.Trim, .Sheet = ComboBox1.Text, .Type = "Agent"}
				SAVE_BUSINESS(db, agent)
				Return agent
			End If
		Catch ex As Exception
			Return New BusinessModel
		End Try
	End Function
	'MANAGER
	Function GetManagerInfo(db As UCommand, ManagerName As String) As BusinessModel
		Try
			If ManagerName.Trim = "" Then Return New BusinessModel
			Dim manager = Businesses.Where(Function(x) x.Name.Trim = ManagerName.Trim And x.Type = "Manager").LastOrDefault()
			If manager IsNot Nothing Then
				Return manager
			Else
				manager = New BusinessModel() With {.Name = ManagerName.Trim, .Sheet = ComboBox1.Text, .Type = "Manager"}
				SAVE_BUSINESS(db, manager)
				Return manager
			End If

		Catch ex As Exception
			Return New BusinessModel
		End Try
	End Function
	'BROKER
	Function GetBrokerInfo(db As UCommand, BrokerName As String) As BusinessModel
		Try
			If BrokerName.Trim = "" Then Return New BusinessModel
			Dim broker = Businesses.Where(Function(x) x.Name.ToLower.Trim = BrokerName.ToLower.Trim And x.Type.ToLower = "BROKER").LastOrDefault()
			If broker IsNot Nothing Then
				Return broker
			Else
				broker = New BusinessModel() With {.Name = BrokerName.Trim, .Sheet = ComboBox1.Text, .Type = "BROKER"}
				SAVE_BUSINESS(db, broker)
				Return broker
			End If
		Catch ex As Exception
			Return New BusinessModel
		End Try
	End Function
	Sub SAVE_SCHEDULE(db As UCommand, i As Integer, Client As BusinessModel, Unit As UnitModel)
		Try
			Dim VIEW = GridView1

			Dim Discount1 As Double = 0, Discount2 As Double = 0
			Dim tc As Double = 0, tcp As Double = 0
			Dim MoveIn As Double = 0
			Dim LTO As Double = 0
			Dim addvat As Double = 0
			Dim lp As Double = 0
			Dim vat As Double = 0
			Dim gross = 0

			Dim DIC As New Dictionary(Of String, String)
			DIC.Add("PRJID", Unit.PROJECTID)
			DIC.Add("UNITID", Unit.UNITID)
			DIC.Add("GUID", TransGuid)
			DIC.Add("CLIENTID", Client.ID)
			DIC.Add("BUYERGUID", TransGuid)
			DIC.Add("trans_date", Tdate.ToMysqlFormat)
			DIC.Add("no_months", 0) ' VIEW.GetRowCellValue(i, "Deferred Months"))

			Double.TryParse(VIEW.GetRowCellValue(i, "List Price *"), lp)
			Double.TryParse(VIEW.GetRowCellValue(i, "Vat"), vat)
			Double.TryParse(VIEW.GetRowCellValue(i, "Gross"), gross)
			Double.TryParse(VIEW.GetRowCellValue(i, "Discount 1"), Discount1)
			Double.TryParse(VIEW.GetRowCellValue(i, "Discount 2"), Discount2)
			Double.TryParse(VIEW.GetRowCellValue(i, "Transfer Charge"), tc)
			Double.TryParse(VIEW.GetRowCellValue(i, "TCP"), tcp)
			Double.TryParse(VIEW.GetRowCellValue(i, "Move-In"), MoveIn)
			Double.TryParse(VIEW.GetRowCellValue(i, "Add Vat"), addvat)
			Double.TryParse(VIEW.GetRowCellValue(i, "Loan Takeout"), LTO)

			DIC.Add("original_sell_price", lp) 'list price 
			If vat > 0 Then
				DIC.Add("IsVat", If(vat > 0, "V", "NV"))
				DIC.Add("vat_percentage", 12)
				DIC.Add("vat_amount", vat)  'Vat 1
			End If
			DIC.Add("total_selling_price", gross) 'Gross
			DIC.Add("discount1_amount", Discount1) 'Discounts
			DIC.Add("discount2_amount", Discount2) 'Discounts
			DIC.Add("SubTotal", lp - Discount1 - Discount2) 'Net Price/Sub Total
			DIC.Add("net_vat_amount", addvat) 'add vat >0
			DIC.Add("total_net_selling_price", tcp) 'TCP
			DIC.Add("transfer_charges_percentage", CDbl(FormatNumber(tc / tcp * 100, 2)))
			DIC.Add("transfer_amount", tc)
			DIC.Add("movein_amount", MoveIn)
			Try
				DIC.Add("move_in_fees_percentage", CDbl(FormatNumber(MoveIn / tcp * 100, 2)))
			Catch ex As Exception
			End Try
			DIC.Add("take_out_loan_amount", LTO)
			DIC.Add("trans_type", "EMI")
			'DIC.Add("equity_percentage",)
			'DIC.Add("equity_amount",)  
			db.Insert("propmanagement.paymentschedmain", DIC)
		Catch ex As Exception
			db.isError = True
			MessageBoxError2(ex)
		End Try
	End Sub
	Function GetSchedGUID(client As BusinessModel, unit As UnitModel) As String
		Try
			Dim item = paymentschedmainRecords.Where(Function(x) x.CLIENTID = client.ID And x.UNITID = unit.UNITID).LastOrDefault
			Return item?.GUID
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
	Function GET_BUYER(CLIENT As BusinessModel, Unit As UnitModel) As BuyersModel
		Try
			Dim buyer = Buyers.Where(Function(x) x.CLIENTID = CLIENT.ID And x.UNITID = Unit.UNITID).LastOrDefault
			Return buyer
			'Return DataObject("SELECT ID FROM propmanagement.buyersinfomain WHERE business_id='" & CLIENTID & "'")
		Catch ex As Exception
			Return Nothing
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

			For I As Integer = 0 To GridView1.RowCount - 1
				Dim view = GridView1
				Dim Unit = GET_UNIT(view.GetRowCellValue(I, "Unit Code *").ToString())
				If Unit Is Nothing Then
					GridView1.FocusedRowHandle = I
					MessageBoxStr($"Unit Code [{view.GetRowCellValue(I, "Unit Code *")}] is not Found. Line:" & I + 1)
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
				If ComboBox1.Text.Trim.ToString.ToLower.Contains("buyer") And View.GetRowCellValue(i, "Buyer Name *").ToString.Trim = "" Then
					MessageBoxStr("Invalid Client Name Line:" & i + 1)
					Return False
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

			Dim DIC As New Dictionary(Of String, String)
			cNTR = 0
			'SAVE CHARGES AND PAYMENTS

			For I As Integer = 0 To GridView1.RowCount - 1
				cNTR += 1
				CashierPK = 0
				Dim CLIENT = GET_BUSINESS(View.GetRowCellValue(I, "Buyer Name *").ToString)
				Dim Unit = GET_UNIT(View.GetRowCellValue(I, "Unit Code *").ToString())
				Dim BuyerInfo = GET_BUYER(CLIENT, Unit)
				TransGuid = GetSchedGUID(CLIENT, Unit)

				DIC = New Dictionary(Of String, String)
				DIC.Add("GUID", TransGuid)
				DIC.Add("BUYERID", BuyerInfo?.ID)
				DIC.Add("BID", CLIENT?.ID)
				DIC.Add("BuyerName", CLIENT?.Name.RSQ)
				DIC.Add("UNITID", Unit?.UNITID)
				DIC.Add("UNITNO", Unit?.UNITNO.RSQ)
				DIC.Add("PRJID", Unit?.PROJECTID)
				DIC.Add("TITLEID", Unit?.COMPANYID)
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

				'for PMT status 
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

				DIC.Add("PMTSTAT", GridView1.GetRowCellValue(I, "Payments Status").ToString)
				db.Insert("propmanagement.allcharges", DIC)
				ChargeID = db.LastPK
				SAVE_SCHED_DETAILS(db, I) 'save schedule 
				'//Update loan takeout amount
				If IsNumeric(GridView1.GetRowCellValue(I, "Due Amount *")) Then
					If GridView1.GetRowCellValue(I, "Charge Name *").ToString.ToLower.Contains("balance") Then
						db.QueryExecNonQuery("Update propmanagement.paymentschedmain set take_out_loan_amount='" & CDbl(GridView1.GetRowCellValue(I, "Due Amount *")) & "' where UNITID='" & Unit.UNITID & "' and CLIENTID='" & CLIENT.ID & "' ")
					ElseIf GridView1.GetRowCellValue(I, "Charge Name *").ToString.ToLower.Contains("loan") Then
						db.QueryExecNonQuery("Update propmanagement.paymentschedmain set take_out_loan_amount='" & CDbl(GridView1.GetRowCellValue(I, "Due Amount *")) & "' where UNITID='" & Unit.UNITID & "' and CLIENTID='" & CLIENT.ID & "' ")
					ElseIf GridView1.GetRowCellValue(I, "Charge Name *").ToString.ToLower.Contains("takeout") Then
						db.QueryExecNonQuery("Update propmanagement.paymentschedmain set take_out_loan_amount='" & CDbl(GridView1.GetRowCellValue(I, "Due Amount *")) & "' where UNITID='" & Unit.UNITID & "' and CLIENTID='" & CLIENT.ID & "' ")
					End If
				End If
				Save_to_cashier(db, I, CLIENT, Unit)
				SAVE_APPLIED_PAYMENT(db, I, CLIENT, Unit)
				SAVE_RESERVATION(db, I, CLIENT, Unit)
			Next
		Catch ex As Exception
			db.isError = True
		End Try
	End Sub
	Sub SAVE_RESERVATION(db As UCommand, I As Integer, Client As BusinessModel, Unit As UnitModel)
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
				db.QueryExecNonQuery("INSERT INTO  reservation SET GUID='" & TransGuid & "', AMT='" & Amt & "', UNITID='" & Unit.UNITID & "', UNITNO='" & Unit.UNITNO.RSQ & "', CLIENTID='" & Client.ID & "', CustName='" & Client.Name.RSQ & "', RSDate='" & CDate(View.GetRowCellValue(I, "Posting Date")).ToMysqlFormat & "'")
			End If
		Catch ex As Exception
		End Try
	End Sub
	Sub SAVE_APPLIED_PAYMENT(db As UCommand, I As Integer, Client As BusinessModel, Unit As UnitModel)
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
			Applied.Add("CMPID", Unit.COMPANYID)
			Applied.Add("PROJECTID", Unit.PROJECTID)
			Applied.Add("UNITID", Unit.UNITID)
			Applied.Add("UNITNO", Unit.UNITNO.RSQ)
			Applied.Add("CUSTOMERID", Client.ID)
			Applied.Add("Customer", Client.Name.RSQ)
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

	Sub Save_to_cashier(db As UCommand, I As Integer, Client As BusinessModel, Unit As UnitModel)
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

			Dim cashier As New Dictionary(Of String, String)
			With cashier
				.Add("PMTRREF", "")
				.Add("Description", GridView1.GetRowCellValue(I, "Particulars *").ToString.Trim.RSQ)
				.Add("UNITID", Unit.UNITID)
				.Add("UNITNO", VIEW.GetRowCellValue(I, "Unit Code *").ToString.Trim.RSQ)
				.Add("PROJECTID", Unit.PROJECTID)
				.Add("CUSTOMERID", Client.ID)
				.Add("Customer", Client.Name.RSQ)
				.Add("CMPID", Unit.COMPANYID)
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
			'db.Insert("propmanagement.cashierpayment_Bk", cashier)
			db.Insert("propmanagement.cashierpayment", cashier)
			CashierPK = db.LastPK
		Catch ex As Exception
		End Try
	End Sub

	Private Async Sub BarButtonItem1_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarButtonItem1.ItemClick
		Try
			SplashScreenManager.ShowDefaultWaitForm()
		Catch ex As Exception
		End Try
		Try
			TransGuid = Guid.NewGuid().ToString()
			Dim View = GridView1
			View.ClearColumnsFilter()
			If ComboBox1.Text.ToLower.Contains("buyers") OrElse ComboBox1.Text.ToLower.Contains("buyer") Then
				Using db As New UCommand
					Remove(db)
					Dim loaded = loadData(db)
					If Not loaded Then
						MessageBoxStr("Cannot continue, Data is not loaded")
						Exit Sub
					End If
					If Validation() = False Then Exit Sub
					SAVE_CLIENT(db)
					db.SaveChanges()
				End Using
			ElseIf ComboBox1.Text.ToLower.Contains("payment") Then
				Using db As New UCommand
					Remove(db)
					Dim loaded = loadData(db)
					If Not loaded Then
						MessageBoxStr("Cannot continue, Data is not loaded")
						Exit Sub
					End If
					If Validation() = False Then Exit Sub
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

			For I As Integer = 0 To GridView1.RowCount - 1
				Dim View = GridView1
				If View.GetRowCellValue(I, "Charge Name *").ToString.Trim = "" Then
					GridView1.FocusedRowHandle = I
					MessageBoxStr("Invalid Charge Name. line: " & I + 1)
					Return False
				End If
				If View.GetRowCellValue(I, "Particulars *").ToString.Trim = "" Then
					GridView1.FocusedRowHandle = I
					MessageBoxStr("Invalid Particulars. line: " & I + 1)
					Return False
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
						GridView1.FocusedRowHandle = I
						MessageBoxStr("Invalid amount @line:  " & I + 1)
						Return False
					End If
				End If
				If View.GetRowCellValue(I, "Charge Name *").ToString.Trim <> "" And View.GetRowCellValue(I, "Particulars *").ToString.Trim = "" Then
					GridView1.FocusedRowHandle = I
					MessageBoxStr("Invalid Particulars @line:  " & I + 1)
					Return False
				End If

				If String.IsNullOrWhiteSpace(GridView1.GetRowCellValue(I, "Payments Status")) Then
					GridView1.FocusedRowHandle = I
					MessageBoxStr("Invalid Payment Status @Line:  " & I + 1)
					Return False
				End If
				If View.GetRowCellValue(I, "MOP *").ToString.Trim <> "" Then
					If View.GetRowCellValue(I, "Check Status").ToString.Trim = "" Then
						GridView1.FocusedRowHandle = I
						MessageBoxStr("Invalid Check Status @Line:  " & I + 1)
						Return False
					End If
				End If

				Select Case View.GetRowCellValue(I, "MOP *").ToString.ToLower.Trim
					Case "check", "cash", "online", "credit card", ""
					Case Else
						MessageBoxStr($"MOP [{View.GetRowCellValue(I, "MOP *").ToString.Trim}]  @ row{I} is  not defined")
						Return False
				End Select

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
						MessageBoxStr("Invalid amount @line:  " & I + 1)
						Return False
					End If
				End If

				Dim Unit = GET_UNIT(View.GetRowCellValue(I, "Unit Code *").ToString.Trim)
				If Unit Is Nothing Then
					GridView1.FocusedRowHandle = I
					MessageBoxStr("Unit Code not Found. Line:" & I + 1)
					Return False
				End If

				Dim CLIENT = GET_BUSINESS(View.GetRowCellValue(I, "Buyer Name *").ToString)
				If CLIENT Is Nothing Then
					GridView1.FocusedRowHandle = I
					MessageBoxStr("Client Not Found. line: " & I + 1)
					Return False
				End If

				Dim Buyer = GET_BUYER(CLIENT, Unit)
				If Buyer Is Nothing Then
					GridView1.FocusedRowHandle = I
					MessageBoxStr("Cannot locate buyer for this Unit:  @row" & I + 1)
					Return False
				End If

				If GetSchedGUID(CLIENT, Unit) = "" Then
					GridView1.FocusedRowHandle = I
					MessageBoxStr("Cannot Locate Payment Schedule. line: " & I + 1)
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
	Function Remove(db As UCommand) As Boolean
		Try
			If ComboBox1.Text.ToLower.Contains("buyers") OrElse ComboBox1.Text.ToLower.Contains("buyer") Then
				db.Delete("general.business_list", "Sheet='" & ComboBox1.Text.RSQ & "'")
				db.QueryExecNonQuery("DELETE FROM propmanagement.buyersinfomain WHERE Sheets='" & ComboBox1.Text.Trim.RSQ & "' and NOT EXISTS(SELECT business_number FROM general.business_list WHERE BUSINESS_NUMBER=business_id)")
				db.QueryExecNonQuery("DELETE FROM propmanagement.buyersinfoagents WHERE NOT EXISTS (SELECT GUID FROM propmanagement.buyersinfomain WHERE buyersinfomain.GUID=buyersinfoagents.GUID)")
				db.QueryExecNonQuery("DELETE FROM propmanagement.buyersinfoproject WHERE NOT EXISTS (SELECT GUID FROM propmanagement.buyersinfomain  WHERE buyersinfomain.GUID=buyersinfoproject.GUID)")
				db.QueryExecNonQuery("DELETE FROM propmanagement.paymentschedmain WHERE NOT EXISTS (SELECT GUID FROM propmanagement.buyersinfomain  WHERE buyersinfomain.GUID=paymentschedmain.BUYERGUID)")
			ElseIf ComboBox1.Text.ToLower.Contains("payment") Then
				'db.Delete("propmanagement.cashierpayment_bk", "GUID NOT IN (SELECT guid FROM propmanagement.cashierpayment)")
				db.Delete("propmanagement.allcharges", "Sheet='" & ComboBox1.Text.Trim.RSQ & "'")
				db.Delete("propmanagement.cashierpayment", "NOT EXISTS(SELECT GUID FROM propmanagement.allcharges   WHERE allcharges.GUID=cashierpayment.AllChargeGUID)")
				db.Delete("propmanagement.appliedpayment", "NOT EXISTS  (SELECT GUID FROM propmanagement.cashierpayment WHERE cashierpayment.GUID=CashierGUID)")
				db.Delete("propmanagement.tbl_payment_details", "NOT EXISTS  (SELECT GUID FROM propmanagement.allcharges WHERE allcharges.GUID=tbl_payment_details.GUID)")
				db.Delete("propmanagement.tbl_othercharges_schedule", "NOT EXISTS  (SELECT GUID FROM propmanagement.allcharges WHERE allcharges.GUID=tbl_othercharges_schedule.GUID)")
				db.Delete("propmanagement.reservation", "NOT EXISTS  (SELECT GUID FROM propmanagement.allcharges WHERE allcharges.GUID=reservation.GUID)")
			End If
			Return True
		Catch ex As Exception
			Return False
		End Try
	End Function
	Private Async Sub SimpleButton2_Click(sender As Object, e As EventArgs) Handles SimpleButton2.Click
		Try
			If Not MessageContinueRemove() Then Exit Sub
			Using db As New UCommand
				Remove(db)
				db.SaveChanges()
			End Using
		Catch ex As Exception
		End Try
	End Sub
	Function loadData(Command As UCommand) As Boolean
		Try
			Units.Clear()
			Businesses.Clear()
			Buyers.Clear()
			paymentschedmainRecords.Clear()

			'//LOAD UNITS
			Dim unitdata = Command.Datasource($"SELECT  UNITID,PRJID,TITLE_ID 'COMPANYID',UNITNO,UNITNAME  ,PROJECT_NO 'PROJECTNO'  FROM propmanagement.tbl_property_unit INNER JOIN general.setup_project  ON PROJECT_ID=PRJID")
			Dim CLIENTDATA = Command.Datasource($"SELECT IFNULL(business_number,0) 'ID',business_fullname ,Type ,Sheet FROM general.`business_list` WHERE `Type`!='SUPPLIER'")
			Dim BUYERSDATA = Command.Datasource($"SELECT ID,business_id,BUYERFULLNAME,PropUnitId FROM propmanagement.buyersinfomain")
			Dim PAYMENTSCHEMDATA = Command.Datasource($"select ID,GUID,UNITID,PRJID,CLIENTID,BUYERGUID from propmanagement.paymentschedmain")

			'Dim t = Await Task.WhenAll(unitdataTask, CLIENTDATATask,
			'					 BUYERSDATATask, PAYMENTSCHEMDATATask)

			'Dim unitdata = Await unitdataTask
			'Dim CLIENTDATA = Await CLIENTDATATask
			'Dim BUYERSDATA = Await BUYERSDATATask
			'Dim PAYMENTSCHEMDATA = Await PAYMENTSCHEMDATATask

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

			If CLIENTDATA.Rows.Count > 0 Then
				For Each i In CLIENTDATA.Rows
					Dim item As New BusinessModel()
					item.ID = i("ID")
					item.Name = i("business_fullname")
					item.Sheet = i("Sheet")
					item.Type = i("Type")
					Businesses.Add(item)
				Next
			End If

			If BUYERSDATA.Rows.Count > 0 Then
				For Each i In BUYERSDATA.Rows
					Dim item As New BuyersModel()
					item.ID = i("ID")
					item.CLIENTID = i("business_id")
					item.UNITID = i("PropUnitId")
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
			Return True
		Catch ex As Exception
			MessageBoxError2(ex)
			Return False
		End Try
	End Function
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
			My.Settings.DefaultCompany = UserInfo.CompanyID
			My.Settings.Save()
			LoginFrms.Hide()
		Catch ex As Exception
		End Try
	End Sub
End Class
