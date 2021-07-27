Module Struct
    Structure Autosigning
        Private XX As Object
        Friend Shared PreparedByIMAGE As Image
        Friend Shared ApprovedbyIMAGE As Image
        Friend Shared NotedByIMAGE As Image
        Friend Shared checkedbyIMAGE As Image
    End Structure
    Structure FORMTYPE
        Friend CV As String
        Friend DV As String
        Friend AP As String
        Friend FS As String
        Friend CASH_V As String
        Friend JE As String
        Friend LIQUIDATION As String
        Friend LEDGER As String
        Friend BILLING As String
        Sub INIT()
            CV = "CHECK VOUCHER"
            DV = "DISBURSEMENT VOUCHER"
            AP = "AP VOUCHER"
            FS = "FINANCIAL STATEMENT"
            CASH_V = "CASH VOUCHER"
            JE = "JOURNAL ENTRY VOUCHER"
            LEDGER = "LEDGERS"
            BILLING = "BILLING"
        End Sub
    End Structure
    Structure Signatorytype
        Friend PreparedBy As String
        Friend Approvedby As String
        Friend NotedBy As String
        Friend checkedby As String
        Sub init()
            PreparedBy = "PREPARED BY"
            Approvedby = "APPROVED BY"
            NotedBy = "NOTED BY"
            checkedby = "CHECKED BY"
        End Sub
    End Structure
    Structure Position
        Friend GeneralManager As String
        Friend OIC As String
        Friend Dept_manager As String
        Friend Reassign_Acct_Analyst As String
        Friend Supervising_prop_sup_officer As String
        Sub Init()
            GeneralManager = "General Manager"
            OIC = "OIC - Finance Division"
            Dept_manager = "Administrative  Department Manager"
            Reassign_Acct_Analyst = "Reassign ,Sr. Corp. Acct. Analyst"
            Supervising_prop_sup_officer = "Supervising Property Supply Officer"
        End Sub
    End Structure
End Module
