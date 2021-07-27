Public Interface IACTIONS
    Inherits Basic.IAction
End Interface
Public Interface IVIEWALL_INFO
    Sub EXECUTE_COMMAND(ByVal DATA As PARAMS)
End Interface
'Public Interface IRELOADS
'    Sub ReloadBusiness(ByVal e As CLSENUMS.BusinessPerson, ByVal setupSuppliername As String)
'End Interface
Public Interface iSHOW
    Enum SignatoryForm As Short
        CV
        DV
        JE
        EXEC_REPORT
        FS_REPORT
    End Enum
    Sub VIEW(ByVal VIEWID As SignatoryForm)
End Interface


