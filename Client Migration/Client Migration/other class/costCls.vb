Public Class costCls
    Property NO As Integer
    Property COMPANYCODE As String
    Property COMPANYNAME As String
    Property ADDRESS As String
    Property CONTACT As String
    Property TIN As String
    Property ZIPCODE As String
    Sub New(no_ As Integer, companyCode_ As String, companyN As String, address_ As String, contact_ As String, tin_ As String, zip_ As String)
        Try
            Me.NO = no_
            Me.COMPANYCODE = companyCode_
            Me.COMPANYNAME = companyN
            Me.ADDRESS = address_
            Me.CONTACT = contact_
            Me.TIN = tin_
            Me.ZIPCODE = zip_
        Catch ex As Exception
            MessageBoxError2(ex)
        End Try
    End Sub
End Class
Public Class branchCls
    Property NO As Integer
    Property COMPANYCODE As String
    Property COMPANYNAME As String
    Property BRANCHCODE As String
    Property BRANCHNAME As String
    Property ADDRESS As String
    Property CONTACTNUMBER As String
    Property BRANCHTYPE As String
    Sub New(no_ As Integer, companyCode_ As String, companyN As String, branchCode_ As String, branchName_ As String, address_ As String, contactNumber_ As String, branchType_ As String)
        Try
            Me.NO = no_
            Me.COMPANYCODE = companyCode_
            Me.COMPANYNAME = companyN
            Me.BRANCHCODE = branchCode_
            Me.BRANCHNAME = branchName_
            Me.ADDRESS = address_
            Me.CONTACTNUMBER = contactNumber_
            Me.BRANCHTYPE = branchType_
        Catch ex As Exception
            MessageBoxError2(ex)
        End Try
    End Sub
End Class
Public Class departmentCls
    Property NO As Integer
    Property COMPANYCODE As String
    Property COMPANYNAME As String
    Property BRANCHCODE As String
    Property BRANCHNAME As String
    Property DEPARTMENTCODE As String
    Property DEPARTMENTNAME As String
    Property ADDRESS As String
    Sub New(no_ As Integer, companyCode_ As String, companyName_ As String, branchCode_ As String, branchN As String, deptCode As String, deptName As String, address_ As String)
        Me.NO = no_
        Me.COMPANYCODE = companyCode_
        Me.COMPANYNAME = companyName_
        Me.BRANCHCODE = branchCode_
        Me.BRANCHNAME = branchN
        Me.DEPARTMENTCODE = deptCode
        Me.DEPARTMENTNAME = deptName
        Me.ADDRESS = address_
    End Sub
End Class
