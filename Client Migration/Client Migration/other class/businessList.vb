Imports System.ComponentModel
Public Class businessList
    Implements INotifyPropertyChanged

    Property No As Integer
    Property CODE As String
    Property FULLNAME As String
    Property FIRSTNAME As String
    Property MIDDLENAME As String
    Property LASTNAME As String
    Property SUFFIX As String
    Property TINNUMBER As String
    Property PAYMENTTERMS As String
    Property ISVAT As String
    Property CREDITLIMIT As String
    Property ATC As String
    Property CATEGORY As String
    Property CLASSIFICATION As String
    Property INHOUSE As String
    Property MANAGERCODE As String
    Property BROKERCODE As String
    Property ACCREDITED As String
    Property LICENSE As String
    Property HLURB As String
    Property EMAIL As String
    Property TEAM As String
    Property SECNUMBER As String
    Property _2303_ As String
    Property CONTRACTTYPE As String
    Property UNITNO As String
    Property CITIZENSHIP As String
    Property PROJECTNAME As String
    Property CIVILSTATUS As String

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Private Sub businessList_PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Handles Me.PropertyChanged

    End Sub
    Protected Sub OnPropertyChanged(propertyName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    End Sub
    Sub New(no_ As Integer, code_ As String, fullName As String, fn_ As String, mn_ As String, ln_ As String, suffix_ As String, termType_ As String,
            vat_ As String, credit_ As String, atc_ As String, category_ As String, clas_ As String, inhouse_ As String,
            managareCode_ As String, brokerCode_ As String, accredited_ As String, license_ As String, hlurb_ As String,
            email_ As String, team_ As String, tinNum As String, sec_ As String, _2303__ As String, contType As String, unit_ As String, citizen_ As String, projectN As String, civilStat As String)
        Try
            Me.No = no_
            Me.CODE = code_
            Me.FULLNAME = fullName
            Me.FIRSTNAME = fn_
            Me.MIDDLENAME = mn_
            Me.LASTNAME = ln_
            Me.SUFFIX = suffix_
            Me.TINNUMBER = tinNum
            Me.PAYMENTTERMS = termType_
            Me.ISVAT = vat_
            Me.CREDITLIMIT = credit_
            Me.ATC = atc_
            Me.CATEGORY = category_
            Me.CLASSIFICATION = clas_
            Me.INHOUSE = inhouse_
            Me.MANAGERCODE = managareCode_
            Me.BROKERCODE = brokerCode_
            Me.ACCREDITED = accredited_
            Me.LICENSE = license_
            Me.HLURB = hlurb_
            Me.EMAIL = email_
            Me.TEAM = team_
            Me.SECNUMBER = sec_
            Me._2303_ = _2303__
            Me.UNITNO = unit_
            Me.CONTRACTTYPE = contType
            Me.CITIZENSHIP = citizen_
            Me.PROJECTNAME = projectN
            Me.CIVILSTATUS = civilStat
        Catch ex As Exception
            MessageBoxError2(ex)
        End Try
    End Sub
End Class
Public Class businessList1
    Implements INotifyPropertyChanged

    Property No As Integer
    Property CODE As String
    Property FULLNAME As String
    Property FIRSTNAME As String
    Property MIDDLENAME As String
    Property LASTNAME As String
    Property SUFFIX As String
    Property TINNUMBER As String
    Property PAYMENTTERMS As String
    Property ISVAT As String
    Property CREDITLIMIT As String
    Property ATC As String
    Property CATEGORY As String
    Property CLASSIFICATION As String
    Property INHOUSE As String
    Property MANAGERCODE As String
    Property BROKERCODE As String
    Property ACCREDITED As String
    Property LICENSE As String
    Property HLURB As String
    Property EMAIL As String
    Property TEAM As String
    Property SECNUMBER As String
    Property _2303_ As String
    Property CONTRACTTYPE As String
    Property UNITNO As String
    Property CITIZENSHIP As String
    Property PROJECTNAME As String
    Property CIVILSTATUS As String
    Property UNITPRICE As Decimal
    Property VATAMNT As Decimal
    Property CONTRACTPRICE As Decimal

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Private Sub businessList_PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Handles Me.PropertyChanged

    End Sub
    Protected Sub OnPropertyChanged(propertyName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    End Sub
    Sub New(no_ As Integer, code_ As String, fullName As String, fn_ As String, mn_ As String, ln_ As String, suffix_ As String, termType_ As String,
            vat_ As String, credit_ As String, atc_ As String, category_ As String, clas_ As String, inhouse_ As String,
            managareCode_ As String, brokerCode_ As String, accredited_ As String, license_ As String, hlurb_ As String,
            email_ As String, team_ As String, tinNum As String, sec_ As String, _2303__ As String, contType As String, unit_ As String, citizen_ As String, projectN As String, civilStat As String,
            unitP As Decimal, vatAmnt_ As Decimal, totalContractP As Decimal)
        Try
            Me.No = no_
            Me.CODE = code_
            Me.FULLNAME = fullName
            Me.FIRSTNAME = fn_
            Me.MIDDLENAME = mn_
            Me.LASTNAME = ln_
            Me.SUFFIX = suffix_
            Me.TINNUMBER = tinNum
            Me.PAYMENTTERMS = termType_
            Me.ISVAT = vat_
            Me.CREDITLIMIT = credit_
            Me.ATC = atc_
            Me.CATEGORY = category_
            Me.CLASSIFICATION = clas_
            Me.INHOUSE = inhouse_
            Me.MANAGERCODE = managareCode_
            Me.BROKERCODE = brokerCode_
            Me.ACCREDITED = accredited_
            Me.LICENSE = license_
            Me.HLURB = hlurb_
            Me.EMAIL = email_
            Me.TEAM = team_
            Me.SECNUMBER = sec_
            Me._2303_ = _2303__
            Me.UNITNO = unit_
            Me.CONTRACTTYPE = contType
            Me.CITIZENSHIP = citizen_
            Me.PROJECTNAME = projectN
            Me.CIVILSTATUS = civilStat
            Me.UNITPRICE = unitP
            Me.VATAMNT = vatAmnt_
            Me.CONTRACTPRICE = totalContractP
        Catch ex As Exception
            MessageBoxError2(ex)
        End Try
    End Sub
End Class

