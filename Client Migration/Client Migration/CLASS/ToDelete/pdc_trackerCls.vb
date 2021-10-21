Imports System.ComponentModel
Public Class pdc_trackerCls
    Implements INotifyPropertyChanged
    Property Bank As String
    Property CheckNumber As String
    Property location As String
    Property ID As Integer
    Property CUSTOMERID As Integer
    Property clientN As String
    Property deposit_slip_number As String
    Property CheckStatus As String
    Property Amount As Decimal
    Property AppliedDate As Date
    Property GUID As String
    Property ORNum As String
    Property CRNum As String
    Property ARNum As String
    Property PRNum As String
    Property ORDate As String
    Property CRDate As String
    Property ARDate As String
    Property PRDate As String
    Property CheckDate As Date
    Property deposit_date As String
    Property UNITID As Integer
    Property PROJECT As Integer
    Property reason As String

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Private Sub pdc_trackerCls_PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Handles Me.PropertyChanged

    End Sub
    Protected Sub OnPropertyChanged(propertyName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    End Sub

    Sub New(bank_ As String, chkDate_ As Date, chkNum_ As String, loc_ As String, id_ As Integer, custID As Integer, clientName As String,
            depo_num As String, depo_Date As String, chkStat As String, amnt As Decimal, appliedDate_ As Date, guid_ As String, orNum_ As String,
            crNum_ As String, ArNum_ As String, prNum_ As String, orDate_ As String, crDate_ As String, ardate_ As String, prDate_ As String, unitID_ As Integer, prjid_ As Integer, reason_ As String)

        Try
            Me.Bank = bank_
            Me.CheckDate = chkDate_
            Me.CheckNumber = chkNum_
            Me.location = loc_
            Me.ID = id_
            Me.CUSTOMERID = custID
            Me.clientN = clientName
            Me.deposit_slip_number = depo_num
            Me.deposit_date = depo_Date
            Me.CheckStatus = chkStat
            Me.Amount = amnt
            Me.AppliedDate = appliedDate_
            Me.GUID = guid_
            Me.ORNum = orNum_
            Me.CRNum = crNum_
            Me.ARNum = ArNum_
            Me.PRNum = prNum_
            Me.ORDate = orDate_
            Me.CRDate = crDate_
            Me.ARDate = ardate_
            Me.PRDate = prDate_
            Me.UNITID = unitID_
            Me.PROJECT = prjid_
            Me.reason = reason_
        Catch ex As Exception
            MessageBoxError2(ex)
        End Try
    End Sub

End Class
