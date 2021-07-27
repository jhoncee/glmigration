Imports System.ComponentModel
Public Class cashDepositCls
    Implements INotifyPropertyChanged
    Property check As Boolean
    Property remarks As String
    Property CUSTOMERID As Integer
    Property ID As Integer
    Property Amount As Decimal
    Property ClientName As String
    Property CheckStat As String
    Property GUID As String
    Property ChargeName As String
    Property descriptions_ As String
    Property UNITID As Integer
    Property PROJECTID As Integer

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Private Sub cashDepositCls_PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Handles Me.PropertyChanged

    End Sub
    Protected Sub OnPropertyChanged(propertyName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    End Sub
    Sub New(check_ As Boolean, remarks_ As String, custID_ As Integer, id_ As Integer, amount_ As Decimal, clientN_ As String, chkStat_ As String, guid_ As String, charge_ As String, desc_ As String, unitid_ As Integer, prjid_ As Integer)
        Me.check = check_
        Me.remarks = remarks_
        Me.CUSTOMERID = custID_
        Me.ID = id_
        Me.Amount = amount_
        Me.ClientName = clientN_
        Me.CheckStat = chkStat_
        Me.GUID = guid_
        Me.ChargeName = charge_
        Me.descriptions_ = desc_
        Me.UNITID = unitid_
        Me.PROJECTID = prjid_
    End Sub
End Class
