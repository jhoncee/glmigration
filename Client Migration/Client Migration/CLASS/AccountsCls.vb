Imports System.Collections
Imports System.ComponentModel
Friend Class AccountsCls
    Implements INotifyPropertyChanged
    Property ACID As Integer
    Property SUBSIDIARY As Integer
    Property SUBTYPE As String
    Property Debit As Decimal
    Property Credit As Decimal
    Property Ref As String

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Sub New()
        Me.ACID = 0
        Me.SUBSIDIARY = ""
        Me.SUBTYPE = ""
        Me.Debit = 0
        Me.Credit = 0
        Me.Ref = ""
    End Sub
    Sub New(ACID As Integer, SUBSIDIARY As Integer, SUBTYPE As String, Debit As Decimal, Credit As Decimal, Ref As String)
        Me.ACID = ACID
        Me.SUBSIDIARY = SUBSIDIARY
        Me.SUBTYPE = SUBTYPE
        Me.Debit = Debit
        Me.Credit = Credit
        Me.Ref = Ref
    End Sub
End Class
