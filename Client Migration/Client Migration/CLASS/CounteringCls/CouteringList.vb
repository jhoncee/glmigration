Imports System.ComponentModel

Public Class CounteringCls
    Implements INotifyPropertyChanged
    Property PONum As String
    Property IsVat As Boolean
    Private _SI As String
    Public Property SI() As String
        Get
            Return _SI
        End Get
        Set(ByVal value As String)
            _SI = value
            OnPropertyChanged("SI")
        End Set
    End Property
    Private _SIDate As Date
    Public Property SIDate() As Date
        Get
            Return _SIDate
        End Get
        Set(ByVal value As Date)
            _SIDate = value
            OnPropertyChanged("SIDate")
        End Set
    End Property
    Private _SIAMT As Decimal
    Public Property SIAMT() As Decimal
        Get
            Return _SIAMT
        End Get
        Set(ByVal value As Decimal)
            _SIAMT = value
            OnPropertyChanged("SIAMT")
        End Set
    End Property

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Protected Sub OnPropertyChanged(propertyName As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    End Sub
    Sub New()

    End Sub
    '    Sub New(
    '      PONum As String,
    '      POAMT As Decimal,
    '      IsVat As String
    ')

    '        Me.PONum = PONum
    '        Me.IsVat = If(IsVat = "Yes", True, False)
    '        SIAMT = POAMT
    '    End Sub

    Private Sub CounteringCls_PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Handles Me.PropertyChanged
        Try
            'If e.PropertyName = "SIAMT" Or e.PropertyName = "SIDate" And e.PropertyName = "SI" Then

            'End If
        Catch ex As Exception
        End Try
    End Sub
End Class
