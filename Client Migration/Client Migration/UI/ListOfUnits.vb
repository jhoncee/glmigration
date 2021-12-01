Public Class ListOfUnits

    Sub New(data As List(Of UnitModel))
        InitializeComponent()
        GridControl1.DataSource = data
    End Sub
End Class