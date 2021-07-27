Public Class DispGridWithParam 
    Private _dtb As DataTable
    Public Property DataTab() As DataTable
        Get
            Return _dtb
        End Get
        Set(ByVal value As DataTable)
            _dtb = value
        End Set
    End Property
    Private _grd As DevExpress.XtraGrid.GridControl
    Public Property GRD() As DevExpress.XtraGrid.GridControl
        Get
            Return _grd
        End Get
        Set(ByVal value As DevExpress.XtraGrid.GridControl)
            _grd = value
        End Set
    End Property 
End Class
