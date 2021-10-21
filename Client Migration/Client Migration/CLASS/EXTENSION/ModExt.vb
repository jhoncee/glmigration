Imports System.Runtime.CompilerServices
Imports DevExpress.XtraGrid.Views.Grid
Imports DevExpress.XtraGrid.Columns
Imports DevExpress.XtraEditors
Imports DevExpress.XtraEditors.Controls
Imports DevExpress.XtraEditors.Repository

Module MOD_EXT

    <Extension()>
    Sub EmptyRepoGrid(Repo As GridLookUpEdit, e As ButtonPressedEventArgs, ByRef Grdview As GridView, Optional GridviewCols As String = "SUBSIDIARY")
        Try
            If e.Button.Caption = "X" Then
                Grdview.SetRowCellValue(Grdview.FocusedRowHandle, GridviewCols, 0)
            End If
        Catch ex As Exception
        End Try
    End Sub

    <Extension()>
    Sub EmptyRepoGrid(Repo As LookUpEdit, e As ButtonPressedEventArgs, ByRef Grdview As GridView, cols As String)
        Try
            If e.Button.Caption = "X" Then
                Grdview.SetRowCellValue(Grdview.FocusedRowHandle, cols, 0)
            End If
        Catch ex As Exception
        End Try
    End Sub
    <Extension()>
    Sub NullRepoDate(Repo As DateEdit, e As ButtonPressedEventArgs, ByRef Grdview As GridView, GridviewCols As String)
        Try
            If e.Button.Caption = "X" Then
                Grdview.SetRowCellValue(Grdview.FocusedRowHandle, GridviewCols, "")
            End If
        Catch ex As Exception
        End Try
    End Sub

    <Extension()>
    Sub EmptyRepoGrid(Repo As GridLookUpEdit, e As ButtonPressedEventArgs)
        Try
            If e.Button.Caption = "X" Then
                Repo.EditValue = 0
            End If
        Catch ex As Exception
        End Try
    End Sub
    <Extension()>
    Function RSQ(ByVal str As String) As String
        If str = "" Then
            Return ""
        End If
        Return str.Replace("'", "\'")
    End Function
    <Extension()>
    Function ToMysqlFormat(ByVal param As Date) As String
        Try
            If IsDate(param) Then
                Return Format(param, "yyyy-MM-dd")
            Else
                Return "NULL"
            End If
        Catch ex As Exception
            Return "NULL"
        End Try
    End Function

End Module
