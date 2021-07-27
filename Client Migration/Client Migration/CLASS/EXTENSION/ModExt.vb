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
    Function REMOVECOMMA(ByVal str As String) As String
        Return str.Replace(",", "")
    End Function
    <Extension()>
    Function RemoveLastChar(ByVal str As String) As String
        Return str.Remove(str.Length - 1, 1)
        Dim s As String
    End Function
    Function CopyRight() As String
        Return getCopyRight()
    End Function
    Function getCopyRight() As String
        Return My.Application.Info.Copyright.ToString
    End Function

    <Extension()>
    Function ToDouble(ByVal Str As String) As String
        Try
            Return FormatCurrency(CDec(Str), 2, TriState.True, TriState.True)
        Catch ex As Exception
            Return 0
        End Try
    End Function
    <Extension()>
    Function ToDecimal(ByVal Str As String) As String
        Try
            Return FormatCurrency(CDec(Str), 2, TriState.True, TriState.True)
        Catch ex As Exception
            Return 0
        End Try
    End Function
    <Extension()>
    Function FormatPercentage(ByVal Str As String) As String
        Try
            Dim d As Double = CDec(Str)
            If Double.IsInfinity(d) Then
                Return FormatPercent(0, 2, TriState.True, TriState.True)
            ElseIf Double.IsNaN(d) Then
                Return FormatPercent(0, 2, TriState.True, TriState.True)
            Else
                Return FormatPercent(CDec(Str), 2, TriState.True, TriState.True)
            End If
        Catch ex As Exception
            Return FormatPercent(0, 2, TriState.True, TriState.True)
        End Try
    End Function
    <Extension()>
    Function FormatPercentage(ByVal Str As Decimal) As String
        Try
            Dim d As Double = CDec(Str)
            If Double.IsInfinity(d) Then
                Return FormatPercent(0, 2, TriState.True, TriState.True)
            ElseIf Double.IsNaN(d) Then
                Return FormatPercent(0, 2, TriState.True, TriState.True)
            Else
                Return FormatPercent(CDec(Str), 2, TriState.True, TriState.True)
            End If

        Catch ex As Exception
            Return FormatPercent(0, 2, TriState.True, TriState.True)
        End Try
    End Function
    <Extension()>
    Function IsNumerics(ByVal Str As String) As Boolean
        Try
            If IsNumeric(Str) = False Then
                Return False
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    <Extension()>
    Function ReplaceChar(ByVal str As String, ByVal CharTobeReplace As String, ByVal ReplaceMentchar As String) As String
        Return str.Replace(CharTobeReplace, ReplaceMentchar)
    End Function
    <Extension()>
    Function ConvertToWord(ByVal str As String) As String
        Dim BC As New Basic.ConvertToWord
        Return BC.SpellNumber(str)
    End Function
    <Extension()>
    Function FormatToDouble(ByVal str As Object) As String
        Try
            Return FormatCurrency(str, 2, , TriState.True, TriState.True)
        Catch ex As Exception
            Return 0
        End Try
    End Function

    <Extension()>
    Public Function GET_RUNNING_BALANCE(ByVal view As GridView, ByVal listSourceRowIndex As Integer, ByVal NORMAL_BALANCE As String) As Decimal
        Try
            Dim CR As Decimal = Convert.ToDecimal(view.GetListSourceRowCellValue(listSourceRowIndex, "Credit"))
            Dim DR As Decimal = Convert.ToDecimal(view.GetListSourceRowCellValue(listSourceRowIndex, "Debit"))
            Dim r As Decimal = 0
            If listSourceRowIndex > 0 Then
                r = Convert.ToDecimal(view.GetListSourceRowCellValue(listSourceRowIndex - 1, "BALANCE"))
            End If

            If NORMAL_BALANCE = "Debit" Then
                Return DR - CR + r
            Else
                Return CR - DR + r
            End If
        Catch ex As Exception
            Return 0
        End Try
    End Function

    <Extension()>
    Public Function GET_RUNNING_BALANCE(ByVal view As GridView, ByVal listSourceRowIndex As Integer) As Decimal
        Try
            Dim BAL As Decimal = Convert.ToDecimal(view.GetListSourceRowCellValue(listSourceRowIndex, "BBAL"))
            Dim r As Decimal = 0
            If listSourceRowIndex > 0 Then
                r = Convert.ToDecimal(view.GetListSourceRowCellValue(listSourceRowIndex - 1, "BALANCE"))
            End If
            Return BAL + r
        Catch ex As Exception
            Return 0
        End Try
    End Function

    <Extension()>
    Public Sub AddCustomColumn(ByRef grd As DevExpress.XtraGrid.GridControl, ByVal VIEW As GridView, ByVal POSITION As Integer)
        Try
            grd.ForceInitialize()
            Dim unbColumn As GridColumn = VIEW.Columns.AddField("BALANCE")
            unbColumn.VisibleIndex = POSITION
            unbColumn.UnboundType = DevExpress.Data.UnboundColumnType.Decimal
            unbColumn.OptionsColumn.AllowEdit = False
            unbColumn.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric
            unbColumn.DisplayFormat.FormatString = "{0:n2}"
            unbColumn.AppearanceCell.BackColor = Color.Cornsilk
            unbColumn.FieldName = "BALANCE"
        Catch ex As Exception
        End Try
    End Sub
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
    '<Extension()>
    'Function ToMysqlFormat(ByVal param As Nullable(Of DateTime)) As String
    '    If param.HasValue Then
    '        Return Format(param, "yyyy-MM-dd")
    '    Else
    '        Return "NULL"
    '    End If
    'End Function
    '<Extension()>
    'Function ToMysqlFormat(ByVal param As Object) As String
    '    Try
    '        If IsDate(param) Then
    '            Return Format(CDate(param), "yyyy-MM-dd")
    '        Else
    '            Return "NULL"
    '        End If
    '    Catch ex As Exception
    '        Return "NULL"
    '    End Try

    'End Function


    <Extension()>
    Sub ToCustomFormat(ByRef param As DevExpress.XtraEditors.DateEdit)
        param.Properties.DisplayFormat.FormatString = "MMM dd,yyyy"
        param.Properties.EditFormat.FormatString = "MMM dd,yyyy"
        param.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Custom
    End Sub


    <Extension()>
    Sub ToCustomFormat(ByRef param As DateTimePicker)
        param.CustomFormat = "MMM dd,yyyy"
        param.Format = DateTimePickerFormat.Custom
    End Sub

    <Extension()>
    Sub ToCustomFormat(ByRef param As RepositoryItemDateEdit)
        Try
            Dim D = param
            D.Properties.DisplayFormat.FormatString = "MMM dd,yyyy"
            D.Properties.EditFormat.FormatString = "MMM dd,yyyy"
            D.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Custom
        Catch ex As Exception
        End Try
    End Sub

    <Extension()>
    Sub AddEmptyRows(ByRef DT As DataTable, ByVal NumRows As Integer)
        For i As Integer = 0 To NumRows
            DT.Rows.Add(Nothing, Nothing)
        Next
    End Sub
    <Extension()>
    Sub AddALLText(ByRef DT As DataTable, Optional ByVal index As Integer = 0)
        Try
            Dim r As DataRow
            r = DT.NewRow
            r("ID") = 0
            r("Name") = "ALL"
            r("PCODE") = ""
            r("ADDRESS") = ""
            r("TYPE") = ""
            DT.Rows.InsertAt(r, index)
        Catch ex As Exception
        End Try
    End Sub
    <Extension()>
    Public Function ToDataTable(ByVal List As System.Collections.IEnumerable) As DataTable
        Return ConvertToDatatable.ToDataTable(List)
    End Function
    <Extension()>
    Sub SAVE_COMMAND(ByVal frm As Form)
        SAVE(CType(frm, IACTIONS))
    End Sub
    Dim o As New Object
    Private Sub SPEAKER(ByVal s As String)
        'Try

        '    Dim synth As New SpeechSynthesizer
        '    Dim voices As System.Collections.ObjectModel.ReadOnlyCollection(Of System.Speech.Synthesis.InstalledVoice) _
        '     = synth.GetInstalledVoices()
        '    synth.SelectVoice(voices(2).VoiceInfo.Name)
        '    synth.Volume = 100

        '    synth.Speak(s)

        'Catch ex As Exception
        'End Try
    End Sub
    <Extension()>
    Public Sub Speak(ByVal s As String)
        'Try
        '    If My.Settings.WITH_AUDIO = True Then
        '        Dim t As New Threading.Thread(AddressOf SPEAKER)
        '        t.Start(s)
        '        't.Join()
        '    End If
        'Catch ex As Exception

        'End Try

    End Sub

    <Extension()>
    Function GET_SIGNATORY(ByVal S As String, ByVal POSITION As String) As String
        Try
            Return GET_SIGNATORY_BY_POSTION(POSITION)
        Catch ex As Exception
            Return ""
        End Try
    End Function

End Module
