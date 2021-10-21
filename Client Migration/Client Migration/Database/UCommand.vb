Imports MySql.Data.MySqlClient

Public Class UCommand
    Implements IDisposable

    Private Connection As MySqlConnection
    Private command As MySqlCommand
    Private Transaction As MySqlTransaction
    Private isError As Boolean
    Private IsTranactional As Boolean
    Private _LPK As Integer
    Public ReadOnly Property LastPK() As Integer
        Get
            Return _LPK
        End Get
    End Property
    Public ReadOnly Property HasError() As Boolean
        Get
            Return isError
        End Get
    End Property
    Sub connect()
        Try
            Connection = DB.Connection
            command = New MySqlCommand
            command.Connection = Connection
            StartTransaction()
        Catch ex As Exception
            isError = True
        End Try
    End Sub
    Sub New()
        connect()
    End Sub
    Sub StartTransaction()
        Transaction = Connection.BeginTransaction
        command.Transaction = Transaction
        IsTranactional = True
    End Sub
    Sub QueryExecNonQuery(ByVal queryString As String)
        Try
            If isError Then Exit Sub
            command.CommandText = queryString
            command.ExecuteNonQuery()
            _LPK = command.LastInsertedId
        Catch ex As Exception
            _LPK = 0
        End Try
    End Sub
    Public Sub Insert(ByVal tableName As String, ByVal data As Dictionary(Of String, String))
        Try
            If isError Then Exit Sub
            Dim columns As String = "", values As String = ""
            For Each Val As KeyValuePair(Of String, String) In data
                columns += String.Format(" {0} ,", Val.Key.ToString())
                If Val.Value.ToUpper = "NULL" Then
                    values += String.Format("{0},", Val.Value)
                Else
                    values += String.Format("'{0}',", Val.Value)
                End If
            Next
            columns = columns.Substring(0, columns.Length - 1)
            values = values.Substring(0, values.Length - 1)
            QueryExecNonQuery(String.Format("insert into  {0}({1}) values({2});", tableName, columns, values))
        Catch ex As Exception
        End Try
    End Sub
    Public Sub Update(ByVal tableName As String, ByVal data As Dictionary(Of String, String), ByVal whereParams As String)
        Try
            If isError Then Exit Sub
            Dim vals As String = ""
            If data.Count = 0 Then Exit Sub
            For Each Val As KeyValuePair(Of String, String) In data
                If Val.Value.ToUpper = "NULL" Then
                    vals += String.Format(" {0} = {1} ,", Val.Key.ToString(), Val.Value.ToString())
                Else
                    vals += String.Format(" {0} = '{1}',", Val.Key.ToString(), Val.Value.ToString())
                End If
            Next
            vals = vals.Substring(0, vals.Length - 1)
            QueryExecNonQuery(String.Format("update {0} set {1} where {2} ;", tableName, vals, whereParams))
        Catch ex As Exception
        End Try
    End Sub
    Public Sub Delete(ByVal tableName As String, Optional ByVal whereParams As String = "")
        Try
            If isError Then Exit Sub
            If whereParams <> "" Then
                QueryExecNonQuery(String.Format("Delete from {0} where {1} ;", tableName, whereParams))
            Else
                QueryExecNonQuery(String.Format("Delete from {0}  ;", tableName))
            End If
        Catch ex As Exception
        End Try
    End Sub
    Public Function Datasource(ByVal QueryString As String, Optional _fromRow As Integer = 0, Optional ReturnRow As Integer = 0) As DataTable
        Try
            If isError Then Return Nothing
            Dim data As New DataTable
            Dim TA = New MySqlDataAdapter(QueryString, Connection)
            If ReturnRow = 0 Then
                TA.Fill(data)
            Else
                TA.Fill(_fromRow, ReturnRow, data)
            End If
            Return data
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function SaveChanges(Optional ByVal SaveCustomMsg As String = "") As Boolean
        Try
            If Not IsTranactional Then Return True
            If isError Then
                Transaction.Rollback()
                MessageBoxStr("Transaction completed with error")
            Else
                Transaction.Commit()
                MessageBoxStr("Transaction completed successfully.")
            End If
            Return Not isError
        Catch ex As Exception
            Return False
        End Try
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        ' TODO: uncomment the following line if Finalize() is overridden above.
        ' GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
