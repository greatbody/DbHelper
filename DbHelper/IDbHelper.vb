Public Interface IDbHelper
    Function ExcuteNonQuery() As Integer
    Function ExcuteNonQuery(ByVal sqlComm As OleDb.OleDbCommand) As Integer

    Function ExecuteScalar(Of T)() As T
    Function ExecuteScalar(Of T)(ByVal sqlComm As OleDb.OleDbCommand) As T

    Function FillDataTable() As DataTable
    Function FillDataTable(ByVal sqlComm As OleDb.OleDbCommand) As DataTable

    Function FillDataSet() As DataSet
    Function FillDataSet(ByVal sqlComm As OleDb.OleDbCommand) As DataSet

    Function From(ByVal sqlStr As String) As AccessDbHelper
End Interface
