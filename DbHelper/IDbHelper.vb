Public Interface IDbHelper
    Function ExcuteNonQuery() As Integer
    Function ExcuteNonQuery(ByVal SqlString As String) As Integer
    Function ExecuteScalar(Of T)() As T
    Function ExecuteScalar(Of T)(ByVal sqlString As String) As T
    Function FillDataTable() As DataTable
    Function FillDataTable(ByVal SqlString As String) As DataTable
    Function FillDataSet() As DataSet
    Function FillDataSet(ByVal SqlString As String) As DataSet
    Function From(ByVal sqlStr As String) As AccessDbHelper
End Interface
