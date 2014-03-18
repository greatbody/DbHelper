Public Interface ISqlDbHelper
    Function ExcuteNonQuery() As Integer
    Function ExcuteNonQuery(ByVal sqlComm As SqlClient.SqlCommand) As Integer

    Function ExecuteScalar(Of T)() As T
    Function ExecuteScalar(Of T)(ByVal sqlComm As SqlClient.SqlCommand) As T

    Function FillDataTable() As DataTable
    Function FillDataTable(ByVal sqlComm As SqlClient.SqlCommand) As DataTable

    Function FillDataSet() As DataSet
    Function FillDataSet(ByVal sqlComm As SqlClient.SqlCommand) As DataSet

    Function From(ByVal sqlStr As String) As SqlDbHelper
End Interface
