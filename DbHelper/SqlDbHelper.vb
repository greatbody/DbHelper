Imports System.Data.SqlClient
Public Class SqlDbHelper
    Implements IDbHelper
    Private _Conn As New System.Data.SqlClient.SqlConnection
    Private _ConnString As String = ""
    Private _SqlString As String = ""

    Public Property ConnectionString() As String
        Get
            Return _ConnString
        End Get
        Set(ByVal value As String)
            _ConnString = value
        End Set
    End Property

    Public Sub New()

    End Sub
    Public Sub New(ByVal connString As String)
        _ConnString = connString
        Create()
    End Sub
    Public Sub Create()
        _Conn.ConnectionString = _ConnString
        _Conn.Open()
    End Sub

    Public Function ExcuteNonQuery() As Integer Implements IDbHelper.ExcuteNonQuery
        Return ExcuteNonQuery(_SqlString)
    End Function

    Public Function ExcuteNonQuery(ByVal SqlString As String) As Integer Implements IDbHelper.ExcuteNonQuery
        Using dbCommand As New Data.SqlClient.SqlCommand
            dbCommand.Connection = _Conn
            dbCommand.CommandText = _SqlString
            dbCommand.CommandText = SqlString
            Return dbCommand.ExecuteNonQuery
        End Using
    End Function

    Public Function ExecuteScalar(Of T)() As T Implements IDbHelper.ExecuteScalar
        Return ExecuteScalar(Of T)(_SqlString)
    End Function
    Public Function ExecuteScalar(Of T)(ByVal sqlString As String) As T Implements IDbHelper.ExecuteScalar
        Dim obj As Object
        _SqlString = sqlString
        Using dbCommand As New Data.SqlClient.SqlCommand
            dbCommand.Connection = _Conn
            dbCommand.CommandText = _SqlString
            obj = dbCommand.ExecuteScalar
            If ((obj Is Nothing) OrElse DBNull.Value.Equals(obj)) Then
                Return CType(Nothing, T)
            End If
            If (obj.GetType Is GetType(T)) Then
                Return DirectCast(obj, T)
            End If
            Return DirectCast(Convert.ChangeType(obj, GetType(T)), T)
        End Using
    End Function


    Public Function FillDataSet() As System.Data.DataSet Implements IDbHelper.FillDataSet
        Return FillDataSet(_SqlString)
    End Function

    Public Function FillDataTable() As System.Data.DataTable Implements IDbHelper.FillDataTable
        Return FillDataTable(_SqlString)
    End Function

    Public Function From(ByVal sqlStr As String) As AccessDbHelper Implements IDbHelper.From
        Return New AccessDbHelper(sqlStr)
    End Function

    Public Function FillDataTable(ByVal SqlString As String) As System.Data.DataTable Implements IDbHelper.FillDataTable
        Dim _DataTable As New DataTable("_sunsoft")
        Using dbCommand As New Data.SqlClient.SqlCommand
            dbCommand.Connection = _Conn
            Using dbReader As Data.SqlClient.SqlDataReader = dbCommand.ExecuteReader
                _DataTable.Load(dbReader)
            End Using
        End Using
        Return _DataTable
    End Function

    Public Function FillDataSet(ByVal SqlString As String) As System.Data.DataSet Implements IDbHelper.FillDataSet
        Dim _DataSet As New DataSet
        Using dbCommand As New SqlClient.SqlCommand
            dbCommand.Connection = _Conn
            dbCommand.CommandText = _SqlString
            dbCommand.CommandText = SqlString
            Dim s = New SqlDataAdapter(dbCommand).Fill(_DataSet)
            Dim sCount As Integer
            For sCount = 0 To _DataSet.Tables.Count - 1
                _DataSet.Tables.Item(sCount).TableName = ("_sunsoft" & sCount.ToString)
            Next sCount
            Return _DataSet
        End Using
    End Function
End Class
