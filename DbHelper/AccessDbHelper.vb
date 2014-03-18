Imports System.Data.OleDb
Imports System.Configuration
Imports System.Text
Imports SunSoft.DAL.StringUtility
Public Class AccessDbHelper
    Implements IDbHelper
    Private _Conn As OleDb.OleDbConnection
    Private _ConnString As String = ""
    Private _SqlString As String = ""
    ' Private _rec As oledb.
    Private _count As Integer '计数参数的个数
    Private _command As OleDbCommand
    Private _sqlBuilder As StringBuilder
    Public Property ConnectionString() As String
        Get
            Return _ConnString
        End Get
        Set(ByVal value As String)
            _ConnString = value
        End Set
    End Property

    Public Sub New()
        'ok at 2014年3月9日19:47:51
        '初始化变量：
        _sqlBuilder = New StringBuilder("")
        Dim connString As String = System.Configuration.ConfigurationManager.AppSettings.Get("ConnectionString")
        If _ConnString.Length > 0 Then
            'Create(_ConnString)
        Else
            'Create(connString)
            _ConnString = connString
        End If
    End Sub
    Public Sub New(ByVal connString As String)
        'ok at 2014年3月9日19:34:24
        _sqlBuilder = New StringBuilder("")
        _command = New OleDbCommand
        CreateConn(connString)
    End Sub
    Public Shared Function Create() As AccessDbHelper
        'ok at
        
        Return New AccessDbHelper
    End Function

    Public Sub CreateConn(ByVal connString As String)
        'ok at 2014年3月9日19:32:27
        '修改：2014年3月16日19:37:17
        _Conn = New OleDb.OleDbConnection(connString)
        _Conn.Open()
    End Sub

    Public Shared Operator +(ByVal meDefault As AccessDbHelper, ByVal strp As String) As AccessDbHelper
        '完成:运算符重载
        meDefault._addSqlText(strp)
        Return meDefault
    End Operator

    Private Sub _addSqlText(ByVal paramStr As String)
        _sqlBuilder.Append(paramStr)
    End Sub
    Public Shared Operator +(ByVal meDefault As AccessDbHelper, ByVal strp As QueryParameter) As AccessDbHelper
        'meDefault._sqlBuilder.Append("@" & _count.ToString)
        meDefault._addParam(strp)
        Return meDefault
    End Operator
    Private Sub _addParam(ByVal p As QueryParameter)
        Me._sqlBuilder.Append("@" & _count.ToString)
        Me.
    End Sub
    Public Function ExcuteNonQuery() As Integer Implements IDbHelper.ExcuteNonQuery
        Return ExcuteNonQuery(_SqlString)
    End Function

    Public Function ExcuteNonQuery(ByVal SqlString As String) As Integer Implements IDbHelper.ExcuteNonQuery
        CreateConn(_ConnString)
        Using dbCom As New OleDb.OleDbCommand(SqlString)
            dbCom.Connection = _Conn
            Return dbCom.ExecuteNonQuery
        End Using
    End Function

    Public Function ExecuteScalar(Of T)() As T Implements IDbHelper.ExecuteScalar
        Dim k As String = "34"

        Return (ExecuteScalar(Of T)(_SqlString))
    End Function
    Public Function ExecuteScalar(Of T)(ByVal sqlString As String) As T Implements IDbHelper.ExecuteScalar
        CreateConn(_ConnString)
        Dim obj As Object
        _SqlString = sqlString
        Using dbCommand As New OleDb.OleDbCommand
            dbCommand.Connection = _Conn
            dbCommand.CommandText = _SqlString
            obj = dbCommand.ExecuteScalar
            If ((obj Is Nothing) OrElse DBNull.Value.Equals(obj)) Then
                Return CType(Nothing, T)
            End If
            If (obj.GetType Is GetType(T)) Then
                Return DirectCast(obj, T)
            End If
        End Using
        Return DirectCast(Convert.ChangeType(obj, GetType(T)), T)
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
        _SqlString = SqlString
        Using dbCommand As New OleDb.OleDbCommand
            dbCommand.Connection = _Conn
            dbCommand.CommandText = _SqlString
            Using dbReader As OleDb.OleDbDataReader = dbCommand.ExecuteReader
                _DataTable.Load(dbReader)
            End Using
        End Using
        Return _DataTable
    End Function

    Public Function FillDataSet(ByVal SqlString As String) As System.Data.DataSet Implements IDbHelper.FillDataSet
        Dim _DataSet As New DataSet
        _SqlString = SqlString
        Using dbCommand As New OleDb.OleDbCommand
            dbCommand.Connection = _Conn
            dbCommand.CommandText = _SqlString
            Dim s = New OleDbDataAdapter(dbCommand).Fill(_DataSet)
            Dim sCount As Integer
            For sCount = 0 To _DataSet.Tables.Count - 1
                _DataSet.Tables.Item(sCount).TableName = ("_sunsoft" & sCount.ToString)
            Next sCount
            Return _DataSet
        End Using
    End Function
    '内部函数……………………
    Private Sub Dispose()
        Me.Dispose()
    End Sub
End Class
