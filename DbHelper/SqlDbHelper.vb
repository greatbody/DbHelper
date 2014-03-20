Imports System.Data.SqlClient
Imports System.Text

Public Class SqlDbHelper
    Implements ISqlDbHelper

    Private _Conn As New SqlClient.SqlConnection
    Private _ConnString As String = ""
    Private _SqlString As String = ""
    Private _command As SqlClient.SqlCommand
    Private _sqlBuilder As StringBuilder
    Private _count As Integer = 0
    Public Property ConnectionString() As String
        Get
            Return _ConnString
        End Get
        Set(ByVal value As String)
            _ConnString = value
        End Set
    End Property

    Public Sub New()
        Dim connString As String = System.Configuration.ConfigurationManager.AppSettings.Get("ConnectionString")
        If connString = Nothing Or connString = "" Or String.IsNullOrEmpty(connString) Then
            Throw New Exception("请配置初始化的连接字符串，否则请提供连接字符串")
        End If
        Create(connString)
    End Sub
    Public Sub New(ByVal connString As String)
        Create(connString)
    End Sub
    Private Sub Create(ByVal connString As String)
        '改动：2014年3月20日22:05:34:
        '功能：将打开连接的功能关闭, 仅在执行的时候打开, 执行后关闭
        _ConnString = connString
        _command = New SqlClient.SqlCommand
        _Conn = New SqlClient.SqlConnection(_ConnString)
        _sqlBuilder = New StringBuilder("")

        _ConnString = connString
        _Conn = New SqlConnection(_ConnString)
        _command = New SqlCommand()
        _command.Connection = _Conn

    End Sub
    Public Function Create() As SqlDbHelper
        Return New SqlDbHelper()
    End Function

    Public Function ExcuteNonQuery() As Integer Implements ISqlDbHelper.ExcuteNonQuery
        Return ExcuteNonQuery(Me._command)
    End Function

    Public Function ExcuteNonQuery(ByVal sqlComm As SqlClient.SqlCommand) As Integer Implements ISqlDbHelper.ExcuteNonQuery
        _Conn.Open()
        Using sqlComm
            sqlComm.CommandText = _sqlBuilder.ToString
            Return sqlComm.ExecuteNonQuery
        End Using
        _Conn.Close()
    End Function

    Public Function ExecuteScalar(Of T)() As T Implements ISqlDbHelper.ExecuteScalar
        Return ExecuteScalar(Of T)(Me._command)
    End Function
    Public Function ExecuteScalar(Of T)(ByVal sqlComm As SqlClient.SqlCommand) As T Implements ISqlDbHelper.ExecuteScalar
        '更新：改为在执行中打开连接，执行后关闭
        Dim obj As Object
        Using sqlComm
            _Conn.Open()
            sqlComm.CommandText = _sqlBuilder.ToString
            obj = sqlComm.ExecuteScalar
            _Conn.Close()
            If ((obj Is Nothing) OrElse DBNull.Value.Equals(obj)) Then
                Return CType(Nothing, T)
            End If
            If (obj.GetType Is GetType(T)) Then
                Return DirectCast(obj, T)
            End If
            Return DirectCast(Convert.ChangeType(obj, GetType(T)), T)
        End Using
    End Function

    Public Function FillDataSet() As System.Data.DataSet Implements ISqlDbHelper.FillDataSet
        Return FillDataSet(Me._command)
    End Function

    Public Function FillDataTable() As System.Data.DataTable Implements ISqlDbHelper.FillDataTable
        Return FillDataTable(Me._command)
    End Function

    Public Function From(ByVal sqlStr As String) As SqlDbHelper Implements ISqlDbHelper.From
        Return New SqlDbHelper(sqlStr)
    End Function

    Public Function FillDataTable(ByVal sqlComm As SqlClient.SqlCommand) As System.Data.DataTable Implements ISqlDbHelper.FillDataTable
        '更新：改为在执行中打开连接，执行后关闭
        Dim _DataTable As New DataTable("_sunsoft")
        Using sqlComm
            sqlComm.CommandText = _sqlBuilder.ToString
            _Conn.Open()
            Using dbReader As Data.SqlClient.SqlDataReader = sqlComm.ExecuteReader
                _DataTable.Load(dbReader)
            End Using
            _Conn.Close()
        End Using
        Return _DataTable
    End Function

    Public Function FillDataSet(ByVal sqlComm As SqlClient.SqlCommand) As System.Data.DataSet Implements ISqlDbHelper.FillDataSet
        '更新：改为在执行中打开连接，执行后关闭
        Dim _DataSet As New DataSet
        Using sqlComm
            sqlComm.CommandText = _sqlBuilder.ToString
            _Conn.Open()
            Dim s = New SqlDataAdapter(sqlComm).Fill(_DataSet)
            _Conn.Close()
            Dim sCount As Integer
            For sCount = 0 To _DataSet.Tables.Count - 1
                _DataSet.Tables.Item(sCount).TableName = ("_sunsoft" & sCount.ToString)
            Next sCount
            Return _DataSet
        End Using
    End Function


    '''后面的代码用于运算符重载相关的代码
    Public Shared Operator +(ByVal meDefault As SqlDbHelper, ByVal strp As String) As SqlDbHelper
        '完成:运算符重载
        meDefault._addSqlText(strp)
        Return meDefault
    End Operator

    Private Sub _addSqlText(ByVal paramStr As String)
        _sqlBuilder.Append(paramStr)
    End Sub
    Public Shared Operator +(ByVal meDefault As SqlDbHelper, ByVal strp As QueryParameter) As SqlDbHelper
        'meDefault._sqlBuilder.Append("@" & _count.ToString)
        meDefault._addParam(strp)
        Return meDefault
    End Operator
    ''' <summary>
    ''' 添加参数，重载，参数类型
    ''' </summary>
    ''' <param name="p">直接数据类型</param>
    ''' <remarks></remarks>
    Private Sub _addParam(ByVal p As QueryParameter)
        Me._sqlBuilder.Append("@p" & _count.ToString)
        Me._command.Parameters.AddWithValue("@p" & _count.ToString, p.Value)
    End Sub
End Class
