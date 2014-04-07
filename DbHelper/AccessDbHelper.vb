﻿Imports System.Data.OleDb
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
    Private _command As OleDbCommand '命令对象
    Private _sqlBuilder As StringBuilder 'sql字符串创建者
    ''' <summary>
    ''' 获取连接字符串
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ConnectionString() As String
        Get
            Return _ConnString
        End Get
        Set(ByVal value As String)
            _ConnString = value
        End Set
    End Property

    '2014年4月7日21:27:54 调整逻辑，采取执行时启动连接的做法
    Public Sub New()
        'ok at 2014年3月9日19:47:51
        '初始化变量：
        _sqlBuilder = New StringBuilder("")
        _command = New OleDbCommand
        '读取连接字符串
        Dim connString As String = System.Configuration.ConfigurationManager.AppSettings.Get("ConnectionString")
        _ConnString = connString
    End Sub

    Public Sub New(ByVal connString As String)
        'ok at 2014年4月7日21:30:16
        '修改代码，修正逻辑
        _ConnString = connString
        _sqlBuilder = New StringBuilder("")
        _command = New OleDbCommand
    End Sub

    ''' <summary>
    ''' 直接从文件建立连接【自动生成连接字符串】
    ''' </summary>
    ''' <param name="FileName"></param>
    ''' <remarks></remarks>
    Public Sub FromFile(ByVal FileName As String)
        '2014年4月7日21:33:23 修正逻辑，之前忘记将取得的路径付给内部变量
        _sqlBuilder = New StringBuilder("")
        _command = New OleDbCommand
        _ConnString = ConnStrFromFile(FileName)
    End Sub

    ''' <summary>
    ''' 根据文件路径生成连接字符串
    ''' </summary>
    ''' <param name="FileName">文件路径</param>
    ''' <returns>数据库连接字符串</returns>
    ''' <remarks></remarks>
    Private Function ConnStrFromFile(ByVal FileName As String) As String
        Dim connStr As String = String.Format("PROVIDER=Microsoft.ACE.OLEDB.12.0;Data Source={0};", FileName)
        Return connStr
    End Function

    Public Shared Function Create() As AccessDbHelper
        '2014年4月7日21:34:18 这个函数直接返回一个这个对象
        Return New AccessDbHelper
    End Function

    Public Sub CreateConn(ByVal connString As String)
        'ok at 2014年3月9日19:32:27
        '修改：2014年3月16日19:37:17
        '注释：这个函数在启动SQL执行前调用，不可任意调用
        _Conn = New OleDb.OleDbConnection(connString)
        _command.Connection = _Conn
        _Conn.Open()
    End Sub

    Public Shared Operator +(ByVal meDefault As AccessDbHelper, ByVal strp As String) As AccessDbHelper
        '完成:运算符重载
        meDefault._addSqlText(strp)
        Return meDefault
    End Operator

    Public Shared Operator +(ByVal meDefault As AccessDbHelper, ByVal strp As QueryParameter) As AccessDbHelper
        'meDefault._sqlBuilder.Append("@" & _count.ToString)
        meDefault._addParam(strp)
        Return meDefault
    End Operator

    Private Sub _addSqlText(ByVal paramStr As String)
        _sqlBuilder.Append(paramStr)
        _command.CommandText = _sqlBuilder.ToString
    End Sub

    ''' <summary>
    ''' 添加参数，重载，参数类型
    ''' </summary>
    ''' <param name="p">直接数据类型</param>
    ''' <remarks></remarks>
    Private Sub _addParam(ByVal p As QueryParameter)
        Me._sqlBuilder.Append("@p" & _count.ToString)
        Me._command.Parameters.AddWithValue("@p" & _count.ToString, p.Value)
    End Sub

    Public Function ExcuteNonQuery(ByVal sqlComm As System.Data.OleDb.OleDbCommand) As Integer Implements IDbHelper.ExcuteNonQuery
        CreateConn(_ConnString)
        Using sqlComm
            Return sqlComm.ExecuteNonQuery
        End Using
    End Function
    Public Function ExcuteNonQuery() As Integer Implements IDbHelper.ExcuteNonQuery
        Return ExcuteNonQuery(Me._command)
    End Function

    Public Function ExecuteScalar(Of T)(ByVal sqlComm As OleDb.OleDbCommand) As T Implements IDbHelper.ExecuteScalar
        CreateConn(_ConnString)
        Dim obj As Object
        Using sqlComm
            obj = sqlComm.ExecuteScalar
            If ((obj Is Nothing) OrElse DBNull.Value.Equals(obj)) Then
                Return CType(Nothing, T)
            End If
            If (obj.GetType Is GetType(T)) Then
                Return DirectCast(obj, T)
            End If
        End Using
        Return DirectCast(Convert.ChangeType(obj, GetType(T)), T)
    End Function
    Public Function ExecuteScalar(Of T)() As T Implements IDbHelper.ExecuteScalar
        Return (ExecuteScalar(Of T)(Me._command))
    End Function

    ''' <summary>
    ''' 从连接字符串创建AccessDbHelper对象
    ''' </summary>
    ''' <param name="ConnectionString">连接字符串</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function From(ByVal ConnectionString As String) As AccessDbHelper Implements IDbHelper.From
        Return New AccessDbHelper(ConnectionString)
    End Function

    Public Function FillDataTable(ByVal sqlComm As OleDb.OleDbCommand) As System.Data.DataTable Implements IDbHelper.FillDataTable
        CreateConn(_ConnString)
        Dim _DataTable As New DataTable("_sunsoft")
        Using sqlComm
            Using dbReader As OleDb.OleDbDataReader = sqlComm.ExecuteReader
                _DataTable.Load(dbReader)
            End Using
        End Using
        Return _DataTable
    End Function
    Public Function FillDataTable() As System.Data.DataTable Implements IDbHelper.FillDataTable
        Return FillDataTable(Me._command)
    End Function

    Public Function FillDataSet(ByVal sqlComm As OleDb.OleDbCommand) As System.Data.DataSet Implements IDbHelper.FillDataSet
        CreateConn(_ConnString)
        Dim _DataSet As New DataSet
        Using sqlComm
            Dim s = New OleDbDataAdapter(sqlComm).Fill(_DataSet)
            Dim sCount As Integer
            For sCount = 0 To _DataSet.Tables.Count - 1
                _DataSet.Tables.Item(sCount).TableName = ("_sunsoft" & sCount.ToString)
            Next sCount
            Return _DataSet
        End Using
    End Function
    Public Function FillDataSet() As System.Data.DataSet Implements IDbHelper.FillDataSet
        Return FillDataSet(Me._command)
    End Function
    '内部函数……………………
    Private Sub Dispose()
        Me.Dispose()
    End Sub

End Class
