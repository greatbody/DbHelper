Imports System.Data.Common
Public Class DbConnFactory
    Public Shared Function GetConn(ByVal connType As String, ByVal connString As String)
        Dim objConn As Object = Nothing
        Select Case connType
            Case "sqlserver"
                objConn = New SqlServerConn
                Return objConn
            Case "mdb"
                objConn = New MdbConn
                Return objConn
            Case "mysql"
                Return Nothing
            Case Else
                Return Nothing
        End Select

    End Function
End Class
Public Class SqlServerConn
    Private _conn As System.Data.Common.DbConnection
    Public Sub New()
        _conn = New SqlClient.SqlConnection


    End Sub


End Class

Public Class MdbConn
    Public Sub New()

    End Sub


End Class