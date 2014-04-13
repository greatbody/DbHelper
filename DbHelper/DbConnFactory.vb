Imports System.Data.Common
Public Class DbConnFactory
    Public Enum DbHelperType
        SqlServer = 1
        AccessDB = 2
        Other = 3
    End Enum

    Public Shared Function GetConn(ByVal connType As DbHelperType, ByVal connString As String) As Object
        Dim objConn As Object = Nothing
        Select Case connType
            Case DbHelperType.SqlServer
                objConn = New SqlDbHelper(connString)
                Return objConn
            Case DbHelperType.AccessDB
                objConn = New AccessDbHelper(connString)
                Return objConn
            Case DbHelperType.Other
                Return Nothing
            Case Else
                Return Nothing
        End Select
    End Function

    Public Shared Function GetConn(ByVal connType As DbHelperType) As Object
        Return GetConn(connType, "")
    End Function
End Class