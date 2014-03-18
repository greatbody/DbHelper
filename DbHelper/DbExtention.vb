Imports System.Runtime.CompilerServices
Namespace StringUtility
    <Extension()> _
    Module StringExtensions
        <Extension()> _
        Public Function AsQueryParameter(ByVal b As String) As QueryParameter
            Return New QueryParameter(b)
        End Function
    End Module
End Namespace









