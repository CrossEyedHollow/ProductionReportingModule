Imports System.Runtime.CompilerServices

Public Module Extenders
    <Extension()>
    Public Function IsNullOrEmpty(ByVal array() As String) As Boolean
        Return IsDBNull(array) OrElse (array.Length < 1)
    End Function

    <Extension()>
    Public Function IsNullOrEmpty(ByVal array() As Integer) As Boolean
        Return IsDBNull(array) OrElse (array.Length < 1)
    End Function

    <Extension()>
    Public Function IsNullOrEmpty(ByVal array() As Decimal) As Boolean
        Return IsDBNull(array) OrElse (array.Length < 1)
    End Function

    <Extension()>
    Public Function IsNullOrEmpty(ByVal str As String) As Boolean
        Return String.IsNullOrEmpty(str) Or str Is Nothing
    End Function
End Module
