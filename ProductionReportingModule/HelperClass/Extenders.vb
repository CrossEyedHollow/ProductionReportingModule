Imports System.Net
Imports System.Runtime.CompilerServices
Imports System.Text

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

    <Extension()>
    Public Function Respond(ByRef context As HttpListenerContext, answer As String, statusCode As Integer) As HttpListenerContext
        'Get the answer as byte array
        Dim byteAnswer As Byte() = Encoding.UTF8.GetBytes(answer)

        'Set needed parameters
        context.Response.StatusCode = statusCode
        context.Response.ContentLength64 = byteAnswer.Length
        context.Response.ContentType = "application/json"

        'Send the answer
        context.Response.OutputStream.Write(byteAnswer, 0, byteAnswer.Count)
        Return context
    End Function

    <Extension()>
    Public Function HasDuplicates(ByRef array As String()) As Boolean
        'If the array is empty or less than 2 elements, return false
        If array.Length < 2 Then Return False
        'Compare array to its Distinct version
        If array.Length <> array.Distinct().Count Then
            Return True
        End If
        Return False
    End Function

    <Extension()>
    Public Function ColumnToArray(ByVal dt As DataTable, columnName As String) As String()
        Return dt.Rows.OfType(Of DataRow).Select(Function(dr) dr.Field(Of String)(columnName)).ToArray()
    End Function
End Module
