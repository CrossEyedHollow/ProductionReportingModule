Imports System.Net
Imports System.Runtime.CompilerServices
Imports System.Text
Imports System

Public Module Extenders
    <Extension()>
    Public Function IsNullOrEmpty(ByVal array() As String) As Boolean
        Return (IsDBNull(array) OrElse array Is Nothing) OrElse (array.Length < 1)
    End Function

    <Extension()>
    Public Function IsNullOrEmpty(ByVal array() As Integer) As Boolean
        Return (IsDBNull(array) OrElse array Is Nothing) OrElse (array.Length < 1)
    End Function

    <Extension()>
    Public Function IsNullOrEmpty(ByVal array() As Decimal) As Boolean
        Return (IsDBNull(array) OrElse array Is Nothing) OrElse (array.Length < 1)
    End Function

    <Extension()>
    Public Function IsNullOrEmpty(ByVal array As Array) As Boolean
        Return (IsDBNull(array) OrElse array Is Nothing) OrElse (array.Length < 1)
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
        If dt.Rows.Count < 1 Then Return Nothing
        Return dt.Rows.OfType(Of DataRow).Select(Function(dr) dr.Field(Of String)(columnName)).ToArray()
    End Function

    <Extension()>
    Public Sub TryAddRange(ByRef list As List(Of String), array As Array)
        If Not array.IsNullOrEmpty() Then
            list.AddRange(array)
        End If
    End Sub

    <Extension()>
    Public Function ToMD5Hash(ByVal input As String) As String
        Using md5 As Security.Cryptography.MD5 = Security.Cryptography.MD5.Create()
            'Get the bytes
            Dim inputBytes As Byte() = Encoding.ASCII.GetBytes(input)
            'Compute the hash
            Dim hashBytes As Byte() = md5.ComputeHash(inputBytes)
            Dim sb As StringBuilder = New StringBuilder()
            'Convert to string
            For i As Integer = 0 To hashBytes.Length - 1
                sb.Append(hashBytes(i).ToString("x2"))
            Next
            Return sb.ToString()
        End Using
    End Function
End Module
