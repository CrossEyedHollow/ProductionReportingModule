Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports RestSharp

Public Class JsonManager
    Public Property Enabled As Boolean
    Public authType As AuthenticationType
    Private client As RestClient
    Private serverAcc As String
    Private serverPass As String
    Private token As AuthenticationToken

    ''' <summary>
    ''' Call this method to initialize the needed internal objects 
    ''' </summary>
    ''' <param name="url"></param>
    Public Sub New(url As String)
        client = New RestClient(url)
        authType = AuthenticationType.NoAuth
    End Sub

    Public Sub New(url As String, username As String, password As String, authenticationType As AuthenticationType, Optional authToken As AuthenticationToken = Nothing)
        client = New RestClient(url)
        serverAcc = username
        serverPass = password
        authType = authenticationType
        token = authToken
    End Sub

    Public Function Post(json As String) As String
        Dim request As RestRequest = New RestRequest(Method.POST)
        Dim byteBody As Byte() = Encoding.UTF8.GetBytes(json)
        Dim hash As String = json.ToMD5Hash()

        Select Case authType
            Case AuthenticationType.Basic
                Dim strBaseCredentials As String = Convert.ToBase64String(Encoding.ASCII.GetBytes(String.Format("{0}:{1}", serverAcc, serverPass)))
                request.AddHeader("Authorization", $"Basic {strBaseCredentials}")

            Case AuthenticationType.Bearer
                'If there is no valid token avaible return
                If Not token.IsValid Then Throw New Exception("No valid token avaible for the operation")

                'Add the headers and body
                request.AddHeader("Authorization", token.Value)

            Case AuthenticationType.NoAuth
                'Add the headers and body
                request.AddHeader("Authorization", "Basic Og==")
            Case Else
                Throw New NotImplementedException($"{authType.ToString()} not implemented yet")
        End Select

        request.AddHeader("Content-Length", byteBody.Length)
        request.AddHeader("X-OriginalHash", hash)
        request.AddHeader("cache-control", "no-cache")
        request.AddHeader("content-type", "application/json; charset=utf-8")
        request.AddParameter("application/json", json, ParameterType.RequestBody)

        'Execute
        Dim response = client.Execute(request)

        'If Not response.IsSuccessful Then
        '    Throw New Exception($"POST operation failed, status code: {response.StatusCode.ToString()}")
        'End If

        Return response.Content
    End Function

    Public Function Post(json As String, url As String) As String
        client = New RestClient(url)
        Return Post(json)
    End Function

    Public Shared Function StandartResponse(code As String, msgType As String, checksum As String, errors As List(Of ValidationError), Optional expireDate As Date = Nothing) As String
        Dim output As JObject = New JObject()
        output("Code") = code
        output("Message_Type") = msgType
        If errors IsNot Nothing Then
            output("Error") = 1
            'Serialize all errors and join then, format: {error1},{error2}
            output("Errors") = JArray.FromObject(errors)

        Else
            output("Error") = 0
            output("Errors") = Nothing
        End If

        output("Checksum") = checksum
        If expireDate <> Nothing Then output("RecallExpiry_Time") = GetTimeLong(expireDate.AddMonths(6))
        Return output.ToString(Formatting.Indented)
    End Function

    'Public Function HashMD5(ByVal input As String) As String
    '    Using md5 As Security.Cryptography.MD5 = Security.Cryptography.MD5.Create()
    '        Dim inputBytes As Byte() = Encoding.ASCII.GetBytes(input)
    '        Dim hashBytes As Byte() = md5.ComputeHash(inputBytes)
    '        Dim sb As StringBuilder = New StringBuilder()

    '        For i As Integer = 0 To hashBytes.Length - 1
    '            sb.Append(hashBytes(i).ToString("x2"))
    '        Next

    '        Return sb.ToString()
    '    End Using
    'End Function


End Class

Public Class AuthenticationToken
    Public Property Value As String = ""
    Public Property IsValid As Boolean = False
    Public Property ExpiresIn As Integer = 0
End Class

Public Enum AuthenticationType
    NoAuth
    Basic
    Bearer
End Enum