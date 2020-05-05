﻿Imports System.Text
Imports Newtonsoft.Json.Linq
Imports RestSharp

Public Module JsonManager

    Private client As RestClient
    Private serverAcc As String
    Private serverPass As String
    Private authType As AuthenticationType
    Private token As AuthenticationToken

    Public Property GlobalURL As String
    Public Property OperationalURL As String
    Public Property TransactionalURL As String
    Public Property RecallURL As String
    Public Property QueryURL As String

    ''' <summary>
    ''' Call this method to initialize the needed internal objects 
    ''' </summary>
    ''' <param name="url"></param>
    Public Sub Init(url As String)
        client = New RestClient(url)
        SetURLs(url)
        authType = AuthenticationType.NoAuth
    End Sub

    Public Sub Init(url As String, username As String, password As String, authenticationType As AuthenticationType, authToken As AuthenticationToken)
        client = New RestClient(url)
        SetURLs(url)
        serverAcc = username
        serverPass = password
        authType = authenticationType
        token = authToken
    End Sub

    Public Function Post(json As String) As String
        Dim request As RestRequest = New RestRequest(Method.POST)

        Dim byteBody As Byte() = Encoding.UTF8.GetBytes(json)
        Dim hash As String = CreateMD5(json)

        Select Case authType
            Case AuthenticationType.Bearer
                'If there is no valid token avaible return
                If Not token.IsValid Then Throw New Exception("No valid token avaible for the operation")
                'Add the headers and body
                request.AddHeader("Authorization", token.Value)
            Case AuthenticationType.NoAuth
            Case Else
                Throw New NotImplementedException($"{authType.ToString()} not implemented yet")
        End Select

        request.AddHeader("cache-control", "no-cache")
        request.AddHeader("Content-Length", byteBody.Length)
        request.AddHeader("content-type", "application/json")
        request.AddHeader("X-OriginalHash", hash)
        request.AddParameter("application/json", json, ParameterType.RequestBody)


        'Execute
        Dim response = client.Execute(request)
        If Not response.IsSuccessful Then
            Throw New Exception($"POST operation failed, status code: {response.StatusCode.ToString()}")
        End If

        'Dim jsonResponse As JObject = JObject.Parse(response.Content)
        'Dim recallCode As String = jsonResponse.Item("Code")
        Return response.Content
    End Function

    Public Function Post(json As String, url As String) As String
        client = New RestClient(url)
        Return Post(json)
    End Function

    Public Function CreateMD5(ByVal input As String) As String
        Using md5 As Security.Cryptography.MD5 = Security.Cryptography.MD5.Create()
            Dim inputBytes As Byte() = Encoding.ASCII.GetBytes(input)
            Dim hashBytes As Byte() = md5.ComputeHash(inputBytes)
            Dim sb As StringBuilder = New StringBuilder()

            For i As Integer = 0 To hashBytes.Length - 1
                sb.Append(hashBytes(i).ToString("x2"))
            Next

            Return sb.ToString()
        End Using
    End Function

    Private Sub SetURLs(url As String)
        GlobalURL = url
        OperationalURL = url & "/operational"
        TransactionalURL = url & "/transactional"
        RecallURL = url & "/recall"
        QueryURL = url & "/query"
    End Sub
End Module

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