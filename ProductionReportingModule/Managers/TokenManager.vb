Imports System.Text
Imports System.Threading
Imports Newtonsoft.Json.Linq
Imports RestSharp
Imports ReportTools

Public Class TokenManager

    Public Property AuthToken As New AuthenticationToken()
    'Public Property TokenState As TokenState = TokenState.Invalid
    'Public Property TokenPeriod As Integer = 10 'This time is in seconds

    Private ReadOnly client As RestClient
    Private ReadOnly serverAcc As String
    Private ReadOnly serverPass As String
    Private ReadOnly authType As AuthenticationType

    Public Sub New(url As String, username As String, password As String, authType As AuthenticationType)
        client = New RestClient(url)
        serverAcc = username
        serverPass = password
        Me.authType = authType
    End Sub

    Public Sub Start()
        Dim tokenThread As New Thread(AddressOf GetToken)
        tokenThread.Start()
    End Sub

    Public Sub GetToken()
        While Main.IsRunning
            AuthToken.IsValid = False

            Dim request As RestRequest = New RestRequest(Method.POST)

            'Assemble request
            request.AddHeader("cache-control", "no-cache")
            request.AddHeader("Connection", "keep-alive")

            Select Case authType
                Case AuthenticationType.NoAuth
                    request.AddHeader("content-length", "74")
                    request.AddHeader("authorization", "Basic Y2xpZW50Og==")
                    request.AddParameter("undefined", $"grant_type=password&username={serverAcc}&password={serverPass}", ParameterType.RequestBody)
                Case AuthenticationType.Basic
                    Dim strBaseCredentials As String = Convert.ToBase64String(Encoding.ASCII.GetBytes(String.Format("{0}:{1}", serverAcc, serverPass)))
                    request.AddHeader("authorization", $"Basic {strBaseCredentials}")
                Case Else
                    Throw New NotImplementedException($"{authType} not aplicable for TokenManager")
            End Select

            request.AddHeader("Content-Type", "application/x-www-form-urlencoded") 'Do not change the content type... magic

            Try
                Dim response = client.Execute(request)

                If response.IsSuccessful Then
                    Try
                        Dim json As JObject = JObject.Parse(response.Content)
                        Dim access_token As String = json.Item("access_token")
                        Dim expires_in As Integer = json.Item("expires_in")
                        Dim token_type As String = json.Item("token_type")

                        AuthToken.IsValid = True

                        AuthToken.ExpiresIn = expires_in - 3
                        AuthToken.Value = access_token

                        Output.ToConsole($"New access token acquired, expires in {expires_in}s")
                    Catch ex As Exception
                        Retry()
                        Output.ToConsole($"Exception occured while aquiring token: {ex.Message}")
                    End Try
                Else
                    Retry()
                    Output.ToConsole("Request failed, status: " & response.StatusCode.ToString())
                End If
            Catch ex As Exception
                Retry()
                Output.ToConsole($"Executing request failed: {ex.Message}")
            End Try

            'Start the new reading 5 seconds before the token expires
            Thread.Sleep(TimeSpan.FromSeconds(AuthToken.ExpiresIn)) 'TimeSpan.FromSeconds(TokenPeriod)
        End While
    End Sub

    Private Sub Retry()
        AuthToken.IsValid = False
        AuthToken.ExpiresIn = 10
    End Sub
End Class
