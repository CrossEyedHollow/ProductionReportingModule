Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Threading
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ReportTools

Public Class JsonListener

    Dim listener As HttpListener
    Shared Property Prefix As String
    Shared Property Users As Dictionary(Of String, String)

    Public Sub New()
        listener = New HttpListener()
        listener.Prefixes.Add("http://localhost:8080/")
        listener.Prefixes.Add("http://127.0.0.1:8080/")
        listener.Prefixes.Add(Prefix)
        listener.AuthenticationSchemes = AuthenticationSchemes.Basic

    End Sub

    ''' <summary>
    ''' Starts the listener
    ''' </summary>
    Public Sub Start()
        Dim thrdListener = New Thread(AddressOf Listen)
        Dim db As New DBManager()
        'Get all authenticated users from db
        'Dim result As DataTable = db.GetAuthenticatedUsers()

        'convert to KeyValuePair
        'For Each row As DataRow In result.Rows
        '    Dim user As String = row("fldUser")
        '    Dim pass As String = row("fldPassword")

        '    Users.Add(user, pass)
        'Next

        listener.Start()
        thrdListener.Start()
    End Sub

    Private Sub Listen()
        Dim context As HttpListenerContext

        While Main.IsRunning
            Try
                'Listen
                context = listener.GetContext()
                Dim id As HttpListenerBasicIdentity = context.User.Identity
                Dim user As String = id.Name
                Dim pass As String = id.Password

                'Proccess the message
                Dim task As Task = Task.Factory.StartNew(Sub() ProccessMessage(context))

            Catch ex As Exception
                'If something fails log the error
                Output.Report("Failed to get context of incoming transaction.")
            End Try
        End While
    End Sub

    Public Sub ProccessMessage(context As HttpListenerContext)
        Try
            'Make a new instance of the db so we can have multiple async connections
            Dim db As New DBManager()

            'Convert message to string
            Dim rawText As String = New StreamReader(context.Request.InputStream, context.Request.ContentEncoding).ReadToEnd()
            Dim answer As String = ""
            Dim responseCode As Integer = 202

            'Parse the incoming msg
            Dim json As JObject = JObject.Parse(rawText)
            Dim msgType As String = json("Message_Type")
            Dim code As String = json("Code")

            'Assemble respective answer
            Select Case msgType.ToUpper()
                Case "IRU"
                    'Assemble response that mimics the secondary repository standary response
                    Dim eventTime As Date = ParseTime(json("Event_Time"))
                    Dim checksum As String = CreateMD5(rawText)
                    answer = StandartResponse(code, msgType, checksum, eventTime)

                    'Save the json in alternative table
                    If db.InsertRawJson("tbljsonsecondary", rawText, msgType.ToUpper()) Then
                        Output.ToConsole("New IRU was received from secondary repository and was sent to the Database")
                    End If
                Case "IRA"
                    Throw New NotImplementedException("IRA message not implemented")
                Case "STA"
                    'Get the response from database
                    Dim result As DataTable = db.CheckForCode(code)

                    'If match is found
                    If result.Rows.Count > 0 Then
                        'Get response
                        Dim response As String = result.Rows(0)("fldResponse")

                        'Validate response
                        If response.IsNullOrEmpty() Then
                            'Return error
                            answer = "Code matched but response field is empty. Message might still be in queue."
                            responseCode = 503
                            Output.ToConsole($"STA message fail: {answer}")
                        Else
                            'Anwer with the response from the secondary
                            answer = response
                            Output.ToConsole($"Responding to STA, code: '{code}'.")
                        End If
                    Else 'No matches
                        answer = $"No matching entities found for code: '{code}'"
                        responseCode = 503
                        Output.ToConsole($"STA message fail: {answer}")
                    End If
                Case Else
                    'Create simple reasponse 
                    answer = MyResponse(code)

                    'Save the json in the db
                    If db.InsertRawJson(DBManager.TableName, rawText, msgType.ToUpper(), code) Then
                        Output.ToConsole("New Json sent to the Database")
                    End If
            End Select

            'Return a response
            Respond(context, answer, responseCode)
        Catch ex As Exception
            'If something fails, respond with error message and log the error
            Dim reason As String = $"Failed to process incoming message: {ex.Message}"
            Output.Report(reason)
            Try
                Respond(context, reason, 500) '500 = Internal server error
            Catch exx As Exception
                Output.Report($"Failed to respond. Error: {exx.Message}")
            End Try
        End Try
    End Sub

    Private Shared Function Respond(context As HttpListenerContext, answer As String, statusCode As Integer) As HttpListenerContext
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

    Private Function MyResponse(guid As String) As String
        Dim output As String = "{" & vbCrLf &
            vbTab & $"""Code"": ""{guid}""" & vbCrLf &
            vbTab & vbCrLf & "}"
        Return output
    End Function

    Private Function StandartResponse(code As String, msgType As String, checksum As String, Optional expireDate As Date = Nothing) As String
        Dim output As JObject = New JObject()
        output("Code") = code
        output("Message_Type") = msgType
        output("Error") = 0
        output("Errors") = Nothing
        output("Checksum") = checksum
        If expireDate <> Nothing Then output("RecallExpiry_Time") = GetTimeLong(expireDate.AddMonths(6))
        Return output.ToString(Formatting.Indented)
    End Function
End Class
