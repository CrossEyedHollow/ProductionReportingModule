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
    Shared Property Users As List(Of User)

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
        Users = New List(Of User)
        Dim db As New DBManager()
        'Get all authenticated users from db
        Dim result As DataTable = db.GetAuthenticatedUsers()

        'convert to KeyValuePair
        For Each row As DataRow In result.Rows
            Dim user As String = row("fldUser")
            Dim pass As String = row("fldPassword")
            Dim fldSecondary As String = CBool(row("fldSecondary"))

            Users.Add(New User() With {.Name = user, .Password = pass, .IsSecondary = fldSecondary})
        Next

        listener.Start()
        thrdListener.Start()
    End Sub

    Private Sub Listen()
        Dim context As HttpListenerContext

        While Main.IsRunning
            Try
                'Listen
                context = listener.GetContext()

                'Proccess the message
                Dim task As Task = Task.Factory.StartNew(Sub() ProccessMessage(context))

            Catch ex As Exception
                'If something fails log the error
                Output.Report($"Failed to process message. Reason: {ex.Message}")
            End Try
        End While
    End Sub


    Public Sub ProccessMessage(context As HttpListenerContext)
        Try
            'Declare variables
            Dim db As New DBManager()
            Dim answer As String = ""
            Dim responseCode As Integer = 202

            'Validation
            Dim vManager As New ValidationManager()
            vManager.Validate(context)

            'Check result
            If (vManager.ValidationResult = ValidationResult.Invalid) Then
                'Get the error response
                responseCode = vManager.ErrorHTTPCode
                answer = vManager.ErrorMessage
                Output.Report($"Validation failed with status: {vManager.ErrorHTTPCode}")
                db.InsertRejected(vManager.msgType, vManager.Content, vManager.ErrorMessage)
            Else
                'Convert message to string
                Dim rawText As String = vManager.Content

                'Parse the incoming msg
                Dim json As JObject = vManager.JSON
                Dim msgType As String = vManager.msgType
                Dim code As String = vManager.msgCode

                'Assemble respective answer
                'If message is coming from the secondary rep
                If vManager.MsgFromSecondary Then
                    'Assemble response that mimics the secondary repository standary response
                    Dim eventTime As Date = ParseTime(json("Event_Time"))
                    Dim checksum As String = rawText.ToMD5Hash()
                    answer = StandartResponse(code, msgType, checksum, Nothing, eventTime)

                    'Save the json in alternative table
                    If db.InsertRawJson("tbljsonsecondary", rawText, msgType.ToUpper()) Then
                        Output.ToConsole("New IRU was received from secondary repository and was sent to the Database")
                    End If
                Else
                    'Message is comming from the local machine
                    Select Case msgType.ToUpper()
                        Case "STA"
                            answer = ProcessSTA(code)
                        Case Else
                            'Create response
                            Dim checksum As String = rawText.ToMD5Hash()
                            answer = StandartResponse(code, msgType.ToUpper(), checksum, Nothing)

                            'Save the json in the db
                            If db.InsertRawJson(DBManager.TableName, rawText, msgType.ToUpper(), code) Then
                                Output.ToConsole("New Json sent to the Database")
                            End If
                    End Select
                End If

                'This is needed to return Warnings
                If vManager.Errors.Count > 0 Then
                    answer = vManager.ErrorMessage
                    responseCode = vManager.ErrorHTTPCode
                End If
            End If
            'Return a response
            context.Respond(answer, responseCode)
        Catch ex As Exception
            'If something fails, respond with error message and log the error
            Dim reason As New ValidationError() With {.Error_Code = "SYSTEM_ERROR", .Error_Descr = $"Failed to process incoming message, reason: {ex.Message}"}
            Dim response As String = StandartResponse(Nothing, Nothing, Nothing, New List(Of ValidationError) From {reason})
            Output.Report(reason.Error_Descr)
            Try
                context.Respond(response, 500) '500 = Internal server error
            Catch exx As Exception
                Output.Report($"Failed to respond. Error: {exx.Message}")
            End Try
        End Try
    End Sub

    Private Shared Function ProcessSTA(code As String) As String
        Dim db As New DBManager()

        'Get the response from database
        Dim result As DataTable = db.CheckForCode(code)

        Dim answer As String
        'If match is found
        If result.Rows.Count > 0 Then
            'Get response
            Dim response As String = If(IsDBNull(result.Rows(0)("fldResponse")), Nothing, result.Rows(0)("fldResponse"))

            'Validate response
            If response.IsNullOrEmpty() Then
                'Return error
                Throw New Exception("Code matched but response field is empty. Message might still be in queue.")
                Output.ToConsole($"STA message fail: {answer}")
            Else
                'Anwer with the response from the secondary
                answer = response
                Output.ToConsole($"Responding to STA, code: '{code}'.")
            End If
        Else 'No matches
            Throw New Exception($"No matching entities found for code: '{code}'")
            Output.ToConsole($"STA message fail: {answer}")
        End If

        Return answer
    End Function
End Class
