Imports Newtonsoft.Json.Linq
Imports ReportTools
Imports System.Threading

Public Class JsonSender
    Public Sub Start()
        Dim thrd = New Thread(AddressOf Work)
        thrd.Start()
    End Sub

    Private Sub Work()
        While Main.IsRunning
            'Check the database
            Dim db = New DBManager()
            Dim result As DataTable = db.CheckForNewRows()

            'If any new jsons are avaible
            If result.Rows.Count > 0 Then
                'Send to repository
                For Each row As DataRow In result.Rows
                    Try
                        Dim rawjson As String = row(DBManager.JsonColumn)
                        Dim json As JObject = JObject.Parse(rawjson)
                        Dim msgType As String = json.Item("Message_Type")
                        Dim sendURL As String

                        Select Case msgType.ToUpper()
                            Case "EPA", "ERP", "IDA", "EUA", "ETL", "EDP", "EUD", "EVR"
                                sendURL = JsonManager.OperationalURL
                            Case "EPR", "EPO", "EIV"
                                sendURL = JsonManager.TransactionalURL
                            Case "RCL"
                                sendURL = JsonManager.RecallURL
                            Case Else
                                Throw New Exception($"Invalid Message_Type column value: '{msgType}'")
                        End Select

                        'Set message time
                        json("Message_Time_long") = GetTimeLong()
                        'Post
                        Dim response As String = JsonManager.Post(json.ToString(), sendURL)
                        'Update the Rep coulumn in the DB
                        db.UpdateDatabase(row("fldIndex"), response)
                        Output.Report($"JSON object with id: {row("fldIndex")} sent to repository, updating database...")
                    Catch ex As Exception
                        Output.Report($"Post operation failed: {ex.Message}")
                    End Try
                Next
            End If
            db.Disconnect()
            'Sleep
            Thread.Sleep(TimeSpan.FromSeconds(DBManager.dbCheckTime))
        End While
    End Sub
End Class
