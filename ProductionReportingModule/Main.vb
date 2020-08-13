﻿Imports System.Threading
Imports ReportTools

Module Main
    Public tokenManager As TokenManager
    Public Property Sender As JsonSender
    Public Property IsRunning As Boolean
    Public Property senderEnabled As Boolean

    Sub Main()
        IsRunning = True

        Try
            Initialize()
        Catch ex As Exception
            Output.Report($"Failed to initialize: {ex.Message}")
            Console.ReadLine()
            Exit Sub
        End Try

        'TESTING GROUND
        'Dim strr As New List(Of String) From {"",""}
        'Dim strTest As String = JsonListener.StandartResponse("ddd", "ggg", "kkk", New List(Of ValidationError) From {New ValidationError() With {.Error_Code = "1", .Error_Descr = "2", .Error_Data = "3"}, New ValidationError() With {.Error_Code = "4", .Error_Descr = "5", .Error_Data = "6"}})
        'END OF TESTING GROUND

        'Start the token manager
        tokenManager.Start()

        Dim listener As New JsonListener()
        listener.Start()

        'Wait
        Thread.Sleep(3000)

        If senderEnabled Then
            Sender.Start()
        End If

        'Stay alive
        While Thread.CurrentThread.IsAlive
            Thread.Sleep(TimeSpan.FromMinutes(5))
        End While

        IsRunning = False
    End Sub

    Private Sub Initialize()
        Setting = New DataSet()
        Setting.ReadXml($"{AppDomain.CurrentDomain.BaseDirectory}Settings.xml")

        'Get the row with the JSON server setting
        Dim tblSecondary = Setting.Tables("tblJSONServer")(0)
        Dim tokenURL As String = tblSecondary("fldTokenAddress")
        Dim globalUrl As String = tblSecondary("fldGlobalURL")
        Dim acc As String = tblSecondary("fldUsername")
        Dim pass As String = tblSecondary("fldPassword")
        Dim authType As String = tblSecondary("fldAuthentication")
        senderEnabled = Convert.ToInt32(tblSecondary("fldEnabled"))

        'Get the facility information
        Dim tblFacility = Setting.Tables("tblFacility")(0)
        Dim fURL As String = tblFacility("fldGlobalURL")
        Dim f_Acc As String = tblFacility("fldUsername")
        Dim f_Pass As String = tblFacility("fldPassword")
        Dim f_AuthType As String = tblFacility("fldAuthentication")

        'Create the Token manager
        tokenManager = New TokenManager(tokenURL, acc, pass, GetAuthType(authType))

        'Instantiate the sender
        Sender = New JsonSender(globalUrl, fURL) With
            {
                .Secondary = New JsonManager(globalUrl, Nothing, Nothing, AuthenticationType.Bearer, tokenManager.AuthToken),
                .Facility = New JsonManager(fURL, f_Acc, f_Pass, GetAuthType(f_AuthType))
            }

        'Get the row with the DataBase settings and initialize
        Dim rowDBSetting = Setting.Tables("tblDBSettings")(0)
        DBBase.DBIP = rowDBSetting("fldServer")
        DBBase.DBPort = Convert.ToInt32(rowDBSetting("fldPort"))
        DBBase.DBName = rowDBSetting("fldDBName")
        DBBase.DBUser = rowDBSetting("fldAccount")
        DBBase.DBPass = rowDBSetting("fldPassword")
        DBManager.JsonColumn = rowDBSetting("fldColJson")
        DBManager.RepColumn = rowDBSetting("fldColRep")
        DBManager.TableName = rowDBSetting("fldTableName")

        'Get the general settings
        Dim generalSettings = Setting.Tables("tblGeneral")(0)
        JsonListener.Prefix = generalSettings("fldListenerPrefix")
        DBManager.dbCheckTime = Convert.ToInt32(generalSettings("fldDBCheckTime"))
    End Sub
End Module
