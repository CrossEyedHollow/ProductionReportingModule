Imports System.Threading
Imports ReportTools

Module Main
    Public tokenManager As TokenManager
    Public Property Sender As JsonSender
    Public Property IsRunning As Boolean

    Sub Main()
        IsRunning = True

        Try
            Initialize()
        Catch ex As Exception
            Output.Report($"Failed to initialize: {ex.Message}. Press any key to exit.")
            Console.ReadLine()
            Exit Sub
        End Try

        'TESTING GROUND
        'END OF TESTING GROUND
        Dim db As New DBManager()
        Dim codes As String() = {"03800001131720SYUD1J3YUI", "03800001131720SysqmUCcH4", "03800001131720Sfsu2ju9G9", "03800001131737B1d753Nd2s", "03800001131720SRI15W8aUW", "03800001131720SL5wUt0HMG", "03800001131720SD6cgumUtA"}
        Dim result As DataTable = db.SelectInvolvedEvents(codes, Tables.tblaggregatedcodes, "fldPrintCode")


        Dim strEvents As String = String.Join(",", result.ColumnToArray("EventList"))
        Dim arr As String() = strEvents.Split(",")
        'Remove duplicates
        Dim distinctEvents As HashSet(Of String) = New HashSet(Of String)(arr.Where(Function(x) Not String.IsNullOrEmpty(x)).ToArray())
        Dim tststr = $"'{String.Join("','", distinctEvents)}'"
        'END TEST

        'Start the token manager
        tokenManager.Start()

        Dim listener As New JsonListener()
        listener.Start()

        'Wait
        Thread.Sleep(3000)

        Sender.Start()

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
        Dim s_TokenURL As String = tblSecondary("fldTokenAddress")
        Dim s_GlobalUrl As String = tblSecondary("fldGlobalURL")
        Dim s_Acc As String = tblSecondary("fldUsername")
        Dim s_Pass As String = tblSecondary("fldPassword")
        Dim s_AuthType As String = tblSecondary("fldAuthentication")
        Dim s_Enabled = Convert.ToInt32(tblSecondary("fldEnabled"))

        'Get the facility information
        Dim tblFacility = Setting.Tables("tblFacility")(0)
        Dim f_URL As String = tblFacility("fldGlobalURL")
        Dim f_Acc As String = tblFacility("fldUsername")
        Dim f_Pass As String = tblFacility("fldPassword")
        Dim f_AuthType As String = tblFacility("fldAuthentication")
        Dim f_Enabled = Convert.ToInt32(tblFacility("fldEnabled"))

        'Create the Token manager
        tokenManager = New TokenManager(s_TokenURL, s_Acc, s_Pass, GetAuthType(s_AuthType))

        'Instantiate the sender
        Sender = New JsonSender(f_URL, s_GlobalUrl) With
            {
                .Secondary = New JsonManager(s_GlobalUrl, Nothing, Nothing, AuthenticationType.Bearer, tokenManager.AuthToken) With {.Enabled = s_Enabled},
                .Facility = New JsonManager(f_URL, f_Acc, f_Pass, GetAuthType(f_AuthType)) With {.Enabled = f_Enabled}
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
