Imports MySql.Data.MySqlClient

Public Class DBBase

    Protected Sub New()
    End Sub

    Public Shared Property DBName As String
    Public Shared Property DBIP As String
    Public Shared Property DBUser As String
    Public Shared Property DBPass As String
    Public Shared Property DBPort As UInteger

    Protected conn As MySqlConnection
    Protected adapter As MySqlDataAdapter
    Protected cmd As MySqlCommand

    Public Sub Init()
        'Generate connection string
        Dim cBuilder As MySqlConnectionStringBuilder = New MySqlConnectionStringBuilder() With {
            .Server = DBIP,
            .UserID = DBUser,
            .Password = DBPass,
            .Port = DBPort,
            .SslMode = MySqlSslMode.None}

        'Instantiate necessary objects
        conn = New MySqlConnection(cBuilder.ConnectionString)
        cmd = New MySqlCommand() With {.Connection = conn}
        adapter = New MySqlDataAdapter()
    End Sub

#Region "Direct access"
    Public Function ReadDatabase(query As String) As DataTable
        cmd.CommandText = query
        adapter.SelectCommand = cmd
        Dim output As New DataTable

        Try
            conn.Open()
            adapter.Fill(output)
        Catch ex As Exception
            ReportTools.Output.Report($"Exception occured while reading from database: '{ex.Message}'")
        End Try

        Disconnect()
        Return output
    End Function

    Public Function Execute(query As String) As Boolean
        If query = String.Empty Then Return False
        Dim output As Boolean = False

        'Execute the query
        cmd.CommandText = query
        Try
            conn.Open()
            cmd.ExecuteNonQuery()
            output = True
        Catch ex As Exception
            ReportTools.Output.Report($"Exception occured while writing to Database: '{ex.Message}'; {Environment.NewLine}Query: {query}")
        End Try

        'Close connection and return the result
        Disconnect()
        Return output
    End Function

    Public Sub Disconnect()
        Try
            If conn.State <> ConnectionState.Closed Then conn.Close()
        Catch
        End Try
    End Sub
#End Region
End Class
