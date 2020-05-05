Imports ReportTools

Public Class DBManager
    Inherits DBBase

    Public Sub New()
        Init()
    End Sub

    Shared Property JsonColumn As String
    Shared Property RepColumn As String
    Shared Property TableName As String
    Shared Property dbCheckTime As Integer = 5

    Public Function CheckForCode(code As String) As DataTable
        Dim query As String = SelectJsonWithCodeQuery(code)
        Return ReadDatabase(query)
    End Function

    Private Function SelectJsonWithCodeQuery(code As String) As String
        Return $"SELECT * FROM `{DBName}`.`tbljson` WHERE fldLocalCode = '{code}';"
    End Function

    Public Function GetAuthenticatedUsers() As DataTable
        Dim query As String = $"SELECT * FROM `{DBName}`.`tblusers`;"
        Return ReadDatabase(query)
    End Function

    Public Function InsertRawJson(table As String, Json As String, type As String, guid As String) As Boolean
        'Generate the query
        Dim query As String = AssembleInsertRawJsonQuery(table, Json, type, guid)
        'Execute it
        Return Execute(query)
    End Function

    Public Function InsertRawJson(table As String, Json As String, type As String) As Boolean
        'Generate the query
        Dim query As String = AssembleInsertRawJsonQuery(table, Json, type)
        'Execute it
        Return Execute(query)
    End Function

    Public Function CheckForNewRows() As DataTable
        Dim query As String = AssembleCheckStatement()
        Return ReadDatabase(query)
    End Function

    Public Sub UpdateDatabase(index As Integer, response As String)
        Dim query As String = AssembleUpdateRepDateQuery(index, response)
        Execute(query)
    End Sub

#Region "Queries"
    Private Function AssembleInsertRawJsonQuery(table As String, Json As String, type As String, guid As String) As String
        Return $"INSERT INTO `{DBName}`.`{table}` ({JsonColumn}, fldLocalCode, fldType) VALUES ('{Json}', '{guid}', '{type}');"
    End Function

    Private Function AssembleInsertRawJsonQuery(table As String, Json As String, type As String) As String
        Return $"INSERT INTO `{DBName}`.`{table}` ({JsonColumn}, fldType) VALUES ('{Json}', '{type}');"
    End Function

    Private Function AssembleCheckStatement() As String
        Return $"SELECT * FROM `{DBName}`.`{TableName}` WHERE {RepColumn} IS NULL;"
    End Function

    Private Function AssembleUpdateRepDateQuery(index As Integer, response As String)
        Return $"UPDATE `{DBName}`.`{TableName}` SET {RepColumn} = NOW(), fldResponse = '{response}' WHERE fldIndex = {index};"
    End Function

#End Region

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
