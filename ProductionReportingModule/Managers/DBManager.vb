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

    Public Function CheckCodesExistence(codes As String(), table As String, column As String) As DataTable
        Dim query As String = $"SELECT * from `{DBName}`.`{table}` WHERE {column} in ('{String.Join("','", codes)}')"
        Return ReadDatabase(query)
    End Function

    Public Function CheckPrimaryCodesExpiration(codes As String()) As DataTable
        Dim query As String = $"SELECT * FROM `{DBName}`.`tblprimarycodes` WHERE fldCode in ('{String.Join("','", codes)}') AND fldIssueDate < NOW() - INTERVAL 6 MONTH;"
        Return ReadDatabase(query)
    End Function

    Public Function CheckAggregatedCodesExpiration(codes As String()) As DataTable
        Dim query As String = $"SELECT * FROM `{DBName}`.`tblaggregatedcodes` WHERE fldPrintCode in ('{String.Join("','", codes)}') AND fldPrintDate < NOW() - INTERVAL 6 MONTH;"
        Return ReadDatabase(query)
    End Function
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
End Class
