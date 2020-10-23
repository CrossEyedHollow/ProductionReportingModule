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
        Dim query As String = $"SELECT * FROM `{DBName}`.`tbljson` WHERE fldLocalCode = '{code}';"
        Return ReadDatabase(query)
    End Function

    Public Function Check_aUI(code As String) As DataTable
        Dim query As String = $"SELECT * FROM `{DBName}`.`tblaggregatedcodes` where fldPrintCode = '{code}' AND fldEUD IS NOT NULL;"
        Return ReadDatabase(query)
    End Function

    Public Function CheckF_ID(id As String) As DataTable
        Dim query As String = $"SELECT fldF_ID FROM `{DBName}`.`tblfacility` WHERE fldF_ID = '{id}';"
        Return ReadDatabase(query)
    End Function

    Public Function GetAuthenticatedUsers() As DataTable
        Dim query As String = $"SELECT * FROM `{DBName}`.`tblusers`;"
        Return ReadDatabase(query)
    End Function

    Public Function SelectInvolvedEvents(codes As String(), table As Tables, codesColumn As String) As DataTable
        Dim query As String
        Select Case table
            Case Tables.tblprimarycodes
                query = $"SELECT T.*, CONCAT_WS(',',fldEUA,fldEPA,fldEDP,fldEIV,fldEPR,fldERP,fldIDA) as EventList FROM `{DBName}`.`{table}` as T WHERE {codesColumn} in ('{String.Join("','", codes)}') ORDER BY T.fldIndex;"
            Case Tables.tblaggregatedcodes
                query = $"SELECT T.*, CONCAT_WS(',',fldEPA,fldEDP,fldEIV,fldEPR,fldERP,fldEUD,fldIDA) as EventList FROM `{DBName}`.`{table}` as T WHERE {codesColumn} in ('{String.Join("','", codes)}') ORDER BY T.fldIndex;"
            Case Else
                Throw New Exception($"Invalid argument for table: '{table}'")
        End Select
        Return ReadDatabase(query)
    End Function

    Public Function SelectMessagesOlderThan(msgDate As Date, msgCodes As HashSet(Of String))
        Dim strCodes = $"'{String.Join("','", msgCodes)}'"
        Dim query As String = ""
        query += $"SELECT fldIndex, fldJson, fldDate, 'tbljson' AS `Table` FROM `{DBName}`.tbljson WHERE fldDate> '{msgDate.ToString(DateTimeFormat)}' AND fldLocalCode in ({strCodes}) "
        query += "UNION "
        query += $"SELECT fldIndex, fldJson, fldDate, 'tbljsonsecondary' AS `Table` FROM `{DBName}`.tbljsonsecondary WHERE fldDate > '{msgDate.ToString(DateTimeFormat)}' AND fldLocalCode in ({strCodes});"
        Return ReadDatabase(query)
    End Function

    Public Function InsertRawJson(table As String, Json As String, type As String, guid As String) As Boolean
        'Generate the query
        Dim query As String = AssembleInsertRawJsonQuery(table, Json, type, guid)
        'Execute it
        Return Execute(query)
    End Function

    Public Sub InsertRejected(type As String, json As String, response As String)
        Dim query As String = $"INSERT INTO `{DBName}`.`tblrejected` (fldType, fldJson, fldRejectReason) VALUES ('{type}','{json}','{response}')".Replace("''", "'-'")
        Execute(query)
    End Sub

    Public Function InsertRawJson(table As String, Json As String, type As String) As Boolean
        'Generate the query
        Dim query As String = AssembleInsertRawJsonQuery(table, Json, type)
        'Execute it
        Return Execute(query)
    End Function

    Public Function CheckForNewRows() As DataTable
        Dim query As String = $"SELECT * FROM `{DBName}`.`{TableName}` WHERE {RepColumn} IS NULL;"
        Return ReadDatabase(query)
    End Function

    Public Function CheckForNewRows(table As String) As DataTable
        Dim query As String = $"SELECT * FROM `{DBName}`.`{table}` WHERE fldRep IS NULL;"
        Return ReadDatabase(query)
    End Function

    Public Function CheckForDeactivated(tableName As String, codes As String(), codesColumnName As String)
        Dim query = $"SELECT * FROM `{DBName}`.`{tableName}` WHERE fldIDA IS NOT NULL AND {codesColumnName} IN ('{String.Join("','", codes)}');"
        Return ReadDatabase(query)
    End Function

    Public Function CheckForDeaggregated(codes As String()) As DataTable
        Dim query = ""
        query += "SELECT A.*, J.fldDate AS fldEUDDate "
        query += $"FROM {DBName}.tblaggregatedcodes AS A "
        query += $"LEFT JOIN (`{DBName}`.tbljson AS J) "
        query += "ON A.fldEUD = J.fldLocalCode "
        query += "WHERE J.fldDate > A.fldAggregatedDate "
        query += "AND A.fldEUD IS NOT NULL "
        query += $"AND A.fldPrintCode IN ('{String.Join("','", codes)}');"
        Return ReadDatabase(query)
    End Function

    Public Sub UpdateDatabase(index As Integer, response As String)
        Dim query As String = AssembleUpdateRepDateQuery(index, response)
        Execute(query)
    End Sub

    Public Sub UpdateFldRep(index As Integer, table As String)
        Dim query As String = $"UPDATE `{DBName}`.`{table}` SET {RepColumn} = NOW() WHERE fldIndex = {index};"
        Execute(query)
    End Sub

    Public Function CheckCodesExistence(codes As String(), table As String, column As String) As DataTable
        Dim query As String = $"SELECT * from `{DBName}`.`{table}` WHERE {column} in ('{String.Join("','", codes)}')"
        Return ReadDatabase(query)
    End Function

    Public Function CheckAUIExistence(codes As String()) As DataTable
        Dim query As String = $"SELECT DISTINCT fldParentCode FROM tblprimarycodes WHERE fldParentCode IN ('{String.Join("','", codes)}');"
        Return ReadDatabase(query)
    End Function

    Public Function CheckApliedCodes(codeList As String()) As DataTable
        'Query that selects all of the codes from the list that already have a printed code
        Dim query As String = $"SELECT fldCode, fldPrintCode FROM `{DBBase.DBName}`.`tblprimarycodes` "
        query += $"WHERE fldCode in ('{String.Join("','", codeList)}') AND fldPrintCode IS NOT NULL;"
        Return ReadDatabase(query)
    End Function

    'Public Function CheckPrintCodes(codeList As String()) As DataTable
    '    ''Query that selects all of the codes from the list that already have a printed code
    '    'Dim query As String = $"SELECT fldCode, fldPrintCode FROM `{DBBase.DBName}`.`tblprimarycodes` "
    '    'query += $"WHERE fldPrintCode in ('{String.Join("','", codeList)}');"
    '    Return ReadDatabase("")
    'End Function

    Public Function CheckPrimaryCodesExpiration(codes As String()) As DataTable
        Dim query As String = $"SELECT * FROM `{DBName}`.`tblprimarycodes` WHERE fldCode in ('{String.Join("','", codes)}') AND fldIssueDate < NOW() - INTERVAL 6 MONTH;"
        Return ReadDatabase(query)
    End Function

    Public Function CheckAggregatedCodesExpiration(codes As String()) As DataTable
        Dim query As String = $"SELECT * FROM `{DBName}`.`tblaggregatedcodes` WHERE fldPrintCode in ('{String.Join("','", codes)}') AND fldPrintDate < NOW() - INTERVAL 6 MONTH;"
        Return ReadDatabase(query)
    End Function

    Public Function CheckCodeLocation(code As String, msgF_ID As String) As DataTable
        Dim query As String = $"SELECT * FROM `{DBName}`.`tblaggregatedcodes` WHERE fldCode = '{code}' AND fldLocation <> '{msgF_ID}';"
        Return ReadDatabase(query)
    End Function
    Public Function CheckCodeLocation(codes() As String, msgF_ID As String, table As String) As DataTable
        Dim query As String = $"SELECT * FROM `{DBName}`.`{table}` WHERE fldCode in ('{String.Join("','", codes)}') AND fldLocation <> '{msgF_ID}';"
        Return ReadDatabase(query)
    End Function

    Public Function SearchUisInJSON(uis As String(), table As String, msgType As String, jField As String) As DataTable
        Dim query As String = $"SELECT jt.* FROM `{DBName}`.{table}, JSON_TABLE(fldJson, '${jField}[*]' COLUMNS (Codes VARCHAR(34) PATH '$')) AS jt "
        query += $"WHERE fldType = '{msgType}' AND Codes IN ('{String.Join("','", uis)}');"
        Return ReadDatabase(query)
    End Function

    Public Function SearchCodeInJSON(ui As String, table As String, msgType As String, jField As String)
        Dim query As String = $"SELECT fldIndex, fldType, fldJson->>'${jField}' AS Code FROM tbljson WHERE fldType = '{msgType}' AND fldJson->>'${jField}' = '{ui}'"
        Return ReadDatabase(query)
    End Function
#Region "Queries"
    Private Function AssembleInsertRawJsonQuery(table As String, Json As String, type As String, guid As String) As String
        Return $"INSERT INTO `{DBName}`.`{table}` ({JsonColumn}, fldLocalCode, fldType) VALUES ('{Json}', '{guid}', '{type}');"
    End Function

    Private Function AssembleInsertRawJsonQuery(table As String, Json As String, type As String) As String
        Return $"INSERT INTO `{DBName}`.`{table}` ({JsonColumn}, fldType) VALUES ('{Json}', '{type}');"
    End Function

    Private Function AssembleUpdateRepDateQuery(index As Integer, response As String)
        Return $"UPDATE `{DBName}`.`{TableName}` SET {RepColumn} = NOW(), fldResponse = '{response.Replace("'", "\'")}' WHERE fldIndex = {index};"
    End Function

#End Region
End Class
