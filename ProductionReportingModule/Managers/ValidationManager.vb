Imports System.IO
Imports System.Net
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class ValidationManager
    Public Sub New()
        Errors = New List(Of ValidationError)
    End Sub

    Public Property Errors As List(Of ValidationError)
    Public Property ValidationResult As ValidationResult
    Public Property ErrorMessage As String
    Public Property ErrorHTTPCode As Integer

    Private Property Context As HttpListenerContext
    Private Property Content As String
    Private Property JSON As JObject
    Private Property msgType As String
    Private Property msgCode As String

    Public Sub Validate(message As HttpListenerContext)
        Context = message
        Content = New StreamReader(Context.Request.InputStream, Context.Request.ContentEncoding).ReadToEnd()

        ValidationResult = ValidationResult.Valid

        'Check Tokens / Credentials
        If (ValidationResult = VAL_SEC_TOKEN()) <> ValidationResult.Valid Then
            ErrorHTTPCode = 401
            ErrorMessage = StandartResponse(Nothing, Nothing, Nothing, Errors)
            Exit Sub
        End If

        'Check message integrity
        If (ValidationResult = VAL_SEC_HASH()) <> ValidationResult.Valid Then
            ErrorHTTPCode = 400
            ErrorMessage = StandartResponse(Nothing, Nothing, Nothing, Errors)
            Exit Sub
        End If

        If (ValidationResult = VAL_MSG_JSON()) <> ValidationResult.Valid Then
            ErrorHTTPCode = 400
            ErrorMessage = StandartResponse(Nothing, Nothing, Nothing, Errors)
            Exit Sub
        End If

        If (ValidationResult = VAL_FIE_MAN()) <> ValidationResult.Valid Then
            ErrorHTTPCode = 400
            ErrorMessage = StandartResponse(Nothing, Nothing, Nothing, Errors)
            Exit Sub
        End If

        If Not (ValidationResult = VAL_MSG_TYPE()) <> ValidationResult.Valid Then
            ErrorHTTPCode = 400
            ErrorMessage = StandartResponse(msgCode, msgType, CreateMD5(Content), Errors)

        End If

        If (ValidationResult = VAL_FIE_FORMAT()) <> ValidationResult.Valid Then
            ErrorHTTPCode = 400
            ErrorMessage = StandartResponse(msgCode, msgType, CreateMD5(Content), Errors)
        End If

        If (ValidationResult = VAL_MSG_CODE_DUPLICATE()) <> ValidationResult.Valid Then
            ErrorHTTPCode = 400
            ErrorMessage = StandartResponse(msgCode, msgType, CreateMD5(Content), Errors)
        End If

        If (ValidationResult = VAL_UI_MULT_MSG()) <> ValidationResult.Valid Then
            ErrorHTTPCode = 400
            ErrorMessage = StandartResponse(msgCode, msgType, CreateMD5(Content), Errors)
        End If

        If (ValidationResult = VAL_UI_EXIST_APP()) <> ValidationResult.Valid Then
            ErrorHTTPCode = 400
            ErrorMessage = StandartResponse(msgCode, msgType, CreateMD5(Content), Errors)
        End If

        If (ValidationResult = VAL_UI_DUPLICATE_APP()) <> ValidationResult.Valid Then
            ErrorHTTPCode = 400
            ErrorMessage = StandartResponse(msgCode, msgType, CreateMD5(Content), Errors)
        End If


    End Sub

#Region "Validations"
    Public Function VAL_SEC_HASH() As ValidationResult
        Dim hash As String = Context.Request.Headers("X-OriginalHash")
        Dim calculatedHash As String = CreateMD5(Content)
        If hash <> calculatedHash Then
            Dim newError As New ValidationError() With {.Error_Code = "INVALID_SIGNATURE",
                .Error_Descr = "Hash information not matching the message signature."}
            Errors.Add(newError)
            Return ValidationResult.Invalid
        End If
        Return ValidationResult.Valid
    End Function

    Public Function VAL_SEC_TOKEN() As ValidationResult
        Dim output As ValidationResult = ValidationResult.Valid
        Dim id As HttpListenerBasicIdentity = Context.User.Identity
        Dim hashPass As String = CreateMD5(id.Password)

        'If the user is not found or the password doesnt match
        If Not JsonListener.Users.Keys.Contains(id.Name) OrElse JsonListener.Users(id.Name) <> hashPass Then
            Dim newError As New ValidationError() With {.Error_Code = "INVALID_OR_EXPIRED_TOKEN", .Error_Descr = "Authentication failed"}
            Errors.Add(newError)
            output = ValidationResult.Invalid
        End If
        Return output
    End Function

    Public Function VAL_MSG_JSON() As ValidationResult
        'Check structure
        Try
            JSON = JObject.Parse(Content)
        Catch ex As Exception
            Dim newError As New ValidationError() With {.Error_Code = "INVALID_INPUT_FORMAT", .Error_Descr = "The body of the message doesn’t contain a valid JSON."}
            Errors.Add(newError)
            Return ValidationResult.Invalid
        End Try
        Return ValidationResult.Valid
    End Function

    Public Function VAL_MSG_TYPE() As ValidationResult
        Dim msgTypes() As String = {"STA", "REO", "REOD", "CEO", "DEO", "RFA", "RFAD", "CFA", "DFA", "RMA", "RMAD", "CMA", "DMA", "ICV", "ICM", "ULO", "ULOD", "PLO", "ISU", "IRU", "ISA", "IRA", "IDA", "EUA", "EPA", "EDP", "ERP", "ETL", "EUD", "EVR", "EIV", "EPO", "EPR", "RCL", "LUP", "LUQ", "CTM"}

        If Not msgType.Contains(msgType) Then
            Dim newError As New ValidationError() With {.Error_Code = "FAILED_VALIDATION", .Error_Descr = "Generic validation error.", .Error_Data = $"Invalid Message_Type: '{msgType}'."}
            Errors.Add(newError)
            Return ValidationResult.Invalid
        End If
        Return ValidationResult.Valid
    End Function

    Public Function VAL_FIE_MAN() As ValidationResult
        'Check required fields
        Try
            msgType = JSON("Message_Type")
            msgCode = JSON("Code")
            'TODO
            'check mandatory fields for current message type
        Catch ex As Exception
            Dim newError As New ValidationError() With {.Error_Code = "REQUIRED_FIELD_FAILED_VALIDATION", .Error_Descr = "Mandatory field is missing", .Error_Data = ex.Message}
            Errors.Add(newError)
            Return ValidationResult.Invalid
        End Try
        Return ValidationResult.Valid
    End Function

    Public Function VAL_FIE_FORMAT() As ValidationResult
        Return ValidationResult.Valid
    End Function

    Public Function VAL_MSG_CODE_DUPLICATE() As ValidationResult
        Dim db As New DBManager()
        Dim dtResult As DataTable = db.CheckForCode(msgCode)
        'Code was found
        If dtResult.Rows.Count > 0 Then
            Dim newError As New ValidationError() With {.Error_Code = "FAILED_VALIDATION", .Error_Descr = "Code already exists in the primary repository.", .Error_Data = msgCode}
            Errors.Add(newError)
            Return ValidationResult.Invalid
        End If
        Return ValidationResult.Valid
    End Function

    Public Function VAL_UI_MULT_MSG() As ValidationResult
        Dim output As Boolean = True
        Select Case msgType
            Case "IRU"
                'upUI
                output = Not CheckForDuplicates("upUI")
            Case "IDA"
                'Deact_upUI
                'Deact_aUI
                output = Not CheckForDuplicates("Deact_upUI")
                output = Not CheckForDuplicates("Deact_aUI")
            Case "EUA"
                'upUI_1
                'upUI_2
                output = Not CheckForDuplicates("upUI_1")
                output = Not CheckForDuplicates("upUI_2")
            Case "EPA"
                'parent code has to be checked too
                'Aggregated_UIs1
                'Aggregated_UIs2
                output = Not CheckForDuplicates("Aggregated_UIs1")
                output = Not CheckForDuplicates("Aggregated_UIs2")
            Case "EDP", "ERP", "ETL", "EVR", "EIV", "EPO", "EPR"
                'upUIs
                'aUIs
                output = Not CheckForDuplicates("upUIs")
                output = Not CheckForDuplicates("aUIs")
            Case "EUD"
                'only has 1 aUI ???
            Case Else
                output = True
        End Select
        Return If(output, ValidationResult.Valid, ValidationResult.Invalid)
    End Function

    Public Function VAL_UI_EXIST_APP() As ValidationResult
        Dim db As New DBManager()
        Select Case msgType
            Case "EUA"
                Dim codes As String() = JSON("upUI_2").ToObject(Of String())
                Dim result As DataTable = db.CheckCodesExistence(codes, "tblprimarycodes", "fldCode")

                'it should not be possible to get Count > Length
                If result.Rows.Count < codes.Length Then 'If there are unexisting codes
                    'Extract them
                    Dim existingCodes As String() = result.ColumnToArray("fldCode")
                    Dim missingCodes As String() = codes.Except(existingCodes)
                    'Create new error
                    Dim newError As New ValidationError() With {.Error_Code = "UIS_APPLICATION_ERROR", .Error_Descr = $"Some of the codes were not found in the primary repository.", .Error_Data = String.Join("#", missingCodes)}
                    Errors.Add(newError)
                    Return ValidationResult.Invalid
                End If
            Case "IDA"
                Dim uiType As Integer = JSON("Deact_Type").ToObject(Of Integer)
                Dim result As DataTable
                Dim codes As String()
                Dim columnName As String
                Dim tableName As String
                Dim tokenName As String

                Select Case uiType
                    Case 1 'UI level
                        columnName = "fldCode"
                        tableName = "tblprimarycodes"
                        tokenName = "Deact_upUI"
                    Case 2 'Aggregated level
                        columnName = "fldPrintCode"
                        tableName = "tblaggregatedcodes"
                        tokenName = "Deact_aUI"
                    Case Else
                        Throw New Exception($"Deact_Type invalid value: {uiType}")
                End Select

                codes = JSON(tokenName).ToObject(Of String())
                result = db.CheckCodesExistence(codes, tableName, columnName)

                'it should not be possible to get Count > Length
                If result.Rows.Count < codes.Length Then 'If there are unexisting codes
                    'Extract them
                    Dim existingCodes As String() = result.ColumnToArray(columnName)
                    Dim missingCodes As String() = codes.Except(existingCodes)
                    'Create new error
                    Dim newError As New ValidationError() With {.Error_Code = "UIS_APPLICATION_ERROR", .Error_Descr = $"Some of the codes were not found in the primary repository.", .Error_Data = String.Join("#", missingCodes)}
                    Errors.Add(newError)
                    Return ValidationResult.Invalid
                End If
                Return ValidationResult.Valid
            Case Else
                Return ValidationResult.Valid
        End Select
        Return True
    End Function

    Public Function VAL_UI_DUPLICATE_APP() As ValidationResult
        'Validation is only ment for EUA messages
        If msgType <> "EUA" Then Return ValidationResult.Valid
        Dim db As New DBManager()
        'Get the codes from the message
        Dim longCodes As String() = JSON("upUI_1").ToObject(Of String())
        Dim shortCodes As String() = JSON("upUI_2").ToObject(Of String())

        'Query that selects all of the codes from the list that already have a printed code
        Dim query As String = $"SELECT fldCode, fldPrintCode FROM `{DBBase.DBName}`.`tblprimarycodes` "
        query += $"WHERE fldCode in ('{String.Join("','", shortCodes)}') AND fldPrintCode IS NOT NULL;"
        Dim result = db.ReadDatabase(query)

        'If codes(s) exist already, fail validation
        If result.Rows.Count > 0 Then
            Dim errorCodes As String() = result.ColumnToArray("fldCode")
            'Create new error
            Dim newError As New ValidationError() With {
                .Error_Code = "UIS_APPLICATION_ERROR",
                .Error_Descr = $"Some upUI(s) have already been applied to a upUI(L)",
                .Error_Data = String.Join("#", errorCodes)}
            Errors.Add(newError)
            Return ValidationResult.Invalid
        End If

        'Else return valid
        Return ValidationResult.Valid
    End Function

    Public Function VAL_UI_FID_APP() As ValidationResult
        'Validation is only ment for EUA messages
        If msgType <> "EUA" Then Return ValidationResult.Valid
        'TODO
        'Check F_ID
        Return ValidationResult.Valid
    End Function

    Public Function VAL_UI_EXIST_UPUI() As ValidationResult
        Select Case msgType
            Case "EPA", "EDP", "ERP", "ETL", "EVR", "EIV", "EPO", "EPR"
                Return ValidationResult.Valid
            Case Else
                Return ValidationResult.Valid
        End Select
    End Function

    Public Function VAL_UI_EXIST_AUI() As ValidationResult
        Select Case msgType
            Case "IDA", "EPA", "EDP", "ERP", "ETL", "EVR", "EIV", "EPO", "EPR"
                Return ValidationResult.Valid
            Case Else
                Return ValidationResult.Valid
        End Select
    End Function

    Public Function VAL_UI_EXIST_UPUI_SEQ() As ValidationResult
        Throw New NotImplementedException("VAL_UI_EXIST_UPUI_SEQ not implemented.")
    End Function

    Public Function VAL_UI_EXIST_AUI_SEQ() As ValidationResult
        Throw New NotImplementedException("VAL_UI_EXIST_UPUI_SEQ not implemented.")
    End Function

    Public Function VAL_UI_EXPIRY() As ValidationResult
        Dim db As New DBManager()
        Select Case msgType
            Case "EUA"
                Dim codes As String() = JSON("upUI_2").ToObject(Of String()) 'fldCode
                'Check if codes are expired (6 months since issuing)
                Dim result As DataTable = db.CheckPrimaryCodesExpiration(codes)
                If result.Rows.Count > 0 Then
                    Dim errorCodes As String() = result.ColumnToArray("fldCode")
                    'Create new error
                    Dim newError As New ValidationError() With {
                        .Error_Code = "UI_EXPIRED",
                        .Error_Descr = $"Application or aggregation date exceeds the 6 months period after the code has been issued.",
                        .Error_Data = String.Join("#", errorCodes)}
                    Errors.Add(newError)
                    Return ValidationResult.Invalid
                End If
                Return ValidationResult.Valid
            Case "EPA"
                Dim aggType As Integer = JSON("Aggregation_Type").ToObject(Of Integer)
                Select Case aggType
                    Case 1 'unit level only
                        Dim codes As String() = JSON("Aggregated_UIs1").ToObject(Of String()) 'fldCode
                        'Check if codes are expired (6 months since issuing)
                        Dim result As DataTable = db.CheckPrimaryCodesExpiration(codes)
                        'If expired codes are found
                        If result.Rows.Count > 0 Then
                            Dim errorCodes As String() = result.ColumnToArray("fldCode")
                            'Create new error
                            Dim newError As New ValidationError() With {
                                .Error_Code = "UI_EXPIRED",
                                .Error_Descr = $"Application or aggregation date exceeds the 6 months period after the code has been issued.",
                                .Error_Data = String.Join("#", errorCodes)}
                            Errors.Add(newError)
                            Return ValidationResult.Invalid
                        End If
                    Case 2 'aggregated level only
                        Dim codes As String() = JSON("Aggregated_UIs2").ToObject(Of String()) 'fldCode
                        'Check if codes are expired (6 months since issuing)
                        Dim result As DataTable = db.CheckAggregatedCodesExpiration(codes)
                        'If expired codes are found
                        If result.Rows.Count > 0 Then
                            Dim errorCodes As String() = result.ColumnToArray("fldPrintCode")
                            'Create new error
                            Dim newError As New ValidationError() With {
                                .Error_Code = "UI_EXPIRED",
                                .Error_Descr = $"Application or aggregation date exceeds the 6 months period after the code has been issued.",
                                .Error_Data = String.Join("#", errorCodes)}
                            Errors.Add(newError)
                            Return ValidationResult.Invalid
                        End If
                    Case 3 'both (rare occasion)
                        Dim pCodes As String() = JSON("Aggregated_UIs1").ToObject(Of String())
                        Dim aCodes As String() = JSON("Aggregated_UIs2").ToObject(Of String())

                        'Check if codes are expired (6 months since issuing)
                        Dim result1 As DataTable = db.CheckPrimaryCodesExpiration(pCodes)
                        Dim result2 As DataTable = db.CheckAggregatedCodesExpiration(aCodes)
                        'If expired codes are found
                        If result1.Rows.Count > 0 OrElse result2.Rows.Count > 0 Then
                            Dim errorCodes As New List(Of String)
                            If result1.Rows.Count > 0 Then errorCodes.AddRange(result1.ColumnToArray("fldCode"))
                            If result2.Rows.Count > 0 Then errorCodes.AddRange(result2.ColumnToArray("fldPrintCode"))
                            'Create new error
                            Dim newError As New ValidationError() With {
                                .Error_Code = "UI_EXPIRED",
                                .Error_Descr = $"Application or aggregation date exceeds the 6 months period after the code has been issued.",
                                .Error_Data = String.Join("#", errorCodes)}
                            Errors.Add(newError)
                            Return ValidationResult.Invalid
                        End If
                    Case Else
                        Throw New Exception($"Unexpected value for Aggregation_Type: '{aggType}'")
                End Select
                Return ValidationResult.Valid
            Case Else
                Return ValidationResult.Valid
        End Select
    End Function
#End Region

    Private Function CheckForDuplicates(uis As String) As Boolean
        'upUI
        Dim upUIs As String() = JSON.Item(uis).ToObject(Of String())
        If upUIs.HasDuplicates() Then
            Dim duplicates As String() = upUIs.GroupBy(Function(s) s).SelectMany(Function(grp) grp.Skip(1)).Distinct()
            Dim newError As New ValidationError() With {.Error_Code = "MULTIPLE_UID", .Error_Descr = $"Multiple duplicate {uis} present in the messages", .Error_Data = String.Join("#", duplicates)}
            Errors.Add(newError)
            Return True
        End If
        Return False
    End Function
End Class

Public Enum ValidationResult
    Invalid
    Valid
    PassWithWarning
End Enum

Public Class ValidationError
    'Public Property Error_InternalID As Integer
    Public Property Error_Code As String
    Public Property Error_Descr As String
    Public Property Error_Data As String
End Class
