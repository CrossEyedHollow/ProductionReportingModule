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
    Public Property Context As HttpListenerContext
    Public Property Content As String
    Public Property JSON As JObject
    Public Property MsgFromSecondary As Boolean

    Public msgType As String
    Public msgCode As String
    Public Property Sender As User

    Public Sub Validate(message As HttpListenerContext)
        Context = message
        Content = New StreamReader(Context.Request.InputStream, Context.Request.ContentEncoding).ReadToEnd()

        Dim currentResult As ValidationResult
        ValidationResult = ValidationResult.Valid

        'Check Tokens / Credentials
        currentResult = VAL_SEC_TOKEN()
        If currentResult <> ValidationResult.Valid Then
            ErrorHTTPCode = 401
            ValidationResult = currentResult
            ErrorMessage = JsonManager.StandartResponse(Nothing, Nothing, Nothing, Errors)
            Exit Sub
        End If

        'Check message integrity
        currentResult = VAL_SEC_HASH()
        If currentResult <> ValidationResult.Valid Then
            ErrorHTTPCode = 400
            ValidationResult = currentResult
            ErrorMessage = JsonManager.StandartResponse(Nothing, Nothing, Nothing, Errors)
            Exit Sub
        End If

        currentResult = VAL_MSG_JSON()
        If currentResult <> ValidationResult.Valid Then
            ErrorHTTPCode = 400
            ValidationResult = currentResult
            ErrorMessage = JsonManager.StandartResponse(Nothing, Nothing, Nothing, Errors)
            Exit Sub
        End If

        currentResult = VAL_FIE_MAN()
        If currentResult <> ValidationResult.Valid Then
            ErrorHTTPCode = 400
            ValidationResult = currentResult
            ErrorMessage = JsonManager.StandartResponse(Nothing, Nothing, Nothing, Errors)
            Exit Sub
        Else 'Save the variables for future use
            msgCode = JSON("Code")
            msgType = JSON("Message_Type")
            'If the message comes from the secondary repository or it's a status message, don't validate
            If Sender.IsSecondary OrElse msgType = "STA" Then Exit Sub
        End If

        currentResult = VAL_EVT_24H()
        If currentResult <> ValidationResult.Valid Then
            ErrorHTTPCode = 299 'Warning
            ValidationResult = currentResult
            ErrorMessage = JsonManager.StandartResponse(msgCode, msgType, ToMD5Hash(Content), Errors)
        End If

        currentResult = VAL_EVT_TIME()
        If currentResult <> ValidationResult.Valid Then
            ErrorHTTPCode = 299 'Warning
            ValidationResult = currentResult
            ErrorMessage = JsonManager.StandartResponse(msgCode, msgType, ToMD5Hash(Content), Errors)
        End If

        currentResult = VAL_MSG_TYPE()
        If currentResult <> ValidationResult.Valid Then
            ErrorHTTPCode = 400
            ValidationResult = currentResult
            ErrorMessage = JsonManager.StandartResponse(msgCode, msgType, ToMD5Hash(Content), Errors)
        End If

        currentResult = VAL_FIE_FORMAT()
        If currentResult <> ValidationResult.Valid Then
            ErrorHTTPCode = 400
            ValidationResult = currentResult
            ErrorMessage = JsonManager.StandartResponse(msgCode, msgType, ToMD5Hash(Content), Errors)
        End If

        currentResult = VAL_MSG_CODE_DUPLICATE()
        If currentResult <> ValidationResult.Valid Then
            ErrorHTTPCode = 400
            ValidationResult = currentResult
            ErrorMessage = JsonManager.StandartResponse(msgCode, msgType, ToMD5Hash(Content), Errors)
        End If

        currentResult = VAL_UI_MULT_MSG()
        If currentResult <> ValidationResult.Valid Then
            ErrorHTTPCode = 400
            ValidationResult = currentResult
            ErrorMessage = JsonManager.StandartResponse(msgCode, msgType, ToMD5Hash(Content), Errors)
        End If

        currentResult = VAL_UI_EXIST_APP()
        If currentResult <> ValidationResult.Valid Then
            ErrorHTTPCode = 400
            ValidationResult = currentResult
            ErrorMessage = JsonManager.StandartResponse(msgCode, msgType, ToMD5Hash(Content), Errors)
        End If

        currentResult = VAL_UI_DUPLICATE_APP()
        If currentResult <> ValidationResult.Valid Then
            ErrorHTTPCode = 400
            ValidationResult = currentResult
            ErrorMessage = JsonManager.StandartResponse(msgCode, msgType, ToMD5Hash(Content), Errors)
        End If

        currentResult = VAL_UI_FID_APP()
        If currentResult <> ValidationResult.Valid Then
            ErrorHTTPCode = 400
            ValidationResult = currentResult
            ErrorMessage = JsonManager.StandartResponse(msgCode, msgType, ToMD5Hash(Content), Errors)
        End If

        currentResult = VAL_UI_EXIST_UPUI()
        If currentResult <> ValidationResult.Valid Then
            ErrorHTTPCode = 400
            ValidationResult = currentResult
            ErrorMessage = JsonManager.StandartResponse(msgCode, msgType, ToMD5Hash(Content), Errors)
        End If

        currentResult = VAL_UI_EXIST_AUI()
        If currentResult <> ValidationResult.Valid Then
            ErrorHTTPCode = 400
            ValidationResult = currentResult
            ErrorMessage = JsonManager.StandartResponse(msgCode, msgType, ToMD5Hash(Content), Errors)
        End If

        currentResult = VAL_UI_EXIST_UPUI_SEQ()
        If currentResult <> ValidationResult.Valid Then
            ErrorHTTPCode = 400
            ValidationResult = currentResult
            ErrorMessage = JsonManager.StandartResponse(msgCode, msgType, ToMD5Hash(Content), Errors)
        End If

        currentResult = VAL_UI_EXIST_AUI_SEQ()
        If currentResult <> ValidationResult.Valid Then
            ErrorHTTPCode = 400
            ValidationResult = currentResult
            ErrorMessage = JsonManager.StandartResponse(msgCode, msgType, ToMD5Hash(Content), Errors)
        End If

        currentResult = VAL_UI_EXPIRY()
        If currentResult <> ValidationResult.Valid Then
            ErrorHTTPCode = 400
            ValidationResult = currentResult
            ErrorMessage = JsonManager.StandartResponse(msgCode, msgType, ToMD5Hash(Content), Errors)
        End If

        currentResult = VAL_UI_ORD_REACTIVATION()
        If currentResult <> ValidationResult.Valid Then
            ErrorHTTPCode = 400
            ValidationResult = currentResult
            ErrorMessage = JsonManager.StandartResponse(msgCode, msgType, ToMD5Hash(Content), Errors)
        End If

        currentResult = VAL_UI_ORD_DEACTIVATED()
        If currentResult <> ValidationResult.Valid Then
            ErrorHTTPCode = 400
            ValidationResult = currentResult
            ErrorMessage = JsonManager.StandartResponse(msgCode, msgType, ToMD5Hash(Content), Errors)
        End If

        currentResult = VAL_UI_ORD_AGG_MULT()
        If currentResult <> ValidationResult.Valid Then
            ErrorHTTPCode = 400
            ValidationResult = currentResult
            ErrorMessage = JsonManager.StandartResponse(msgCode, msgType, ToMD5Hash(Content), Errors)
        End If

        currentResult = VAL_UI_ORD_DISAGG()
        If currentResult <> ValidationResult.Valid Then
            ErrorHTTPCode = 400
            ValidationResult = currentResult
            ErrorMessage = JsonManager.StandartResponse(msgCode, msgType, ToMD5Hash(Content), Errors)
        End If
        'Untested
        currentResult = VAL_UI_ORD_AGG_FID()
        If currentResult <> ValidationResult.Valid Then
            ErrorHTTPCode = 400
            ValidationResult = currentResult
            ErrorMessage = JsonManager.StandartResponse(msgCode, msgType, ToMD5Hash(Content), Errors)
        End If

        currentResult = VAL_UI_ORD_DISPATCH()
        If currentResult <> ValidationResult.Valid Then
            ErrorHTTPCode = 400
            ValidationResult = currentResult
            ErrorMessage = JsonManager.StandartResponse(msgCode, msgType, ToMD5Hash(Content), Errors)
        End If

        currentResult = VAL_UI_ORD_ARRIVAL()
        If currentResult <> ValidationResult.Valid Then
            ErrorHTTPCode = 400
            ValidationResult = currentResult
            ErrorMessage = JsonManager.StandartResponse(msgCode, msgType, ToMD5Hash(Content), Errors)
        End If

        currentResult = VAL_UI_ORD_ARRIVAL_RETURN()
        If currentResult <> ValidationResult.Valid Then
            ErrorHTTPCode = 400
            ValidationResult = currentResult
            ErrorMessage = JsonManager.StandartResponse(msgCode, msgType, ToMD5Hash(Content), Errors)
        End If

        currentResult = VAL_RECALL_EXIST()
        If currentResult <> ValidationResult.Valid Then
            ErrorHTTPCode = 400
            ValidationResult = currentResult
            ErrorMessage = JsonManager.StandartResponse(msgCode, msgType, ToMD5Hash(Content), Errors)
        End If

        currentResult = VAL_RECALL_LAST()
        If currentResult <> ValidationResult.Valid Then
            ErrorHTTPCode = 400
            ValidationResult = currentResult
            ErrorMessage = JsonManager.StandartResponse(msgCode, msgType, ToMD5Hash(Content), Errors)
        End If
    End Sub

#Region "Validations"
    Public Function VAL_SEC_HASH() As ValidationResult
        Dim hash As String = Context.Request.Headers("X-OriginalHash")
        Dim calculatedHash As String = Content.ToMD5Hash()
        If hash <> calculatedHash Then
            Dim newError As New ValidationError() With {.Error_Code = "INVALID_SIGNATURE",
                .Error_Descr = "Hash information not matching the message signature."}
            Errors.Add(newError)
            Return ValidationResult.Invalid
        End If
        Return ValidationResult.Valid
    End Function

    Public Function VAL_SEC_TOKEN() As ValidationResult
        Dim id As HttpListenerBasicIdentity = Context.User.Identity
        Dim hashPass As String = ToMD5Hash(id.Password)

        Sender = JsonListener.Users.FirstOrDefault(Function(x) x.Name = id.Name)

        'If the user is not found or the password doesnt match
        If Sender Is Nothing OrElse Sender.Password <> hashPass Then
            ReportTools.Output.Report($"Bad user or password, user: '{id.Name}', pass: '{id.Password}'.")
            Dim newError As New ValidationError() With {.Error_Code = "INVALID_OR_EXPIRED_TOKEN", .Error_Descr = "Authentication failed"}
            Errors.Add(newError)
            Return ValidationResult.Invalid
        End If

        MsgFromSecondary = Sender.IsSecondary
        Return ValidationResult.Valid
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
        Dim output As ValidationResult = ValidationResult.Valid
        Dim mandatoryFields As String() = New String() {"Code", "Message_Type"}

        'Check fields
        For Each field As String In mandatoryFields
            'If the field is missing generate error
            If Not JSON.ContainsKey(field) Then
                Dim newError As New ValidationError() With {.Error_Code = "REQUIRED_FIELD_FAILED_VALIDATION", .Error_Descr = "Mandatory field is missing", .Error_Data = "Message_Type"}
                Errors.Add(newError)
                output = ValidationResult.Invalid
            End If
        Next
        Return output
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
        Dim output As Boolean
        Select Case msgType
            Case "IRU"
                'upUI
                output = Not CheckForDuplicates("upUI")
            Case "IDA"
                'Deact_upUI
                'Deact_aUI
                Dim aggType As Integer = JSON("Deact_Type").ToObject(Of Integer)
                Select Case aggType
                    Case 1 'ui only
                        output = Not CheckForDuplicates("Deact_upUI")
                    Case 2 'Aggregated only
                        output = Not CheckForDuplicates("Deact_aUI")
                    Case Else
                        Throw New Exception($"Unexpected value for Deact_Type: {aggType}")
                End Select
            Case "EUA"
                'upUI_1
                'upUI_2
                output = Not (CheckForDuplicates("upUI_1") OrElse CheckForDuplicates("upUI_2"))
            Case "EPA"
                'parent code has to be checked too
                'Aggregated_UIs1
                'Aggregated_UIs2
                Dim aggType As Integer = JSON("Aggregation_Type").ToObject(Of Integer)
                Select Case aggType
                    Case 1 'ui only
                        output = Not CheckForDuplicates("Aggregated_UIs1")
                    Case 2 'Aggregated only
                        output = Not CheckForDuplicates("Aggregated_UIs2")
                    Case 3 'Both
                        output = Not (CheckForDuplicates("Aggregated_UIs1") OrElse CheckForDuplicates("Aggregated_UIs2"))
                    Case Else
                        Throw New Exception($"Unexpected value for Aggregation_Type: {aggType}")
                End Select

            Case "EDP", "ERP", "ETL", "EVR", "EIV", "EPO", "EPR"
                'upUIs
                'aUIs
                Dim aggType As Integer = JSON("UI_Type").ToObject(Of Integer)
                Select Case aggType
                    Case 1 'ui only
                        output = Not CheckForDuplicates("upUIs")
                    Case 2 'Aggregated only
                        output = Not CheckForDuplicates("aUIs")
                    Case 3 'Both
                        output = (Not CheckForDuplicates("upUIs")) And (Not CheckForDuplicates("aUIs"))
                    Case Else
                        Throw New Exception($"Unexpected value for UI_Type: {aggType}")
                End Select
            Case "EUD"
                'only has 1 aUI ???
                output = True
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
                    ' Extract them
                    Dim existingCodes As String() = result.ColumnToArray("fldCode")
                    Dim missingCodes As String() = codes.Except(existingCodes).ToArray()
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
            Case Else
                Return ValidationResult.Valid
        End Select
        Return ValidationResult.Valid
    End Function

    Public Function VAL_UI_DUPLICATE_APP() As ValidationResult
        'Validation is only ment for EUA messages
        If msgType <> "EUA" Then Return ValidationResult.Valid
        Dim db As New DBManager()
        'Get the codes from the message
        Dim longCodes As String() = JSON("upUI_1").ToObject(Of String())
        Dim shortCodes As String() = JSON("upUI_2").ToObject(Of String())


        Dim result = db.CheckApliedCodes(shortCodes)

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
        Dim db As New DBManager()
        'Validation is only ment for EUA messages
        If msgType <> "EUA" Then Return ValidationResult.Valid
        'Check F_ID
        Dim id As String = JSON("F_ID")
        Dim result = db.CheckF_ID(id)

        If result.Rows.Count < 1 Then
            'Generate error
            Dim newError As New ValidationError() With {.Error_Code = "FID_MISMATCH", .Error_Descr = "F_ID not found in the primary repository."}
            Errors.Add(newError)
            Return ValidationResult.Invalid
        Else Return ValidationResult.Valid
        End If
    End Function

    Public Function VAL_UI_EXIST_UPUI() As ValidationResult
        Dim db As New DBManager()
        Dim jAggColumn As String
        Dim jCodesColumn As String
        Dim codeList As String()

        'Make adjusments based on the message type
        Select Case msgType
            Case "EIV", "EPR", "EDP", "ERP", "ETL", "EVR", "EPO"
                jAggColumn = "UI_Type"
                jCodesColumn = "upUIs"
            Case "EPA"
                jAggColumn = "Aggregation_Type"
                jCodesColumn = "Aggregated_UIs1"
            Case Else
                Return ValidationResult.Valid
        End Select

        'Get the codes from the right column
        Dim aggType As Integer = JSON(jAggColumn).ToObject(Of Integer)
        Select Case aggType
            Case 1, 3 'Unit level or Both
                codeList = JSON.Item(jCodesColumn).ToObject(Of String())
            Case Else 'This validation only checks unit level uis
                Return ValidationResult.Valid
        End Select

        'Check db
        Dim result As DataTable = db.CheckCodesExistence(codeList, Tables.tblprimarycodes.ToString(), "fldPrintCode")

        'If some of the codes weren't found
        If result.Rows.Count <> codeList.Length Then
            'Generate error
            Dim errCodes As String() = If(result.Rows.Count < 1, codeList, codeList.Except(result.ColumnToArray("fldPrintCode")).ToArray())

            'Create new error
            Dim newError As New ValidationError() With {
                .Error_Code = "UI_NOT_EXIST",
                .Error_Descr = $"Some of the UIs were not found in the primary repository.",
                .Error_Data = String.Join("#", errCodes)}
            Errors.Add(newError)
            Return ValidationResult.Invalid
        Else
            Return ValidationResult.Valid
        End If
    End Function

    Public Function VAL_UI_EXIST_AUI() As ValidationResult
        Dim db As New DBManager()
        Dim jAggColumn As String
        Dim jCodesColumn As String
        Dim codeList As String()

        'Make adjusments based on the message type
        Select Case msgType
            Case "IDA"
                jAggColumn = "Deact_Type"
                jCodesColumn = "Deact_aUI"
            Case "EPA"
                jAggColumn = "Aggregation_Type"
                jCodesColumn = "Aggregated_UIs2"
            Case "EDP", "ERP", "ETL", "EVR", "EIV", "EPO", "EPR"
                jAggColumn = "UI_Type"
                jCodesColumn = "aUIs"
            Case Else
                Return ValidationResult.Valid
        End Select

        'Get the codes from the right column
        Dim aggType As Integer = JSON(jAggColumn).ToObject(Of Integer)
        Select Case aggType
            Case 2, 3 'aggregated level or Both
                codeList = JSON.Item(jCodesColumn).ToObject(Of String())
            Case Else 'This validation only checks aggregated level uis
                Return ValidationResult.Valid
        End Select

        'Check db
        'Dim result As DataTable = db.CheckAUIExistence(codeList)
        Dim result As DataTable = db.CheckCodesExistence(codeList, Tables.tblaggregatedcodes.ToString(), "fldPrintCode")

        'If some of the codes weren't found
        If result.Rows.Count <> codeList.Length Then
            'Generate error
            Dim errCodes As String() = If(result.Rows.Count < 1, codeList, codeList.Except(result.ColumnToArray("fldPrintCode")).ToArray())
            'Dim errCodes As String() = codeList.Except(result.ColumnToArray("fldParentCode"))

            'Create new error
            Dim newError As New ValidationError() With {
                .Error_Code = "UI_NOT_EXIST",
                .Error_Descr = $"Some of the UIs were not found in the primary repository.",
                .Error_Data = String.Join("#", errCodes)}
            Errors.Add(newError)
            Return ValidationResult.Invalid
        Else
            Return ValidationResult.Valid
        End If
    End Function

    Public Function VAL_UI_EXIST_UPUI_SEQ() As ValidationResult
        Dim db As New DBManager()

        Dim jAggColumn As String
        Dim jCodesColumn As String
        Dim codeList As String()

        Select Case msgType
            Case "EPA" 'Children ui only
                jAggColumn = "Aggregation_Type"
                jCodesColumn = "Aggregated_UIs1"
            Case "EDP", "ERP", "ETL", "EVR"
                jAggColumn = "UI_Type"
                jCodesColumn = "upUIs"
            Case Else
                Return ValidationResult.Valid
        End Select

        'Get the codes from the right column
        Dim aggType As Integer = JSON(jAggColumn).ToObject(Of Integer)

        Select Case aggType
            Case 1, 3 'Unit level or Both
                codeList = JSON.Item(jCodesColumn).ToObject(Of String())
            Case Else 'This validation only checks unit level uis
                Return ValidationResult.Valid
        End Select

        'Check db
        Dim result As DataTable = db.CheckForDeactivated(Tables.tblprimarycodes.ToString(), codeList, "fldPrintCode")

        'If there are any deactivated
        If result.Rows.Count > 0 Then
            'Generate error
            Dim errCodes As String() = result.ColumnToArray("fldPrintCode")

            'Create new error
            Dim newError As New ValidationError() With {
                .Error_Code = "UI_NOT_VALID",
                .Error_Descr = $"Some of the UIs have been a part of deactivation(IDA) message.",
                .Error_Data = String.Join("#", errCodes)}
            Errors.Add(newError)
            Return ValidationResult.Invalid
        Else
            Return ValidationResult.Valid
        End If
    End Function

    Public Function VAL_UI_EXIST_AUI_SEQ() As ValidationResult
        Dim db As New DBManager()
        Dim jAggColumn As String
        Dim jCodesColumn As String
        Dim codeList As String()

        'Make adjusments based on the message type
        Select Case msgType
            Case "IDA"
                jAggColumn = "Deact_Type"
                jCodesColumn = "Deact_aUI"
            Case "EPA"
                jAggColumn = "Aggregation_Type"
                jCodesColumn = "Aggregated_UIs2"
            Case "EDP", "ERP", "ETL", "EVR"
                jAggColumn = "UI_Type"
                jCodesColumn = "aUIs"
            Case Else
                Return ValidationResult.Valid
        End Select

        'Get the codes from the right column
        Dim aggType As Integer = JSON(jAggColumn).ToObject(Of Integer)
        Select Case aggType
            Case 2, 3 'aggregated level or Both
                codeList = JSON.Item(jCodesColumn).ToObject(Of String())
            Case Else 'This validation only checks aggregated level uis
                Return ValidationResult.Valid
        End Select

        'Check db for deactivated/disagreggated
        Dim resultDeact As DataTable = db.CheckForDeactivated(Tables.tblaggregatedcodes.ToString(), codeList, "fldPrintCode")
        Dim resultDeagg As DataTable = db.CheckForEUD(codeList)

        'If any of the codes are deactivated/deaggregated
        If resultDeact.Rows.Count > 0 OrElse resultDeagg.Rows.Count > 0 Then
            'Generate error
            Dim errCodes As List(Of String) = New List(Of String)
            'Add both code lists to the error list
            errCodes.TryAddRange(resultDeact.ColumnToArray("fldPrintCode"))
            errCodes.TryAddRange(resultDeagg.ColumnToArray("fldPrintCode"))

            'Create new error
            Dim newError As New ValidationError() With {
                .Error_Code = "UI_NOT_EXIST",
                .Error_Descr = $"Some of the UIs were not found in the primary repository.",
                .Error_Data = String.Join("#", errCodes)}
            Errors.Add(newError)
            Return ValidationResult.Invalid
        Else
            Return ValidationResult.Valid
        End If
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

    Public Function VAL_UI_ORD_REACTIVATION() As ValidationResult
        Dim db As New DBManager()
        Select Case msgType
            Case "EUA"
                'Check the db
                Dim codes As String() = JSON("upUI_2").ToObject(Of String())
                Dim result As DataTable = db.CheckForDeactivated("tblprimarycodes", codes, "fldCode")
                'If any deactivated codes are found
                If result.Rows.Count > 0 Then
                    'Create new error
                    Dim deactivatedUIs As String() = result.ColumnToArray("fldCode")
                    Dim newError As New ValidationError() With {
                      .Error_Code = "UI_DEACTIVATED",
                      .Error_Descr = $"upUI(s) that have been deactivated should not participate in any application event (EUA).",
                      .Error_Data = $"{String.Join("#", deactivatedUIs)}"}
                    Errors.Add(newError)
                    Return ValidationResult.Invalid
                End If
                Return ValidationResult.Valid
            Case Else
                Return ValidationResult.Valid
        End Select
    End Function

    Public Function VAL_UI_ORD_DEACTIVATED() As ValidationResult
        Dim db As New DBManager()
        Dim output As Boolean
        Select Case msgType
            Case "IDA"
                'Deact_upUI
                'Deact_aUI
                Dim aUIs As String() = JSON.Item("Deact_aUI").ToObject(Of String())
                Dim upUIs As String() = JSON.Item("Deact_upUI").ToObject(Of String())
                Dim aggType As Integer = JSON("Deact_Type").ToObject(Of Integer)
                Select Case aggType
                    Case 1 'ui only
                        output = Not CheckForDeactivated("tblprimarycodes", upUIs, "fldCode")
                    Case 2 'Aggregated only
                        output = Not CheckForDeactivated("tblaggregatedcodes", aUIs, "fldPrintCode")
                    Case Else
                        Throw New Exception($"Unexpected value for Deact_Type: {aggType}")
                End Select
            Case "EPA"
                'Aggregated_UIs1 = upUI(L)
                'Aggregated_UIs2
                Dim upUIs As String() = JSON.Item("Aggregated_UIs1").ToObject(Of String())
                Dim aUIs As String() = JSON.Item("Aggregated_UIs2").ToObject(Of String())
                Dim aggType As Integer = JSON("Aggregation_Type").ToObject(Of Integer)
                Select Case aggType
                    Case 1 'unit level ui
                        output = Not CheckForDeactivated("tblprimarycodes", upUIs, "fldPrintCode")
                    Case 2 'Aggregated ui
                        output = Not CheckForDeactivated("tblaggregatedcodes", aUIs, "fldPrintCode")
                    Case 3 'Both
                        output = Not (CheckForDeactivated("tblprimarycodes", upUIs, "fldPrintCode") OrElse CheckForDeactivated("tblaggregatedcodes", aUIs, "fldPrintCode"))
                    Case Else
                        Throw New Exception($"Unexpected value for Aggregation_Type: {aggType}")
                End Select
            Case "EDP", "ERP", "ETL", "EVR"
                'upUI(L)
                'aUIs
                Dim upUIs As String() = JSON.Item("upUIs").ToObject(Of String())
                Dim aUIs As String() = JSON.Item("aUIs").ToObject(Of String())
                Dim aggType As Integer = JSON("UI_Type").ToObject(Of Integer)
                Select Case aggType
                    Case 1 'unit level ui
                        output = Not CheckForDeactivated("tblprimarycodes", upUIs, "fldPrintCode")
                    Case 2 'Aggregated ui
                        output = Not CheckForDeactivated("tblaggregatedcodes", aUIs, "fldPrintCode")
                    Case 3 'Both
                        output = Not (CheckForDeactivated("tblprimarycodes", upUIs, "fldPrintCode") OrElse CheckForDeactivated("tblaggregatedcodes", aUIs, "fldPrintCode"))
                    Case Else
                        Throw New Exception($"Unexpected value for UI_Type: {aggType}")
                End Select
            Case "EUD"
                'aUI
                Return ValidationResult.Valid
            Case Else
                Return ValidationResult.Valid
        End Select
        Return If(output, ValidationResult.Valid, ValidationResult.Invalid)
    End Function

    Public Function VAL_UI_ORD_AGG_MULT() As ValidationResult
        Dim db As New DBManager()
        Select Case msgType
            Case "EPA"
                Dim aUI As String = JSON("aUI")
                Dim result As DataTable = db.Check_aUI(aUI)
                If result.Rows.Count > 0 Then
                    'Create new error
                    Dim newError As New ValidationError() With {
                        .Error_Code = "MULTIPLE_AGGREGATION",
                        .Error_Descr = $"ERROR: VAL_UI_ORD_AGG_MULT, Packaging out of sequence for components aggregation by manufacturers/importers is not expected nor allowed when aggregation is in progress with components produced in EU (AWAITING_IN_STOCK/EPA_EU_NOT_IN_STOCK_PRIMARY) -",
                        .Error_Data = aUI}
                    Errors.Add(newError)
                    Return ValidationResult.Invalid
                End If
            Case Else
                'If the UI exists and is not deaggregated, validation fails
                Return ValidationResult.Valid
        End Select
        Return ValidationResult.Valid
    End Function

    Public Function VAL_EVT_24H() As ValidationResult
        Select Case msgType
            Case "EUA", "EPA", "EVR", "EIV", "EPO", "EPR"
                Dim eventTime As Date = ParseTime(JSON("Event_Time"))
                If Date.UtcNow - eventTime > TimeSpan.FromHours(24) Then
                    'It's an older code, sir, but it checks out. 
                    Dim newError As New ValidationError() With {
                       .Error_Code = "OPERATION_WITHIN_24_HOURS",
                       .Error_Descr = $"Events should be reported within 24 hours from the occurrence of the event"}
                    Errors.Add(newError)
                    Return ValidationResult.PassWithWarning
                Else Return ValidationResult.Valid
                End If
            Case Else
                Return ValidationResult.Valid
        End Select
    End Function

    Public Function VAL_EVT_TIME() As ValidationResult
        Select Case msgType
            Case "EDP", "ETL"
                Dim eventTime As Date = ParseTime(JSON("Event_Time"))
                If Date.UtcNow - eventTime > TimeSpan.FromHours(24) Then
                    'It's an older code, sir, but it checks out. 
                    Dim newError As New ValidationError() With {
                       .Error_Code = "SHIPMENT_WITHIN_24_HOURS",
                       .Error_Descr = $"Events should be reported within 24 hours from the occurrence of the event"}
                    Errors.Add(newError)
                    Return ValidationResult.PassWithWarning
                Else Return ValidationResult.Valid
                End If
            Case Else
                Return ValidationResult.Valid
        End Select
    End Function

    Public Function VAL_UI_ORD_DISAGG() As ValidationResult
        Dim db As New DBManager()
        Select Case msgType
            Case "EDP", "ERP", "ETL", "EVR" 'aUIs
                'Check the codes for deaggregated
                Dim codes As String() = JSON("aUIs").ToObject(Of String())
                Dim result As DataTable = db.CheckForDeaggregated(codes)

                'If any deaggregated codes are found
                If result.Rows.Count > 0 Then
                    'Create new error
                    Dim errorUIs As String() = result.ColumnToArray("fldPrintCode")
                    Dim newError As New ValidationError() With {
                                      .Error_Code = "UI_ALREADY_DISAGGREGATED",
                                      .Error_Descr = $"An aUI that has been disaggregated cannot be part on any product movement prior of being aggregated.",
                                      .Error_Data = $"{String.Join("#", errorUIs)}"}
                    Errors.Add(newError)
                    Return ValidationResult.Invalid
                End If
            Case Else 'Skip                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     ip
                Return ValidationResult.Valid
        End Select

        Return ValidationResult.Valid
    End Function

    Public Function VAL_UI_ORD_IMPLDISAGG() As ValidationResult
        Return ValidationResult.Valid
    End Function

    Public Function VAL_UI_ORD_AGG_FID() As ValidationResult
        Dim db As New DBManager()

        Select Case msgType
            Case "EPA"
                'Get necessary vars
                Dim F_ID As String = JSON("F_ID")
                Dim aggType As Integer = JSON("Aggregation_Type").ToObject(Of Integer)
                Dim uis As String() = JSON.Item("Aggregated_UIs1").ToObject(Of String())
                Dim aUIs As String() = JSON.Item("Aggregated_UIs2").ToObject(Of String())
                Dim err As Boolean = False
                Dim errCodes As List(Of String) = New List(Of String)

                Select Case aggType
                    Case 1
                        'Check for codes with non matching location
                        Dim result As DataTable = db.CheckCodeLocation(uis, F_ID, Tables.tblprimarycodes.ToString(), "fldPrintCode")
                        'If there are any
                        If result.Rows.Count > 0 Then
                            'Flag error
                            err = True
                            errCodes.AddRange(result.ColumnToArray("fldPrintCode"))
                        End If
                    Case 2
                        'Check for codes with non matching location
                        Dim result As DataTable = db.CheckCodeLocation(aUIs, F_ID, Tables.tblaggregatedcodes.ToString(), "fldPrintCode")
                        'If there are any
                        If result.Rows.Count > 0 Then
                            'Flag error
                            err = True
                            errCodes.AddRange(result.ColumnToArray("fldPrintCode"))
                        End If
                    Case 3
                        'Check both arrays
                        Dim result As DataTable = db.CheckCodeLocation(uis, F_ID, Tables.tblprimarycodes.ToString(), "fldPrintCode")
                        Dim result2 As DataTable = db.CheckCodeLocation(aUIs, F_ID, Tables.tblaggregatedcodes.ToString(), "fldPrintCode")
                        'If any one the arrays have non matching locations
                        If result.Rows.Count > 0 OrElse result2.Rows.Count > 0 Then
                            err = True
                            errCodes.TryAddRange(result.ColumnToArray("fldPrintCode"))
                            errCodes.TryAddRange(result2.ColumnToArray("fldPrintCode"))
                        End If
                End Select
                'If erronous data was found
                If err Then
                    'Generate error
                    Dim newError As New ValidationError() With {
                        .Error_Code = "LOCATION_MISMATCH",
                        .Error_Descr = "Aggregation and the disaggregation events must happen at the same facility (FID) where the products have been either created or arrived",
                        .Error_Data = String.Join("#", errCodes)}
                    Errors.Add(newError)
                    Return ValidationResult.Invalid
                Else Return ValidationResult.Valid
                End If

            Case "EUD"
                Dim F_ID As String = JSON("F_ID")
                Dim aUI As String = JSON("aUI")
                'Check in db
                Dim result As DataTable = db.CheckCodeLocation(aUI, F_ID)

                'If the code is found AND the location is different from the sent F_ID
                If result.Rows.Count > 0 Then
                    'Generate error
                    Dim newError As New ValidationError() With {.Error_Code = "LOCATION_MISMATCH", .Error_Descr = "Aggregation and the disaggregation events must happen at the same facility (FID) where the products have been either created or arrived"}
                    Errors.Add(newError)
                    Return ValidationResult.Invalid
                Else
                    Return ValidationResult.Valid
                End If

            Case Else 'Do not check message
                Return ValidationResult.Valid
        End Select
    End Function

    Public Function VAL_UI_ORD_DISPATCH() As ValidationResult
        Dim db As New DBManager()
        Select Case msgType
            Case "EDP"
                Dim F_ID As String = JSON("F_ID")
                Dim uiType As Integer = JSON("UI_Type").ToObject(Of Integer)
                Dim uis As String() = JSON.Item("upUIs").ToObject(Of String())
                Dim aUIs As String() = JSON.Item("aUIs").ToObject(Of String())
                Dim err As Boolean = False
                Dim errCodes As List(Of String) = New List(Of String)

                Select Case uiType
                    Case 1
                        'Check for codes with non matching location
                        Dim result As DataTable = db.CheckCodeLocation(uis, F_ID, Tables.tblprimarycodes.ToString(), "fldPrintCode")
                        'If there are any
                        If result.Rows.Count > 0 Then
                            'Flag error
                            err = True
                            errCodes.AddRange(result.ColumnToArray("fldPrintCode"))
                        End If
                    Case 2
                        'Check for codes with non matching location
                        Dim result As DataTable = db.CheckCodeLocation(aUIs, F_ID, Tables.tblaggregatedcodes.ToString(), "fldPrintCode")
                        'If there are any
                        If result.Rows.Count > 0 Then
                            'Flag error
                            err = True
                            errCodes.AddRange(result.ColumnToArray("fldPrintCode"))
                        End If
                    Case 3
                        'Check both arrays
                        Dim result As DataTable = db.CheckCodeLocation(uis, F_ID, Tables.tblprimarycodes.ToString(), "fldPrintCode")
                        Dim result2 As DataTable = db.CheckCodeLocation(aUIs, F_ID, Tables.tblaggregatedcodes.ToString(), "fldPrintCode")
                        'If any one the arrays have non matching locations
                        If result.Rows.Count > 0 OrElse result2.Rows.Count > 0 Then
                            err = True
                            errCodes.TryAddRange(result.ColumnToArray("fldPrintCode"))
                            errCodes.TryAddRange(result2.ColumnToArray("fldPrintCode"))
                        End If
                End Select

                'If erronous data was found
                If err Then
                    'Generate error
                    Dim newError As New ValidationError() With {
                        .Error_Code = "LOCATION_MISMATCH",
                        .Error_Descr = "UI last location (FID) must matche the source location (FID) of the dispatch event",
                        .Error_Data = String.Join("#", errCodes)}
                    Errors.Add(newError)
                    Return ValidationResult.Invalid
                Else Return ValidationResult.Valid
                End If
            Case Else
                Return ValidationResult.Valid
        End Select
    End Function

    Public Function VAL_UI_ORD_ARRIVAL() As ValidationResult
        Dim db As New DBManager()
        Select Case msgType
            Case "ERP"
                'Get variables
                Dim product_return As Integer = JSON("Product_Return").ToObject(Of Integer)
                If product_return = 0 Then
                    'Get the arriving codes
                    Dim ui_type = JSON("UI_Type").ToObject(Of Integer)
                    Dim upUIs As String() = JSON("upUIs").ToObject(Of String())
                    Dim aUIs As String() = JSON("aUIs").ToObject(Of String())

                    Dim err As Boolean = False
                    Dim errCodes As List(Of String) = New List(Of String)

                    'Search the database
                    Select Case ui_type
                        Case 1
                            Dim result As DataTable = db.SearchUisInJSON(upUIs, Tables.tbljsonsecondary.ToString(), msgType, ".upUIs")

                            'If some of the codes are missing
                            If result.Rows.Count <> upUIs.Length Then
                                'Get the missing codes and return error
                                Dim missingCodes = upUIs.Except(result.ColumnToArray("Codes"))
                                err = True
                                errCodes.AddRange(missingCodes)
                            End If
                        Case 2
                            Dim result As DataTable = db.SearchUisInJSON(aUIs, Tables.tbljsonsecondary.ToString(), msgType, ".aUIs")

                            'If some of the codes are missing
                            If result.Rows.Count <> aUIs.Length Then
                                'Get the missing codes and return error
                                Dim missingCodes = aUIs.Except(result.ColumnToArray("Codes"))
                                err = True
                                errCodes.AddRange(missingCodes)
                            End If
                        Case 3
                            Dim result1 As DataTable = db.SearchUisInJSON(upUIs, Tables.tbljsonsecondary.ToString(), msgType, ".upUIs")
                            Dim result2 As DataTable = db.SearchUisInJSON(aUIs, Tables.tbljsonsecondary.ToString(), msgType, ".aUIs")

                            'If some of the codes are missing
                            If result1.Rows.Count <> upUIs.Length OrElse result2.Rows.Count <> aUIs.Length Then
                                'Get the missing codes and return error
                                Dim missingUIs = upUIs.Except(result1.ColumnToArray("Codes"))
                                Dim missingAUIs = aUIs.Except(result2.ColumnToArray("Codes"))

                                err = True
                                errCodes.AddRange(missingUIs)
                                errCodes.AddRange(missingAUIs)
                            End If
                        Case Else
                            Throw New Exception($"Exception in VAL_UI_ORD_ARRIVAL(). Bad value for ui_type: '{ui_type}'")
                    End Select

                    'If search resulted in error
                    If err Then
                        'Generate error
                        Dim newError As New ValidationError() With {
                                  .Error_Code = "ARRIVAL_NOTALLOWED",
                                  .Error_Descr = "Some or all of the UIs have not been part of a prior reported dispatch or transloading event (EDP, ETL).",
                                  .Error_Data = String.Join("#", errCodes)}
                        Errors.Add(newError)
                        Return ValidationResult.Invalid
                    Else
                        Return ValidationResult.Valid
                    End If
                Else 'Skip
                    Return ValidationResult.Valid
                End If
            Case Else 'Skip
                Return ValidationResult.Valid
        End Select
    End Function

    Public Function VAL_UI_ORD_ARRIVAL_RETURN() As ValidationResult
        Dim db As New DBManager()
        Select Case msgType
            Case "ERP"
                'Get variables
                Dim product_return As Integer = JSON("Product_Return").ToObject(Of Integer)

                'Is of type RETURN
                If product_return = 1 Then
                    'Get the arriving codes
                    Dim ui_type = JSON("UI_Type").ToObject(Of Integer)
                    Dim upUIs As String() = JSON("upUIs").ToObject(Of String())
                    Dim aUIs As String() = JSON("aUIs").ToObject(Of String())

                    Dim err As Boolean = False
                    Dim errCodes As List(Of String) = New List(Of String)

                    'Search the database
                    Select Case ui_type
                        Case 1
                            Dim result As DataTable = db.SearchUisInJSON(upUIs, Tables.tbljsonsecondary.ToString(), msgType, ".upUIs")

                            'If some of the codes are missing
                            If result.Rows.Count <> upUIs.Length Then
                                'Get the missing codes and return error
                                Dim missingCodes = upUIs.Except(result.ColumnToArray("Codes"))
                                err = True
                                errCodes.AddRange(missingCodes)
                            End If
                        Case 2
                            Dim result As DataTable = db.SearchUisInJSON(aUIs, Tables.tbljsonsecondary.ToString(), msgType, ".aUIs")

                            'If some of the codes are missing
                            If result.Rows.Count <> aUIs.Length Then
                                'Get the missing codes and return error
                                Dim missingCodes = aUIs.Except(result.ColumnToArray("Codes"))
                                err = True
                                errCodes.AddRange(missingCodes)
                            End If
                        Case 3
                            Dim result1 As DataTable = db.SearchUisInJSON(upUIs, Tables.tbljsonsecondary.ToString(), msgType, ".upUIs")
                            Dim result2 As DataTable = db.SearchUisInJSON(aUIs, Tables.tbljsonsecondary.ToString(), msgType, ".aUIs")

                            'If some of the codes are missing
                            If result1.Rows.Count <> upUIs.Length OrElse result2.Rows.Count <> aUIs.Length Then
                                'Get the missing codes and return error
                                Dim missingUIs = upUIs.Except(result1.ColumnToArray("Codes"))
                                Dim missingAUIs = aUIs.Except(result2.ColumnToArray("Codes"))

                                err = True
                                errCodes.TryAddRange(missingUIs)
                                errCodes.TryAddRange(missingAUIs)
                            End If
                        Case Else
                            Throw New Exception($"Exception in VAL_UI_ORD_ARRIVAL(). Bad value for ui_type: '{ui_type}'")
                    End Select

                    'If search resulted in error
                    If err Then
                        'Generate error
                        Dim newError As New ValidationError() With {
                                  .Error_Code = "ARRIVAL_NOTALLOWED",
                                  .Error_Descr = "Some or all of the UIs have not been part of a prior reported dispatch or transloading event (EDP, ETL).",
                                  .Error_Data = String.Join("#", errCodes)}
                        Errors.Add(newError)
                        Return ValidationResult.Invalid
                    Else
                        Return ValidationResult.Valid
                    End If
                Else 'Skip
                    Return ValidationResult.Valid
                End If
            Case Else 'Skip
                Return ValidationResult.Valid
        End Select
    End Function

    Public Function VAL_RECALL_EXIST() As ValidationResult
        Dim db As New DBManager()
        Select Case msgType
            Case "RCL"
                Dim targetCode As String = JSON("Recall_CODE")
                'Check if it exists
                Dim result As DataTable = db.CheckForCode(targetCode)
                'If it doesn't exist
                If result.Rows.Count < 1 Then
                    'Generate error
                    Dim newError As New ValidationError() With {
                              .Error_Code = "CODE_NOT_EXIST",
                              .Error_Descr = "Recall code was not found in the repository",
                              .Error_Data = targetCode}
                    Errors.Add(newError)
                    Return ValidationResult.Invalid
                Else
                    Return ValidationResult.Valid
                End If
            Case Else
                Return ValidationResult.Valid
        End Select
    End Function

    Public Function VAL_RECALL_LAST() As ValidationResult
        Dim db As New DBManager()
        Select Case msgType
            Case "RCL"
                'Get the index of the recalled JSON
                Dim targetID As String = JSON("Recall_CODE")
                'Get it from the db
                Dim result As DataTable = db.CheckForCode(targetID)
                'If its there

                'Convert to JObject
                Dim resultJSON As JObject = JObject.Parse(result.Rows(0)("fldJson"))
                Dim type As String = resultJSON("Message_Type")
                Dim jsonDate As Date = CDate(result.Rows(0)("fldDate"))
                Dim err As Boolean = False

                'Get the UIs from the message
                Select Case type
                    Case "IDA"
                        'Deact_upUI
                        'Deact_aUI
                        Dim aggType As Integer = resultJSON("Deact_Type").ToObject(Of Integer)
                        Select Case aggType
                            Case 1 'ui only
                                'Get the codes
                                Dim uis As String() = resultJSON("Deact_upUI").ToObject(Of String())
                                Dim afterEvents As DataTable = CheckForAfterEvents(db, jsonDate, uis, "fldPrintCode", Tables.tblprimarycodes)

                                'If any events are found for these codes, fail validation
                                If afterEvents.Rows.Count > 0 Then err = True
                            Case 2 'Aggregated only
                                'Get the codes
                                Dim uis As String() = resultJSON("Deact_aUI").ToObject(Of String())
                                Dim afterEvents As DataTable = CheckForAfterEvents(db, jsonDate, uis, "fldPrintCode", Tables.tblaggregatedcodes)

                                'If any events are found for these codes, fail validation
                                If afterEvents.Rows.Count > 0 Then err = True
                            Case Else
                                Throw New Exception($"Unexpected value for Deact_Type: {aggType}")
                        End Select
                    Case "EUA"
                        'upUI_1 = upUI(L)
                        Dim uis As String() = resultJSON("upUI_1").ToObject(Of String())
                        Dim afterEvents As DataTable = CheckForAfterEvents(db, jsonDate, uis, "fldPrintCode", Tables.tblprimarycodes)

                        'If any events are found for these codes, fail validation
                        If afterEvents.Rows.Count > 0 Then err = True
                    Case "EPA"
                        'parent code has to be checked too
                        'Aggregated_UIs1
                        'Aggregated_UIs2
                        Dim aggType As Integer = resultJSON("Aggregation_Type").ToObject(Of Integer)
                        Select Case aggType
                            Case 1 'ui only
                                'Get the codes
                                Dim uis As String() = resultJSON("Aggregated_UIs1").ToObject(Of String())
                                Dim afterEvents As DataTable = CheckForAfterEvents(db, jsonDate, uis, "fldPrintCode", Tables.tblprimarycodes)

                                'If any events are found for these codes, fail validation
                                If afterEvents.Rows.Count > 0 Then err = True
                            Case 2 'Aggregated only
                                'Get the codes
                                Dim uis As String() = resultJSON("Aggregated_UIs2").ToObject(Of String())
                                Dim afterEvents As DataTable = CheckForAfterEvents(db, jsonDate, uis, "fldPrintCode", Tables.tblaggregatedcodes)

                                'If any events are found for these codes, fail validation
                                If afterEvents.Rows.Count > 0 Then err = True
                            Case 3 'Both
                                'Get the uis
                                Dim uis As String() = resultJSON("Aggregated_UIs1").ToObject(Of String())
                                Dim afterEvents As DataTable = CheckForAfterEvents(db, jsonDate, uis, "fldPrintCode", Tables.tblprimarycodes)
                                'Get the aUIs
                                Dim ais As String() = resultJSON("Aggregated_UIs2").ToObject(Of String())
                                Dim afterEvents2 As DataTable = CheckForAfterEvents(db, jsonDate, uis, "fldPrintCode", Tables.tblaggregatedcodes)
                                'If any events are found for these codes, fail validation
                                If afterEvents.Rows.Count > 0 OrElse afterEvents2.Rows.Count > 0 Then err = True
                            Case Else
                                Throw New Exception($"Unexpected value for Aggregation_Type: {aggType}")
                        End Select

                    Case "EDP", "ERP", "ETL", "EVR", "EIV", "EPO", "EPR"
                        'upUIs
                        'aUIs
                        Dim aggType As Integer = resultJSON("UI_Type").ToObject(Of Integer)
                        Select Case aggType
                            Case 1 'ui only
                                'Get the codes
                                Dim uis As String() = resultJSON("upUIs").ToObject(Of String())
                                Dim afterEvents As DataTable = CheckForAfterEvents(db, jsonDate, uis, "fldPrintCode", Tables.tblprimarycodes)

                                'If any events are found for these codes, fail validation
                                If afterEvents.Rows.Count > 0 Then err = True
                            Case 2 'Aggregated only
                                'Get the codes
                                Dim uis As String() = resultJSON("aUIs").ToObject(Of String())
                                Dim afterEvents As DataTable = CheckForAfterEvents(db, jsonDate, uis, "fldPrintCode", Tables.tblaggregatedcodes)

                                'If any events are found for these codes, fail validation
                                If afterEvents.Rows.Count > 0 Then err = True
                            Case 3 'Both
                                'Get the uis
                                Dim uis As String() = resultJSON("upUIs").ToObject(Of String())
                                Dim afterEvents As DataTable = CheckForAfterEvents(db, jsonDate, uis, "fldPrintCode", Tables.tblprimarycodes)
                                'Get the aUIs
                                Dim ais As String() = resultJSON("aUIs").ToObject(Of String())
                                Dim afterEvents2 As DataTable = CheckForAfterEvents(db, jsonDate, uis, "fldPrintCode", Tables.tblaggregatedcodes)
                                'If any events are found for these codes, fail validation
                                If afterEvents.Rows.Count > 0 OrElse afterEvents2.Rows.Count > 0 Then err = True
                            Case Else
                                Throw New Exception($"Unexpected value for UI_Type: {aggType}")
                        End Select
                    Case Else
                End Select

                'If any events after the eventDate were found
                If err Then
                    'Generate error
                    Dim newError As New ValidationError() With {
                          .Error_Code = "RECALL_NOT_LAST_EVENT",
                          .Error_Descr = "RecallCode must the very last unrecalled event occurred on all UIs"}
                    Errors.Add(newError)
                    Return ValidationResult.Invalid
                Else
                    Return ValidationResult.Valid
                End If

            Case Else
                Return ValidationResult.Valid
        End Select
    End Function

#End Region
    Private Function CheckForAfterEvents(db As DBManager, eventDate As Date, eventUIs As String(), sColumn As String, sTable As Tables) As DataTable
        'Get the events for those codes
        Dim involvedEvents As DataTable = db.SelectInvolvedEvents(eventUIs, sTable, sColumn)
        'Convert to one dimentional array
        Dim strEvents As String = String.Join(",", involvedEvents.ColumnToArray("EventList"))
        'Remove null values
        Dim arrEvents = strEvents.Split(",").Where(Function(x) Not String.IsNullOrEmpty(x)).ToArray()
        'Remove duplicates
        Dim distinctEvents As HashSet(Of String) = New HashSet(Of String)(arrEvents)
        'Check for events that happened after the eventDate
        Dim afterEvents As DataTable = db.SelectMessagesOlderThan(eventDate, distinctEvents)
        Return afterEvents
    End Function
    Private Function CheckForDeactivated(table As String, codes As String(), codesField As String) As Boolean
        Dim db As New DBManager()
        Dim result As DataTable = db.CheckForDeactivated(table, codes, codesField)
        If result.Rows.Count > 0 Then
            'Create new error
            Dim errorUIs As String() = result.ColumnToArray(codesField)
            Dim newError As New ValidationError() With {
                              .Error_Code = "UI_DEACTIVATED",
                              .Error_Descr = $"Presence of UIs in a message after being deactivated.",
                              .Error_Data = $"{String.Join("#", errorUIs)}"}
            Errors.Add(newError)
            Return True
        Else
            Return False
        End If
    End Function

    Private Function CheckForDuplicates(uis As String) As Boolean
        'upUI
        Dim codes As String() = JSON.Item(uis).ToObject(Of String())
        If codes.HasDuplicates() Then
            Dim duplicates As String() = codes.GroupBy(Function(s) s).SelectMany(Function(grp) grp.Skip(1)).Distinct().ToArray()
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
