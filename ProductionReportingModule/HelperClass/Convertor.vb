Public Module Convertor
    Public Function GetAuthType(type As String) As AuthenticationType
        Select Case type
            Case "0"
                Return AuthenticationType.NoAuth
            Case "1"
                Return AuthenticationType.Basic
            Case Else
                Throw New NotImplementedException($"Authentication type not implemented: {type}")
        End Select
    End Function
End Module
