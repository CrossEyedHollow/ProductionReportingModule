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
        Dim connString As String = ConnectionTools.DataBaseTools.GenerateConnectionString(DBIP, DBUser, DBPass, DBPort)
        conn = New MySqlConnection(connString)
        cmd = New MySqlCommand() With {.Connection = conn}
        adapter = New MySqlDataAdapter()
    End Sub
End Class
