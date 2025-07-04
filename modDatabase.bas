Attribute VB_Name = "modDatabase"
Public conn As ADODB.Connection

Public Sub OpenConnection()
    Set conn = New ADODB.Connection
    conn.Open "Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=ColombusDB;User ID=sa;Password=123;"
End Sub

Public Sub CloseConnection()
    If Not conn Is Nothing Then
        conn.Close
        Set conn = Nothing
    End If
End Sub