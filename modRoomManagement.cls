VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "modRoomManagement"
Public Sub UpdateRoomStatus(roomID As Integer, status As String)
    Call OpenConnection
    Dim sql As String
    sql = "UPDATE Kamar SET Status='" & status & "' WHERE ID=" & roomID
    conn.Execute sql
    Call CloseConnection
End Sub

Public Function GetRoomChart() As ADODB.Recordset
    Call OpenConnection
    Dim rs As New ADODB.Recordset
    rs.Open "SELECT * FROM Kamar ORDER BY NoKamar", conn, adOpenStatic
    Set GetRoomChart = rs
End Function