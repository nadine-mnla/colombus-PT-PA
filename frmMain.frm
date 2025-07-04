VERSION 5.00
Begin VB.Form frmMain
   Caption         =   "Colombus - Room Chart"
   ClientHeight    =   4000
   ClientLeft      =   100
   ClientTop       =   100
   ClientWidth     =   6000
   Begin VB.ListBox lstRoomChart
      Height          =   2000
      Left            =   200
      Top             =   300
      Width           =   5500
   End
   Begin VB.TextBox txtRoomID
      Height          =   300
      Left            =   200
      Top             =   2500
      Width           =   1000
   End
   Begin VB.ComboBox cmbStatus
      Height          =   300
      Left            =   1300
      Top             =   2500
      Width           =   1500
   End
   Begin VB.CommandButton cmdUpdate
      Caption         =   "Update Status"
      Height          =   350
      Left            =   200
      Top             =   2900
      Width           =   1500
   End
   Begin VB.CommandButton cmdRefresh
      Caption         =   "Refresh"
      Height          =   350
      Left            =   1800
      Top             =   2900
      Width           =   1500
   End
End
Attribute VB_Name = "frmMain"

Private Sub Form_Load()
    Call LoadRoomChart
End Sub

Private Sub LoadRoomChart()
    Dim rs As ADODB.Recordset
    Set rs = GetRoomChart

    lstRoomChart.Clear
    Do While Not rs.EOF
        lstRoomChart.AddItem rs!ID & " - " & rs!NoKamar & " - " & rs!Tipe & " - " & rs!Status
        rs.MoveNext
    Loop
    rs.Close
End Sub

Private Sub cmdUpdate_Click()
    Dim id As Integer
    id = Val(txtRoomID.Text)
    Call UpdateRoomStatus(id, cmbStatus.Text)
    MsgBox "Status kamar diperbarui."
    Call LoadRoomChart
End Sub

Private Sub cmdRefresh_Click()
    Call LoadRoomChart
End Sub